/**
 * PSD Builder — uses ag-psd to construct PSD files from Figma data
 * PSD Reader — uses ag-psd to extract layers from PSD files
 * 
 * Enhanced: supports editable text layers, real image layers, 
 * and groups with clipping masks.
 */
import { writePsdUint8Array, readPsd, initializeCanvas } from 'ag-psd';

// Initialize ag-psd with canvas support for the browser
initializeCanvas(
    (width, height) => {
        const canvas = document.createElement('canvas');
        canvas.width = width;
        canvas.height = height;
        return canvas;
    }
);

/**
 * Build a PSD from Figma nodes + rendered images + text data + image fills.
 * 
 * Node types handled:
 * - TEXT → editable text layer (font, size, color preserved)
 * - IMAGE fill → real image layer with original image data
 * - FRAME with clipsContent → group with clipping mask
 * - Other (vectors, shapes) → rasterized PNG layer
 * 
 * @param {Array} nodes - Flat list of node info from collectNodes
 * @param {Object} rasterImages - Map of nodeId → { imageData, width, height, canvas }
 * @param {Object} imageFillCanvases - Map of imageRef → canvas (original fill images)
 * @param {Object} fileInfo - { name, width, height }
 * @returns {Uint8Array} PSD file bytes
 */
export function buildPsdFromFigma(nodes, rasterImages, imageFillCanvases, fileInfo) {
    let canvasWidth = fileInfo.width || 1920;
    let canvasHeight = fileInfo.height || 1080;

    // Find origin offset (Figma uses absolute coordinates)
    let originX = Infinity, originY = Infinity;
    for (const node of nodes) {
        if (node.bounds && !node.isGroup && !node.isGroupEnd) {
            originX = Math.min(originX, node.bounds.x);
            originY = Math.min(originY, node.bounds.y);
        }
    }
    if (originX === Infinity) originX = 0;
    if (originY === Infinity) originY = 0;

    // Build layer tree
    const rootChildren = [];
    const groupStack = [{ children: rootChildren }];

    for (const node of nodes) {
        if (node.isGroupEnd) {
            groupStack.pop();
            continue;
        }

        const currentGroup = groupStack[groupStack.length - 1];

        // ── GROUP / FRAME ──
        if (node.isGroup) {
            const group = {
                name: node.name,
                children: [],
                opened: true,
                blendMode: mapBlendMode(node.blendMode),
                opacity: node.opacity,
            };

            // If clipsContent is true, create a clipping mask group
            // In PSD, this means the group has a mask derived from its bounds
            if (node.clipsContent && node.bounds) {
                const maskLeft = Math.round(node.bounds.x - originX);
                const maskTop = Math.round(node.bounds.y - originY);
                const maskW = Math.round(node.bounds.width);
                const maskH = Math.round(node.bounds.height);

                // Create a mask canvas (white = visible, transparent = hidden)
                const maskCanvas = document.createElement('canvas');
                maskCanvas.width = canvasWidth;
                maskCanvas.height = canvasHeight;
                const mCtx = maskCanvas.getContext('2d');
                // Start fully transparent (hidden)
                mCtx.clearRect(0, 0, canvasWidth, canvasHeight);
                // White rectangle where content should show
                mCtx.fillStyle = '#FFFFFF';
                mCtx.fillRect(maskLeft, maskTop, maskW, maskH);

                group.mask = {
                    left: maskLeft,
                    top: maskTop,
                    right: maskLeft + maskW,
                    bottom: maskTop + maskH,
                    canvas: maskCanvas,
                    defaultColor: 0,    // 0 = hide by default
                    positionRelativeToLayer: false,
                };
            }

            currentGroup.children.push(group);
            groupStack.push(group);
            continue;
        }

        // ── TEXT LAYER (editable) ──
        if (node.isText && node.characters) {
            const layer = createTextLayer(node, originX, originY);
            currentGroup.children.push(layer);
            continue;
        }

        // ── IMAGE FILL LAYER (original image) ──
        if (node.isImageFill && node.imageRef && imageFillCanvases[node.imageRef]) {
            const layer = createImageFillLayer(node, imageFillCanvases[node.imageRef], originX, originY);
            currentGroup.children.push(layer);
            continue;
        }

        // ── RASTER LAYER (rendered PNG fallback) ──
        const imgData = rasterImages[node.id];
        if (!imgData) continue;

        const layer = createRasterLayer(node, imgData, originX, originY);
        currentGroup.children.push(layer);
    }

    // Build PSD object
    const psd = {
        width: canvasWidth,
        height: canvasHeight,
        children: rootChildren,
    };

    return writePsdUint8Array(psd);
}

/**
 * Create an editable text layer for PSD.
 * ag-psd's text layer format:
 *   - layer.text.text = the string content
 *   - layer.text.style = font/size/color info
 */
function createTextLayer(node, originX, originY) {
    const left = node.bounds ? Math.round(node.bounds.x - originX) : 0;
    const top = node.bounds ? Math.round(node.bounds.y - originY) : 0;
    const width = node.bounds ? Math.round(node.bounds.width) : 200;
    const height = node.bounds ? Math.round(node.bounds.height) : 40;

    const fontSize = node.textStyle?.fontSize || 16;
    const fontFamily = node.textStyle?.fontFamily || 'Inter';
    const color = node.textColor || { r: 0, g: 0, b: 0, a: 1 };

    // We need to render the text on a canvas for the composite image
    // (PSD needs both the text data AND a rendered preview)
    const canvas = document.createElement('canvas');
    canvas.width = Math.max(width, 1);
    canvas.height = Math.max(height, 1);
    const ctx = canvas.getContext('2d');

    ctx.fillStyle = `rgba(${color.r}, ${color.g}, ${color.b}, ${color.a})`;

    const fontWeight = node.textStyle?.fontWeight || 400;
    const fontStyle = fontWeight >= 700 ? 'bold' : 'normal';
    ctx.font = `${fontStyle} ${fontSize}px ${fontFamily}, Inter, sans-serif`;
    ctx.textBaseline = 'top';

    // Render text lines
    const lines = node.characters.split('\n');
    const lineHeight = node.textStyle?.lineHeightPx || fontSize * 1.4;
    for (let i = 0; i < lines.length; i++) {
        ctx.fillText(lines[i], 0, i * lineHeight);
    }

    // Font weight name mapping for PSD
    const weightNames = {
        100: 'Thin', 200: 'ExtraLight', 300: 'Light', 400: 'Regular',
        500: 'Medium', 600: 'SemiBold', 700: 'Bold', 800: 'ExtraBold', 900: 'Black',
    };
    const closestWeight = Object.keys(weightNames).reduce((prev, curr) =>
        Math.abs(curr - fontWeight) < Math.abs(prev - fontWeight) ? curr : prev
    );

    return {
        name: node.name,
        left,
        top,
        right: left + width,
        bottom: top + height,
        opacity: node.opacity,
        blendMode: mapBlendMode(node.blendMode),
        canvas,
        text: {
            text: node.characters,
            antiAlias: 'smooth',
            gridding: 'none',
            orientation: 'horizontal',
            warp: { style: 'none', value: 0, perspective: 0, perspectiveOther: 0, rotate: 'horizontal' },
            style: {
                font: { name: fontFamily },
                fontSize: fontSize,
                fillColor: { r: color.r, g: color.g, b: color.b },
                fauxBold: fontWeight >= 700,
                fauxItalic: false,
                tracking: node.textStyle?.letterSpacing ? Math.round(node.textStyle.letterSpacing * 1000 / fontSize) : 0,
                autoLeading: !node.textStyle?.lineHeightPx,
                leading: node.textStyle?.lineHeightPx || undefined,
            },
            paragraphStyle: {
                justification: mapTextAlign(node.textStyle?.textAlignHorizontal),
            },
        },
    };
}

/**
 * Create an image fill layer using the ORIGINAL uploaded image.
 * This preserves the actual image data (not rasterized with the frame).
 */
function createImageFillLayer(node, imageCanvas, originX, originY) {
    const left = node.bounds ? Math.round(node.bounds.x - originX) : 0;
    const top = node.bounds ? Math.round(node.bounds.y - originY) : 0;
    const width = node.bounds ? Math.round(node.bounds.width) : imageCanvas.width;
    const height = node.bounds ? Math.round(node.bounds.height) : imageCanvas.height;

    // Scale the original image to fit the node bounds
    const canvas = document.createElement('canvas');
    canvas.width = Math.max(width, 1);
    canvas.height = Math.max(height, 1);
    const ctx = canvas.getContext('2d');

    // Use FILL scaling (cover the area, crop excess)
    const scaleX = width / imageCanvas.width;
    const scaleY = height / imageCanvas.height;
    const scale = Math.max(scaleX, scaleY); // cover

    const srcW = width / scale;
    const srcH = height / scale;
    const srcX = (imageCanvas.width - srcW) / 2;
    const srcY = (imageCanvas.height - srcH) / 2;

    ctx.drawImage(imageCanvas, srcX, srcY, srcW, srcH, 0, 0, width, height);

    return {
        name: node.name,
        left,
        top,
        right: left + width,
        bottom: top + height,
        opacity: node.opacity,
        blendMode: mapBlendMode(node.blendMode),
        canvas,
    };
}

/**
 * Create a raster layer from a rendered PNG.
 */
function createRasterLayer(node, imgData, originX, originY) {
    const left = node.bounds ? Math.round(node.bounds.x - originX) : 0;
    const top = node.bounds ? Math.round(node.bounds.y - originY) : 0;

    const canvas = document.createElement('canvas');
    canvas.width = imgData.width;
    canvas.height = imgData.height;
    const ctx = canvas.getContext('2d');
    ctx.putImageData(imgData.imageData, 0, 0);

    return {
        name: node.name,
        left,
        top,
        right: left + imgData.width,
        bottom: top + imgData.height,
        opacity: node.opacity,
        blendMode: mapBlendMode(node.blendMode),
        canvas,
    };
}

/**
 * Read a PSD file and extract layer tree
 */
export function readPsdFile(buffer) {
    const psd = readPsd(new Uint8Array(buffer), {
        skipThumbnail: true,
    });
    return psd;
}

/**
 * Extract all layers from a PSD as flat list with PNG blobs
 */
export async function extractPsdLayers(psd, onProgress) {
    const layers = [];
    let processed = 0;

    function processLayer(layer, parentPath = '') {
        const path = parentPath ? `${parentPath}/${layer.name}` : (layer.name || 'Layer');

        if (layer.children && layer.children.length > 0) {
            // Check for mask (indicates clipsContent in original)
            const hasMask = !!layer.mask;

            layers.push({
                name: layer.name || 'Group',
                path,
                type: 'group',
                left: layer.left || 0,
                top: layer.top || 0,
                opacity: layer.opacity !== undefined ? layer.opacity : 1,
                blendMode: layer.blendMode || 'normal',
                hasMask,
                children: [],
            });

            for (const child of layer.children) {
                processLayer(child, path);
            }
        } else {
            let pngBlob = null;

            if (layer.canvas) {
                const canvas = layer.canvas;
                pngBlob = canvasToBlob(canvas);
            }

            // Check if it's a text layer
            const isTextLayer = !!layer.text;

            layers.push({
                name: layer.name || `Layer ${processed + 1}`,
                path,
                type: isTextLayer ? 'text' : 'layer',
                left: layer.left || 0,
                top: layer.top || 0,
                right: layer.right || 0,
                bottom: layer.bottom || 0,
                width: (layer.right || 0) - (layer.left || 0),
                height: (layer.bottom || 0) - (layer.top || 0),
                opacity: layer.opacity !== undefined ? layer.opacity : 1,
                blendMode: layer.blendMode || 'normal',
                hidden: layer.hidden || false,
                canvas: layer.canvas || null,
                pngBlob,
                // Preserve text data if available
                textData: layer.text ? {
                    text: layer.text.text || '',
                    fontSize: layer.text.style?.fontSize,
                    fontFamily: layer.text.style?.font?.name,
                } : null,
            });

            processed++;
            if (onProgress) onProgress(processed);
        }
    }

    if (psd.children) {
        for (const child of psd.children) {
            processLayer(child);
        }
    }

    return layers;
}

function canvasToBlob(canvas) {
    return new Promise((resolve) => {
        canvas.toBlob((blob) => resolve(blob), 'image/png');
    });
}

/**
 * Map Figma blend modes to PSD blend modes
 */
function mapBlendMode(figmaMode) {
    const map = {
        'NORMAL': 'normal',
        'MULTIPLY': 'multiply',
        'SCREEN': 'screen',
        'OVERLAY': 'overlay',
        'DARKEN': 'darken',
        'LIGHTEN': 'lighten',
        'COLOR_DODGE': 'color dodge',
        'COLOR_BURN': 'color burn',
        'HARD_LIGHT': 'hard light',
        'SOFT_LIGHT': 'soft light',
        'DIFFERENCE': 'difference',
        'EXCLUSION': 'exclusion',
        'HUE': 'hue',
        'SATURATION': 'saturation',
        'COLOR': 'color',
        'LUMINOSITY': 'luminosity',
    };
    return map[figmaMode] || 'normal';
}

/**
 * Map Figma text alignment to PSD justification
 */
function mapTextAlign(align) {
    const map = {
        'LEFT': 'left',
        'CENTER': 'center',
        'RIGHT': 'right',
        'JUSTIFIED': 'justifyAll',
    };
    return map[align] || 'left';
}
