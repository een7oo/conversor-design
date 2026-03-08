/**
 * Figma API Service
 * Fetches file structure and renders layers as PNGs
 * Enhanced: collects text content, image fills, and clip masks
 */

const FIGMA_API = 'https://api.figma.com/v1';

/**
 * Parse a Figma URL to extract file key and optional node ID
 */
export function parseFigmaUrl(url) {
    const match = url.match(/figma\.com\/(?:file|design)\/([a-zA-Z0-9]+)/);
    if (!match) throw new Error('URL do Figma inválida. Cole o link do arquivo.');

    const fileKey = match[1];
    const urlObj = new URL(url);
    const nodeId = urlObj.searchParams.get('node-id');

    return { fileKey, nodeId };
}

/**
 * Fetch the full file tree from Figma API
 */
export async function fetchFigmaFile(fileKey, token) {
    const res = await fetch(`${FIGMA_API}/files/${fileKey}?geometry=paths`, {
        headers: { 'X-Figma-Token': token },
    });

    if (!res.ok) {
        if (res.status === 403) throw new Error('Token de API inválido ou sem permissão para este arquivo.');
        if (res.status === 404) throw new Error('Arquivo não encontrado. Verifique o link.');
        throw new Error(`Erro da API Figma: ${res.status}`);
    }

    return res.json();
}

/**
 * Collect all renderable node IDs from the file tree.
 * Enhanced to preserve text data, image fills, and clipsContent.
 * 
 * Returns flat array of node info objects.
 */
export function collectNodes(fileData, opts = {}) {
    const { allPages = true, flattenGroups = false } = opts;
    const nodes = [];

    const pages = fileData.document.children || [];

    for (const page of pages) {
        if (!allPages && pages.indexOf(page) > 0) break;

        const frames = page.children || [];

        for (const frame of frames) {
            collectNodeRecursive(frame, page.name, nodes, flattenGroups, 0);
        }
    }

    return nodes;
}

function collectNodeRecursive(node, pageName, nodes, flattenGroups, depth) {
    if (node.visible === false) return;

    const nodeInfo = {
        id: node.id,
        name: node.name,
        type: node.type,
        pageName,
        depth,
        bounds: node.absoluteBoundingBox || null,
        opacity: node.opacity !== undefined ? node.opacity : 1,
        blendMode: node.blendMode || 'NORMAL',
    };

    // ── TEXT node: capture editable text data ──
    if (node.type === 'TEXT') {
        nodeInfo.isText = true;
        nodeInfo.characters = node.characters || '';
        nodeInfo.textStyle = {
            fontFamily: node.style?.fontFamily || 'Inter',
            fontWeight: node.style?.fontWeight || 400,
            fontSize: node.style?.fontSize || 16,
            letterSpacing: node.style?.letterSpacing || 0,
            lineHeightPx: node.style?.lineHeightPx || null,
            textAlignHorizontal: node.style?.textAlignHorizontal || 'LEFT',
        };
        // Text fill color
        if (node.fills && node.fills.length > 0 && node.fills[0].type === 'SOLID') {
            const c = node.fills[0].color;
            nodeInfo.textColor = {
                r: Math.round((c.r || 0) * 255),
                g: Math.round((c.g || 0) * 255),
                b: Math.round((c.b || 0) * 255),
                a: c.a !== undefined ? c.a : 1,
            };
        } else {
            nodeInfo.textColor = { r: 0, g: 0, b: 0, a: 1 };
        }
        nodes.push(nodeInfo);
        return;
    }

    // ── IMAGE fill detection ──
    const hasImageFill = node.fills?.some(f => f.type === 'IMAGE' && f.visible !== false);
    if (hasImageFill) {
        const imgFill = node.fills.find(f => f.type === 'IMAGE');
        nodeInfo.isImageFill = true;
        nodeInfo.imageRef = imgFill.imageRef; // hash for /v1/images endpoint
        nodeInfo.scaleMode = imgFill.scaleMode || 'FILL';
    }

    // ── Container nodes (FRAME, GROUP, COMPONENT, etc.) ──
    const isContainer = ['FRAME', 'GROUP', 'COMPONENT', 'COMPONENT_SET', 'INSTANCE', 'SECTION'].includes(node.type);

    if (isContainer && !flattenGroups && node.children?.length > 0) {
        // Check if this frame clips its content (becomes group+mask in PSD)
        const clipsContent = node.clipsContent === true;

        nodes.push({
            ...nodeInfo,
            isGroup: true,
            clipsContent,
            childCount: node.children.length,
        });

        for (const child of node.children) {
            collectNodeRecursive(child, pageName, nodes, flattenGroups, depth + 1);
        }

        nodes.push({ id: node.id + '_end', name: node.name, isGroupEnd: true, depth });
    } else {
        // Leaf node (vector, rectangle, ellipse, etc.) — will be rendered as PNG
        nodeInfo.isGroup = false;
        nodeInfo.isLeafRaster = !nodeInfo.isText; // everything non-text renders as raster
        nodes.push(nodeInfo);
    }
}

/**
 * Render nodes as PNG images via Figma Image API.
 * Only renders leaf raster nodes (skips text nodes which we handle natively).
 * Batches requests (max 50 IDs per request).
 */
export async function renderNodes(fileKey, token, nodeIds, onProgress, scale = 2) {
    const BATCH_SIZE = 30;
    const results = {};
    let totalDone = 0;

    for (let i = 0; i < nodeIds.length; i += BATCH_SIZE) {
        const batch = nodeIds.slice(i, i + BATCH_SIZE);
        const idsForApi = batch.join(',');

        const res = await fetch(
            `${FIGMA_API}/images/${fileKey}?ids=${idsForApi}&format=png&scale=${scale}`,
            { headers: { 'X-Figma-Token': token } }
        );

        if (!res.ok) {
            console.warn(`Batch ${i}-${i + batch.length} falhou (${res.status}), pulando...`);
            totalDone += batch.length;
            if (onProgress) onProgress(totalDone, nodeIds.length);
            continue;
        }

        const data = await res.json();
        if (data.err) {
            console.warn(`Erro do Figma no batch: ${data.err}`);
            totalDone += batch.length;
            if (onProgress) onProgress(totalDone, nodeIds.length);
            continue;
        }

        // Download all images from this batch IN PARALLEL
        const entries = Object.entries(data.images || {}).filter(([, url]) => url);
        const downloads = entries.map(async ([nodeId, imageUrl]) => {
            try {
                const imgRes = await fetch(imageUrl);
                const blob = await imgRes.blob();
                const arrayBuffer = await blob.arrayBuffer();
                results[nodeId] = new Uint8Array(arrayBuffer);
            } catch (err) {
                console.warn(`Falha ao baixar imagem do node ${nodeId}:`, err);
            }
        });

        await Promise.allSettled(downloads);
        totalDone += batch.length;
        if (onProgress) onProgress(totalDone, nodeIds.length);
    }

    return results;
}

/**
 * Fetch original image fills from Figma (the actual uploaded images, not renders).
 * Uses the /v1/files/:key/images endpoint to get image download URLs.
 */
export async function fetchImageFills(fileKey, token) {
    const res = await fetch(`${FIGMA_API}/files/${fileKey}/images`, {
        headers: { 'X-Figma-Token': token },
    });

    if (!res.ok) return {};

    const data = await res.json();
    return data.meta?.images || {};
}

/**
 * Download an image from URL and return as Uint8Array
 */
export async function downloadImage(url) {
    const res = await fetch(url);
    const blob = await res.blob();
    return new Uint8Array(await blob.arrayBuffer());
}

/**
 * Load image data into canvas pixel data
 */
export function loadImageData(uint8Array) {
    return new Promise((resolve, reject) => {
        const blob = new Blob([uint8Array], { type: 'image/png' });
        const url = URL.createObjectURL(blob);
        const img = new Image();

        img.onload = () => {
            const canvas = document.createElement('canvas');
            canvas.width = img.naturalWidth;
            canvas.height = img.naturalHeight;
            const ctx = canvas.getContext('2d');
            ctx.drawImage(img, 0, 0);
            const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
            URL.revokeObjectURL(url);
            resolve({ imageData, width: canvas.width, height: canvas.height, canvas });
        };

        img.onerror = () => {
            URL.revokeObjectURL(url);
            reject(new Error('Falha ao decodificar imagem'));
        };

        img.src = url;
    });
}

/**
 * Export a Figma frame/page as SVG via the Images API.
 * Returns the SVG string content.
 */
export async function exportFrameAsSvg(fileKey, token, nodeId) {
    const res = await fetch(
        `${FIGMA_API}/images/${fileKey}?ids=${nodeId}&format=svg&svg_include_id=true&svg_simplify_stroke=false`,
        { headers: { 'X-Figma-Token': token } }
    );

    if (!res.ok) throw new Error(`Erro ao exportar SVG: ${res.status}`);

    const data = await res.json();
    const svgUrl = data.images?.[nodeId];
    if (!svgUrl) throw new Error('Figma não retornou URL do SVG.');

    const svgRes = await fetch(svgUrl);
    const svgText = await svgRes.text();
    return svgText;
}

/**
 * Collect all top-level frames (artboards) from the file tree.
 * Returns array of { id, name, bounds }.
 */
export function collectFrames(fileData, allPages = true) {
    const frames = [];
    const pages = fileData.document.children || [];

    for (const page of pages) {
        if (!allPages && pages.indexOf(page) > 0) break;
        for (const frame of (page.children || [])) {
            frames.push({
                id: frame.id,
                name: frame.name,
                pageName: page.name,
                bounds: frame.absoluteBoundingBox,
            });
        }
    }
    return frames;
}

/**
 * Walk the file tree and collect all non-Normal blend modes.
 * Returns Map of layerName → PS BlendMode string.
 */
const FIGMA_TO_PS = {
    'PASS_THROUGH': 'BlendMode.PASSTHROUGH',
    'DARKEN': 'BlendMode.DARKEN',
    'MULTIPLY': 'BlendMode.MULTIPLY',
    'LINEAR_BURN': 'BlendMode.LINEARBURN',
    'COLOR_BURN': 'BlendMode.COLORBURN',
    'LIGHTEN': 'BlendMode.LIGHTEN',
    'SCREEN': 'BlendMode.SCREEN',
    'LINEAR_DODGE': 'BlendMode.LINEARDODGE',
    'COLOR_DODGE': 'BlendMode.COLORDODGE',
    'OVERLAY': 'BlendMode.OVERLAY',
    'SOFT_LIGHT': 'BlendMode.SOFTLIGHT',
    'HARD_LIGHT': 'BlendMode.HARDLIGHT',
    'DIFFERENCE': 'BlendMode.DIFFERENCE',
    'EXCLUSION': 'BlendMode.EXCLUSION',
    'HUE': 'BlendMode.HUE',
    'SATURATION': 'BlendMode.SATURATION',
    'COLOR': 'BlendMode.COLORBLEND',
    'LUMINOSITY': 'BlendMode.LUMINOSITY',
};

export function collectBlendModes(fileData) {
    const blendModes = new Map();

    function walk(node) {
        if (node.blendMode && node.blendMode !== 'NORMAL' && node.blendMode !== 'PASS_THROUGH') {
            const ps = FIGMA_TO_PS[node.blendMode];
            if (ps && node.name) {
                blendModes.set(node.name, ps);
            }
        }
        if (node.children) {
            for (const child of node.children) walk(child);
        }
    }

    for (const page of (fileData.document.children || [])) {
        for (const frame of (page.children || [])) {
            walk(frame);
        }
    }

    return blendModes;
}

/**
 * Collect all TEXT nodes with their styling from the file tree.
 * Used to recreate editable text in Photopea after SVG import.
 */
export function collectTextNodes(fileData) {
    const texts = [];

    function walk(node) {
        if (node.type === 'TEXT' && node.characters && node.absoluteBoundingBox) {
            const style = node.style || {};
            const fills = node.fills || [];
            const solidFill = fills.find(f => f.type === 'SOLID' && f.visible !== false);
            const color = solidFill?.color || { r: 0, g: 0, b: 0 };

            texts.push({
                id: node.id,
                name: node.name,
                text: node.characters,
                bounds: node.absoluteBoundingBox,
                fontSize: style.fontSize || 16,
                fontFamily: style.fontFamily || 'Arial',
                fontWeight: style.fontWeight || 400,
                textAlignHorizontal: node.style?.textAlignHorizontal || 'LEFT',
                color: {
                    r: Math.round(color.r * 255),
                    g: Math.round(color.g * 255),
                    b: Math.round(color.b * 255),
                },
                opacity: node.opacity !== undefined ? node.opacity : 1,
            });
        }
        if (node.children) {
            for (const child of node.children) walk(child);
        }
    }

    for (const page of (fileData.document.children || [])) {
        for (const frame of (page.children || [])) {
            walk(frame);
        }
    }

    return texts;
}

/**
 * Collect mask groups: only GROUP nodes containing an isMask child.
 * Figma creates GROUP type when user applies "Use as mask".
 * FRAME, INSTANCE, COMPONENT nodes are ignored to avoid false detections.
 */
export function collectMaskedGroups(fileData) {
    const masked = [];

    function walk(node) {
        // Never recurse into instances or components
        if (node.type === 'INSTANCE' || node.type === 'COMPONENT') {
            return;
        }

        // Only GROUP type can be actual user mask groups
        if (node.type === 'GROUP' && node.children && node.absoluteBoundingBox) {
            const maskChildIndex = node.children.findIndex(child => child.isMask === true);
            if (maskChildIndex !== -1) {
                const maskChild = node.children[maskChildIndex];
                const contentChildren = node.children.slice(maskChildIndex + 1);
                if (contentChildren.length > 0) {
                    masked.push({
                        id: node.id,
                        name: node.name,
                        bounds: node.absoluteBoundingBox,
                        maskChildId: maskChild.id,
                        maskChildName: maskChild.name,
                        contentChildIds: contentChildren.map(c => c.id),
                    });
                }
            }
        }

        if (node.children) {
            for (const child of node.children) walk(child);
        }
    }

    for (const page of (fileData.document.children || [])) {
        for (const frame of (page.children || [])) {
            walk(frame);
        }
    }

    return masked;
}
