/**
 * PPTX Service — parse PPTX files and extract slides/shapes/images
 * Also creates PPTX from PSD layers
 */
import JSZip from 'jszip';

/**
 * Parse a PPTX file and extract slide content
 * 
 * @param {ArrayBuffer} buffer - PPTX file data
 * @param {Function} onProgress - Progress callback
 * @returns {Object} Parsed presentation data
 */
export async function parsePptx(buffer, onProgress) {
    const zip = await JSZip.loadAsync(buffer);

    // Find slide files
    const slideFiles = Object.keys(zip.files)
        .filter(name => /^ppt\/slides\/slide\d+\.xml$/.test(name))
        .sort((a, b) => {
            const numA = parseInt(a.match(/slide(\d+)/)[1]);
            const numB = parseInt(b.match(/slide(\d+)/)[1]);
            return numA - numB;
        });

    // Get presentation dimensions
    const presXml = await zip.file('ppt/presentation.xml')?.async('text');
    const presDims = parsePresentationDimensions(presXml);

    const slides = [];

    for (let i = 0; i < slideFiles.length; i++) {
        const slideXml = await zip.file(slideFiles[i]).async('text');
        const slideNum = parseInt(slideFiles[i].match(/slide(\d+)/)[1]);

        // Get relationships for this slide
        const relsPath = `ppt/slides/_rels/slide${slideNum}.xml.rels`;
        const relsXml = await zip.file(relsPath)?.async('text');
        const relationships = parseRelationships(relsXml);

        // Parse slide content
        const slideData = parseSlideXml(slideXml, relationships);

        // Extract images referenced by this slide
        for (const shape of slideData.shapes) {
            if (shape.type === 'image' && shape.imageRef) {
                const rel = relationships.find(r => r.id === shape.imageRef);
                if (rel) {
                    const imagePath = `ppt/slides/${rel.target}`.replace(/\/\.\.\//, '/').replace('slides/../', '');
                    const imageFile = zip.file(imagePath) || zip.file(`ppt/${rel.target.replace('../', '')}`);
                    if (imageFile) {
                        shape.imageData = await imageFile.async('uint8array');
                        shape.imageMime = getImageMime(rel.target);
                    }
                }
            }
        }

        slides.push({
            number: slideNum,
            shapes: slideData.shapes,
        });

        if (onProgress) onProgress(i + 1, slideFiles.length);
    }

    return {
        width: presDims.width,
        height: presDims.height,
        slides,
    };
}

/**
 * Parse presentation.xml for slide dimensions
 */
function parsePresentationDimensions(xml) {
    if (!xml) return { width: 914400 * 10 / 914400 * 96, height: 914400 * 7.5 / 914400 * 96 }; // Default 10x7.5 inches

    const parser = new DOMParser();
    const doc = parser.parseFromString(xml, 'text/xml');
    const sldSz = doc.querySelector('sldSz');

    if (sldSz) {
        // EMU to pixels (1 inch = 914400 EMU, 96 DPI)
        const cx = parseInt(sldSz.getAttribute('cx')) || 9144000;
        const cy = parseInt(sldSz.getAttribute('cy')) || 6858000;
        return {
            width: Math.round(cx / 914400 * 96),
            height: Math.round(cy / 914400 * 96),
            emuWidth: cx,
            emuHeight: cy,
        };
    }

    return { width: 960, height: 720, emuWidth: 9144000, emuHeight: 6858000 };
}

/**
 * Parse relationship file
 */
function parseRelationships(xml) {
    if (!xml) return [];

    const parser = new DOMParser();
    const doc = parser.parseFromString(xml, 'text/xml');
    const rels = doc.querySelectorAll('Relationship');

    return Array.from(rels).map(rel => ({
        id: rel.getAttribute('Id'),
        type: rel.getAttribute('Type'),
        target: rel.getAttribute('Target'),
    }));
}

/**
 * Parse slide XML for shapes and images
 */
function parseSlideXml(xml, relationships) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(xml, 'text/xml');
    const shapes = [];

    // Find all shape elements
    const spElements = doc.querySelectorAll('sp');
    const picElements = doc.querySelectorAll('pic');
    const grpSpElements = doc.querySelectorAll('grpSp');

    // Process shapes (text boxes, rectangles, etc.)
    for (const sp of spElements) {
        const shape = parseShapeElement(sp);
        if (shape) shapes.push(shape);
    }

    // Process images
    for (const pic of picElements) {
        const imageShape = parsePicElement(pic);
        if (imageShape) shapes.push(imageShape);
    }

    return { shapes };
}

/**
 * Parse a <sp> element (shape / text)
 */
function parseShapeElement(sp) {
    const nvSpPr = sp.querySelector('nvSpPr');
    const name = nvSpPr?.querySelector('cNvPr')?.getAttribute('name') || 'Shape';

    const spPr = sp.querySelector('spPr');
    const xfrm = spPr?.querySelector('xfrm');
    const bounds = parseTransform(xfrm);

    // Check for text
    const txBody = sp.querySelector('txBody');
    let text = '';
    if (txBody) {
        const paragraphs = txBody.querySelectorAll('p');
        const textParts = [];
        for (const p of paragraphs) {
            const runs = p.querySelectorAll('r');
            for (const r of runs) {
                const t = r.querySelector('t');
                if (t) textParts.push(t.textContent);
            }
        }
        text = textParts.join('\n');
    }

    // Check for fill color
    const solidFill = spPr?.querySelector('solidFill');
    let fillColor = null;
    if (solidFill) {
        const srgb = solidFill.querySelector('srgbClr');
        if (srgb) fillColor = '#' + srgb.getAttribute('val');
    }

    return {
        type: text ? 'text' : 'shape',
        name,
        text,
        fillColor,
        ...bounds,
    };
}

/**
 * Parse a <pic> element (image)
 */
function parsePicElement(pic) {
    const nvPicPr = pic.querySelector('nvPicPr');
    const name = nvPicPr?.querySelector('cNvPr')?.getAttribute('name') || 'Image';

    const spPr = pic.querySelector('spPr');
    const xfrm = spPr?.querySelector('xfrm');
    const bounds = parseTransform(xfrm);

    // Get image reference
    const blipFill = pic.querySelector('blipFill');
    const blip = blipFill?.querySelector('blip');
    const embedRef = blip?.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed');

    return {
        type: 'image',
        name,
        imageRef: embedRef,
        ...bounds,
    };
}

/**
 * Parse transform (position and size) from xfrm element
 */
function parseTransform(xfrm) {
    if (!xfrm) return { left: 0, top: 0, width: 100, height: 100 };

    const off = xfrm.querySelector('off');
    const ext = xfrm.querySelector('ext');

    const emuToPixels = (emu) => Math.round(parseInt(emu || '0') / 914400 * 96);

    return {
        left: emuToPixels(off?.getAttribute('x')),
        top: emuToPixels(off?.getAttribute('y')),
        width: emuToPixels(ext?.getAttribute('cx')),
        height: emuToPixels(ext?.getAttribute('cy')),
    };
}

function getImageMime(filename) {
    const ext = filename.split('.').pop().toLowerCase();
    const mimes = {
        png: 'image/png',
        jpg: 'image/jpeg',
        jpeg: 'image/jpeg',
        gif: 'image/gif',
        bmp: 'image/bmp',
        webp: 'image/webp',
        svg: 'image/svg+xml',
        emf: 'image/emf',
        wmf: 'image/wmf',
        tiff: 'image/tiff',
        tif: 'image/tiff',
    };
    return mimes[ext] || 'image/png';
}

/**
 * Load image data from Uint8Array into canvas pixel data
 */
export function loadImageFromBytes(bytes, mime) {
    return new Promise((resolve, reject) => {
        const blob = new Blob([bytes], { type: mime });
        const url = URL.createObjectURL(blob);
        const img = new Image();

        img.onload = () => {
            const canvas = document.createElement('canvas');
            canvas.width = img.naturalWidth;
            canvas.height = img.naturalHeight;
            const ctx = canvas.getContext('2d');
            ctx.drawImage(img, 0, 0);
            URL.revokeObjectURL(url);
            resolve(canvas);
        };

        img.onerror = () => {
            URL.revokeObjectURL(url);
            reject(new Error('Failed to load image'));
        };

        img.src = url;
    });
}
