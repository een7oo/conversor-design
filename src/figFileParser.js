/**
 * Figma .fig File Parser
 * 
 * Uses fig-kiwi's FigmaArchiveParser + kiwi-schema to parse .fig files
 * and extract layer blend modes that Photopea doesn't preserve.
 */
import { FigmaArchiveParser } from 'fig-kiwi';
import { decodeBinarySchema, compileSchema } from 'kiwi-schema';
import { inflateRaw } from 'pako';

/**
 * Figma blend mode → Photoshop BlendMode mapping
 */
const FIGMA_TO_PS_BLEND_MODE = {
    'PASS_THROUGH': 'BlendMode.PASSTHROUGH',
    'NORMAL': 'BlendMode.NORMAL',
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

/**
 * Parse a .fig file and extract all non-Normal/non-PassThrough blend modes.
 * 
 * @param {ArrayBuffer} arrayBuffer - The .fig file data
 * @returns {Map<string, string>} Map of layer name → PS BlendMode string
 */
export function extractBlendModes(arrayBuffer) {
    try {
        const uint8 = new Uint8Array(arrayBuffer);
        const { files } = FigmaArchiveParser.parseArchive(uint8);

        // files[0] = compressed schema, files[1] = compressed data
        const schema = decodeBinarySchema(inflateRaw(files[0]));
        const compiledSchema = compileSchema(schema);
        const message = compiledSchema.decodeMessage(inflateRaw(files[1]));

        const blendModes = new Map();

        if (message && message.nodeChanges) {
            for (const node of message.nodeChanges) {
                if (node.blendMode && node.blendMode !== 'NORMAL' && node.blendMode !== 'PASS_THROUGH') {
                    const psMode = FIGMA_TO_PS_BLEND_MODE[node.blendMode];
                    if (psMode && node.name) {
                        blendModes.set(node.name, psMode);
                    }
                }
            }
        }

        console.log(`📊 fig-kiwi: ${message?.nodeChanges?.length || 0} nodes parseados, ${blendModes.size} com blend mode não-Normal`);
        return blendModes;
    } catch (err) {
        console.warn('⚠️ Falha ao parsear .fig para blend modes:', err);
        return new Map();
    }
}
