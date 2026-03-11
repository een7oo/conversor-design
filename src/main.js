/**
 * Design Converter — Main Entry Point
 * Wires UI events to conversion services
 */
import { parseFigmaUrl, fetchFigmaFile, collectNodes, collectFrames, collectBlendModes, collectTextNodes, collectMaskedGroups, exportFrameAsSvg, renderNodes, loadImageData, fetchImageFills, downloadImage } from './figmaService.js';
import { buildPsdFromFigma, readPsdFile, extractPsdLayers } from './psdService.js';
import { parsePptx, loadImageFromBytes } from './pptxService.js';
import { initPhotopea, loadFileInPhotopea, getDocumentStructure, exportAsPsd, runScript, destroyPhotopea } from './photopeaService.js';
import { writePsdUint8Array } from 'ag-psd';
import JSZip from 'jszip';


// ==========================================
// State
// ==========================================
let currentMode = 'figma-to-psd';
let selectedFile = null;
let resultBlob = null;
let resultFilename = '';

// ==========================================
// DOM References
// ==========================================
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

// Mode buttons
const modeButtons = $$('.mode-btn');

// Panels
const panels = $$('.panel');

// Progress
const progressContainer = $('#progress-container');
const progressTitle = $('#progress-title');
const progressPercent = $('#progress-percent');
const progressFill = $('#progress-fill');
const progressStatus = $('#progress-status');
const progressLayers = $('#progress-layers');
const progressLayersList = $('#progress-layers-list');

// Result
const resultContainer = $('#result-container');
const resultDesc = $('#result-desc');

// Error
const errorContainer = $('#error-container');
const errorDesc = $('#error-desc');

// Load saved Figma token from localStorage
const savedToken = localStorage.getItem('figma-api-token');
if (savedToken) {
    const tokenInput = $('#figma-token');
    if (tokenInput) tokenInput.value = savedToken;
    const tokenStatus = $('#token-status');
    if (tokenStatus) tokenStatus.textContent = '(salvo)';
    // Keep details closed since token is saved
} else {
    // Open details if no token saved yet
    const tokenDetails = $('#token-details');
    if (tokenDetails) tokenDetails.open = true;
}

// ==========================================
// Mode Switching
// ==========================================
modeButtons.forEach(btn => {
    btn.addEventListener('click', () => {
        const mode = btn.dataset.mode;
        switchMode(mode);
    });
});

function switchMode(mode) {
    currentMode = mode;

    modeButtons.forEach(btn => {
        btn.classList.toggle('active', btn.dataset.mode === mode);
    });

    panels.forEach(panel => {
        panel.classList.toggle('active', panel.id === `panel-${mode}`);
    });

    // Hide progress/result/error
    hideAll();
}

// ==========================================
// Sub-tab Switching (Figma panel)
// ==========================================
const subTabs = document.querySelectorAll('.sub-tab');
const subPanels = document.querySelectorAll('.sub-panel');

subTabs.forEach(tab => {
    tab.addEventListener('click', () => {
        const targetId = tab.dataset.subtab;
        subTabs.forEach(t => t.classList.toggle('active', t.dataset.subtab === targetId));
        subPanels.forEach(p => p.classList.toggle('active', p.id === `subpanel-${targetId}`));
    });
});

// ==========================================
// Upload Zones
// ==========================================
function setupUploadZone(zoneId, inputId, fileInfoId, fileNameId, removeId, extensions, onFileSelected) {
    const zone = $(zoneId);
    const input = $(inputId);
    const fileInfo = $(fileInfoId);
    const fileName = $(fileNameId);
    const removeBtn = $(removeId);

    if (!zone || !input) return;

    zone.addEventListener('click', () => input.click());

    zone.addEventListener('dragover', (e) => {
        e.preventDefault();
        zone.classList.add('drag-over');
    });

    zone.addEventListener('dragleave', () => {
        zone.classList.remove('drag-over');
    });

    zone.addEventListener('drop', (e) => {
        e.preventDefault();
        zone.classList.remove('drag-over');
        const files = e.dataTransfer.files;
        if (files.length > 0) handleFile(files[0]);
    });

    input.addEventListener('change', () => {
        if (input.files.length > 0) handleFile(input.files[0]);
    });

    if (removeBtn) {
        removeBtn.addEventListener('click', () => {
            selectedFile = null;
            fileInfo.classList.add('hidden');
            zone.classList.remove('hidden');
            onFileSelected(null);
        });
    }

    function handleFile(file) {
        const ext = file.name.split('.').pop().toLowerCase();
        if (!extensions.includes(ext)) {
            showError(`Formato inválido. Selecione um arquivo ${extensions.map(e => '.' + e).join(' ou ')}.`);
            return;
        }

        selectedFile = file;
        if (fileName) fileName.textContent = `📄 ${file.name} (${formatSize(file.size)})`;
        if (fileInfo) fileInfo.classList.remove('hidden');
        zone.classList.add('hidden');
        onFileSelected(file);
    }
}

// .fig upload
setupUploadZone(
    '#upload-zone-fig', '#fig-file-input', '#fig-file-info', '#fig-file-name', '#btn-remove-fig',
    ['fig'],
    (file) => {
        const btn = $('#btn-convert-fig');
        if (file) {
            btn.disabled = false;
            btn.classList.remove('disabled');
        } else {
            btn.disabled = true;
            btn.classList.add('disabled');
        }
    }
);

// PSD → Figma upload
setupUploadZone(
    '#upload-zone-psd', '#psd-file-input', '#psd-file-info', '#psd-file-name', '#btn-remove-psd',
    ['psd'],
    (file) => {
        const btn = $('#btn-convert-psd');
        if (file) {
            btn.disabled = false;
            btn.classList.remove('disabled');
        } else {
            btn.disabled = true;
            btn.classList.add('disabled');
        }
    }
);

// PPTX → PSD upload
setupUploadZone(
    '#upload-zone-pptx', '#pptx-file-input', '#pptx-file-info', '#pptx-file-name', '#btn-remove-pptx',
    ['pptx'],
    (file) => {
        const btn = $('#btn-convert-pptx');
        if (file) {
            btn.disabled = false;
            btn.classList.remove('disabled');
        } else {
            btn.disabled = true;
            btn.classList.add('disabled');
        }
    }
);

// PSD → PPTX upload
setupUploadZone(
    '#upload-zone-psd-pptx', '#psd-pptx-file-input', '#psd-pptx-file-info', '#psd-pptx-file-name', '#btn-remove-psd-pptx',
    ['psd'],
    (file) => {
        const btn = $('#btn-convert-psd-pptx');
        if (file) {
            btn.disabled = false;
            btn.classList.remove('disabled');
        } else {
            btn.disabled = true;
            btn.classList.add('disabled');
        }
    }
);

// ==========================================
// Conversion: .fig → PSD (via Photopea — FAST)
// ==========================================
$('#btn-convert-fig')?.addEventListener('click', async () => {
    if (!selectedFile) return;

    try {
        showProgress('Iniciando Photopea...');

        // Step 1: Initialize Photopea iframe
        updateProgress(5, 'Carregando Photopea (pode levar alguns segundos)...');
        await initPhotopea();

        // Step 2: Load .fig file into Photopea
        updateProgress(25, 'Abrindo arquivo .fig no Photopea...');
        const buffer = await selectedFile.arrayBuffer();
        await loadFileInPhotopea(buffer);

        // Step 3: Fix pattern fills → Smart Objects
        updateProgress(50, 'Convertendo imagens para Smart Objects...');
        await runScript(`
            function fixLayers(parent) {
                for (var i = parent.layers.length - 1; i >= 0; i--) {
                    var layer = parent.layers[i];
                    if (layer.typename === "LayerSet") {
                        fixLayers(layer);
                    } else if (layer.typename === "ArtLayer") {
                        if (layer.kind === LayerKind.PATTERNFILL) {
                            try {
                                var savedOpacity = layer.opacity;
                                var savedName = layer.name;

                                app.activeDocument.activeLayer = layer;
                                layer.rasterize(RasterizeType.ENTIRELAYER);

                                // Convert to Smart Object
                                var desc = new ActionDescriptor();
                                executeAction(stringIDToTypeID("newPlacedLayer"), desc, DialogModes.NO);

                                // Restore opacity and name
                                var smartLayer = app.activeDocument.activeLayer;
                                smartLayer.opacity = savedOpacity;
                                smartLayer.name = savedName;
                            } catch(e) {}
                        }
                    }
                }
            }
            fixLayers(app.activeDocument);
        `);

        // Step 3b: Clean up Smart Objects — keep only the top layer inside each
        // Figma stacks multiple fills as layers inside a Smart Object, but we only
        // want the topmost (visible) one. This opens each SO, removes bottom layers,
        // flattens, saves and closes.
        updateProgress(55, 'Limpando camadas internas dos Smart Objects...');
        await runScript(`
            function cleanSmartObjects(parent) {
                for (var i = parent.layers.length - 1; i >= 0; i--) {
                    var layer = parent.layers[i];
                    if (layer.typename === "LayerSet") {
                        cleanSmartObjects(layer);
                    } else if (layer.typename === "ArtLayer" && layer.kind === LayerKind.SMARTOBJECT) {
                        try {
                            var mainDoc = app.activeDocument;
                            app.activeDocument.activeLayer = layer;

                            // Open the Smart Object (edits its contents)
                            var desc = new ActionDescriptor();
                            executeAction(stringIDToTypeID("placedLayerEditContents"), desc, DialogModes.NO);

                            var soDoc = app.activeDocument;
                            // Only clean if there are multiple layers
                            if (soDoc.layers.length > 1) {
                                // Hide all layers except the topmost one
                                for (var j = 1; j < soDoc.layers.length; j++) {
                                    soDoc.layers[j].visible = false;
                                }
                                // Flatten to keep only the visible top layer
                                soDoc.flattenImage();
                            }

                            // Save and close the Smart Object
                            soDoc.save();
                            soDoc.close();
                        } catch(e) {}
                    }
                }
            }
            cleanSmartObjects(app.activeDocument);
        `);

        // Step 4: Keep paragraph text as-is (do NOT convert to point text)
        // Converting to POINTTEXT was truncating multi-line text to only the first line
        // and shifting position upward. Paragraph text preserves all lines correctly.
        updateProgress(65, 'Verificando camadas de texto...');

        // Step 4b: Force re-render all text layers so the cached bitmap
        // matches the actual font data (fixes wrong visual font style)
        updateProgress(70, 'Corrigindo renderização de fontes...');
        await runScript(`
            function touchTextLayers(parent) {
                for (var i = 0; i < parent.layers.length; i++) {
                    var layer = parent.layers[i];
                    if (layer.typename === "LayerSet") {
                        touchTextLayers(layer);
                    } else if (layer.typename === "ArtLayer" && layer.kind === LayerKind.TEXT) {
                        try {
                            var t = layer.textItem;
                            t.contents = t.contents;
                        } catch(e) {}
                    }
                }
            }
            touchTextLayers(app.activeDocument);
        `);

        // Step 5: Export entire document as PSD directly from Photopea
        updateProgress(80, 'Exportando como PSD...');
        const psdArrayBuffer = await exportAsPsd();

        // Step 5: Cleanup Photopea
        destroyPhotopea();

        updateProgress(100, 'Concluído!');

        resultBlob = new Blob([psdArrayBuffer], { type: 'application/octet-stream' });
        resultFilename = `${selectedFile.name.replace('.fig', '')}.psd`;

        showResult('PSD exportado com sucesso! Texto editável, blend modes e imagens preservados.');

    } catch (err) {
        console.error(err);
        destroyPhotopea();
        showError(err.message);
    }
});

// ==========================================
// Conversion: Figma → PSD (Hybrid: SVG + Photopea)
// ==========================================
$('#btn-convert-figma')?.addEventListener('click', async () => {
    const url = $('#figma-url').value.trim();
    const token = $('#figma-token').value.trim();
    const allPages = $('#opt-all-pages').checked;

    if (!url) return showError('Cole o link do arquivo Figma.');
    if (!token) return showError('Insira seu token de API. Clique em "Token de API" para expandir.');

    // Save token to localStorage for next time
    localStorage.setItem('figma-api-token', token);
    const tokenStatus = $('#token-status');
    if (tokenStatus) tokenStatus.textContent = '(salvo)';

    try {
        showProgress('Conectando ao Figma...');

        // Step 1: Parse URL and fetch file structure
        const { fileKey } = parseFigmaUrl(url);

        updateProgress(5, 'Baixando estrutura do arquivo...');
        const fileData = await fetchFigmaFile(fileKey, token);

        // Step 2: Collect frames, blend modes, text nodes, and masked groups
        updateProgress(10, 'Analisando frames, textos, blend modes e mascaras...');
        const frames = collectFrames(fileData, allPages);
        const blendModeMap = collectBlendModes(fileData);
        const textNodes = collectTextNodes(fileData);
        const maskedGroups = collectMaskedGroups(fileData);

        if (frames.length === 0) {
            throw new Error('Nenhum frame encontrado no arquivo.');
        }

        console.log(`Frames: ${frames.length}, Texts: ${textNodes.length}, Blend modes: ${blendModeMap.size}, Masked groups: ${maskedGroups.length}`);

        // Step 3: Export each frame as SVG from Figma API
        updateProgress(15, `Exportando ${frames.length} frames como SVG...`);
        const svgDataArray = [];
        for (let fi = 0; fi < frames.length; fi++) {
            const frame = frames[fi];
            const pct = 15 + Math.round((fi / frames.length) * 25);
            updateProgress(pct, `Exportando SVG: ${frame.name} (${fi + 1}/${frames.length})...`);

            try {
                const svgText = await exportFrameAsSvg(fileKey, token, frame.id);
                svgDataArray.push({ frame, svgText });
            } catch (e) {
                console.warn(`Falha ao exportar frame ${frame.name}:`, e);
            }
        }

        if (svgDataArray.length === 0) {
            throw new Error('Nenhum frame foi exportado com sucesso.');
        }

        // Step 4: Initialize Photopea
        updateProgress(42, 'Iniciando Photopea...');
        await initPhotopea();

        // Step 5: Open the first SVG in Photopea
        const firstSvg = svgDataArray[0];
        updateProgress(48, `Abrindo "${firstSvg.frame.name}" no Photopea...`);

        // PRE-PROCESS SVG: Remove vectorized text groups so we don't get duplicate shapes
        let finalSvgText = firstSvg.svgText;
        if (textNodes.length > 0) {
            try {
                const parser = new DOMParser();
                const svgDoc = parser.parseFromString(firstSvg.svgText, 'image/svg+xml');
                let removedCount = 0;
                for (const tn of textNodes) {
                    // Figma SVG export with svg_include_id=true uses "id" attribute matching the node id
                    const el = svgDoc.getElementById(tn.id);
                    if (el) {
                        el.remove();
                        removedCount++;
                    }
                }
                if (removedCount > 0) {
                    finalSvgText = new XMLSerializer().serializeToString(svgDoc);
                    console.log(`Removed ${removedCount} vectorized text groups from SVG before import.`);
                }
            } catch (e) {
                console.warn('Failed to pre-process SVG text removal:', e);
            }
        }

        const svgBlob = new Blob([finalSvgText], { type: 'image/svg+xml' });
        const svgBuffer = await svgBlob.arrayBuffer();
        await loadFileInPhotopea(svgBuffer);

        // Step 5.5: Create proper clipping masks for mask groups
        if (maskedGroups.length > 0) {
            updateProgress(52, `Corrigindo ${maskedGroups.length} mascaras de corte...`);
            console.log('Mask groups found:', maskedGroups.map(g => g.name));

            // Render mask shapes + content for each mask group
            const allIds = [];
            for (const g of maskedGroups) {
                allIds.push(g.maskChildId);
                allIds.push(g.id); // renders the full group (correctly masked by Figma)
            }
            const rendered = await renderNodes(fileKey, token, allIds, null, 2);

            const frameX = firstSvg.frame.bounds?.x || 0;
            const frameY = firstSvg.frame.bounds?.y || 0;

            for (let gi = 0; gi < maskedGroups.length; gi++) {
                const g = maskedGroups[gi];
                const maskBytes = rendered[g.maskChildId];
                const contentBytes = rendered[g.id];
                if (!maskBytes || !contentBytes) continue;

                updateProgress(53 + gi, `Mascara ${gi + 1}/${maskedGroups.length}: ${g.name}...`);
                const relX = Math.round((g.bounds.x - frameX) * 2);
                const relY = Math.round((g.bounds.y - frameY) * 2);
                const eName = g.name.replace(/"/g, '\\"');

                // Remove the broken Clip-Path group from SVG import
                await runScript(`
                    function removeGroup(parent, name) {
                        for (var i = parent.layers.length - 1; i >= 0; i--) {
                            var l = parent.layers[i];
                            if ((l.name === "${eName}" || l.name === "Clip-Path") && l.typename === "LayerSet") {
                                l.remove(); return true;
                            }
                            if (l.typename === "LayerSet" && removeGroup(l, name)) return true;
                        }
                        return false;
                    }
                    removeGroup(app.activeDocument, "${eName}");
                `);

                // Place mask shape (bottom layer) — Smart Object with transparency
                await loadFileInPhotopea(await new Blob([maskBytes], { type: 'image/png' }).arrayBuffer());
                await runScript(`
                    if (app.documents.length > 1) {
                        var d = app.activeDocument;
                        d.selection.selectAll();
                        d.activeLayer.copy(true);
                        d.close(SaveOptions.DONOTSAVECHANGES);
                        app.activeDocument.paste();
                        var ml = app.activeDocument.activeLayer;
                        ml.name = "${g.maskChildName?.replace(/"/g, '\\"') || eName + ' shape'}";
                        var b = ml.bounds;
                        ml.translate(${relX} - b[0].as("px"), ${relY} - b[1].as("px"));
                        // Convert to Smart Object
                        var desc = new ActionDescriptor();
                        executeAction(stringIDToTypeID("newPlacedLayer"), desc, DialogModes.NO);
                        app.activeDocument.activeLayer.name = "${g.maskChildName?.replace(/"/g, '\\"') || eName + ' shape'}";
                    }
                `);

                // Place content (top layer, clipped to mask)
                await loadFileInPhotopea(await new Blob([contentBytes], { type: 'image/png' }).arrayBuffer());
                await runScript(`
                    if (app.documents.length > 1) {
                        var d = app.activeDocument;
                        d.selection.selectAll();
                        d.activeLayer.copy(true);
                        d.close(SaveOptions.DONOTSAVECHANGES);
                        app.activeDocument.paste();
                        var cl = app.activeDocument.activeLayer;
                        cl.name = "${eName}";
                        var b = cl.bounds;
                        cl.translate(${relX} - b[0].as("px"), ${relY} - b[1].as("px"));
                        cl.grouped = true;
                    }
                `);

                await new Promise(r => setTimeout(r, 200));
            }
        }

        // Step 6: Create editable text layers (SVG outlines text as vectors)
        if (textNodes.length > 0) {
            updateProgress(58, `Criando ${textNodes.length} camadas de texto editavel...`);
            const frameX = firstSvg.frame.bounds?.x || 0;
            const frameY = firstSvg.frame.bounds?.y || 0;

            // Process in chunks to avoid timeout
            const chunkSize = 20;
            for (let ti = 0; ti < textNodes.length; ti += chunkSize) {
                const chunk = textNodes.slice(ti, Math.min(ti + chunkSize, textNodes.length));
                updateProgress(58 + Math.round((ti / textNodes.length) * 5), `Textos: ${ti}/${textNodes.length}...`);

                // Build a single script for this chunk of text nodes
                let script = '';
                for (const tn of chunk) {
                    const relX = Math.round((tn.bounds.x - frameX) * 2);
                    const relY = Math.round((tn.bounds.y - frameY) * 2);
                    const fontSize = Math.round(tn.fontSize * 2);
                    const escaped = tn.text.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\n/g, '\\r');
                    const nameEsc = tn.name.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
                    const fontFamily = tn.fontFamily.replace(/"/g, '\\"');

                    script += `
                    try {
                        var textLayer = app.activeDocument.artLayers.add();
                        textLayer.kind = LayerKind.TEXT;
                        textLayer.textItem.contents = "${escaped}";
                        textLayer.textItem.size = ${fontSize};
                        textLayer.textItem.font = "${fontFamily}";
                        textLayer.textItem.color = new SolidColor();
                        textLayer.textItem.color.rgb.red = ${tn.color.r};
                        textLayer.textItem.color.rgb.green = ${tn.color.g};
                        textLayer.textItem.color.rgb.blue = ${tn.color.b};
                        textLayer.textItem.position = [${relX}, ${relY + fontSize}];
                        textLayer.textItem.kind = TextType.POINTTEXT;
                        textLayer.name = "${nameEsc}";
                        textLayer.opacity = ${Math.round(tn.opacity * 100)};
                    } catch(e) {}
                    `;
                }

                await runScript(script);
                await new Promise(r => setTimeout(r, 100));
            }
            console.log(`${textNodes.length} text layers criados`);
        }

        // Step 6: Convert PatternFill layers to Smart Objects (actual images)
        updateProgress(58, 'Convertendo imagens para Smart Objects...');
        await runScript(`
            function fixPatternFills(parent) {
                for (var i = parent.layers.length - 1; i >= 0; i--) {
                    var layer = parent.layers[i];
                    if (layer.typename === "LayerSet") {
                        fixPatternFills(layer);
                    } else if (layer.typename === "ArtLayer" && layer.kind === LayerKind.PATTERNFILL) {
                        try {
                            var savedName = layer.name;
                            var savedOpacity = layer.opacity;
                            app.activeDocument.activeLayer = layer;
                            layer.rasterize(RasterizeType.ENTIRELAYER);
                            var desc = new ActionDescriptor();
                            executeAction(stringIDToTypeID("newPlacedLayer"), desc, DialogModes.NO);
                            var so = app.activeDocument.activeLayer;
                            so.name = savedName;
                            so.opacity = savedOpacity;
                        } catch(e) {}
                    }
                }
            }
            fixPatternFills(app.activeDocument);
        `);

        // Step 7: Apply blend modes from Figma API
        if (blendModeMap.size > 0) {
            updateProgress(65, `Aplicando ${blendModeMap.size} blend modes...`);
            const blendEntries = Array.from(blendModeMap.entries());
            const blendMapJson = JSON.stringify(blendEntries);

            await runScript(`
                // Direct mapping without eval()
                var blendModeRef = {
                    "DARKEN": BlendMode.DARKEN,
                    "MULTIPLY": BlendMode.MULTIPLY,
                    "LINEARBURN": BlendMode.LINEARBURN,
                    "COLORBURN": BlendMode.COLORBURN,
                    "LIGHTEN": BlendMode.LIGHTEN,
                    "SCREEN": BlendMode.SCREEN,
                    "LINEARDODGE": BlendMode.LINEARDODGE,
                    "COLORDODGE": BlendMode.COLORDODGE,
                    "OVERLAY": BlendMode.OVERLAY,
                    "SOFTLIGHT": BlendMode.SOFTLIGHT,
                    "HARDLIGHT": BlendMode.HARDLIGHT,
                    "DIFFERENCE": BlendMode.DIFFERENCE,
                    "EXCLUSION": BlendMode.EXCLUSION,
                    "HUE": BlendMode.HUE,
                    "SATURATION": BlendMode.SATURATION,
                    "COLORBLEND": BlendMode.COLORBLEND,
                    "LUMINOSITY": BlendMode.LUMINOSITY,
                    "PASSTHROUGH": BlendMode.PASSTHROUGH
                };

                var blendMap = ${blendMapJson};
                var nameToMode = {};
                for (var i = 0; i < blendMap.length; i++) {
                    // blendMap[i][1] is like "BlendMode.MULTIPLY" -> extract "MULTIPLY"
                    var key = blendMap[i][1].replace("BlendMode.", "");
                    nameToMode[blendMap[i][0]] = blendModeRef[key];
                }

                function applyBlendModes(parent) {
                    for (var i = 0; i < parent.layers.length; i++) {
                        var layer = parent.layers[i];
                        if (nameToMode[layer.name]) {
                            try {
                                layer.blendMode = nameToMode[layer.name];
                            } catch(e) {}
                        }
                        if (layer.typename === "LayerSet") {
                            applyBlendModes(layer);
                        }
                    }
                }
                applyBlendModes(app.activeDocument);
            `);
        }

        // Step 8: Convert text to point text
        updateProgress(75, 'Convertendo textos...');
        await runScript(`
            function fixTextLayers(parent) {
                for (var i = 0; i < parent.layers.length; i++) {
                    var layer = parent.layers[i];
                    if (layer.typename === "LayerSet") {
                        fixTextLayers(layer);
                    } else if (layer.typename === "ArtLayer" && layer.kind === LayerKind.TEXT) {
                        try {
                            layer.textItem.kind = TextType.POINTTEXT;
                        } catch(e) {}
                    }
                }
            }
            fixTextLayers(app.activeDocument);
        `);

        // Step 9: Export as PSD
        updateProgress(85, 'Exportando como PSD...');
        const psdArrayBuffer = await exportAsPsd();

        // Step 10: Cleanup
        destroyPhotopea();
        updateProgress(100, 'Concluido!');

        resultBlob = new Blob([psdArrayBuffer], { type: 'application/octet-stream' });
        resultFilename = `${fileData.name || 'design'}.psd`;

        const parts = [`${frames.length} frames`];
        if (maskedGroups.length > 0) parts.push(`${maskedGroups.length} mascaras corrigidas`);
        if (blendModeMap.size > 0) parts.push(`${blendModeMap.size} blend modes`);
        showResult(`Convertido com sucesso! ${parts.join(', ')}.`);

    } catch (err) {
        console.error(err);
        destroyPhotopea();
        showError(err.message);
    }
});

// ==========================================
// Conversion: PSD → Figma (export as ZIP)
// ==========================================
$('#btn-convert-psd')?.addEventListener('click', async () => {
    if (!selectedFile) return;

    try {
        showProgress('Lendo arquivo PSD...');

        const buffer = await selectedFile.arrayBuffer();

        updateProgress(15, 'Decodificando camadas...');
        const psd = readPsdFile(buffer);

        updateProgress(30, 'Extraindo layers...');
        showLayersList();

        let layerCount = 0;
        const layers = await extractPsdLayers(psd, (count) => {
            layerCount = count;
            updateProgress(30 + Math.round((count / (psd.children?.length || 1)) * 40), `Processando layer ${count}...`);
        });

        const layerLayers = layers.filter(l => l.type === 'layer');

        updateProgress(75, 'Criando ZIP com PNGs...');

        // Build a ZIP with:
        // - /layers/ — each layer as PNG
        // - manifest.json — layer positions and metadata for Figma import
        const zip = new JSZip();
        const manifest = {
            canvasWidth: psd.width,
            canvasHeight: psd.height,
            layers: [],
        };

        for (let i = 0; i < layerLayers.length; i++) {
            const layer = layerLayers[i];
            const safeName = layer.name.replace(/[\\/:*?"<>|]/g, '_');
            const filename = `${String(i + 1).padStart(3, '0')}_${safeName}.png`;

            if (layer.pngBlob) {
                const pngBuffer = await layer.pngBlob;
                if (pngBuffer) {
                    const arrayBuffer = await pngBuffer.arrayBuffer();
                    zip.file(`layers/${filename}`, arrayBuffer);
                    addLayerToList(layer.name);
                }
            } else if (layer.canvas) {
                const blob = await new Promise(res => layer.canvas.toBlob(res, 'image/png'));
                if (blob) {
                    const arrayBuffer = await blob.arrayBuffer();
                    zip.file(`layers/${filename}`, arrayBuffer);
                    addLayerToList(layer.name);
                }
            }

            manifest.layers.push({
                filename,
                name: layer.name,
                path: layer.path,
                x: layer.left,
                y: layer.top,
                width: layer.width,
                height: layer.height,
                opacity: layer.opacity,
                blendMode: layer.blendMode,
                hidden: layer.hidden,
            });

            updateProgress(75 + Math.round((i / layerLayers.length) * 20), `Exportando: ${i + 1}/${layerLayers.length}`);
        }

        // Add manifest
        zip.file('manifest.json', JSON.stringify(manifest, null, 2));

        // Add README with Figma import instructions
        zip.file('LEIAME.txt', `Design Converter — PSD para Figma
=============================================

Este ZIP contém as camadas do PSD como imagens PNG separadas
com um arquivo manifest.json contendo as posições de cada camada.

COMO IMPORTAR NO FIGMA:
1. Abra o Figma e crie um novo arquivo
2. Crie um Frame com ${psd.width}x${psd.height}px
3. Arraste todas as PNGs da pasta "layers/" para dentro do Frame
4. Use o manifest.json para reposicionar cada camada (x, y)

DICA: Para posicionamento automático, instale o plugin 
"Import from ZIP" ou use a Figma API com o manifest.json.

Dimensões do canvas: ${psd.width} x ${psd.height}px
Total de camadas: ${layerLayers.length}
`);

        updateProgress(98, 'Comprimindo ZIP...');
        const zipBlob = await zip.generateAsync({ type: 'blob' });

        updateProgress(100, 'Concluído!');

        resultBlob = zipBlob;
        resultFilename = `${selectedFile.name.replace('.psd', '')}_figma_layers.zip`;

        showResult(`${layerLayers.length} camadas extraídas. ZIP com PNGs + manifest para Figma.`);

    } catch (err) {
        console.error(err);
        showError(err.message);
    }
});

// ==========================================
// Conversion: PPTX → PSD
// ==========================================
$('#btn-convert-pptx')?.addEventListener('click', async () => {
    if (!selectedFile) return;

    try {
        showProgress('Lendo arquivo PPTX...');

        const buffer = await selectedFile.arrayBuffer();

        updateProgress(15, 'Descompactando slides...');
        const presentation = await parsePptx(buffer, (done, total) => {
            updateProgress(15 + Math.round((done / total) * 30), `Processando slide ${done}/${total}...`);
        });

        showLayersList();

        updateProgress(50, 'Construindo PSD...');

        // Build PSD with each slide as a group, shapes as layers
        const psdChildren = [];

        for (let si = 0; si < presentation.slides.length; si++) {
            const slide = presentation.slides[si];
            const slideGroup = {
                name: `Slide ${slide.number}`,
                children: [],
                opened: true,
            };

            for (const shape of slide.shapes) {
                if (shape.type === 'image' && shape.imageData) {
                    // Create image layer
                    try {
                        const canvas = await loadImageFromBytes(shape.imageData, shape.imageMime);

                        // Scale image to fit the shape bounds
                        const scaledCanvas = document.createElement('canvas');
                        scaledCanvas.width = shape.width;
                        scaledCanvas.height = shape.height;
                        const ctx = scaledCanvas.getContext('2d');
                        ctx.drawImage(canvas, 0, 0, shape.width, shape.height);

                        slideGroup.children.push({
                            name: shape.name,
                            left: shape.left,
                            top: shape.top,
                            right: shape.left + shape.width,
                            bottom: shape.top + shape.height,
                            canvas: scaledCanvas,
                        });

                        addLayerToList(`Slide ${slide.number} / ${shape.name}`);
                    } catch (e) {
                        console.warn('Failed to load image:', e);
                    }
                } else if (shape.type === 'text' && shape.text) {
                    // For text, render on a canvas
                    const textCanvas = renderTextToCanvas(shape);
                    if (textCanvas) {
                        slideGroup.children.push({
                            name: shape.name || shape.text.substring(0, 30),
                            left: shape.left,
                            top: shape.top,
                            right: shape.left + shape.width,
                            bottom: shape.top + shape.height,
                            canvas: textCanvas,
                        });
                        addLayerToList(`Slide ${slide.number} / ${shape.name}`);
                    }
                } else if (shape.type === 'shape' && shape.fillColor) {
                    // Render filled shape
                    const shapeCanvas = document.createElement('canvas');
                    shapeCanvas.width = Math.max(shape.width, 1);
                    shapeCanvas.height = Math.max(shape.height, 1);
                    const ctx = shapeCanvas.getContext('2d');
                    ctx.fillStyle = shape.fillColor;
                    ctx.fillRect(0, 0, shape.width, shape.height);

                    slideGroup.children.push({
                        name: shape.name,
                        left: shape.left,
                        top: shape.top,
                        right: shape.left + shape.width,
                        bottom: shape.top + shape.height,
                        canvas: shapeCanvas,
                    });
                    addLayerToList(`Slide ${slide.number} / ${shape.name}`);
                }
            }

            if (slideGroup.children.length > 0) {
                psdChildren.push(slideGroup);
            }

            updateProgress(50 + Math.round(((si + 1) / presentation.slides.length) * 40),
                `Convertendo slide ${si + 1}/${presentation.slides.length}...`);
        }

        const psd = {
            width: presentation.width,
            height: presentation.height,
            children: psdChildren,
        };

        updateProgress(95, 'Gerando arquivo PSD...');

        const { writePsdUint8Array: writePsd } = await import('ag-psd');
        const psdBytes = writePsd(psd);

        updateProgress(100, 'Concluído!');

        resultBlob = new Blob([psdBytes], { type: 'application/octet-stream' });
        resultFilename = `${selectedFile.name.replace('.pptx', '')}.psd`;

        showResult(`${presentation.slides.length} slides convertidos para PSD com ${psdChildren.reduce((a, g) => a + g.children.length, 0)} camadas.`);

    } catch (err) {
        console.error(err);
        showError(err.message);
    }
});

// ==========================================
// Conversion: PSD → PPTX
// ==========================================
$('#btn-convert-psd-pptx')?.addEventListener('click', async () => {
    if (!selectedFile) return;

    try {
        showProgress('Lendo arquivo PSD...');

        const buffer = await selectedFile.arrayBuffer();

        updateProgress(15, 'Decodificando camadas...');
        const psd = readPsdFile(buffer);

        updateProgress(30, 'Extraindo layers...');
        showLayersList();

        const layers = await extractPsdLayers(psd, (count) => {
            updateProgress(30 + Math.round((count / (psd.children?.length || 1)) * 30), `Layer ${count}...`);
        });

        const layerLayers = layers.filter(l => l.type === 'layer' && l.canvas);

        updateProgress(65, 'Construindo PPTX...');

        // Build a PPTX manually with JSZip
        const zip = new JSZip();

        // Width/height in EMU (English Metric Units)
        const slideW = Math.round(psd.width * 914400 / 96);
        const slideH = Math.round(psd.height * 914400 / 96);

        // Content Types
        zip.file('[Content_Types].xml', buildContentTypes(layerLayers.length));

        // Relationships
        zip.file('_rels/.rels', buildRels());

        // Presentation
        zip.file('ppt/presentation.xml', buildPresentation(slideW, slideH));
        zip.file('ppt/_rels/presentation.xml.rels', buildPresentationRels());

        // Slide Layout (minimal)
        zip.file('ppt/slideLayouts/slideLayout1.xml', buildSlideLayout(slideW, slideH));
        zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', buildSlideLayoutRels());

        // Slide Master
        zip.file('ppt/slideMasters/slideMaster1.xml', buildSlideMaster(slideW, slideH));
        zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', buildSlideMasterRels());

        // Theme
        zip.file('ppt/theme/theme1.xml', buildTheme());

        // Build slide with all layers as images
        const imageFiles = [];
        for (let i = 0; i < layerLayers.length; i++) {
            const layer = layerLayers[i];
            const blob = await new Promise(res => layer.canvas.toBlob(res, 'image/png'));
            if (blob) {
                const arrayBuffer = await blob.arrayBuffer();
                const filename = `image${i + 1}.png`;
                zip.file(`ppt/media/${filename}`, arrayBuffer);
                imageFiles.push({
                    filename,
                    layer,
                    relId: `rId${i + 3}`, // rId1=slideLayout, rId2=reserved
                });
                addLayerToList(layer.name);
            }
            updateProgress(65 + Math.round((i / layerLayers.length) * 25), `Exportando layer ${i + 1}/${layerLayers.length}`);
        }

        // Slide XML
        zip.file('ppt/slides/slide1.xml', buildSlideXml(imageFiles, slideW, slideH, psd.width, psd.height));
        zip.file('ppt/slides/_rels/slide1.xml.rels', buildSlideRels(imageFiles));

        updateProgress(95, 'Comprimindo PPTX...');

        const pptxBlob = await zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });

        updateProgress(100, 'Concluído!');

        resultBlob = pptxBlob;
        resultFilename = `${selectedFile.name.replace('.psd', '')}.pptx`;

        showResult(`${layerLayers.length} camadas convertidas para PPTX.`);

    } catch (err) {
        console.error(err);
        showError(err.message);
    }
});

// ==========================================
// PPTX XML Builders
// ==========================================

function buildContentTypes(imageCount) {
    let imageOverrides = '';
    for (let i = 1; i <= imageCount; i++) {
        // PNG extension is already covered by Default
    }
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`;
}

function buildRels() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`;
}

function buildPresentation(w, h) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>
  <p:sldSz cx="${w}" cy="${h}"/>
  <p:notesSz cx="${h}" cy="${w}"/>
</p:presentation>`;
}

function buildPresentationRels() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`;
}

function buildSlideLayout(w, h) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>
</p:sldLayout>`;
}

function buildSlideLayoutRels() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`;
}

function buildSlideMaster(w, h) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
</p:sldMaster>`;
}

function buildSlideMasterRels() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`;
}

function buildTheme() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Design Converter">
  <a:themeElements>
    <a:clrScheme name="Custom">
      <a:dk1><a:srgbClr val="000000"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="191919"/></a:dk2>
      <a:lt2><a:srgbClr val="F7F6F3"/></a:lt2>
      <a:accent1><a:srgbClr val="2B2B2B"/></a:accent1>
      <a:accent2><a:srgbClr val="555555"/></a:accent2>
      <a:accent3><a:srgbClr val="888888"/></a:accent3>
      <a:accent4><a:srgbClr val="AAAAAA"/></a:accent4>
      <a:accent5><a:srgbClr val="CCCCCC"/></a:accent5>
      <a:accent6><a:srgbClr val="EEEEEE"/></a:accent6>
      <a:hlink><a:srgbClr val="2B2B2B"/></a:hlink>
      <a:folHlink><a:srgbClr val="555555"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Custom"><a:majorFont><a:latin typeface="Inter"/></a:majorFont><a:minorFont><a:latin typeface="Inter"/></a:minorFont></a:fontScheme>
    <a:fmtScheme name="Custom"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>
  </a:themeElements>
</a:theme>`;
}

function buildSlideXml(imageFiles, slideW, slideH, pxW, pxH) {
    let shapesXml = '';
    const pxToEmu = (px) => Math.round(px * 914400 / 96);

    for (let i = 0; i < imageFiles.length; i++) {
        const { layer, relId } = imageFiles[i];
        const x = pxToEmu(layer.left || 0);
        const y = pxToEmu(layer.top || 0);
        const cx = pxToEmu(layer.width || 100);
        const cy = pxToEmu(layer.height || 100);

        shapesXml += `
      <p:pic>
        <p:nvPicPr>
          <p:cNvPr id="${i + 2}" name="${escapeXml(layer.name)}"/>
          <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
          <p:nvPr/>
        </p:nvPicPr>
        <p:blipFill>
          <a:blip r:embed="${relId}"/>
          <a:stretch><a:fillRect/></a:stretch>
        </p:blipFill>
        <p:spPr>
          <a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
      </p:pic>`;
    }

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      ${shapesXml}
    </p:spTree>
  </p:cSld>
</p:sld>`;
}

function buildSlideRels(imageFiles) {
    let rels = '';
    for (const img of imageFiles) {
        rels += `<Relationship Id="${img.relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${img.filename}"/>`;
    }
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  ${rels}
</Relationships>`;
}

function escapeXml(str) {
    return (str || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ==========================================
// Text Rendering Helper
// ==========================================
function renderTextToCanvas(shape) {
    if (!shape.text || !shape.width || !shape.height) return null;

    const canvas = document.createElement('canvas');
    canvas.width = Math.max(shape.width, 1);
    canvas.height = Math.max(shape.height, 1);
    const ctx = canvas.getContext('2d');

    if (shape.fillColor) {
        ctx.fillStyle = shape.fillColor;
        ctx.fillRect(0, 0, canvas.width, canvas.height);
    }

    ctx.fillStyle = '#191919';
    ctx.font = `${Math.min(canvas.height * 0.6, 24)}px Inter, sans-serif`;
    ctx.textBaseline = 'middle';
    ctx.textAlign = 'left';

    // Simple text wrapping
    const lines = shape.text.split('\n');
    const lineHeight = Math.min(canvas.height * 0.7, 28);
    const startY = canvas.height / 2 - (lines.length - 1) * lineHeight / 2;

    for (let i = 0; i < lines.length; i++) {
        ctx.fillText(lines[i], 8, startY + i * lineHeight);
    }

    return canvas;
}

// ==========================================
// UI Helpers
// ==========================================
function hideAll() {
    progressContainer.classList.add('hidden');
    resultContainer.classList.add('hidden');
    errorContainer.classList.add('hidden');
}

function showProgress(title) {
    hideAll();
    progressContainer.classList.remove('hidden');
    progressTitle.textContent = title;
    progressPercent.textContent = '0%';
    progressFill.style.width = '0%';
    progressStatus.textContent = 'Iniciando...';
    progressLayers.classList.add('hidden');
    progressLayersList.innerHTML = '';
}

function updateProgress(percent, status) {
    progressPercent.textContent = `${percent}%`;
    progressFill.style.width = `${percent}%`;
    if (status) progressStatus.textContent = status;
}

function showLayersList() {
    progressLayers.classList.remove('hidden');
}

function addLayerToList(name) {
    const li = document.createElement('li');
    li.textContent = name;
    progressLayersList.appendChild(li);
    // Auto-scroll
    progressLayersList.scrollTop = progressLayersList.scrollHeight;
}

function showResult(desc) {
    progressContainer.classList.add('hidden');
    resultContainer.classList.remove('hidden');
    resultDesc.textContent = desc;
}

function showError(message) {
    progressContainer.classList.add('hidden');
    errorContainer.classList.remove('hidden');
    errorDesc.textContent = message;
}

function formatSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / 1024 / 1024).toFixed(1) + ' MB';
}

// ==========================================
// Download & New Conversion
// ==========================================
$('#btn-download')?.addEventListener('click', () => {
    if (!resultBlob) return;
    const url = URL.createObjectURL(resultBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = resultFilename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
});

$('#btn-new-conversion')?.addEventListener('click', () => {
    hideAll();
    resultBlob = null;
    resultFilename = '';
    // Reset file inputs
    document.querySelectorAll('.upload-zone').forEach(z => z.classList.remove('hidden'));
    document.querySelectorAll('.file-info').forEach(f => f.classList.add('hidden'));
    document.querySelectorAll('.btn-primary[id^="btn-convert"]').forEach(b => {
        if (b.id !== 'btn-convert-figma') {
            b.disabled = true;
            b.classList.add('disabled');
        }
    });
    selectedFile = null;
});

$('#btn-retry')?.addEventListener('click', () => {
    hideAll();
});
