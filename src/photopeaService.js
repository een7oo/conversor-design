/**
 * Photopea Service
 * 
 * Embeds Photopea in a hidden iframe and uses postMessage API
 * to load .fig files, extract each layer as PNG, and collect metadata.
 * 
 * Flow:
 *   1. Create hidden iframe with Photopea
 *   2. Wait for "done" message (ready)
 *   3. Send .fig file as ArrayBuffer
 *   4. Wait for "done" (file loaded)
 *   5. Run script to get layer tree structure (names, bounds, types)
 *   6. For each layer, export as PNG via saveToOE
 *   7. Return all layer data to caller
 */

let photopeaFrame = null;
let photopeaWindow = null;
let messageQueue = [];
let isReady = false;

function isArrayBuffer(val) {
  if (!val) return false;
  if (Object.prototype.toString.call(val) === '[object ArrayBuffer]') return true;
  // Photopea may send Uint8Array instead of raw ArrayBuffer
  if (ArrayBuffer.isView(val) && !(val instanceof DataView)) return true;
  return false;
}

function toArrayBuffer(val) {
  if (val instanceof ArrayBuffer) return val;
  if (ArrayBuffer.isView(val)) return val.buffer.slice(val.byteOffset, val.byteOffset + val.byteLength);
  return val;
}

/**
 * Initialize Photopea in a hidden iframe.
 * Returns a promise that resolves when Photopea is ready.
 */
export function initPhotopea() {
  return new Promise((resolve, reject) => {
    // Remove existing iframe if any
    if (photopeaFrame) {
      photopeaFrame.remove();
    }

    photopeaFrame = document.createElement('iframe');
    photopeaFrame.id = 'photopea-frame';
    photopeaFrame.style.cssText = 'position:fixed;left:-9999px;top:0;width:1px;height:1px;opacity:0;pointer-events:none;';

    // Configure Photopea: hide UI, no initial file
    const config = {
      environment: {
        customIO: {},
        theme: 0,
        lang: 'pt',
      }
    };

    photopeaFrame.src = `https://www.photopea.com#${encodeURIComponent(JSON.stringify(config))}`;

    document.body.appendChild(photopeaFrame);
    photopeaWindow = photopeaFrame.contentWindow;

    // Listen for messages from Photopea
    const timeout = setTimeout(() => {
      reject(new Error('Photopea não carregou em 30 segundos. Verifique sua conexão.'));
    }, 30000);

    const readyHandler = (e) => {
      if (e.source !== photopeaWindow) return;
      if (e.data === 'done') {
        clearTimeout(timeout);
        isReady = true;
        window.removeEventListener('message', readyHandler);
        resolve();
      }
    };

    window.addEventListener('message', readyHandler);
  });
}

/**
 * Send a file (ArrayBuffer) to Photopea to open.
 * Returns a promise that resolves when the file is loaded.
 */
export function loadFileInPhotopea(arrayBuffer) {
  return new Promise((resolve, reject) => {
    if (!photopeaWindow) {
      reject(new Error('Photopea não inicializado'));
      return;
    }

    const timeout = setTimeout(() => {
      reject(new Error('Timeout ao carregar arquivo no Photopea.'));
    }, 60000);

    const handler = (e) => {
      if (e.source !== photopeaWindow) return;
      if (e.data === 'done') {
        clearTimeout(timeout);
        window.removeEventListener('message', handler);
        resolve();
      }
    };

    window.addEventListener('message', handler);
    photopeaWindow.postMessage(arrayBuffer, '*');
  });
}

/**
 * Execute a script in Photopea and collect all responses.
 * Photopea sends results via echoToOE, then "done" when finished.
 * 
 * @param {string} script - JavaScript code to run in Photopea
 * @returns {Promise<Array>} Array of responses (strings or ArrayBuffers)
 */
export function runScript(script) {
  return new Promise((resolve, reject) => {
    if (!photopeaWindow) {
      reject(new Error('Photopea não inicializado'));
      return;
    }

    const responses = [];
    const timeout = setTimeout(() => {
      window.removeEventListener('message', handler);
      reject(new Error('Timeout ao executar script no Photopea.'));
    }, 300000); // 5 minutes for complex documents

    const handler = (e) => {
      if (e.source !== photopeaWindow) return;

      if (e.data === 'done') {
        clearTimeout(timeout);
        window.removeEventListener('message', handler);
        resolve(responses);
      } else {
        responses.push(e.data);
      }
    };

    window.addEventListener('message', handler);
    photopeaWindow.postMessage(script, '*');
  });
}

/**
 * Get the document structure from Photopea (layer tree).
 * Returns JSON with layer names, bounds, types, visibility.
 */
export async function getDocumentStructure() {
  const script = `
    function getLayerInfo(layer) {
      var info = {
        name: layer.name,
        kind: layer.kind,
        visible: layer.visible,
        opacity: layer.opacity,
        bounds: {
          left: layer.bounds[0].value,
          top: layer.bounds[1].value,
          right: layer.bounds[2].value,
          bottom: layer.bounds[3].value
        },
        isGroup: layer.typename === "LayerSet",
        children: []
      };
      
      if (layer.typename === "LayerSet") {
        for (var i = 0; i < layer.layers.length; i++) {
          info.children.push(getLayerInfo(layer.layers[i]));
        }
      }
      
      return info;
    }
    
    var doc = app.activeDocument;
    var result = {
      name: doc.name,
      width: doc.width.value,
      height: doc.height.value,
      layers: []
    };
    
    for (var i = 0; i < doc.layers.length; i++) {
      result.layers.push(getLayerInfo(doc.layers[i]));
    }
    
    app.echoToOE(JSON.stringify(result));
  `;

  const responses = await runScript(script);
  const jsonStr = responses.find(r => typeof r === 'string' && r.startsWith('{'));
  if (!jsonStr) throw new Error('Não foi possível ler a estrutura do documento.');
  return JSON.parse(jsonStr);
}

/**
 * Export each layer individually as PNG from Photopea.
 * Uses a script that isolates each layer, exports it, then restores visibility.
 * 
 * @param {Function} onLayerExported - Callback(layerIndex, layerName, pngArrayBuffer)
 * @param {Function} onProgress - Callback(done, total)
 * @returns {Array} Array of { name, bounds, png: ArrayBuffer }
 */
export async function exportLayersAsPng(onProgress) {
  // First, get the count and names
  const countScript = `
    function countLayers(parent) {
      var count = 0;
      for (var i = 0; i < parent.layers.length; i++) {
        var layer = parent.layers[i];
        if (layer.typename === "LayerSet") {
          count += countLayers(layer);
        } else {
          count++;
        }
      }
      return count;
    }
    var total = countLayers(app.activeDocument);
    app.echoToOE("LAYER_COUNT:" + total);
  `;

  const countResult = await runScript(countScript);
  const countStr = countResult.find(r => typeof r === 'string' && r.startsWith('LAYER_COUNT:'));
  const totalLayers = parseInt(countStr?.split(':')[1] || '0');

  if (totalLayers === 0) throw new Error('Nenhuma camada encontrada.');

  if (onProgress) onProgress(0, totalLayers);

  // Export each layer one by one
  // Strategy: hide all layers, then show one at a time and export
  const layers = [];

  for (let i = 0; i < totalLayers; i++) {
    // Script to: hide all, show only layer i, get its info, export as PNG
    const exportScript = `
      // Flatten layer index counter
      var flatIndex = 0;
      var targetIndex = ${i};
      var targetLayer = null;
      
      function findLayer(parent) {
        for (var j = 0; j < parent.layers.length; j++) {
          var layer = parent.layers[j];
          if (layer.typename === "LayerSet") {
            findLayer(layer);
          } else {
            if (flatIndex === targetIndex) {
              targetLayer = layer;
            }
            flatIndex++;
          }
        }
      }
      
      findLayer(app.activeDocument);
      
      if (targetLayer) {
        // Send layer info
        var info = {
          index: ${i},
          name: targetLayer.name,
          kind: targetLayer.kind,
          bounds: {
            left: targetLayer.bounds[0].value,
            top: targetLayer.bounds[1].value,
            right: targetLayer.bounds[2].value,
            bottom: targetLayer.bounds[3].value
          },
          opacity: targetLayer.opacity,
          visible: targetLayer.visible
        };
        app.echoToOE("LAYER_INFO:" + JSON.stringify(info));
        
        // Duplicate, isolate, and export
        var origDoc = app.activeDocument;
        var w = targetLayer.bounds[2].value - targetLayer.bounds[0].value;
        var h = targetLayer.bounds[3].value - targetLayer.bounds[1].value;
        
        if (w > 0 && h > 0) {
          // Select and duplicate the layer to a new document
          app.activeDocument.activeLayer = targetLayer;
          targetLayer.duplicate(null, ElementPlacement.PLACEATBEGINNING);
          
          // The duplicated layer is now in a new temp doc
          // Actually, let's use a simpler approach: just export the layer directly
          // by making only this layer visible
          
          // Save all visibility states
          var visStates = [];
          function saveVis(parent) {
            for (var j = 0; j < parent.layers.length; j++) {
              visStates.push(parent.layers[j].visible);
              parent.layers[j].visible = false;
              if (parent.layers[j].typename === "LayerSet") saveVis(parent.layers[j]);
            }
          }
          // Actually this approach is too slow for many layers.
          // Instead, duplicate just this layer to a new document.
          
          origDoc.activeLayer = targetLayer;
          // Duplicate layer to new document
          targetLayer.duplicate();
          // Now we have a new doc with just this layer - but that's not right either
          
          // Simplest reliable approach: use trim + saveToOE on isolated layer
          // Let's use the "copy merged" approach instead:
          // Select the layer bounds, copy, paste into new doc
          
          app.activeDocument = origDoc;
          origDoc.activeLayer = targetLayer;
          
          // Just export the whole doc as PNG with only this layer visible
          // Hide all
          function setAllVis(parent, vis) {
            for (var j = 0; j < parent.layers.length; j++) {
              parent.layers[j].visible = vis;
              if (parent.layers[j].typename === "LayerSet") {
                setAllVis(parent.layers[j], vis);
              }
            }
          }
          
          // Make sure parent groups are visible
          function showParents(layer) {
            var p = layer.parent;
            while (p && p.typename === "LayerSet") {
              p.visible = true;
              p = p.parent;
            }
          }
          
          setAllVis(origDoc, false);
          targetLayer.visible = true;
          showParents(targetLayer);
          
          // Crop to layer bounds and export
          origDoc.crop([
            new UnitValue(info.bounds.left, "px"),
            new UnitValue(info.bounds.top, "px"),
            new UnitValue(info.bounds.right, "px"),
            new UnitValue(info.bounds.bottom, "px")
          ]);
          
          app.activeDocument.saveToOE("png");
          
          // Undo the crop
          origDoc.activeHistoryState = origDoc.historyStates[origDoc.historyStates.length - 2];
          
          // Restore all layers visibility
          setAllVis(origDoc, true);
        } else {
          app.echoToOE("LAYER_SKIP:" + ${i});
        }
      }
    `;

    const result = await runScript(exportScript);

    // Parse results
    let layerInfo = null;
    let pngData = null;

    for (const r of result) {
      if (typeof r === 'string' && r.startsWith('LAYER_INFO:')) {
        layerInfo = JSON.parse(r.replace('LAYER_INFO:', ''));
      } else if (isArrayBuffer(r)) {
        pngData = toArrayBuffer(r);
      }
    }

    if (layerInfo && pngData) {
      layers.push({
        name: layerInfo.name,
        bounds: layerInfo.bounds,
        opacity: layerInfo.opacity,
        kind: layerInfo.kind,
        png: new Uint8Array(pngData),
      });
    }

    if (onProgress) onProgress(i + 1, totalLayers);
  }

  // Restore all layer visibility
  await runScript(`
    function showAll(parent) {
      for (var i = 0; i < parent.layers.length; i++) {
        parent.layers[i].visible = true;
        if (parent.layers[i].typename === "LayerSet") showAll(parent.layers[i]);
      }
    }
    showAll(app.activeDocument);
  `);

  return layers;
}

/**
 * Export the document as PSD from Photopea.
 * Uses a dedicated listener so binary data is captured even if it
 * arrives after the "done" signal (which can happen with large files).
 */
export function exportAsPsd() {
  return new Promise((resolve, reject) => {
    if (!photopeaWindow) {
      reject(new Error('Photopea não inicializado'));
      return;
    }

    let receivedDone = false;
    let psdData = null;

    const timeout = setTimeout(() => {
      window.removeEventListener('message', handler);
      reject(new Error('Timeout ao exportar PSD do Photopea.'));
    }, 120000);

    const tryFinish = () => {
      if (receivedDone && psdData) {
        clearTimeout(timeout);
        window.removeEventListener('message', handler);
        resolve(toArrayBuffer(psdData));
      } else if (receivedDone && !psdData) {
        // "done" arrived before binary — wait up to 10s more
        const fallbackTimer = setTimeout(() => {
          window.removeEventListener('message', handler);
          clearTimeout(timeout);
          reject(new Error('Falha ao exportar PSD do Photopea.'));
        }, 10000);
        const originalHandler = handler;
        window.removeEventListener('message', handler);
        const waitHandler = (e) => {
          if (e.source !== photopeaWindow) return;
          if (isArrayBuffer(e.data)) {
            clearTimeout(fallbackTimer);
            clearTimeout(timeout);
            window.removeEventListener('message', waitHandler);
            resolve(toArrayBuffer(e.data));
          }
        };
        window.addEventListener('message', waitHandler);
      }
    };

    const handler = (e) => {
      if (e.source !== photopeaWindow) return;
      if (isArrayBuffer(e.data)) {
        psdData = e.data;
        if (receivedDone) tryFinish();
      } else if (e.data === 'done') {
        receivedDone = true;
        tryFinish();
      }
    };

    window.addEventListener('message', handler);
    photopeaWindow.postMessage('app.activeDocument.saveToOE("psd");', '*');
  });
}

/**
 * Get layer group structure (for groups with clipping masks, etc.)
 */
export async function getGroupStructure() {
  const script = `
    function getGroups(parent, path) {
      var groups = [];
      for (var i = 0; i < parent.layers.length; i++) {
        var layer = parent.layers[i];
        var layerPath = path ? path + "/" + layer.name : layer.name;
        if (layer.typename === "LayerSet") {
          groups.push({
            name: layer.name,
            path: layerPath,
            bounds: {
              left: layer.bounds[0].value,
              top: layer.bounds[1].value,
              right: layer.bounds[2].value,
              bottom: layer.bounds[3].value
            },
            opacity: layer.opacity,
            childCount: layer.layers.length
          });
          var subGroups = getGroups(layer, layerPath);
          for (var j = 0; j < subGroups.length; j++) {
            groups.push(subGroups[j]);
          }
        }
      }
      return groups;
    }
    
    var result = getGroups(app.activeDocument, "");
    app.echoToOE("GROUPS:" + JSON.stringify(result));
  `;

  const responses = await runScript(script);
  const str = responses.find(r => typeof r === 'string' && r.startsWith('GROUPS:'));
  if (!str) return [];
  return JSON.parse(str.replace('GROUPS:', ''));
}

/**
 * Cleanup: remove the Photopea iframe
 */
export function destroyPhotopea() {
  if (photopeaFrame) {
    photopeaFrame.remove();
    photopeaFrame = null;
    photopeaWindow = null;
    isReady = false;
  }
}
