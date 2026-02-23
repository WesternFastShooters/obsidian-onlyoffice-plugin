/**
 * editor.js - Outer iframe editor logic.
 * Loaded inside editor.html. Communicates with Obsidian plugin via postMessage.
 */

(function () {
  "use strict";

  const EXT_TO_TYPE = {
    docx: "word", doc: "word",
    xlsx: "cell", xls: "cell",
    pptx: "slide", ppt: "slide",
  };

  const FILE_TYPE_CODE_MAP = {
    65: "DOCX", 66: "DOC", 67: "ODT", 68: "RTF", 69: "TXT",
    257: "XLSX", 258: "XLS", 259: "ODS", 260: "CSV",
    129: "PPTX", 130: "PPT", 131: "ODP",
    513: "PDF",
  };

  let editorInstance = null;
  let currentFileName = "";
  let currentFileExt = "";
  let mediaMap = {};
  let apiLoaded = false;

  function loadEditorApi() {
    if (apiLoaded && window.DocsAPI) return Promise.resolve();
    return new Promise(function (resolve, reject) {
      if (window.DocsAPI) { apiLoaded = true; resolve(); return; }
      var s = document.createElement("script");
      s.src = "./web-apps/apps/api/documents/api.js";
      s.onload = function () { apiLoaded = true; resolve(); };
      s.onerror = function () { reject(new Error("Failed to load api.js")); };
      document.head.appendChild(s);
    });
  }

  function hideLoading() {
    var el = document.getElementById("loading");
    if (el) el.classList.add("hidden");
  }

  function sendToPlugin(msg, transfer) {
    window.parent.postMessage(msg, "*", transfer || []);
  }

  function getExtFromName(name) {
    var parts = (name || "").split(".");
    return parts.length > 1 ? parts.pop().toLowerCase() : "";
  }

  /**
   * Open a document from ArrayBuffer data.
   */
  async function openDocument(requestId, fileName, fileData, readonly) {
    try {
      currentFileName = fileName;
      currentFileExt = getExtFromName(fileName);

      var file = new File([fileData], fileName);
      var result = await window.X2TConverter.convertDocument(file);

      await loadEditorApi();

      if (editorInstance) {
        try { editorInstance.destroyEditor(); } catch (_) {}
        editorInstance = null;
      }

      var rootEl = document.getElementById("editor-root");
      rootEl.innerHTML = '<div id="editor-frame"></div>';

      editorInstance = new window.DocsAPI.DocEditor("editor-frame", {
        document: {
          title: fileName,
          url: fileName,
          fileType: currentFileExt,
          permissions: {
            edit: !readonly,
            chat: false,
            protect: false,
          },
        },
        editorConfig: {
          lang: "zh",
          customization: {
            help: false,
            about: false,
            hideRightMenu: true,
            features: {
              spellcheck: { change: false },
            },
            anonymous: {
              request: false,
              label: "Guest",
            },
          },
        },
        events: {
          onAppReady: function () {
            if (result.media && Object.keys(result.media).length > 0) {
              mediaMap = result.media;
              editorInstance.sendCommand({
                command: "asc_setImageUrls",
                data: { urls: mediaMap },
              });
            }

            editorInstance.sendCommand({
              command: "asc_openDocument",
              data: { buf: result.bin },
            });
          },
          onDocumentReady: function () {
            hideLoading();
            sendToPlugin({ type: "oo:opened", requestId: requestId });
          },
          onSave: function (event) {
            handleSave(event);
          },
          writeFile: function (event) {
            handleWriteFile(event);
          },
        },
      });
    } catch (err) {
      console.error("[editor] openDocument failed:", err);
      sendToPlugin({
        type: "oo:error",
        requestId: requestId,
        payload: { code: "OPEN_FAILED", message: err.message || String(err) },
      });
    }
  }

  /**
   * Handle the editor's internal save event (Ctrl+S in editor).
   */
  async function handleSave(event) {
    try {
      if (!event || !event.data || !event.data.data) {
        console.warn("[editor] onSave: no data in event");
        notifySaveCallback();
        return;
      }

      var data = event.data.data;
      var option = event.data.option || {};
      var outputformat = option.outputformat;
      var targetExt = FILE_TYPE_CODE_MAP[outputformat] || currentFileExt.toUpperCase() || "DOCX";

      var binData;
      if (data.data instanceof Uint8Array) {
        binData = data.data;
      } else if (data instanceof Uint8Array) {
        binData = data;
      } else if (typeof data === "object" && data.data) {
        binData = new Uint8Array(data.data);
      } else {
        console.warn("[editor] onSave: unexpected data structure", typeof data);
        notifySaveCallback();
        return;
      }

      var result = await window.X2TConverter.convertBinToDocument(
        binData,
        currentFileName,
        targetExt,
      );

      var arrayBuffer = result.data.buffer.slice(
        result.data.byteOffset,
        result.data.byteOffset + result.data.byteLength,
      );

      sendToPlugin(
        {
          type: "oo:saved",
          requestId: "",
          payload: { fileData: arrayBuffer, fileName: currentFileName },
        },
        [arrayBuffer],
      );

      notifySaveCallback();
    } catch (err) {
      console.error("[editor] handleSave failed:", err);
      notifySaveCallback();
    }
  }

  function notifySaveCallback() {
    if (editorInstance) {
      editorInstance.sendCommand({
        command: "asc_onSaveCallback",
        data: { err_code: 0 },
      });
    }
  }

  /**
   * Handle pasted images (writeFile event from editor).
   */
  function handleWriteFile(event) {
    try {
      if (!event || !event.data) return;
      var eventData = event.data;
      var imageData = eventData.data;
      var fileName = eventData.file;

      if (!(imageData instanceof Uint8Array)) return;
      if (!fileName) return;

      var ext = getExtFromName(fileName) || "png";
      var mimeMap = {
        png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg",
        gif: "image/gif", bmp: "image/bmp", webp: "image/webp",
        svg: "image/svg+xml",
      };
      var blob = new Blob([imageData], { type: mimeMap[ext] || "image/png" });
      var objectUrl = URL.createObjectURL(blob);

      mediaMap["media/" + fileName] = objectUrl;
      editorInstance.sendCommand({
        command: "asc_setImageUrls",
        data: { urls: mediaMap },
      });
      editorInstance.sendCommand({
        command: "asc_writeFileCallback",
        data: { path: objectUrl, imgName: fileName },
      });
    } catch (err) {
      console.error("[editor] handleWriteFile failed:", err);
    }
  }

  /**
   * Handle export request from plugin.
   */
  async function handleExport(requestId, targetExt) {
    try {
      // TODO: Need a way to get current bin from editor, for now this is a placeholder.
      sendToPlugin({
        type: "oo:error",
        requestId: requestId,
        payload: { code: "NOT_IMPLEMENTED", message: "Export not yet implemented" },
      });
    } catch (err) {
      sendToPlugin({
        type: "oo:error",
        requestId: requestId,
        payload: { code: "EXPORT_FAILED", message: err.message || String(err) },
      });
    }
  }

  // Listen for messages from the Obsidian plugin.
  window.addEventListener("message", function (e) {
    var msg = e.data;
    if (!msg || !msg.type || !msg.type.startsWith("oo:")) return;

    switch (msg.type) {
      case "oo:open":
        openDocument(
          msg.requestId,
          msg.payload.fileName,
          msg.payload.fileData,
          msg.payload.readonly,
        );
        break;

      case "oo:save":
        // Programmatic save trigger from plugin - not yet implemented
        break;

      case "oo:export":
        handleExport(msg.requestId, msg.payload.targetExt);
        break;
    }
  });

  // Notify parent that the iframe is ready.
  sendToPlugin({ type: "oo:ready" });
})();
