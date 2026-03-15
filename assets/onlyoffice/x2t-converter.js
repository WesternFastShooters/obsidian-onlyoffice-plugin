/**
 * X2T Converter - extracted and adapted from onlyoffice-web-local/src/utils/x2t.ts
 * Handles document format conversion via x2t WASM module.
 */

const DOCUMENT_TYPE_MAP = {
  docx: "word", doc: "word", odt: "word", rtf: "word", txt: "word",
  xlsx: "cell", xls: "cell", ods: "cell", csv: "cell",
  pptx: "slide", ppt: "slide", odp: "slide",
};

const WORKING_DIRS = ["/working", "/working/media", "/working/fonts", "/working/themes"];
const INIT_TIMEOUT = 30000;

let x2tModule = null;
let isReady = false;
let initPromise = null;
let scriptLoaded = false;

function loadScript(src) {
  if (scriptLoaded) return Promise.resolve();
  return new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = src;
    script.onload = () => { scriptLoaded = true; resolve(); };
    script.onerror = (e) => reject(new Error("Failed to load x2t script: " + src));
    document.head.appendChild(script);
  });
}

function initialize(scriptPath) {
  if (isReady && x2tModule) return Promise.resolve(x2tModule);
  if (initPromise) return initPromise;

  initPromise = (async () => {
    try {
      await loadScript(scriptPath);
      return new Promise((resolve, reject) => {
        const mod = window.Module;
        if (!mod) { reject(new Error("X2T Module not found after script load")); return; }

        const timer = setTimeout(() => {
          if (!isReady) reject(new Error("X2T init timeout"));
        }, INIT_TIMEOUT);

        mod.onRuntimeInitialized = () => {
          clearTimeout(timer);
          WORKING_DIRS.forEach((dir) => {
            try { mod.FS.mkdir(dir); } catch (_) {}
          });
          x2tModule = mod;
          isReady = true;
          console.log("[x2t] initialized");
          resolve(mod);
        };
      });
    } catch (e) {
      initPromise = null;
      throw e;
    }
  })();
  return initPromise;
}

function sanitizeFileName(input) {
  if (!input || typeof input !== "string") return "file.bin";
  const parts = input.split(".");
  const ext = parts.pop() || "bin";
  let name = parts.join(".")
    .replace(/[\/\?<>\\:\*\|"]/g, "")
    .replace(/[\x00-\x1f\x80-\x9f]/g, "")
    .replace(/^\.+$/, "")
    .replace(/[&'%!"{}[\]]/g, "")
    .trim() || "file";
  return name.slice(0, 200) + "." + ext;
}

function createConversionParams(fromPath, toPath, additionalParams) {
  return `<?xml version="1.0" encoding="utf-8"?>
<TaskQueueDataConvert xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
  <m_sFileFrom>${fromPath}</m_sFileFrom>
  <m_sThemeDir>/working/themes</m_sThemeDir>
  <m_sFileTo>${toPath}</m_sFileTo>
  <m_bIsNoBase64>false</m_bIsNoBase64>
  ${additionalParams || ""}
</TaskQueueDataConvert>`;
}

function executeConversion(paramsPath) {
  const result = x2tModule.ccall("main1", "number", ["string"], [paramsPath]);
  if (result !== 0) throw new Error("Conversion failed with code: " + result);
}

function readMediaFiles() {
  if (!x2tModule) return {};
  const media = {};
  try {
    const files = x2tModule.FS.readdir("/working/media/");
    files.filter((f) => f !== "." && f !== "..").forEach((f) => {
      try {
        const data = x2tModule.FS.readFile("/working/media/" + f, { encoding: "binary" });
        media["media/" + f] = URL.createObjectURL(new Blob([data]));
      } catch (_) {}
    });
  } catch (_) {}
  return media;
}

/**
 * Convert a File object to OnlyOffice internal bin format.
 * @param {File} file
 * @returns {Promise<{fileName: string, type: string, bin: Uint8Array, media: Object}>}
 */
async function convertDocument(file) {
  await initialize("./wasm/x2t/x2t.js");
  const fileName = file.name;
  const ext = fileName.split(".").pop().toLowerCase();
  const docType = DOCUMENT_TYPE_MAP[ext];
  if (!docType) throw new Error("Unsupported format: " + ext);

  const arrayBuffer = await file.arrayBuffer();
  const data = new Uint8Array(arrayBuffer);
  const safeName = sanitizeFileName(fileName);
  const inputPath = "/working/" + safeName;
  const outputPath = inputPath + ".bin";

  x2tModule.FS.writeFile(inputPath, data);
  const params = createConversionParams(inputPath, outputPath);
  x2tModule.FS.writeFile("/working/params.xml", params);
  executeConversion("/working/params.xml");

  const bin = x2tModule.FS.readFile(outputPath);
  const media = readMediaFiles();

  return { fileName: safeName, type: docType, bin, media };
}

/**
 * Convert bin data back to a document format. Returns Uint8Array, does NOT download.
 * @param {Uint8Array} bin
 * @param {string} originalFileName
 * @param {string} targetExt - e.g. "DOCX", "XLSX", "PDF"
 * @returns {Promise<{fileName: string, data: Uint8Array}>}
 */
async function convertBinToDocument(bin, originalFileName, targetExt) {
  targetExt = targetExt || "DOCX";
  await initialize("./wasm/x2t/x2t.js");

  const sanitizedBase = sanitizeFileName(originalFileName).replace(/\.[^/.]+$/, "");
  const binFileName = sanitizedBase + ".bin";
  const outputFileName = sanitizedBase + "." + targetExt.toLowerCase();

  x2tModule.FS.writeFile("/working/" + binFileName, bin);

  let additionalParams = "";
  if (targetExt.toUpperCase() === "PDF") {
    additionalParams = "<m_sFontDir>/working/fonts/</m_sFontDir>";
  }

  const params = createConversionParams(
    "/working/" + binFileName,
    "/working/" + outputFileName,
    additionalParams,
  );
  x2tModule.FS.writeFile("/working/params.xml", params);
  executeConversion("/working/params.xml");

  const result = x2tModule.FS.readFile("/working/" + outputFileName);
  return { fileName: outputFileName, data: result };
}

// File type code to name mapping (from OnlyOffice)
const FILE_TYPE_MAP = {
  65: "DOCX", 66: "DOC", 67: "ODT", 68: "RTF", 69: "TXT",
  257: "XLSX", 258: "XLS", 259: "ODS", 260: "CSV",
  129: "PPTX", 130: "PPT", 131: "ODP",
  513: "PDF",
};

function getFileTypeNameByCode(code) {
  return FILE_TYPE_MAP[code] || "DOCX";
}

// Expose to global scope for editor.js
window.X2TConverter = {
  initialize,
  convertDocument,
  convertBinToDocument,
  getFileTypeNameByCode,
};
