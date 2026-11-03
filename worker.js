/* Markdoc converter worker.
   Runs the real Microsoft MarkItDown (Python) inside the browser via
   Pyodide (CPython compiled to WebAssembly). The document bytes arrive
   from the page over postMessage and never leave the device. */

const PYODIDE_VERSION = "0.29.4";
const PYODIDE_INDEX = "https://cdn.jsdelivr.net/pyodide/v" + PYODIDE_VERSION + "/full/";

let pyodide = null;
let readyPromise = null;
let sheetPromise = null;

function progress(stage, detail) {
  self.postMessage({ type: "progress", stage: stage, detail: detail || "" });
}

function sleep(ms) {
  return new Promise(function (resolve) { setTimeout(resolve, ms); });
}

/* Load the Pyodide script with retries: a single dropped connection
   should not kill the converter for the whole session. */
async function loadRuntimeScript() {
  let lastError = null;
  for (let attempt = 1; attempt <= 5; attempt++) {
    try {
      importScripts(PYODIDE_INDEX + "pyodide.js");
      return;
    } catch (err) {
      lastError = err;
      if (attempt < 5) {
        progress("runtime", "Koneksi untuk runtime terputus, mencoba lagi (" + attempt + "/5)…");
        await sleep(1500 * attempt);
      }
    }
  }
  throw lastError;
}

async function init() {
  progress("runtime", "Memuat runtime Python (Pyodide)…");
  await loadRuntimeScript();
  pyodide = await loadPyodide({ indexURL: PYODIDE_INDEX });

  progress("packages", "Memuat paket bawaan Pyodide…");
  await pyodide.loadPackage(["micropip", "lxml", "pillow", "cryptography"]);

  progress("install", "Memasang MarkItDown dan dependensinya dari PyPI…");
  await pyodide.runPythonAsync(`
import micropip
await micropip.install([
    "markitdown==0.1.9",
    "beautifulsoup4",
    "soupsieve",
    "markdownify",
    "six",
    "charset-normalizer",
    "defusedxml",
    "requests",
    "urllib3",
    "certifi",
    "idna",
    "mammoth",
    "cobble",
    "python-pptx",
    "XlsxWriter",
    "pdfminer.six",
    "olefile",
], deps=False)
`);

  progress("prepare", "Menyiapkan MarkItDown…");
  await pyodide.runPythonAsync(`
import sys
import types

# --- Stub: magika -------------------------------------------------------
# MarkItDown imports magika (file-type detection backed by onnxruntime,
# which has no WebAssembly build) and calls identify_stream() on every
# conversion. Markdoc always passes the file extension, filename and
# mimetype explicitly via StreamInfo, so detection can safely report
# "unknown": MarkItDown then falls back to the extension/mimetype guess,
# which is exactly the converter selection we want.
_magika = types.ModuleType("magika")

class _MagikaOutput:
    label = "unknown"
    is_text = False
    mime_type = "application/octet-stream"
    extensions = []

class _MagikaPrediction:
    output = _MagikaOutput()

class _MagikaResult:
    status = "error"
    prediction = _MagikaPrediction()

class _Magika:
    def identify_stream(self, stream):
        return _MagikaResult()
    def identify_bytes(self, data):
        return _MagikaResult()
    def identify_path(self, path):
        return _MagikaResult()

_magika.Magika = _Magika
sys.modules["magika"] = _magika

# --- Stub: pdfplumber -----------------------------------------------------
# MarkItDown's PdfConverter imports pdfplumber (which needs the binary
# pypdfium2) and falls back to pdfminer.six whenever pdfplumber raises.
# This stub makes the import succeed and open() raise immediately, so
# every PDF takes MarkItDown's own documented pdfminer fallback path.
_pdfplumber = types.ModuleType("pdfplumber")

def _pdfplumber_open(*args, **kwargs):
    raise RuntimeError("pdfplumber is not available in the browser build")

_pdfplumber.open = _pdfplumber_open
sys.modules["pdfplumber"] = _pdfplumber

# --- Converter ------------------------------------------------------------
import io
from markitdown import MarkItDown
from markitdown._stream_info import StreamInfo

_converter = MarkItDown(enable_plugins=False)

def markdoc_reload_converter():
    # The spreadsheet converters record missing pandas/openpyxl at import
    # time, so after installing them the markitdown modules must be
    # re-imported and the converter rebuilt before .xlsx/.xls will work.
    global _converter
    for _name in [m for m in sys.modules if m == "markitdown" or m.startswith("markitdown.")]:
        del sys.modules[_name]
    import importlib
    _md = importlib.import_module("markitdown")
    _si = importlib.import_module("markitdown._stream_info")
    globals()["MarkItDown"] = _md.MarkItDown
    globals()["StreamInfo"] = _si.StreamInfo
    _converter = _md.MarkItDown(enable_plugins=False)

def markdoc_convert(data, filename, mimetype):
    name = filename or "document"
    ext = ""
    if "." in name:
        ext = "." + name.rsplit(".", 1)[-1].lower()
    info = StreamInfo(
        extension=ext or None,
        filename=name,
        mimetype=mimetype or None,
    )
    result = _converter.convert_stream(io.BytesIO(bytes(data)), stream_info=info)
    return result.markdown
`);

  progress("ready", "Mesin konversi siap.");
}

function ensureReady() {
  if (!readyPromise) {
    readyPromise = init().catch(function (err) {
      readyPromise = null;
      throw err;
    });
  }
  return readyPromise;
}

/* Spreadsheet support (pandas + openpyxl) is heavy, so it loads on
   demand the first time an .xlsx/.xls file is converted. */
function ensureSheets() {
  if (!sheetPromise) {
    sheetPromise = (async function () {
      progress("packages", "Memuat pandas untuk file spreadsheet…");
      await pyodide.loadPackage(["pandas"]);
      progress("install", "Memasang openpyxl untuk file spreadsheet…");
      await pyodide.runPythonAsync(`
import micropip
await micropip.install(["openpyxl", "et-xmlfile", "xlrd"], deps=False)
markdoc_reload_converter()
`);
    })().catch(function (err) {
      sheetPromise = null;
      throw err;
    });
  }
  return sheetPromise;
}

self.onmessage = async function (event) {
  const msg = event.data || {};
  if (msg.type === "warmup") {
    try {
      await ensureReady();
      self.postMessage({ type: "ready" });
    } catch (err) {
      self.postMessage({ type: "fatal", error: String(err && err.message ? err.message : err) });
    }
    return;
  }
  if (msg.type === "convert") {
    const id = msg.id;
    try {
      await ensureReady();
      const lower = String(msg.filename || "").toLowerCase();
      if (lower.endsWith(".xlsx") || lower.endsWith(".xls")) {
        await ensureSheets();
      }
      progress("convert", "Mengonversi " + msg.filename + "…");
      pyodide.globals.set("_md_data", pyodide.toPy(msg.bytes));
      pyodide.globals.set("_md_name", msg.filename);
      pyodide.globals.set("_md_mime", msg.mimetype || "");
      const markdown = await pyodide.runPythonAsync(
        "markdoc_convert(_md_data, _md_name, _md_mime)"
      );
      self.postMessage({ type: "result", id: id, markdown: markdown });
    } catch (err) {
      self.postMessage({
        type: "error",
        id: id,
        error: String(err && err.message ? err.message : err),
      });
    }
  }
};
