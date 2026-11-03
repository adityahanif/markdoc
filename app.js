/* Markdoc app: file handling, worker management, output rendering.
   Conversion itself happens in worker.js (Pyodide + MarkItDown). */
(function () {
  "use strict";

  var SUPPORTED = {
    pdf: "application/pdf",
    docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    xls: "application/vnd.ms-excel",
    html: "text/html",
    htm: "text/html",
    epub: "application/epub+zip",
    csv: "text/csv",
    txt: "text/plain",
    md: "text/markdown",
    markdown: "text/markdown",
    json: "application/json",
    xml: "application/xml",
    ipynb: "application/x-ipynb+json"
  };
  var MAX_BYTES = 50 * 1024 * 1024;

  var dropzone = document.getElementById("dropzone");
  var fileInput = document.getElementById("file-input");
  var fileList = document.getElementById("file-list");
  var convertBtn = document.getElementById("convert-btn");
  var clearBtn = document.getElementById("clear-btn");
  var statusEl = document.getElementById("workbench-status");
  var engineEl = document.getElementById("engine-status");
  var outputEl = document.getElementById("output");
  var outputSelect = document.getElementById("output-select");
  var outputCode = document.getElementById("output-code");
  var outputRaw = document.getElementById("output-raw");
  var outputPreview = document.getElementById("output-preview");
  var outputBody = outputRaw.parentElement;
  var outputMeta = document.getElementById("output-meta");
  var copyBtn = document.getElementById("copy-btn");
  var downloadBtn = document.getElementById("download-btn");
  var viewRawBtn = document.getElementById("view-raw");
  var viewPreviewBtn = document.getElementById("view-preview");

  var entries = [];
  var nextId = 1;
  var busy = false;
  var viewMode = "raw";
  var selectedId = null;

  /* ---------------- Worker management ---------------- */
  var worker = null;
  var workerReady = false;
  var workerFailed = false;
  var seq = 0;
  var pending = new Map();

  function setEngine(text) { engineEl.textContent = text; }
  function setStatus(text, isError) {
    statusEl.textContent = text;
    statusEl.classList.toggle("is-error", !!isError);
  }

  function createWorker() {
    workerReady = false;
    workerFailed = false;
    worker = new Worker("./worker.js?v=1");
    worker.onmessage = function (event) {
      var m = event.data || {};
      if (m.type === "progress") {
        setEngine(m.detail || m.stage);
        if (busy) setStatus(m.detail || m.stage, false);
      } else if (m.type === "ready") {
        workerReady = true;
        setEngine("siap (MarkItDown 0.1.9 di Pyodide)");
      } else if (m.type === "fatal") {
        workerFailed = true;
        setEngine("gagal dimuat. Periksa koneksi lalu coba tambahkan file lagi.");
        setStatus("Mesin gagal dimuat: " + m.error, true);
      } else if (m.type === "result") {
        workerReady = true;
        var p = pending.get(m.id);
        if (p) { pending.delete(m.id); p.resolve(m.markdown); }
      } else if (m.type === "error") {
        var q = pending.get(m.id);
        if (q) { pending.delete(m.id); q.reject(new Error(m.error)); }
      }
    };
    worker.onerror = function () {
      workerFailed = true;
      setEngine("worker bermasalah. Muat ulang halaman untuk mencoba lagi.");
    };
    return worker;
  }

  function ensureWorker() {
    if (!worker || workerFailed) createWorker();
    return worker;
  }

  function warmup() {
    var w = ensureWorker();
    if (!workerReady && !workerFailed) {
      setEngine("sedang disiapkan…");
      w.postMessage({ type: "warmup" });
    }
  }

  function convertWithWorker(entry) {
    var w = ensureWorker();
    return entry.file.arrayBuffer().then(function (buf) {
      return new Promise(function (resolve, reject) {
        var id = ++seq;
        pending.set(id, { resolve: resolve, reject: reject });
        w.postMessage({
          type: "convert",
          id: id,
          filename: entry.file.name,
          mimetype: entry.file.type || SUPPORTED[entry.ext] || "",
          bytes: new Uint8Array(buf)
        });
      });
    });
  }

  /* ---------------- File list ---------------- */
  function extOf(name) {
    var i = name.lastIndexOf(".");
    return i >= 0 ? name.slice(i + 1).toLowerCase() : "";
  }
  function formatSize(bytes) {
    if (bytes < 1024) return bytes + " B";
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1).replace(".", ",") + " KB";
    return (bytes / (1024 * 1024)).toFixed(1).replace(".", ",") + " MB";
  }
  var STATUS_TEXT = {
    queued: "menunggu",
    converting: "mengonversi…",
    done: "selesai",
    error: "gagal"
  };

  function renderList() {
    fileList.innerHTML = "";
    entries.forEach(function (entry) {
      var li = document.createElement("li");
      li.className = "file-row";
      var name = document.createElement("span");
      name.className = "file-name";
      name.textContent = entry.name;
      var size = document.createElement("span");
      size.className = "file-size";
      size.textContent = formatSize(entry.file.size);
      var st = document.createElement("span");
      st.className = "file-status" + (entry.status === "done" ? " is-done" : entry.status === "error" ? " is-error" : "");
      st.textContent = entry.status === "error" ? ("gagal: " + entry.error) : STATUS_TEXT[entry.status];
      st.title = entry.status === "error" ? entry.error : "";
      var rm = document.createElement("button");
      rm.type = "button";
      rm.className = "file-remove";
      rm.textContent = "hapus";
      rm.disabled = busy;
      rm.addEventListener("click", function () {
        entries = entries.filter(function (e) { return e.id !== entry.id; });
        if (selectedId === entry.id) selectedId = null;
        renderList(); renderOutput(); refreshButtons();
      });
      li.appendChild(name); li.appendChild(size); li.appendChild(st); li.appendChild(rm);
      fileList.appendChild(li);
    });
  }

  function refreshButtons() {
    var hasQueued = entries.some(function (e) { return e.status === "queued" || e.status === "error"; });
    convertBtn.disabled = busy || !hasQueued;
    clearBtn.disabled = busy || entries.length === 0;
  }

  function addFiles(fileArray) {
    if (busy) { setStatus("Tunggu konversi yang berjalan selesai dulu.", false); return; }
    var added = 0;
    Array.prototype.forEach.call(fileArray, function (file) {
      var ext = extOf(file.name);
      var entry = {
        id: nextId++, file: file, name: file.name, ext: ext,
        status: "queued", markdown: null, error: ""
      };
      if (!SUPPORTED.hasOwnProperty(ext)) {
        entry.status = "error";
        entry.error = "format ." + (ext || "?") + " tidak didukung";
      } else if (file.size > MAX_BYTES) {
        entry.status = "error";
        entry.error = "file lebih besar dari 50 MB";
      } else if (file.size === 0) {
        entry.status = "error";
        entry.error = "file kosong";
      } else {
        added++;
      }
      entries.push(entry);
    });
    renderList(); refreshButtons();
    if (added > 0) {
      setStatus(added + " file siap dikonversi. Mesin sedang disiapkan di latar belakang.", false);
      warmup();
    } else {
      setStatus("Tidak ada file valid yang ditambahkan. Cek format dan ukuran file.", true);
    }
  }

  dropzone.addEventListener("click", function () { fileInput.click(); });
  dropzone.addEventListener("keydown", function (e) {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); fileInput.click(); }
  });
  fileInput.addEventListener("change", function () {
    addFiles(fileInput.files);
    fileInput.value = "";
  });
  ["dragenter", "dragover"].forEach(function (t) {
    dropzone.addEventListener(t, function (e) { e.preventDefault(); dropzone.classList.add("is-over"); });
  });
  ["dragleave", "drop"].forEach(function (t) {
    dropzone.addEventListener(t, function (e) { e.preventDefault(); dropzone.classList.remove("is-over"); });
  });
  dropzone.addEventListener("drop", function (e) {
    if (e.dataTransfer && e.dataTransfer.files.length) addFiles(e.dataTransfer.files);
  });

  clearBtn.addEventListener("click", function () {
    entries = []; selectedId = null;
    renderList(); renderOutput(); refreshButtons();
    setStatus("Daftar file dibersihkan.", false);
  });

  /* ---------------- Conversion ---------------- */
  convertBtn.addEventListener("click", function () {
    var queue = entries.filter(function (e) { return e.status === "queued" || e.status === "error"; });
    if (!queue.length || busy) return;
    busy = true; refreshButtons(); renderList();
    var chain = Promise.resolve();
    queue.forEach(function (entry) {
      chain = chain.then(function () {
        entry.status = "converting"; entry.error = "";
        renderList();
        setStatus("Mengonversi " + entry.name + "…", false);
        return convertWithWorker(entry).then(function (markdown) {
          entry.status = "done";
          entry.markdown = markdown;
          selectedId = entry.id;
        }).catch(function (err) {
          entry.status = "error";
          entry.error = shortenError(err);
        }).then(function () { renderList(); renderOutput(); });
      });
    });
    chain.then(function () {
      busy = false; refreshButtons(); renderList(); renderOutput();
      if (workerReady) setEngine("siap (MarkItDown 0.1.9 di Pyodide)");
      var done = entries.filter(function (e) { return e.status === "done"; }).length;
      var failed = entries.filter(function (e) { return e.status === "error"; }).length;
      if (failed === 0) setStatus("Selesai. " + done + " file berhasil dikonversi.", false);
      else setStatus("Selesai dengan catatan: " + done + " berhasil, " + failed + " gagal. File yang gagal bisa dicoba lagi dengan tombol Convert.", failed > 0 && done === 0);
    });
  });

  function shortenError(err) {
    var msg = String(err && err.message ? err.message : err);
    if (/MissingDependencyException/.test(msg)) return "dependensi format ini tidak tersedia di build browser";
    if (/UnsupportedFormatException/.test(msg)) return "format tidak didukung MarkItDown";
    if (/FileConversionException/.test(msg)) return "MarkItDown gagal membaca file ini";
    if (msg.length > 110) msg = msg.slice(0, 110) + "…";
    return msg;
  }

  /* ---------------- Output ---------------- */
  function doneEntries() { return entries.filter(function (e) { return e.status === "done"; }); }

  function renderOutput() {
    var done = doneEntries();
    if (!done.length) { outputEl.hidden = true; return; }
    outputEl.hidden = false;
    if (!done.some(function (e) { return e.id === selectedId; })) selectedId = done[done.length - 1].id;
    outputSelect.innerHTML = "";
    done.forEach(function (e) {
      var opt = document.createElement("option");
      opt.value = String(e.id);
      opt.textContent = e.name.replace(/\.[^.]+$/, "") + ".md";
      if (e.id === selectedId) opt.selected = true;
      outputSelect.appendChild(opt);
    });
    showSelected();
  }

  function selectedEntry() {
    var done = doneEntries();
    for (var i = 0; i < done.length; i++) if (done[i].id === selectedId) return done[i];
    return done[done.length - 1] || null;
  }

  function showSelected() {
    var entry = selectedEntry();
    if (!entry) return;
    outputCode.textContent = entry.markdown;
    outputPreview.innerHTML = renderMarkdown(entry.markdown);
    var chars = entry.markdown.length;
    var words = entry.markdown.trim() ? entry.markdown.trim().split(/\s+/).length : 0;
    outputMeta.textContent = chars.toLocaleString("id-ID") + " karakter, " + words.toLocaleString("id-ID") + " kata";
    applyViewMode();
  }

  outputSelect.addEventListener("change", function () {
    selectedId = parseInt(outputSelect.value, 10);
    showSelected();
  });

  function applyViewMode() {
    var raw = viewMode === "raw";
    outputRaw.hidden = !raw;
    outputPreview.hidden = raw;
    outputBody.classList.toggle("is-preview", !raw);
    viewRawBtn.classList.toggle("is-active", raw);
    viewPreviewBtn.classList.toggle("is-active", !raw);
  }
  viewRawBtn.addEventListener("click", function () { viewMode = "raw"; applyViewMode(); });
  viewPreviewBtn.addEventListener("click", function () { viewMode = "preview"; applyViewMode(); });

  copyBtn.addEventListener("click", function () {
    var entry = selectedEntry();
    if (!entry) return;
    function ok() { copyBtn.textContent = "Tersalin"; setTimeout(function () { copyBtn.textContent = "Salin Markdown"; }, 1600); }
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(entry.markdown).then(ok, function () { fallbackCopy(entry.markdown); ok(); });
    } else { fallbackCopy(entry.markdown); ok(); }
  });
  function fallbackCopy(text) {
    var ta = document.createElement("textarea");
    ta.value = text; ta.style.position = "fixed"; ta.style.opacity = "0";
    document.body.appendChild(ta); ta.select();
    try { document.execCommand("copy"); } catch (e) {}
    document.body.removeChild(ta);
  }

  downloadBtn.addEventListener("click", function () {
    var entry = selectedEntry();
    if (!entry) return;
    var blob = new Blob([entry.markdown], { type: "text/markdown;charset=utf-8" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = entry.name.replace(/\.[^.]+$/, "") + ".md";
    document.body.appendChild(a); a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 4000);
  });

  /* ---------------- Minimal Markdown preview ---------------- */
  function escapeHtml(s) {
    return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }
  function inline(text) {
    var spans = [];
    text = text.replace(/`([^`]+)`/g, function (_, code) {
      spans.push("<code>" + code + "</code>");
      return "\u0000" + (spans.length - 1) + "\u0000";
    });
    text = text.replace(/!\[([^\]]*)\]\(([^)\s]+)(?:\s+"[^"]*")?\)/g, function (_, alt, src) {
      if (/^(data:image\/|https?:\/\/)/i.test(src)) return '<img src="' + src + '" alt="' + alt + '" loading="lazy">';
      return alt ? "[" + alt + "]" : "";
    });
    text = text.replace(/\[([^\]]+)\]\(([^)\s]+)(?:\s+"[^"]*")?\)/g, function (_, label, href) {
      if (/^(https?:\/\/|mailto:|#)/i.test(href)) return '<a href="' + href + '">' + label + "</a>";
      return label;
    });
    text = text.replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>");
    text = text.replace(/__([^_]+)__/g, "<strong>$1</strong>");
    text = text.replace(/(^|[^*])\*([^*\n]+)\*/g, "$1<em>$2</em>");
    text = text.replace(/~~([^~]+)~~/g, "<del>$1</del>");
    text = text.replace(/\u0000(\d+)\u0000/g, function (_, i) { return spans[Number(i)]; });
    return text;
  }
  function renderMarkdown(md) {
    var lines = escapeHtml(md).split("\n");
    var out = [];
    var i = 0;
    function isTableSep(line) { return /^\s*\|?[\s:|-]+\|?\s*$/.test(line) && line.indexOf("-") >= 0; }
    while (i < lines.length) {
      var line = lines[i];
      if (/^```/.test(line)) {
        var buf = [];
        i++;
        while (i < lines.length && !/^```/.test(lines[i])) { buf.push(lines[i]); i++; }
        i++;
        out.push("<pre><code>" + buf.join("\n") + "</code></pre>");
        continue;
      }
      if (!line.trim()) { i++; continue; }
      var h = line.match(/^(#{1,6})\s+(.*)$/);
      if (h) { out.push("<h" + h[1].length + ">" + inline(h[2]) + "</h" + h[1].length + ">"); i++; continue; }
      if (/^\s*(-{3,}|\*{3,}|_{3,})\s*$/.test(line)) { out.push("<hr>"); i++; continue; }
      if (line.indexOf("|") >= 0 && i + 1 < lines.length && isTableSep(lines[i + 1])) {
        var head = line.split("|").map(function (c) { return c.trim(); }).filter(function (c, idx, a) { return !(idx === 0 && c === "") && !(idx === a.length - 1 && c === ""); });
        i += 2;
        var rows = [];
        while (i < lines.length && lines[i].indexOf("|") >= 0 && lines[i].trim()) {
          rows.push(lines[i].split("|").map(function (c) { return c.trim(); }).filter(function (c, idx, a) { return !(idx === 0 && c === "") && !(idx === a.length - 1 && c === ""); }));
          i++;
        }
        var t = "<table><thead><tr>" + head.map(function (c) { return "<th>" + inline(c) + "</th>"; }).join("") + "</tr></thead><tbody>";
        rows.forEach(function (r) { t += "<tr>" + r.map(function (c) { return "<td>" + inline(c) + "</td>"; }).join("") + "</tr>"; });
        out.push(t + "</tbody></table>");
        continue;
      }
      if (/^\s*&gt;\s?/.test(line)) {
        var quote = [];
        while (i < lines.length && /^\s*&gt;\s?/.test(lines[i])) { quote.push(lines[i].replace(/^\s*&gt;\s?/, "")); i++; }
        out.push("<blockquote>" + quote.map(function (q) { return "<p>" + inline(q) + "</p>"; }).join("") + "</blockquote>");
        continue;
      }
      if (/^\s*([-*+]|\d+[.)])\s+/.test(line)) {
        var ordered = /^\s*\d+[.)]\s+/.test(line);
        var items = [];
        while (i < lines.length && /^\s*([-*+]|\d+[.)])\s+/.test(lines[i])) {
          items.push(lines[i].replace(/^\s*([-*+]|\d+[.)])\s+/, "")); i++;
        }
        var tag = ordered ? "ol" : "ul";
        out.push("<" + tag + ">" + items.map(function (it) { return "<li>" + inline(it) + "</li>"; }).join("") + "</" + tag + ">");
        continue;
      }
      var para = [line];
      i++;
      while (i < lines.length && lines[i].trim() && !/^(#{1,6})\s/.test(lines[i]) && !/^```/.test(lines[i]) && !/^\s*([-*+]|\d+[.)])\s+/.test(lines[i]) && !/^\s*&gt;\s?/.test(lines[i])) {
        para.push(lines[i]); i++;
      }
      out.push("<p>" + inline(para.join(" ")) + "</p>");
    }
    return out.join("\n");
  }

  /* ---------------- Cobalt ASCII field (static) ---------------- */
  (function drawField() {
    var canvas = document.getElementById("field");
    if (!canvas || !canvas.getContext) return;
    var styles = getComputedStyle(document.documentElement);
    var colors = [
      styles.getPropertyValue("--color-glyph-1").trim(),
      styles.getPropertyValue("--color-glyph-2").trim(),
      styles.getPropertyValue("--color-glyph-3").trim(),
      styles.getPropertyValue("--color-glyph-4").trim(),
      styles.getPropertyValue("--color-accent-deep").trim()
    ];
    var glyphs = " .:-=+*#%@";
    var syntax = "#>-*_`";
    var seed = 85;
    function rnd() { seed = (seed * 16807) % 2147483647; return seed / 2147483647; }
    function draw() {
      var hero = canvas.parentElement;
      var w = hero.clientWidth, hgt = hero.clientHeight;
      if (!w || !hgt) return;
      canvas.width = w; canvas.height = hgt;
      var ctx = canvas.getContext("2d");
      ctx.clearRect(0, 0, w, hgt);
      ctx.font = '13px "JetBrains Mono", ui-monospace, monospace';
      ctx.textBaseline = "top";
      ctx.globalAlpha = 0.6;
      seed = 85;
      var cols = Math.ceil(w / 8), rows = Math.ceil(hgt / 16);
      for (var y = 0; y < rows; y++) {
        for (var x = 0; x < cols; x++) {
          /* Sparse on the left where the copy sits, denser toward the
             right margin; most of the field hides behind the workbench. */
          var rise = 0.15 + 0.85 * (x / cols);
          var density = Math.pow(rnd(), 3.2) * rise;
          if (density < 0.30) continue;
          var idx = Math.min(glyphs.length - 1, Math.floor(density * glyphs.length * 1.6));
          var ch = rnd() < 0.10 ? syntax[Math.floor(rnd() * syntax.length)] : glyphs[idx];
          var ci = Math.min(colors.length - 2, Math.floor(density * colors.length * 1.5));
          ctx.fillStyle = colors[ci];
          ctx.fillText(ch, x * 8, y * 16);
        }
      }
      ctx.globalAlpha = 1;
    }
    var t;
    window.addEventListener("resize", function () { clearTimeout(t); t = setTimeout(draw, 180); });
    if (document.fonts && document.fonts.ready) document.fonts.ready.then(draw);
    draw();
  })();

  refreshButtons();
})();
