/* Markdoc app: queue UI plus the JavaScript conversion engine.
   Every library loads lazily from ./vendor/ on this same origin,
   the first time a file of that kind is converted. */
(function () {
  'use strict';

  /* ------------------------------------------------ engine */
  var VENDOR = 'vendor/';
  var scriptCache = {};
  function loadScript(name) {
    if (!scriptCache[name]) {
      scriptCache[name] = new Promise(function (resolve, reject) {
        var s = document.createElement('script');
        s.src = VENDOR + name;
        s.onload = resolve;
        s.onerror = function () { reject(new Error('Could not load ' + name)); };
        document.head.appendChild(s);
      });
    }
    return scriptCache[name];
  }

  var tdPromise = null;
  function getTurndown() {
    if (!tdPromise) {
      tdPromise = Promise.all([loadScript('turndown.js'), loadScript('turndown-plugin-gfm.js')]).then(function () {
        var td = new TurndownService({ headingStyle: 'atx', codeBlockStyle: 'fenced', bulletListMarker: '-' });
        td.use(turndownPluginGfm.gfm);
        return td;
      });
    }
    return tdPromise;
  }

  function escCell(s) {
    return String(s == null ? '' : s).replace(/\|/g, '\\|').replace(/\s*\n\s*/g, ' ');
  }
  function rowsToTable(rows) {
    if (!rows.length) return '';
    var width = 0;
    rows.forEach(function (r) { width = Math.max(width, r.length); });
    var norm = rows.map(function (r) {
      var c = r.slice();
      while (c.length < width) c.push('');
      return c;
    });
    var out = ['| ' + norm[0].map(escCell).join(' | ') + ' |',
               '| ' + norm[0].map(function () { return '---'; }).join(' | ') + ' |'];
    for (var i = 1; i < norm.length; i++) out.push('| ' + norm[i].map(escCell).join(' | ') + ' |');
    return out.join('\n');
  }
  function parseCsv(text) {
    var rows = [], row = [], cell = '', inQ = false;
    for (var i = 0; i < text.length; i++) {
      var c = text[i];
      if (inQ) {
        if (c === '"') {
          if (text[i + 1] === '"') { cell += '"'; i++; } else inQ = false;
        } else cell += c;
      } else if (c === '"') inQ = true;
      else if (c === ',') { row.push(cell); cell = ''; }
      else if (c === '\n' || c === '\r') {
        if (c === '\r' && text[i + 1] === '\n') i++;
        row.push(cell); cell = '';
        if (row.length > 1 || row[0] !== '') rows.push(row);
        row = [];
      } else cell += c;
    }
    row.push(cell);
    if (row.length > 1 || row[0] !== '') rows.push(row);
    return rows;
  }
  function byTag(root, local) {
    return Array.prototype.slice.call(root.getElementsByTagNameNS('*', local));
  }
  function parseXml(text) {
    return new DOMParser().parseFromString(text, 'text/xml');
  }

  /* --- DOCX: mammoth to HTML, then Turndown. Tables get a real
         header row first, because mammoth emits td-only tables. */
  function fixDocxTables(html) {
    var doc = new DOMParser().parseFromString(html, 'text/html');
    Array.prototype.slice.call(doc.querySelectorAll('table')).forEach(function (t) {
      var firstRow = t.querySelector('tr');
      if (firstRow) {
        firstRow.querySelectorAll('td').forEach(function (td) {
          var th = doc.createElement('th');
          th.innerHTML = td.innerHTML;
          td.replaceWith(th);
        });
        var thead = doc.createElement('thead');
        t.insertBefore(thead, t.firstChild);
        thead.appendChild(firstRow);
      }
      t.querySelectorAll('td p, th p').forEach(function (p) {
        p.replaceWith.apply(p, Array.prototype.slice.call(p.childNodes));
      });
    });
    return doc.body.innerHTML;
  }
  function convertDocx(file) {
    return Promise.all([loadScript('mammoth.browser.min.js'), getTurndown()]).then(function (vals) {
      var td = vals[1];
      return file.arrayBuffer().then(function (buf) {
        return mammoth.convertToHtml({ arrayBuffer: buf }, {
          convertImage: mammoth.images.imgElement(function () {
            return Promise.resolve({ src: '' });
          })
        });
      }).then(function (result) {
        var cleaned = result.value.replace(/<img[^>]*src=""[^>]*>/g, '');
        return td.turndown(fixDocxTables(cleaned)).trim();
      });
    });
  }

  /* --- XLSX / XLS: SheetJS straight to Markdown tables. */
  function convertSheet(file) {
    return loadScript('xlsx.full.min.js').then(function () {
      return file.arrayBuffer();
    }).then(function (buf) {
      var wb = XLSX.read(buf, { type: 'array' });
      var parts = [];
      wb.SheetNames.forEach(function (name) {
        var rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: '' });
        var clean = rows.filter(function (r) { return r.some(function (c) { return c !== ''; }); });
        if (!clean.length) return;
        if (wb.SheetNames.length > 1) parts.push('## ' + name);
        parts.push(rowsToTable(clean.map(function (r) {
          return r.map(function (c) { return String(c); });
        })));
      });
      return parts.join('\n\n').trim();
    });
  }

  /* --- PDF: pdf.js text items grouped into lines by position.
         Items are joined into words by their actual horizontal
         gaps, so words the PDF split into fragments come back
         together and punctuation reattaches. Lines that repeat
         on most pages (signature stamps, running footers) are
         dropped. Remaining lines become paragraphs (adaptive
         joining) and simple tables (geometric detection). */
  function convertPdf(file) {
    return loadScript('pdf.min.js').then(function () {
      pdfjsLib.GlobalWorkerOptions.workerSrc = VENDOR + 'pdf.worker.min.js';
      return file.arrayBuffer();
    }).then(function (buf) {
      return pdfjsLib.getDocument({ data: buf }).promise;
    }).then(function (pdf) {
      var pageLines = [];
      var chain = Promise.resolve();
      for (var p = 1; p <= pdf.numPages; p++) {
        (function (pageNum) {
          chain = chain.then(function () {
            return pdf.getPage(pageNum);
          }).then(function (page) {
            return page.getTextContent();
          }).then(function (tc) {
            pageLines.push(pdfExtractLines(tc.items));
          });
        })(p);
      }
      return chain.then(function () {
        pdfDropBoilerplate(pageLines);
        var pagesMd = pageLines.map(pdfLinesToBlocks).filter(function (md) { return md; });
        /* A paragraph that continues over a page break arrives as
           two blocks; rejoin it when the first ends mid-sentence
           and the next opens with a lowercase letter. */
        var merged = [];
        pagesMd.forEach(function (md) {
          if (!merged.length) { merged.push(md); return; }
          var prev = merged[merged.length - 1];
          var prevParas = prev.split('\n\n');
          var nextParas = md.split('\n\n');
          var lastBlock = prevParas[prevParas.length - 1];
          var firstBlock = nextParas[0];
          var midSentence = !/[.!?\u2026]["')]*$/.test(lastBlock) &&
            lastBlock.charAt(0) !== '|' && !/^#/.test(lastBlock);
          if (midSentence && /^[a-z]/.test(firstBlock)) {
            prevParas[prevParas.length - 1] = lastBlock + ' ' + firstBlock;
            merged[merged.length - 1] = prevParas.concat(nextParas.slice(1)).join('\n\n');
          } else {
            merged.push(md);
          }
        });
        return merged.join('\n\n').trim();
      });
    });
  }

  function pdfExtractLines(items) {
    var linesMap = new Map();
    items.forEach(function (item) {
      if (!item.str || !item.str.trim()) return;
      var y = Math.round(item.transform[5]);
      if (!linesMap.has(y)) linesMap.set(y, []);
      linesMap.get(y).push({
        x: item.transform[4],
        end: item.transform[4] + (item.width || 0),
        str: item.str,
        h: item.height || 10
      });
    });
    var lines = Array.from(linesMap.entries()).sort(function (a, b) { return b[0] - a[0]; })
      .map(function (entry) {
        var parts = entry[1].sort(function (a, b) { return a.x - b.x; });
        var cells = pdfSplitCells(parts);
        return {
          y: entry[0],
          h: parts[0].h,
          x0: parts[0].x,
          x1: Math.max.apply(null, parts.map(function (p) { return p.end; })),
          cells: cells,
          text: cells.map(function (c) { return c.text; }).join(' ')
        };
      }).filter(function (l) { return l.text; });
    lines.forEach(function (l) { l.w = Math.max(0, l.x1 - l.x0); });
    return lines;
  }

  /* Join split fragments: a space is inserted only where the
     horizontal gap says there is one. Producers that split words
     into pieces leave gaps near zero; real spaces are several
     points wide (or arrive as separate space items, which are
     filtered out before this and leave their width as the gap). */
  function pdfJoinParts(parts) {
    var out = '';
    var prevEnd = null;
    parts.forEach(function (p) {
      if (prevEnd !== null && p.x - prevEnd > 1.2) out += ' ';
      out += p.str;
      prevEnd = Math.max(prevEnd === null ? p.end : prevEnd, p.end);
    });
    return out.replace(/\s+/g, ' ').trim();
  }

  function pdfSplitCells(parts) {
    var groups = [], cur = null;
    parts.forEach(function (p) {
      if (!cur) { cur = { x: p.x, end: p.end, parts: [p] }; return; }
      var gap = p.x - cur.end;
      if (gap > Math.max(12, (p.h || 10) * 1.05)) {
        groups.push(cur);
        cur = { x: p.x, end: p.end, parts: [p] };
      } else {
        cur.parts.push(p);
        cur.end = Math.max(cur.end, p.end);
      }
    });
    if (cur) groups.push(cur);
    return groups.map(function (g) {
      return { x: g.x, end: g.end, text: pdfJoinParts(g.parts) };
    }).filter(function (c) { return c.text; });
  }

  /* Lines whose text shows up on most pages of the document are
     stamps and running footers, not content. */
  function pdfDropBoilerplate(pageLines) {
    if (pageLines.length < 3) return;
    var counts = new Map();
    pageLines.forEach(function (lines) {
      var seen = {};
      lines.forEach(function (l) {
        var norm = l.text.toLowerCase();
        if (!seen[norm]) { seen[norm] = true; counts.set(norm, (counts.get(norm) || 0) + 1); }
      });
    });
    var threshold = Math.max(3, Math.ceil(pageLines.length * 0.6));
    pageLines.forEach(function (lines, idx) {
      pageLines[idx] = lines.filter(function (l) {
        return (counts.get(l.text.toLowerCase()) || 0) < threshold;
      });
    });
  }

  function pdfLinesToBlocks(lines) {
    if (!lines.length) return '';
    var pageText = lines.map(function (l) { return l.text; }).join('\n');
    /* Split into text segments and table segments. Consecutive
       multi-cell lines form a table segment while their column
       anchors stay compatible; a form block followed by a real
       table (different anchors) becomes two segments. */
    var blocks = [], textRun = [], i = 0;
    function flushText() {
      if (textRun.length) {
        var md = pdfLinesToParagraphs(textRun, pageText);
        if (md) blocks.push(md);
        textRun = [];
      }
    }
    var pageWidths = lines.map(function (l) { return l.w; }).sort(function (a, b) { return a - b; });
    var pageFullW = pageWidths.length ? pageWidths[Math.floor(pageWidths.length * 0.8)] : 0;
    var pageLeft = Math.min.apply(null, lines.map(function (l) { return l.x0; }));
    while (i < lines.length) {
      if (lines[i].cells.length >= 2) {
        /* A table run starts at a multi-cell line. It continues
           while lines carry cells on the run's column anchors,
           including single-cell lines (fragments of wrapped
           cells), and stops at full-width prose lines. */
        var runLines = [lines[i]];
        var runAnchors = [];
        lines[i].cells.forEach(function (c) { runAnchors.push(c.x); });
        i++;
        while (i < lines.length) {
          var cand = lines[i];
          var onAnchor = cand.cells.some(function (c) {
            return runAnchors.some(function (a) { return Math.abs(a - c.x) <= 8; });
          });
          /* Merged header lines carry disjoint column subsets, so
             a multi-cell line packed tightly against the previous
             one also belongs to the run, anchor match or not. */
          var prevY = runLines[runLines.length - 1].y;
          var tight = cand.cells.length >= 2 && (prevY - cand.y) > 0 &&
            (prevY - cand.y) <= (cand.h || 12) * 1.8 + 4;
          var isProse = cand.cells.length === 1 && pageFullW > 0 &&
            cand.w > pageFullW * 0.6 && Math.abs(cand.x0 - pageLeft) <= 8;
          if ((!onAnchor && !tight) || isProse) break;
          runLines.push(cand);
          cand.cells.forEach(function (c) {
            if (!runAnchors.some(function (a) { return Math.abs(a - c.x) <= 8; })) runAnchors.push(c.x);
          });
          i++;
        }
        var segments = [runLines];
        var placed = false;
        segments.forEach(function (sg) {
          var table = pdfBuildTable(sg);
          if (table) { flushText(); blocks.push(table); placed = true; }
          else textRun = textRun.concat(sg);
        });
        if (!placed) { /* all fell back to text */ }
      } else {
        textRun.push(lines[i]);
        i++;
      }
    }
    flushText();
    return blocks.join('\n\n');
  }

  function pdfBuildTable(seg) {
    if (seg.length < 2) return null;
    var anchors = [];
    seg.forEach(function (line) {
      line.cells.forEach(function (c) { anchors.push(c.x); });
    });
    anchors.sort(function (a, b) { return a - b; });
    var cols = [];
    anchors.forEach(function (x) {
      if (!cols.length || x - cols[cols.length - 1] > 6) cols.push(x);
    });
    if (cols.length < 2 || cols.length > 12) return null;
    var colUse = cols.map(function () { return 0; });
    seg.forEach(function (line) {
      var seen = {};
      line.cells.forEach(function (c) {
        var idx = pdfColumnIndex(cols, c.x);
        if (idx >= 0) seen[idx] = true;
      });
      Object.keys(seen).forEach(function (k) { colUse[Number(k)]++; });
    });
    var shared = colUse.filter(function (n) { return n >= 2; }).length;
    if (shared < 2) return null;

    /* Split the segment into chunks at wider vertical gaps
       (header block vs body), then decide per chunk: sparse
       chunks (most lines fill few columns) are one logical row
       with wrapped cell text; dense chunks are a row per line. */
    var gaps = [];
    for (var gi = 1; gi < seg.length; gi++) {
      var g = seg[gi - 1].y - seg[gi].y;
      if (g > 0) gaps.push(g);
    }
    gaps.sort(function (a, b) { return a - b; });
    var medGap = gaps.length ? gaps[Math.floor(gaps.length / 2)] : 0;
    var chunks = [[seg[0]]];
    for (var li = 1; li < seg.length; li++) {
      var gap = seg[li - 1].y - seg[li].y;
      if (medGap > 0 && gap > medGap * 1.6) chunks.push([]);
      chunks[chunks.length - 1].push(seg[li]);
    }
    var grid = [];
    chunks.forEach(function (chunk) {
      var fillTotal = 0;
      chunk.forEach(function (line) {
        var filled = {};
        line.cells.forEach(function (c) {
          var idx = pdfColumnIndex(cols, c.x);
          if (idx >= 0) filled[idx] = true;
        });
        fillTotal += Object.keys(filled).length;
      });
      var avgFill = chunk.length ? (fillTotal / chunk.length) / cols.length : 1;
      if (avgFill < 0.75 && chunk.length > 1) {
        var row = cols.map(function () { return ''; });
        chunk.forEach(function (line) {
          line.cells.forEach(function (c) {
            var idx = pdfColumnIndex(cols, c.x);
            if (idx < 0) idx = 0;
            row[idx] = row[idx] ? row[idx] + ' ' + c.text : c.text;
          });
        });
        grid.push(row);
      } else {
        chunk.forEach(function (line) {
          var row = cols.map(function () { return ''; });
          line.cells.forEach(function (c) {
            var idx = pdfColumnIndex(cols, c.x);
            if (idx < 0) idx = 0;
            row[idx] = row[idx] ? row[idx] + ' ' + c.text : c.text;
          });
          grid.push(row);
        });
      }
    });
    if (grid.length < 2) return null;
    /* Drop columns that are empty in every row. */
    var keep = cols.map(function (_, ci) {
      return grid.some(function (row) { return row[ci]; });
    });
    if (keep.filter(Boolean).length < 2) return null;
    grid = grid.map(function (row) {
      return row.filter(function (_, ci) { return keep[ci]; });
    });
    /* Two-column prose (journal layout) also aligns on anchors;
       its cells are long sentences, table cells are short. */
    if (grid[0].length === 2) {
      var words = 0, cellCount = 0;
      grid.forEach(function (row) {
        row.forEach(function (c) {
          if (c) { words += c.split(/\s+/).length; cellCount++; }
        });
      });
      if (cellCount && words / cellCount > 6) return null;
    }
    return rowsToTable(grid);
  }

  function pdfColumnIndex(cols, x) {
    var idx = -1;
    for (var i = 0; i < cols.length; i++) {
      if (cols[i] <= x + 6) idx = i;
      else break;
    }
    return idx;
  }

  var PDF_LIST_START = /^\s*(\d{1,3}[.)]|\u2022|\u25aa|\u25cf)\s/;
  function pdfLinesToParagraphs(lines, pageText) {
    /* Paragraph detection adapts to the document itself, because
       line spacing varies too much between PDF producers for one
       fixed threshold. Signals: the font size changes, the gap
       grows far beyond the segment's median line gap, the
       previous line is short (it ended a paragraph), the next
       line is indented (a first-line indent), or the next line
       opens a numbered or bulleted list item. */
    var gaps = [];
    for (var i = 1; i < lines.length; i++) {
      var g = lines[i - 1].y - lines[i].y;
      if (g > 0) gaps.push(g);
    }
    gaps.sort(function (a, b) { return a - b; });
    var medGap = gaps.length ? gaps[Math.floor(gaps.length / 2)] : 0;
    var widths = lines.map(function (l) { return l.w; }).sort(function (a, b) { return a - b; });
    var fullW = widths.length ? widths[Math.floor(widths.length * 0.8)] : 0;
    var leftEdge = Math.min.apply(null, lines.map(function (l) { return l.x0; }));

    var paragraphs = [], current = '', paraH = 0, prevLine = null, paraIsList = false;
    lines.forEach(function (line) {
      if (!current) {
        current = line.text; paraH = line.h; prevLine = line;
        paraIsList = PDF_LIST_START.test(line.text);
        return;
      }
      var gap = prevLine.y - line.y;
      var fontChanged = paraH > 0 && Math.abs(line.h - paraH) / paraH > 0.3;
      var gapBreak = medGap > 0 && gap > medGap * 1.4;
      var shortBreak = fullW > 0 && prevLine.w < fullW * 0.82 && line.x0 <= leftEdge + fullW * 0.08;
      var indentBreak = fullW > 0 && (line.x0 - prevLine.x0) > fullW * 0.04 && (line.x0 - prevLine.x0) < fullW * 0.3;
      var listBreak = PDF_LIST_START.test(line.text);
      /* A list item whose last line ends mid-sentence continues
         on the next line, whatever the spacing says. */
      var listContinuation = paraIsList && !/[.!?]["')]*$/.test(current);
      if (!fontChanged && !listBreak && (listContinuation || (!gapBreak && !shortBreak && !indentBreak))) {
        var hyph = current.match(/([A-Za-z]+)-$/);
        if (hyph && /^[a-z]/.test(line.text)) {
          var fragB = (line.text.match(/^[A-Za-z]+/) || [''])[0];
          if (pageText && pageText.indexOf(hyph[1] + '-' + fragB) >= 0) {
            current += line.text;
          } else {
            current = current.slice(0, -1) + line.text;
          }
        } else {
          current += ' ' + line.text;
        }
      } else {
        paragraphs.push(current);
        current = line.text;
        paraH = line.h;
        paraIsList = PDF_LIST_START.test(line.text);
      }
      prevLine = line;
    });
    if (current) paragraphs.push(current);
    return paragraphs.join('\n\n');
  }

  /* --- PPTX: own parser. ZIP, then slide XML in order.
         Title placeholder becomes a heading, tables become
         Markdown tables, other paragraphs become lines. */
  function convertPptx(file) {
    return loadScript('jszip.min.js').then(function () {
      return file.arrayBuffer();
    }).then(function (buf) {
      return JSZip.loadAsync(buf);
    }).then(function (zip) {
      var slideNames = Object.keys(zip.files)
        .filter(function (n) { return /^ppt\/slides\/slide\d+\.xml$/.test(n); })
        .sort(function (a, b) {
          return parseInt(a.match(/(\d+)\.xml$/)[1], 10) - parseInt(b.match(/(\d+)\.xml$/)[1], 10);
        });
      var out = [];
      var chain = Promise.resolve();
      slideNames.forEach(function (name, idx) {
        chain = chain.then(function () {
          return zip.file(name).async('string');
        }).then(function (xml) {
          var md = pptxSlideToMarkdown(parseXml(xml), idx + 1);
          if (md) out.push(md);
        });
      });
      return chain.then(function () { return out.join('\n\n').trim(); });
    });
  }
  function pptxSlideToMarkdown(doc, num) {
    var parts = ['<!-- Slide number: ' + num + ' -->'];
    var spTree = byTag(doc, 'spTree')[0];
    if (!spTree) return '';
    Array.prototype.slice.call(spTree.children).forEach(function (node) {
      if (node.localName === 'sp') {
        var ph = byTag(node, 'ph')[0];
        var type = ph ? (ph.getAttribute('type') || 'body') : 'body';
        var paras = byTag(node, 'p').map(function (p) {
          return byTag(p, 't').map(function (t) { return t.textContent; }).join('');
        }).filter(function (s) { return s.trim(); });
        if (!paras.length) return;
        if (type === 'title' || type === 'ctrTitle') {
          parts.push('# ' + paras.join(' '));
        } else {
          paras.forEach(function (s) { parts.push(s); });
        }
      } else if (node.localName === 'graphicFrame') {
        var rows = byTag(node, 'tr').map(function (tr) {
          return byTag(tr, 'tc').map(function (tc) {
            return byTag(tc, 't').map(function (t) { return t.textContent; }).join('');
          });
        }).filter(function (r) { return r.length; });
        if (rows.length) parts.push(rowsToTable(rows));
      }
    });
    return parts.length > 1 ? parts.join('\n\n') : '';
  }

  /* --- EPUB: own parser. container.xml to the OPF, then walk
         the spine and Turndown each chapter document. */
  function convertEpub(file) {
    return Promise.all([loadScript('jszip.min.js'), getTurndown()]).then(function (vals) {
      var td = vals[1];
      return file.arrayBuffer().then(function (buf) {
        return JSZip.loadAsync(buf);
      }).then(function (zip) {
        return zip.file('META-INF/container.xml').async('string').then(function (container) {
          var rootfile = byTag(parseXml(container), 'rootfile')[0].getAttribute('full-path');
          var base = rootfile.indexOf('/') >= 0 ? rootfile.slice(0, rootfile.lastIndexOf('/') + 1) : '';
          return zip.file(rootfile).async('string').then(function (opfText) {
            var opf = parseXml(opfText);
            var titleEl = byTag(opf, 'title')[0];
            var manifest = {};
            byTag(opf, 'item').forEach(function (item) {
              manifest[item.getAttribute('id')] = item.getAttribute('href');
            });
            var spine = byTag(opf, 'itemref').map(function (ref) {
              return manifest[ref.getAttribute('idref')];
            }).filter(Boolean);
            var parts = [];
            if (titleEl && titleEl.textContent.trim()) parts.push('# ' + titleEl.textContent.trim());
            var chain = Promise.resolve();
            spine.forEach(function (href) {
              var entry = zip.file(base + href);
              if (!entry) return;
              chain = chain.then(function () { return entry.async('string'); })
                .then(function (html) {
                  var doc = new DOMParser().parseFromString(html, 'text/html');
                  if (doc.body) {
                    var md = td.turndown(doc.body.innerHTML).trim();
                    if (md) parts.push(md);
                  }
                });
            });
            return chain.then(function () { return parts.join('\n\n').trim(); });
          });
        });
      });
    });
  }

  /* --- IPYNB: plain JSON. Markdown cells pass through, code
         cells become fenced blocks, outputs are skipped. */
  function convertIpynb(file) {
    return file.text().then(function (text) {
      var nb = JSON.parse(text);
      var lang = (nb.metadata && nb.metadata.language_info && nb.metadata.language_info.name) || 'python';
      var parts = [];
      (nb.cells || []).forEach(function (cell) {
        var src = Array.isArray(cell.source) ? cell.source.join('') : (cell.source || '');
        if (!src.trim()) return;
        if (cell.cell_type === 'code') parts.push('```' + lang + '\n' + src.replace(/\n$/, '') + '\n```');
        else parts.push(src.trim());
      });
      return parts.join('\n\n').trim();
    });
  }

  function convertHtml(file) {
    return getTurndown().then(function (td) {
      return file.text().then(function (text) {
        var doc = new DOMParser().parseFromString(text, 'text/html');
        return td.turndown(doc.body ? doc.body.innerHTML : text).trim();
      });
    });
  }

  function convertFile(file) {
    var ext = (file.name.split('.').pop() || '').toLowerCase();
    switch (ext) {
      case 'docx': return convertDocx(file);
      case 'xlsx': case 'xls': return convertSheet(file);
      case 'pdf': return convertPdf(file);
      case 'pptx': return convertPptx(file);
      case 'epub': return convertEpub(file);
      case 'ipynb': return convertIpynb(file);
      case 'html': case 'htm': return convertHtml(file);
      case 'csv': return file.text().then(function (t) { return rowsToTable(parseCsv(t)).trim(); });
      default: return file.text().then(function (t) { return t.trim(); });
    }
  }
  var SUPPORTED = ['pdf', 'docx', 'pptx', 'xlsx', 'xls', 'html', 'htm', 'epub', 'csv', 'txt', 'md', 'markdown', 'json', 'xml', 'ipynb'];

  /* ------------------------------------------------ UI */
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

  function setEngine(text) { engineEl.textContent = text; }
  function setStatus(text, isError) {
    statusEl.textContent = text;
    statusEl.classList.toggle("is-error", !!isError);
  }

  /* ---------------- File list ---------------- */
  function extOf(name) {
    var i = name.lastIndexOf(".");
    return i >= 0 ? name.slice(i + 1).toLowerCase() : "";
  }
  function formatSize(bytes) {
    if (bytes < 1024) return bytes + " B";
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
    return (bytes / (1024 * 1024)).toFixed(1) + " MB";
  }
  var STATUS_TEXT = {
    queued: "waiting",
    converting: "converting…",
    done: "done",
    error: "failed"
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
      st.textContent = entry.status === "error" ? ("failed: " + entry.error) : STATUS_TEXT[entry.status];
      st.title = entry.status === "error" ? entry.error : "";
      var rm = document.createElement("button");
      rm.type = "button";
      rm.className = "file-remove";
      rm.textContent = "remove";
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
    if (busy) { setStatus("Wait for the conversion in progress to finish first.", false); return; }
    var added = 0;
    Array.prototype.forEach.call(fileArray, function (file) {
      var ext = extOf(file.name);
      var entry = {
        id: nextId++, file: file, name: file.name, ext: ext,
        status: "queued", markdown: null, error: ""
      };
      if (SUPPORTED.indexOf(ext) < 0) {
        entry.status = "error";
        entry.error = "format ." + (ext || "?") + " is not supported";
      } else if (file.size > MAX_BYTES) {
        entry.status = "error";
        entry.error = "file is larger than 50 MB";
      } else if (file.size === 0) {
        entry.status = "error";
        entry.error = "file is empty";
      } else {
        added++;
      }
      entries.push(entry);
    });
    renderList(); refreshButtons();
    if (added > 0) {
      setStatus(added + (added === 1 ? " file ready" : " files ready") + " to convert.", false);
    } else {
      setStatus("No valid files were added. Check the format and size.", true);
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
    setStatus("File list cleared.", false);
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
        setStatus("Converting " + entry.name + "…", false);
        return convertFile(entry.file).then(function (markdown) {
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
      setEngine("ready");
      var done = entries.filter(function (e) { return e.status === "done"; }).length;
      var failed = entries.filter(function (e) { return e.status === "error"; }).length;
      if (failed === 0) setStatus("Done. " + done + (done === 1 ? " file" : " files") + " converted.", false);
      else setStatus("Finished with issues: " + done + " converted, " + failed + " failed. Failed files can be retried with the Convert button.", failed > 0 && done === 0);
    });
  });

  function shortenError(err) {
    var msg = String(err && err.message ? err.message : err);
    if (/Invalid PDF/i.test(msg)) return "this PDF could not be read";
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
    outputMeta.textContent = chars.toLocaleString("en-US") + " characters, " + words.toLocaleString("en-US") + " words";
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
    function ok() { copyBtn.textContent = "Copied"; setTimeout(function () { copyBtn.textContent = "Copy Markdown"; }, 1600); }
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
