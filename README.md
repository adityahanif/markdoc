# Markdoc

Convert documents to Markdown, entirely in your browser. No uploads: files are read locally and never leave the device.

**Live:** https://adityahanif.github.io/markdoc/

## How it works

Markdoc is a static site (vanilla HTML, CSS, and JS, no build step). Each format is converted by a JavaScript library or a small built-in parser, loaded lazily from the `vendor/` folder on this same origin the first time that format is used. Nothing is fetched from a CDN, and document contents are never sent anywhere.

| Format | Engine |
| --- | --- |
| .docx | Mammoth 1.13 to HTML, then Turndown 7 (GFM). Tables get a real header row, because mammoth emits td-only tables. |
| .xlsx, .xls | SheetJS 0.18.5, one Markdown table per sheet. |
| .pdf | Built on pdf.js 3.11 text items; see "PDF pipeline" below. |
| .pptx | Built-in parser: JSZip opens the file, slide XML is read in order; titles become headings, tables become Markdown tables. Speaker notes and diagram text are not extracted. |
| .epub | Built-in parser: JSZip plus the OPF spine; each chapter document goes through Turndown, and the book title becomes the top heading. |
| .html, .htm | Turndown (GFM). |
| .csv | Built-in CSV parser to a Markdown table. |
| .ipynb | Built-in JSON parser: Markdown cells pass through, code cells become fenced blocks, outputs are skipped. |
| .txt, .md, .markdown, .json, .xml | Passed through as plain text. |

File size limit: 50 MB per file. Multiple files can be queued and converted in sequence.

### PDF pipeline

PDFs store positioned text fragments, not paragraphs or tables, so the converter reconstructs structure from geometry:

1. **Lines and words.** Items are grouped into lines by their y position. Within a line, a space is inserted only where the horizontal gap says there is one (a couple of points or more). Producers that split words into fragments leave near-zero gaps, so split words, digits, and punctuation come back together (`202 6` becomes `2026`, `word .` becomes `word.`).
2. **Boilerplate removal.** A line whose text appears on at least 60% of the pages is treated as a stamp or running footer and dropped. This removes, for example, the electronic signature (BSrE) stamp block that Indonesian government documents repeat on every page.
3. **Tables.** Lines are split into cells at wide horizontal gaps. A run of multi-cell lines becomes a table when its columns line up on shared start positions; runs continue through single-cell lines that sit on those anchors, which is how cells whose text wraps over several visual lines are reassembled. Sparse chunks of a run (most lines fill few columns) are treated as one logical row, dense chunks as one row per line, and chunks are separated at wider vertical gaps, so merged header rows are rebuilt too. Columns that end up empty everywhere are dropped.
4. **Paragraphs.** Remaining lines are joined into paragraphs adaptively, using the segment's own median line gap, short paragraph-ending lines, first-line indents, and font size changes as signals. Numbered and bulleted lines start a new paragraph, and a list item that ends mid-sentence continues onto the next line. Hyphenated line breaks are rejoined; a hyphen is kept when the hyphenated form occurs intact elsewhere on the page (reduplicated compounds such as `masing-masing`). Paragraphs split by a page break are rejoined when the first part ends mid-sentence and the next page opens with a lowercase letter.

## Not supported

Scanned PDFs and image OCR, audio transcription, and embedded images (images inside documents are skipped). PDF tables with cells spanning multiple columns, borderless tables with very tight column gaps, and two-column page layouts (journals) can still come out wrong or as plain text lines.

## Development

Everything is static. Serve the folder with any web server, for example:

    python3 -m http.server

Then bump the `?v=` query on the stylesheet and script tags in `index.html` whenever `style.css` or `app.js` changes, so GitHub Pages visitors skip stale caches. The vendored libraries in `vendor/` are loaded with plain script tags (no bundler): mammoth.browser.min.js, turndown.js, turndown-plugin-gfm.js, xlsx.full.min.js, pdf.min.js and pdf.worker.min.js, jszip.min.js.

## History

The first version of Markdoc (2026-11-03) ran Microsoft MarkItDown 0.1.9 in the browser via Pyodide and WebAssembly. It worked, but the first load was tens of megabytes and took a long while. The JavaScript build replaced it the same day: output quality matched or beat the Python build on the formats tested (DOCX, PDF, XLSX, HTML, CSV compared side by side), at roughly 800 KB of libraries in total, loaded per format on demand. The PDF pipeline was then tuned against a real signed government report whose first conversion showed scrambled tables, split words, and a signature stamp repeated on all eight pages; the current pipeline is the result of fixing those failure modes one by one from the document's actual text geometry.

## License

Site code: MIT. Vendored libraries keep their own licenses (Mammoth BSD-2, Turndown MIT, SheetJS Apache-2.0, pdf.js Apache-2.0, JSZip MIT).
