# Markdoc

Convert documents to Markdown, entirely in your browser. No uploads: files are read locally and never leave the device.

**Live:** https://adityahanif.github.io/markdoc/

## How it works

Markdoc is a static site (vanilla HTML, CSS, and JS, no build step). Each format is converted by a JavaScript library or a small built-in parser, loaded lazily from the `vendor/` folder on this same origin the first time that format is used. Nothing is fetched from a CDN, and document contents are never sent anywhere.

| Format | Engine |
| --- | --- |
| .docx | Mammoth 1.13 to HTML, then Turndown 7 (GFM). Tables get a real header row, because mammoth emits td-only tables. |
| .xlsx, .xls | SheetJS 0.18.5, one Markdown table per sheet. |
| .pdf | pdf.js 3.11 text extraction. Lines are grouped by position, then joined into paragraphs adaptively: the page median line gap, short paragraph-ending lines, first-line indents, and font size changes are the signals, so wrapped lines merge back into sentences even in loosely spaced documents. Hyphenated line breaks are rejoined too. |
| .pptx | Built-in parser: JSZip opens the file, slide XML is read in order; titles become headings, tables become Markdown tables. Speaker notes and diagram text are not extracted. |
| .epub | Built-in parser: JSZip plus the OPF spine; each chapter document goes through Turndown, and the book title becomes the top heading. |
| .html, .htm | Turndown (GFM). |
| .csv | Built-in CSV parser to a Markdown table. |
| .ipynb | Built-in JSON parser: Markdown cells pass through, code cells become fenced blocks, outputs are skipped. |
| .txt, .md, .markdown, .json, .xml | Passed through as plain text. |

File size limit: 50 MB per file. Multiple files can be queued and converted in sequence.

## Not supported

Scanned PDFs and image OCR, audio transcription, and embedded images (images inside documents are skipped). PDF tables come out as plain text lines, not Markdown tables.

## Development

Everything is static. Serve the folder with any web server, for example:

    python3 -m http.server

Then bump the `?v=` query on the stylesheet and script tags in `index.html` whenever `style.css` or `app.js` changes, so GitHub Pages visitors skip stale caches. The vendored libraries in `vendor/` are loaded with plain script tags (no bundler): mammoth.browser.min.js, turndown.js, turndown-plugin-gfm.js, xlsx.full.min.js, pdf.min.js and pdf.worker.min.js, jszip.min.js.

## History

The first version of Markdoc (2026-11-03) ran Microsoft MarkItDown 0.1.9 in the browser via Pyodide and WebAssembly. It worked, but the first load was tens of megabytes and took a long while. The JavaScript build replaced it the same day: output quality matched or beat the Python build on the formats tested (DOCX, PDF, XLSX, HTML, CSV compared side by side), at roughly 800 KB of libraries in total, loaded per format on demand.

## License

Site code: MIT. Vendored libraries keep their own licenses (Mammoth BSD-2, Turndown MIT, SheetJS Apache-2.0, pdf.js Apache-2.0, JSZip MIT).
