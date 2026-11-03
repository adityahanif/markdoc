# Markdoc

Convert dokumen ke Markdown langsung di browser, tanpa upload.

Markdoc menjalankan library **MarkItDown** asli dari Microsoft (Python, versi 0.1.9) di dalam browser memakai **Pyodide** (CPython yang dikompilasi ke WebAssembly). File yang dipilih pengguna dibaca lewat File API dan dikonversi di Web Worker pada tab yang sama. Dokumen tidak pernah dikirim ke server mana pun.

Live: https://adityahanif.github.io/markdoc/

## Format yang didukung

Sudah diuji end-to-end di browser dengan file contoh nyata:

| Format | Ekstensi |
| --- | --- |
| PDF teks | .pdf |
| Word | .docx |
| PowerPoint | .pptx |
| Excel | .xlsx, .xls |
| Web dan e-book | .html, .htm, .epub |
| Teks dan data | .csv, .txt, .md, .json, .xml |
| Notebook | .ipynb |

Belum didukung di build browser ini:

- PDF hasil scan dan OCR gambar (butuh model/tool di luar WebAssembly)
- Transkripsi audio dan deskripsi gambar otomatis
- Tabel PDF kompleks bisa kurang rapi, karena jalur pdfplumber tidak tersedia di browser (lihat catatan teknis)

## Cara kerja

1. Pengguna memilih atau menjatuhkan file. File dibaca sebagai ArrayBuffer lalu diteruskan ke `worker.js` lewat `postMessage`.
2. Worker memuat Pyodide dari CDN, memasang MarkItDown dan dependensinya dari PyPI dengan micropip, lalu mengonversi file dengan `MarkItDown.convert_stream()`.
3. Hasil Markdown ditampilkan sebagai teks mentah di panel graphite atau sebagai pratinjau, lalu bisa disalin atau diunduh sebagai `.md`.

Pandas dan openpyxl untuk file spreadsheet dimuat terpisah, hanya saat file `.xlsx`/`.xls` pertama dikonversi, supaya load awal lebih ringan.

Yang diunduh browser pada pemakaian pertama hanyalah kode (runtime Pyodide dari CDN jsDelivr dan wheel Python dari PyPI), totalnya puluhan megabyte, dan tersimpan di cache browser untuk kunjungan berikutnya.

## Catatan teknis: dua stub kecil

Dua dependensi MarkItDown tidak punya build WebAssembly, jadi `worker.js` memasang stub sebelum MarkItDown diimpor:

- **magika** (deteksi tipe file berbasis onnxruntime). Markdoc selalu meneruskan ekstensi, nama file, dan mimetype secara eksplisit lewat `StreamInfo`, jadi stub cukup melaporkan "unknown": MarkItDown akan memakai tebakan dari ekstensi/mimetype, dan pemilihan converter tetap tepat.
- **pdfplumber** (butuh binary pypdfium2). Stub membuat `open()` langsung gagal, sehingga PdfConverter memakai jalur fallback resminya sendiri, yaitu pdfminer.six. Hasil PDF teks tetap sama; ekstraksi tabel terstruktur dari PDF yang biasanya ditangani pdfplumber tidak ikut.

Converter spreadsheet mencatat ketiadaan pandas/openpyxl saat modul diimpor, jadi setelah pandas dipasang, `worker.js` memuat ulang modul `markitdown` dan membangun ulang converter (`markdoc_reload_converter()`).

## Menjalankan lokal

Cukup server statis biasa:

```bash
python3 -m http.server 8000
# buka http://localhost:8000
```

Tidak ada proses build. Semua file adalah HTML/CSS/JS vanilla.

## Struktur

- `index.html` — halaman dan panel converter
- `style.css`, `tokens.css`, `base.css` — gaya (tema Hallmark Cobalt)
- `app.js` — logika UI, antrean file, salin/unduh, pratinjau Markdown
- `worker.js` — worker Pyodide + MarkItDown, termasuk kedua stub di atas

## Lisensi

Kode situs ini MIT. MarkItDown adalah proyek Microsoft berlisensi MIT; Pyodide berlisensi MPL-2.0.
