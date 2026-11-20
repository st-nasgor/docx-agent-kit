> **Local environment note.** `pypdf`, `pdfplumber`, `Pillow` and `pymupdf`
> are installed. `pdf2image` needs the Poppler binary, which may not be
> installed — rasterize with `pymupdf` instead
> (`import pymupdf; page.get_pixmap(dpi=110).save(png)`). Tesseract and qpdf
> are typically absent, so OCR and qpdf-specific repair are unavailable.
> LibreOffice CLI: `soffice.com` on Windows
> (`%ProgramFiles%\LibreOffice\program\soffice.com`), `soffice` on macOS/Linux.
