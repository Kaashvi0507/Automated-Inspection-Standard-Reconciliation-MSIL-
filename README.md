# 🔍 GSIS-P / SMIR Reconciliation

[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![Streamlit App](https://img.shields.io/badge/streamlit-1.45.0+-red.svg)](https://streamlit.io)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Code style: black](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)

A **fully local, CPU-only reconciliation system** for matching Excel (GSIS-P) against PDF (SMIR) documents with human-in-the-loop verification, OCR correction, and intelligent fuzzy matching. Designed for Maruti Suzuki vendor quality inspection workflows.

> **No internet required at runtime** — all OCR models and embeddings are bundled locally.

## ✨ Features

### 🔄 Three-Phase Reconciliation Pipeline
- **Phase 1: Metadata Extraction** — Vendor Code, Part Number, Model No., Issue Date from PDF headers
- **Phase 2: Context-Aware OCR Correction** — Merged-cell deduplication + semantic embeddings for noise cleanup
- **Phase 3: Intelligent Matching & Validation** — Fuzzy item matching, numeric tolerance cross-checks, MIC category validation

### 🧠 Advanced Capabilities
- **Confidence Scoring** — Row-level health metrics with detailed reasoning (OCR artifacts, format violations, missing data)
- **Semantic Similarity** — Sentence embeddings (all-MiniLM-L6-v2) for tolerant item & criterion matching
- **Fuzzy Matching** — RapidFuzz-powered string similarity for OCR error recovery
- **Manual Alignment** — Override automatic matches for edge cases
- **Table Extraction** — Multi-engine OCR (RapidOCR, optional: Tesseract, Docling) with PDF pdfplumber fallback

### 📊 UI / UX
- **Real-time Comparison Grid** — Side-by-side PDF ↔ Excel in 13-column layout
- **Color-Coded Severity** — Visual indicators for critical, warning, info issues
- **Confidence Dashboard** — Track data quality with min/max/avg confidence per row
- **Pagination & Search** — Browse large reconciliations (50–200 rows per page)
- **Role-based Editing** — Read-only PDF data, full control over Excel corrections
- **Bulk Operations** — Renumber Operation №, Reorder by MIC, Re-apply automation

### 💾 Export & Traceability
- **Corrected Excel** — Updated GSIS-P with all edits applied
- **Rebuilt-from-PDF** — Clean Excel reconstructed directly from PDF truth
- **Diff CSV** — Row-by-row change log (before/after/status)

## 📋 19-Column Schema

Strict GSIS-P format validation:

```
VENDOR CODE | Part number | Model No. | Operation number | MIC Name | 
Inspection Item | Parameter | Lower Limit | Target Value | Upper Limit | 
Decimal Places | UOM | Inspection Method | Inspection Tool | 
Info Field 1 | Info Field 2 | Info Field 3 | Issue date | Long Text
```

- **Numeric validation** for limits and decimal places
- **MIC categories**: Appearance, Dimension, Material, Performance, Revalidation
- **Date format**: DD-MM-YYYY

## 🚀 Quick Start

### Prerequisites
- **Python 3.8+**
- **Windows / macOS / Linux** (fully tested on Windows with bundled OCR)
- **No GPU required** — CPU-only; runs on budget hardware

### Installation

1. **Clone & navigate:**
   ```bash
   git clone https://github.com/GouravKim/gsis-reconciliation.git
   cd gsis-reconciliation
   ```

2. **Create virtual environment:**
   ```bash
   # Windows
   python -m venv venv
   venv\Scripts\activate
   
   # macOS / Linux
   python3 -m venv venv
   source venv/bin/activate
   ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Prefetch heavy models (optional, recommended):**
   ```bash
   # Bundles Docling models for fallback OCR
   python -c "from docling import Document; Document.from_pdf_bytes(b'')" 2>/dev/null || true
   ```

### Running the App

**Local only:**
```bash
streamlit run reconcile.py
```
Opens `http://localhost:8502` in your browser.

**Network access (shared PC on LAN):**

**Windows:**
```bash
start.bat
```

**macOS / Linux:**
```bash
python -m streamlit run reconcile.py --server.port 8502 --server.address 0.0.0.0 --server.headless true
```

Then on any machine, open: `http://<your-pc-ip>:8502`

## 📖 Usage

### Step 1: Load Documents
1. Upload a **PDF** (SMIR source of truth)
2. Upload an **Excel** (GSIS-P to validate)
3. App auto-extracts metadata and starts reconciliation

### Step 2: Review Matches
- **Comparison Tab**: Side-by-side grid of PDF vs. Excel rows
- **Filters**: Show all, rows with PDF match, rows missing PDF match
- **Confidence scores**: Hover to see OCR risk factors
- Search by inspection item or parameter

### Step 3: Manual Fixes
- **Review Tab**: Edit cells, renumber operations, reorder by MIC
- **Manual Alignment**: Override auto-matches for unmatched rows
- **PDF Tab**: View source PDF pages (original evidence)

### Step 4: Export
- **Corrected Excel**: Your edits merged back
- **Rebuilt-from-PDF**: Pure PDF truth in Excel format
- **Diff CSV**: Change tracking for audit trail

## ⚙️ Configuration

Environment variables (optional, set in `start.bat` or shell):

```bash
# Disable Hugging Face cloud fallback (required for offline-only)
HF_HUB_OFFLINE=1
TRANSFORMERS_OFFLINE=1
HF_HUB_DISABLE_SYMLINKS=1

# Streamlit performance (avoids torch introspection crashes)
STREAMLIT_SERVER_FILE_WATCHER_TYPE=none
```

All settings are **already configured in `start.bat`** and `reconcile.py` — no manual setup needed.

## 📦 Dependencies

### Core (always bundled)
- `streamlit` 1.45.0 — Web UI framework
- `pandas` 2.2.3 — Data manipulation
- `pdfplumber` 0.11.6 — PDF extraction
- `PyMuPDF` 1.28.0 — PDF rendering
- `rapidocr-onnxruntime` 1.4.4 — Fast OCR (bundled models)
- `sentence-transformers` 3.4.1 — Semantic embeddings

### Optional (fallback)
- `docling` 2.114.0 — Heavy OCR engine (models auto-download on first use)
- `pytesseract` / `cv2` — OpenCV table detection
- `rapidfuzz` — Advanced fuzzy matching

**→ No internet at runtime** — all models are local.

## 🏗️ Architecture

```
reconcile.py
├── Config & Constants
│   ├── 19-column SCHEMA
│   ├── Numeric/Required/Optional columns
│   └── MIC category validation
├── Phase 1: Metadata Extraction
│   ├── PDF header parsing (vendor, part, model, date)
│   └── Confidence scoring
├── Phase 2: OCR & Table Extraction
│   ├── RapidOCR + Docling fallback
│   ├── Context-aware correction (embeddings)
│   └── Merged-cell deduplication
├── Phase 3: Matching & Validation
│   ├── Fuzzy item matching (RapidFuzz)
│   ├── Tolerance cross-checks
│   ├── MIC category validation
│   └── Severity classification (critical/warning/info)
└── UI & Export
    ├── Streamlit multi-tab dashboard
    ├── Real-time Excel editor
    └── Download (XLSX/CSV)
```

## 🔍 Key Functions

| Function | Purpose |
|----------|---------|
| `extract_metadata_from_pdf()` | Parse vendor/part/model from PDF |
| `extract_tables_from_pdf()` | OCR + table extraction with multi-engine fallback |
| `calculate_cell_confidence()` | Confidence scoring for data quality |
| `fuzzy_match()` | Item matching with similarity threshold |
| `validate_schema()` | Enforce required columns & numeric formats |
| `build_workbook()` | Reconstruct corrected Excel from edits |
| `build_diff()` | Generate change log CSV |

## 🐛 Troubleshooting

### Blank screen on startup
- **Issue**: Streamlit file watcher crashes on PyTorch.
- **Fix**: `STREAMLIT_SERVER_FILE_WATCHER_TYPE=none` (already set in `start.bat`)

### "RapidOCR not found"
- **Issue**: `rapidocr-onnxruntime` wheel not installed.
- **Fix**: `pip install rapidocr-onnxruntime==1.4.4`

### PDF pages don't render
- **Issue**: PyMuPDF version conflict.
- **Fix**: `pip install PyMuPDF==1.28.0 --force-reinstall`

### "NaTType does not support strftime"
- **Issue**: Date validation on mixed types.
- **Fix**: Already patched in `validate_schema()` via `pd.isna()` checks

### Models downloading on first run
- **Issue**: First use of Docling triggers HuggingFace download.
- **Fix**: Run `python -c "from docling import Document; Document.from_pdf_bytes(b'')"` once offline to prefetch.

## 📊 Example Workflow

**Scenario**: Reconcile a vendor quality inspection spreadsheet against the official PDF:

1. Open app → `http://localhost:8502`
2. Upload `SMIR_Vendor_PN_17855.pdf` (source of truth)
3. Upload `GSIS-P_Corrections_v2.xlsx` (to validate)
4. Review **Comparison** tab:
   - ✅ Green rows: Perfect matches
   - ⚠️ Yellow rows: Minor OCR issues (auto-corrected)
   - 🔴 Red rows: Misaligns or missing PDF reference
5. Click on red rows → **Manual Alignment** → pick correct PDF row
6. Fix any data inconsistencies → **Save Changes**
7. Export:
   - `corrected.xlsx` → Send back to vendor
   - `diff.csv` → Audit trail (what changed & why)

## 🤝 Contributing

We welcome contributions! Please:

1. **Fork** the repo
2. **Create a branch**: `git checkout -b feature/your-feature`
3. **Commit**: `git commit -m "Add feature"` (follow [conventional commits](https://www.conventionalcommits.org/))
4. **Test**: Ensure `reconcile.py` runs without errors
5. **Push**: `git push origin feature/your-feature`
6. **Open a Pull Request** with a clear description

### Code Style
- **Python 3.8+** type hints preferred
- **Black** formatting: `black reconcile.py`
- **Docstrings** for all functions (Google style)

## 📄 License

This project is licensed under the **MIT License** — see [LICENSE](LICENSE) for details.

## 🙋 Support

- **Issues?** Open a GitHub issue with:
  - OS & Python version
  - Exact error message + traceback
  - Sample PDF & Excel (if shareable)
- **Feature request?** Describe use case & expected outcome

## 📚 Further Reading

- [Streamlit Docs](https://docs.streamlit.io)
- [pdfplumber](https://github.com/jsvine/pdfplumber)
- [Sentence Transformers](https://www.sbert.net/)
- [RapidFuzz](https://rapidfuzz.github.io/)

## 👨‍💻 Author

**Gourav** — Software Engineer @ AI/ML
- GitHub: [@GouravKim](https://github.com/GouravKim)
- LinkedIn: [Gourav](https://linkedin.com/in/gouravkim)

---

**Built with ❤️ for Maruti Suzuki vendor quality workflows.**
