<p align="center">
  <img src="https://img.shields.io/badge/Python-3.12+-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python 3.12+"/>
  <img src="https://img.shields.io/badge/Flask-3.0+-000000?style=for-the-badge&logo=flask&logoColor=white" alt="Flask"/>
  <img src="https://img.shields.io/badge/Gemini_AI-2.0_Flash-4285F4?style=for-the-badge&logo=google&logoColor=white" alt="Gemini AI"/>
  <img src="https://img.shields.io/badge/License-MIT-green?style=for-the-badge" alt="License MIT"/>
</p>

<h1 align="center">📄 PDF to Word Converter</h1>

<p align="center">
  <b>High-fidelity PDF → DOCX conversion with AI-powered quality assurance</b><br/>
  Preserves formatting, images, tables, and layout with near-pixel-perfect accuracy.
</p>

<p align="center">
  <a href="#-quick-start">Quick Start</a> •
  <a href="#-features">Features</a> •
  <a href="#-architecture">Architecture</a> •
  <a href="#-ai-strategies">AI Strategies</a> •
  <a href="#-web-ui">Web UI</a> •
  <a href="#-cli-reference">CLI Reference</a> •
  <a href="#-configuration">Configuration</a>
</p>

---

## ✨ Features

| | Feature | Description |
|---|---|---|
| 📝 | **Full Text Formatting** | Font family, size, color, bold, italic, underline, strikethrough |
| 🖼️ | **Image Extraction** | SMask / alpha compositing with transparency flattened onto white; display sizing from PDF bbox; overlapping image merging |
| 📊 | **Table Detection** | Cell shading, column collapsing, row heights, cell padding, border detection |
| 📐 | **Layout Analysis** | Multi-column detection, content-based page margins, exact paragraph spacing (`space_before` / `space_after` from PDF line gaps) |
| 🔖 | **Headers & Footers** | Repeating-text detection across pages |
| 📑 | **Page Breaks** | Orientation-aware page break insertion with section properties |
| 🔍 | **OCR Support** | Tesseract integration for scanned / image-based PDFs |
| 🤖 | **AI Quality Assurance** | Two Gemini-powered strategies for near-perfect output (see [AI Strategies](#-ai-strategies)) |
| 📈 | **Visual SSIM Scoring** | Pixel-level comparison: PDF → PNG vs DOCX → LibreOffice → PDF → PNG |
| 🌐 | **Web Interface** | Drag-and-drop upload, SSE progress, side-by-side preview, conversion history |
| 🔒 | **Encrypted PDFs** | Password-protected document support |
| 📦 | **Batch Conversion** | Convert entire directories of PDFs in one command |

---

## 🚀 Quick Start

### Prerequisites

| Requirement | Purpose | Required? |
|---|---|---|
| **Python 3.12+** | Runtime | ✅ Yes |
| **LibreOffice** | DOCX → PDF rendering for visual diff | ⚠️ For `--visual-validate` |
| **Tesseract OCR** | Scanned PDF text recognition | ⚠️ For `--ocr` |
| **Gemini API Key** | AI-powered comparison & correction | ⚠️ For `--ai-compare` |

### Installation

```bash
# Clone the repository
git clone https://github.com/hfzizz/PDFtoWord.git
cd PDFtoWord

# Create virtual environment
python -m venv .venv
```

**Activate the environment:**

<table>
<tr><th>Windows</th><th>macOS / Linux</th></tr>
<tr>
<td>

```powershell
.venv\Scripts\activate
```

</td>
<td>

```bash
source .venv/bin/activate
```

</td>
</tr>
</table>

```bash
# Install Python dependencies
pip install -r requirements.txt
```

**Install external tools (optional):**

<details>
<summary><b>LibreOffice</b> — required for visual validation</summary>

| OS | Command |
|---|---|
| **Windows** | Download from [libreoffice.org](https://www.libreoffice.org/download/) and install |
| **macOS** | `brew install --cask libreoffice` |
| **Ubuntu / Debian** | `sudo apt install libreoffice` |
| **Fedora** | `sudo dnf install libreoffice` |

</details>

<details>
<summary><b>Tesseract OCR</b> — required for scanned PDFs</summary>

| OS | Command |
|---|---|
| **Windows** | Download from [UB Mannheim](https://github.com/UB-Mannheim/tesseract/wiki) and add to PATH |
| **macOS** | `brew install tesseract` |
| **Ubuntu / Debian** | `sudo apt install tesseract-ocr` |
| **Fedora** | `sudo dnf install tesseract` |

</details>

### Set up Gemini API Key (optional)

```bash
# Linux / macOS
export GEMINI_API_KEY="your_api_key_here"

# Windows (PowerShell)
$env:GEMINI_API_KEY = "your_api_key_here"

# Windows (CMD)
set GEMINI_API_KEY=your_api_key_here
```

Or add `"api_key": "your_key"` under `"ai_comparison"` in `config/settings.json`.

Get a key from [Google AI Studio](https://aistudio.google.com/apikey).

### Convert your first PDF

```bash
python pdf2docx.py input.pdf
```

---

## 🏗️ Architecture

```
PDFtoWord/
│
├── pdf2docx.py                  # Main orchestrator + CLI entry point
├── config/
│   └── settings.json            # All configuration options
├── requirements.txt             # Python dependencies
│
├── extractors/                  # ── Stage 1-3: Data Extraction ──
│   ├── text_extractor.py        # Text + font info per page
│   ├── image_extractor.py       # Images with SMask/transparency handling
│   ├── table_extractor.py       # Table detection + column collapsing
│   └── metadata_extractor.py    # Per-page dimensions, rotation
│
├── analyzers/                   # ── Stage 4-5: Analysis ──
│   ├── font_analyzer.py         # Font mapping & style detection
│   ├── layout_analyzer.py       # Layout structure detection
│   └── semantic_analyzer.py     # Semantic ordering, page breaks, headers/footers
│
├── builders/                    # ── Stage 6: Document Construction ──
│   └── docx_builder.py          # DOCX construction with formatting overrides
│
├── quality/                     # ── Stage 7-9: Quality Assurance ──
│   ├── visual_diff.py           # SSIM scoring (PyMuPDF + LibreOffice render)
│   ├── ai_comparator.py         # Gemini vision page-by-page comparison
│   ├── ai_layout_analyzer.py    # Gemini pre-build layout analysis (Strategy B)
│   └── correction_engine.py     # Auto-fix engine for Strategy A corrections
│
├── utils/                       # ── Shared Utilities ──
│   ├── ocr_handler.py           # Tesseract OCR integration
│   ├── pdf_info.py              # PDF metadata analysis
│   ├── progress.py              # Progress tracking
│   └── validator.py             # Output validation + quality scoring
│
├── web/                         # ── Web Interface ──
│   ├── app.py                   # Flask application factory
│   ├── routes/                  # upload, status (SSE), download endpoints
│   ├── services/                # Converter service (background threads)
│   ├── templates/index.html     # Single-page application
│   └── static/                  # CSS + JS assets
│
└── tests/
    └── test_ai_comparison.py    # 24 tests (AI comparator + correction engine)
```

---

## 🔄 Conversion Pipeline

The converter processes each PDF through a multi-stage pipeline:

```
┌─────────────┐     ┌─────────────┐     ┌──────────────┐     ┌───────────────┐
│  Extract     │     │  Extract     │     │  Extract      │     │  Extract       │
│  Text +      │────▶│  Images +    │────▶│  Tables +     │────▶│  Metadata      │
│  Fonts       │     │  SMask       │     │  Shading      │     │  (dimensions)  │
└─────────────┘     └─────────────┘     └──────────────┘     └───────┬───────┘
                                                                      │
                    ┌─────────────┐     ┌──────────────┐              ▼
                    │  Font        │     │  Semantic     │     ┌───────────────┐
                    │  Analysis    │◀────│  Analysis     │◀────│  Layout        │
                    │  & Mapping   │     │  & Ordering   │     │  Detection     │
                    └──────┬──────┘     └──────────────┘     └───────────────┘
                           │
                           ▼
                    ┌─────────────────────────────────────────────────┐
                    │              DOCX Builder                       │
                    │  (sections, paragraphs, tables, images, styles) │
                    └──────────────────────┬──────────────────────────┘
                                           │
                              ┌────────────┴────────────┐
                              ▼                         ▼
                    ┌─────────────────┐       ┌─────────────────┐
                    │  Strategy A      │       │  Strategy B      │
                    │  (Post-Build)    │       │  (Pre-Build)     │
                    │  AI Correction   │       │  AI-Guided       │
                    │  Loop            │       │  Layout          │
                    └─────────────────┘       └─────────────────┘
```

---

## 🤖 AI Strategies

PDFtoWord features two distinct AI-powered quality strategies using **Google Gemini 2.0 Flash** via the `google-genai` SDK:

### Strategy A — Post-Build Correction Loop

```
PDF ──▶ Build DOCX ──▶ Render Both ──▶ AI Compare ──▶ Generate Fixes ──▶ Rebuild ──┐
                              ▲                                                      │
                              └──────────── repeat up to 3 rounds ◀──────────────────┘
```

1. Build the DOCX using the standard extraction pipeline
2. Render both PDF and DOCX as page images (via PyMuPDF + LibreOffice)
3. Send page pairs to Gemini Vision for difference detection
4. Correction engine parses AI feedback into concrete formatting fixes
5. Rebuild DOCX with corrections applied
6. Repeat until quality target is met or 3 rounds exhausted

**Token usage:** ~18,000 tokens per conversion (multi-round image comparison)

### Strategy B — Pre-Build AI-Guided Layout ⭐ Recommended

```
PDF ──▶ AI Analyze Pages ──▶ Extract Formatting Overrides ──▶ Build DOCX (with overrides)
```

1. Render PDF pages as images
2. Send to Gemini Vision for layout analysis in a **single call**
3. AI returns structured formatting overrides (margins, font sizes, spacing, alignment)
4. Standard extraction pipeline runs with overrides applied during build
5. Result is produced in one pass — no iterative loop

**Token usage:** ~2,000–3,000 tokens per conversion (single analysis call)

### Strategy Comparison

| Aspect | Strategy A | Strategy B |
|---|---|---|
| **Approach** | Post-build correction loop | Pre-build AI-guided layout |
| **Rounds** | Up to 3 iterations | 1 pass |
| **Token Usage** | ~18K tokens | ~2–3K tokens |
| **Speed** | Slower (multiple renders) | Faster (single AI call) |
| **Accuracy** | Iteratively improves | Good first-pass accuracy |
| **Best For** | Complex layouts needing fine-tuning | Most documents |
| **CLI Flag** | `--ai-strategy A` | `--ai-strategy B` |
| **Recommended** | — | ⭐ Yes |

---

## 📈 Quality Metrics — SSIM Scoring

Visual validation uses **Structural Similarity Index (SSIM)** to quantify how closely the output DOCX matches the original PDF:

| SSIM Score | Quality | Interpretation |
|---|---|---|
| **0.95 – 1.00** | 🟢 Excellent | Near-identical rendering |
| **0.85 – 0.94** | 🟡 Good | Minor differences (spacing, font substitution) |
| **0.70 – 0.84** | 🟠 Fair | Noticeable layout differences |
| **< 0.70** | 🔴 Poor | Significant structural differences |

**How it works:**

1. Render each PDF page to PNG at 150 DPI (via PyMuPDF)
2. Convert DOCX → PDF via LibreOffice, then render to PNG
3. Compute per-page SSIM using `scikit-image`
4. Report mean score across all pages

---

## 🌐 Web UI

Start the web interface:

```bash
python -m web.app
# Server starts at http://localhost:5000
```

### Features

- **Drag-and-Drop Upload** — Drop a PDF (up to 50 MB) onto the page
- **Real-Time Progress** — Server-Sent Events (SSE) stream conversion status live
- **AI Strategy Selector** — Choose between Strategy A and Strategy B from the UI
- **Side-by-Side Preview** — Compare original PDF pages with converted DOCX pages
- **Visual Diff Report** — SSIM scores and difference heatmaps per page
- **Conversion History** — Browse and re-download previous conversions
- **One-Click Download** — Download the converted `.docx` file

---

## ⌨️ CLI Reference

```
usage: pdf2docx [-h] [--ocr] [--password PASSWORD] [--config CONFIG]
                [--verbose] [--validate] [--batch BATCH]
                [--output-dir OUTPUT_DIR] [--skip-watermarks] [--force]
                [--visual-validate] [--ai-compare] [--ai-strategy {A,B}]
                [input] [output]
```

### Arguments

| Argument | Description |
|---|---|
| `input` | Input PDF file path |
| `output` | Output `.docx` file path (default: same name as input with `.docx` extension) |

### Options

| Flag | Description |
|---|---|
| `--ocr` | Enable OCR for scanned / image-based pages (requires Tesseract) |
| `--password PASSWORD` | Password for encrypted PDFs |
| `--config CONFIG` | Path to a custom configuration JSON file |
| `-v`, `--verbose` | Enable verbose (DEBUG) logging |
| `--validate` | Run validation checks on the output document |
| `--batch BATCH` | Directory containing PDF files for batch conversion |
| `--output-dir DIR` | Output directory for batch conversion results |
| `--skip-watermarks` | Attempt to skip watermark / background images |
| `--force` | Overwrite existing output files without prompting |
| `--visual-validate` | Render PDF and DOCX, then compute SSIM visual similarity scores |
| `--ai-compare` | Use Gemini AI to detect and fix visual differences (requires `GEMINI_API_KEY`) |
| `--ai-strategy {A,B}` | AI strategy: `A` = post-build correction loop, `B` = pre-build AI-guided layout (default: `A`) |

### Examples

```bash
# Basic conversion
python pdf2docx.py document.pdf

# Specify output path
python pdf2docx.py document.pdf output.docx

# Convert a scanned PDF with OCR
python pdf2docx.py scanned.pdf --ocr

# Full quality pipeline with AI Strategy B (recommended)
python pdf2docx.py report.pdf --visual-validate --ai-compare --ai-strategy B

# Post-build AI correction loop (Strategy A)
python pdf2docx.py report.pdf --visual-validate --ai-compare --ai-strategy A

# Batch convert all PDFs in a directory
python pdf2docx.py --batch ./pdfs/ --output-dir ./output/

# Encrypted PDF
python pdf2docx.py secure.pdf --password "s3cret"

# Verbose output with validation
python pdf2docx.py input.pdf --verbose --validate
```

---

## ⚙️ Configuration

All settings are stored in `config/settings.json`:

```json
{
  "page_size": "auto",
  "ocr_enabled": false,
  "ocr_language": "eng",
  "preserve_fonts": true,
  "image_quality": "original",
  "table_detection": "auto",
  "heading_detection": "both",
  "skip_watermarks": false,
  "verbose": false,
  "fallback_font": "Arial",
  "max_image_dpi": 300,
  "skip_ocr_if_no_tesseract": true
}
```

### Settings Reference

| Setting | Type | Default | Description |
|---|---|---|---|
| `page_size` | string | `"auto"` | Page size detection mode. `"auto"` reads from PDF metadata |
| `ocr_enabled` | bool | `false` | Enable Tesseract OCR for scanned pages |
| `ocr_language` | string | `"eng"` | Tesseract language code (e.g., `"eng"`, `"fra"`, `"deu"`) |
| `preserve_fonts` | bool | `true` | Attempt to map PDF fonts to system fonts |
| `image_quality` | string | `"original"` | Image quality mode: `"original"`, `"high"`, `"medium"`, `"low"` |
| `table_detection` | string | `"auto"` | Table detection mode: `"auto"`, `"strict"`, `"none"` |
| `heading_detection` | string | `"both"` | Heading detection: `"both"`, `"font_size"`, `"bold"`, `"none"` |
| `skip_watermarks` | bool | `false` | Attempt to skip watermark / background images |
| `verbose` | bool | `false` | Enable verbose logging |
| `fallback_font` | string | `"Arial"` | Font used when PDF font cannot be mapped |
| `max_image_dpi` | int | `300` | Maximum DPI for extracted images |
| `skip_ocr_if_no_tesseract` | bool | `true` | Silently skip OCR if Tesseract is not installed |

### AI Configuration

AI settings can be provided via environment variables or in `config/settings.json` under an `"ai_comparison"` key:

| Setting | Env Variable | Default | Description |
|---|---|---|---|
| `api_key` | `GEMINI_API_KEY` | — | Google Gemini API key |
| `model` | — | `gemini-2.0-flash` | Gemini model to use |
| `strategy` | — | `B` | Default AI strategy (`A` or `B`) |
| `max_rounds` | — | `3` | Maximum correction rounds (Strategy A only) |
| `max_tokens_per_conversion` | — | `50000` | Token budget per conversion |

---

## 📦 Dependencies

### Python Packages

| Package | Version | Purpose |
|---|---|---|
| **PyMuPDF** | ≥ 1.23.0 | PDF rendering and text/image extraction |
| **python-docx** | ≥ 1.1.0 | DOCX document creation and manipulation |
| **pypdf** | ≥ 4.0.0 | PDF utilities and metadata |
| **Pillow** | ≥ 10.0.0 | Image processing and compositing |
| **tqdm** | ≥ 4.66.0 | Progress bars for CLI |
| **Flask** | ≥ 3.0.0 | Web interface server |
| **scikit-image** | ≥ 0.24.0 | SSIM visual similarity computation |
| **google-genai** | ≥ 1.0.0 | Gemini AI API (**new SDK**, not `google-generativeai`) |

### External Tools

| Tool | Purpose | Required? |
|---|---|---|
| **LibreOffice** | DOCX → PDF rendering for visual diff | Optional |
| **Tesseract OCR** | Text recognition in scanned PDFs | Optional |

Install all Python packages: `pip install -r requirements.txt`

---

## 🔬 Module Details

### Extractors

| Module | Responsibility |
|---|---|
| **text_extractor.py** | Extracts text blocks with full formatting metadata (font family, size, color, weight, style, decoration) per page using PyMuPDF |
| **image_extractor.py** | Extracts embedded images, composites SMask alpha channels onto white backgrounds, calculates display dimensions from PDF bounding boxes, and merges overlapping images |
| **table_extractor.py** | Detects tables via ruled-line analysis, handles column collapsing for merged cells, extracts cell shading/background colors, row heights, and cell padding |
| **metadata_extractor.py** | Extracts per-page dimensions, rotation, crop boxes, and document-level metadata |

### Analyzers

| Module | Responsibility |
|---|---|
| **font_analyzer.py** | Maps PDF font names to system-available fonts, detects bold/italic variants, builds a font substitution table |
| **layout_analyzer.py** | Detects multi-column layouts, reading order, and content regions |
| **semantic_analyzer.py** | Orders elements semantically, detects headers/footers (repeating text across pages), inserts page breaks with correct orientation |

### Builders

| Module | Responsibility |
|---|---|
| **docx_builder.py** | Constructs the final DOCX: creates sections with correct page geometry, adds paragraphs with exact spacing, inserts tables with shading/borders, places images at correct sizes, applies AI formatting overrides (Strategy B) |

### Quality

| Module | Responsibility |
|---|---|
| **visual_diff.py** | Renders PDF and DOCX to PNG, computes per-page SSIM scores, generates difference heatmaps |
| **ai_comparator.py** | Sends rendered page pairs to Gemini Vision, parses structured difference reports |
| **ai_layout_analyzer.py** | Sends PDF page images to Gemini for pre-build layout analysis (Strategy B), returns formatting overrides |
| **correction_engine.py** | Parses AI comparator feedback into concrete DOCX modifications (spacing, font size, margins, alignment) for Strategy A rebuild cycles |

### Utilities

| Module | Responsibility |
|---|---|
| **ocr_handler.py** | Wraps Tesseract OCR, detects image-only pages, extracts text with bounding boxes |
| **pdf_info.py** | Quick PDF metadata dump (page count, dimensions, encryption status) |
| **progress.py** | Progress bar tracking for multi-stage pipeline |
| **validator.py** | Post-conversion validation: checks page count, text coverage, image count, and overall quality scoring |

---

## 🗺️ Roadmap

### Known Limitations

- Per-page margins currently only applied to the first section (page 0)
- Mixed page sizes (different dimensions per page) not fully supported
- Cross-page table continuity (tables split across pages)
- Headers/footers not carried across section breaks
- Table border styles (thick/thin/colored/none) not fully differentiated

### Planned Improvements

- [ ] **Per-page geometry** — Apply unique margins and page sizes to each section
- [ ] **Cross-page tables** — Detect and merge tables that span page boundaries
- [ ] **Advanced borders** — Thick, thin, colored, and invisible border detection
- [ ] **Floating images** — Anchored / floating image positioning
- [ ] **Page backgrounds** — Background colors and watermark layers
- [ ] **RTL & CJK support** — Right-to-left text and CJK font handling
- [ ] **Vector graphics** — SVG / EMF vector content preservation
- [ ] **Embedded fonts** — Extract and embed PDF fonts into DOCX
- [ ] **Form fields** — Interactive form field conversion
- [ ] **Annotations** — Comment and markup preservation
- [ ] **Digital signatures** — Signature field handling
- [ ] **Render caching** — Cache rendered pages for faster re-processing
- [ ] **Parallel processing** — Multi-threaded per-page extraction and rendering
- [ ] **Desktop application** — Standalone GUI (Electron / Tauri)
- [ ] **Cloud service** — Hosted API with queue-based processing

---

## ❓ Troubleshooting

<details>
<summary><b>LibreOffice not found</b></summary>

**Symptom:** `--visual-validate` fails with "LibreOffice not found" error.

**Fix:** Install LibreOffice and ensure `soffice` is on your system PATH:
```bash
# Verify installation
soffice --version
```
On Windows, you may need to add the LibreOffice program directory to your PATH:
```
C:\Program Files\LibreOffice\program\
```
</details>

<details>
<summary><b>Tesseract not found</b></summary>

**Symptom:** `--ocr` fails with "Tesseract not found" error.

**Fix:** Install Tesseract and ensure `tesseract` is on your PATH. If you don't need OCR, the converter will silently skip it when `skip_ocr_if_no_tesseract` is `true` (default).
</details>

<details>
<summary><b>GEMINI_API_KEY not set</b></summary>

**Symptom:** `--ai-compare` fails with an API key error.

**Fix:** Set the environment variable or add it to `config/settings.json`:
```json
{
  "ai_comparison": {
    "api_key": "your_key_here"
  }
}
```
Get a key from [Google AI Studio](https://aistudio.google.com/apikey).
</details>

<details>
<summary><b>Low SSIM scores</b></summary>

**Symptom:** Visual validation reports scores below 0.70.

**Possible causes:**
- Complex multi-column layouts → Try `--ai-compare --ai-strategy B`
- Missing system fonts → Install fonts that match the PDF, or adjust `fallback_font` in config
- Scanned PDF without OCR → Add `--ocr` flag
- Tables with complex borders → Known limitation; improvements planned
</details>

<details>
<summary><b>Port 5000 already in use</b></summary>

**Symptom:** Web server fails to start with "Address already in use."

**Fix:**
```bash
# Find and kill the process using port 5000
# Windows
netstat -ano | findstr :5000
taskkill /PID <pid> /F

# Linux / macOS
lsof -i :5000
kill -9 <pid>
```
</details>

<details>
<summary><b>Import errors after installation</b></summary>

**Symptom:** `ModuleNotFoundError` when running.

**Fix:** Ensure you're in the virtual environment and dependencies are installed:
```bash
# Activate venv first, then:
pip install -r requirements.txt
```
> **Note:** This project uses `google-genai` (the new Gemini SDK), **not** `google-generativeai` (legacy). Make sure you have the correct package installed.
</details>

---

## 🧪 Testing

```bash
# Run the test suite
python -m pytest tests/ -v

# Run with coverage
python -m pytest tests/ --cov=quality --cov-report=term-missing
```

The test suite includes **24 tests** covering the AI comparator, correction engine, and quality pipeline. Tests work without an API key using mock objects.

---

## 🤝 Contributing

Contributions are welcome! Here's how to get started:

1. **Fork** the repository
2. **Create** a feature branch:
   ```bash
   git checkout -b feature/your-feature-name
   ```
3. **Make** your changes with clear, descriptive commits
4. **Test** your changes:
   ```bash
   python -m pytest tests/ -v
   ```
5. **Submit** a Pull Request with a clear description of what you changed and why

### Development Guidelines

- Follow PEP 8 style conventions
- Add type hints to all function signatures
- Write docstrings for public functions and classes
- Add tests for new features in `tests/`
- Keep the pipeline modular — extractors, analyzers, builders, and quality modules should remain independent
- Use `logging` instead of `print()` for diagnostic output

---

## 📜 License

This project is licensed under the **MIT License**. See the [LICENSE](LICENSE) file for details.

---

## 🙏 Acknowledgments

- **[PyMuPDF (fitz)](https://pymupdf.readthedocs.io/)** — PDF rendering and text extraction
- **[python-docx](https://python-docx.readthedocs.io/)** — DOCX document creation
- **[Google Gemini](https://ai.google.dev/)** — AI-powered visual comparison and layout analysis
- **[scikit-image](https://scikit-image.org/)** — SSIM computation for visual quality scoring
- **[Flask](https://flask.palletsprojects.com/)** — Web interface framework
- **[LibreOffice](https://www.libreoffice.org/)** — DOCX → PDF rendering for validation
- **[Tesseract OCR](https://github.com/tesseract-ocr/tesseract)** — Optical character recognition for scanned PDFs
- **[Pillow](https://pillow.readthedocs.io/)** — Image processing and compositing

---

<p align="center">
  <b>Phase 6 In Progress</b> — Image transparency fixed, polish items underway.<br/>
  See <a href="PLAN.md">PLAN.md</a> for the full project roadmap.
</p>
