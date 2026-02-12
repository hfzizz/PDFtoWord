# PDF → Word Converter — Project Plan

## Status

| Phase | Description | Status |
|-------|-------------|--------|
| 1 | Core MVP — text, images, tables → DOCX | ✅ Done |
| 2 | Enhanced Features — spacing, links, run-level formatting, dedup | ✅ Done |
| 3 | Advanced — multi-column, OCR, headers/footers, rotation, scoring | ✅ Done |
| 4 | Web UI + Quality — Flask app, SSE progress, table collapsing | ✅ Done |
| 5 | Visual Fidelity + AI Comparison — SSIM scoring, Gemini correction | ✅ Done |
| **6** | **Image & Polish — transparency fix done, color accuracy, testing** | **← Current** |
| 7 | Future Improvements | Planned |

### Phase Summaries

**Phase 1–2:** Built the core extraction pipeline (PyMuPDF) and DOCX builder (python-docx) with text blocks, images, tables, hyperlinks, and per-run font styling.

**Phase 3:** Added multi-column layout detection, optional OCR via Tesseract, header/footer extraction, rotated text handling, and a structure-based quality score.

**Phase 4:** Shipped a Flask web UI with drag-and-drop upload, SSE progress streaming, download routes, and fixed table column collapsing and cell-merge bugs.

**Phase 5:** Implemented visual fidelity scoring (SSIM via scikit-image) and two AI correction strategies using Google Gemini. Strategy A is a post-build 3-round correction loop (~18K tokens, +0.7% SSIM). Strategy B is a pre-build single-call AI layout analysis (~2–3K tokens, recommended). Also added exact column widths, content-based page margins, image display sizing, table row heights, cell padding, font colors, cell shading, and a 24-test suite.

---

## Architecture

```
PDFtoWord/
├── pdf2docx.py                # Main orchestrator + CLI
├── config/
│   └── settings.json          # Extraction & AI settings
├── requirements.txt
├── analyzers/
│   ├── font_analyzer.py       # Font mapping & style detection
│   ├── layout_analyzer.py     # Page layout & column detection
│   └── semantic_analyzer.py   # Element ordering & structure
├── builders/
│   └── docx_builder.py        # DOCX construction (python-docx)
├── extractors/
│   ├── image_extractor.py     # Image extraction & placement
│   ├── metadata_extractor.py  # PDF metadata extraction
│   ├── table_extractor.py     # Table detection & column collapsing
│   └── text_extractor.py      # Text block extraction
├── quality/
│   ├── visual_diff.py         # Render + SSIM scoring
│   ├── ai_comparator.py       # Gemini vision comparison (Strategy A)
│   ├── ai_layout_analyzer.py  # Gemini pre-build analysis (Strategy B)
│   └── correction_engine.py   # Auto-fix from AI feedback
├── utils/
│   ├── ocr_handler.py         # Tesseract OCR integration
│   ├── pdf_info.py            # PDF metadata utilities
│   ├── progress.py            # Progress reporting
│   └── validator.py           # Input validation
├── web/
│   ├── app.py                 # Flask application entry
│   ├── routes/                # Upload, download, status endpoints
│   ├── services/              # Converter & file manager
│   ├── templates/             # Jinja2 HTML templates
│   └── static/                # CSS + JS assets
└── tests/
    └── test_ai_comparison.py  # 24 tests for AI components
```

---

## Pipeline

### Strategy A — Post-Build Correction (9 stages)

```
PDF → Analyze (fonts, layout) → Extract (text, images, tables)
    → Build DOCX → Validate structure
    → Render PDF pages (PyMuPDF → PNG)
    → Render DOCX pages (LibreOffice → PDF → PNG)
    → SSIM score
    → AI Compare (Gemini vision, if API key set)
    → Auto-Fix → rebuild → loop (max 3 rounds)
    → DOCX Output + Quality Report + Diff Images
```

### Strategy B — AI-Guided Build (7 stages, recommended)

```
PDF → Render pages (PyMuPDF → PNG)
    → AI Analyze (Gemini → structured layout spec, 1 call/page)
    → Extract (text, images, tables)
    → Build DOCX with AI formatting overrides
    → Render DOCX (LibreOffice → PDF → PNG)
    → SSIM score (validation only, no correction loop)
    → Validate structure
    → DOCX Output + Quality Report + Diff Images
```

| | Strategy A (Post-Patch) | Strategy B (Pre-Build) |
|--|------------------------|----------------------|
| AI calls | 3+ (per round) | 1 (before build) |
| Token cost | ~18K | ~2–3K |
| Render cycles | 6 | 2 |
| SSIM gain | +0.7% | Higher (fixes at source) |
| Time | ~90s | ~30s |

---

## Phase 6 — Image Transparency Fix ← Current

**Goal:** Fix transparent images rendering with black backgrounds in converted DOCX.

### Problem

Logos and icons with transparent backgrounds appear with solid black fills. Three causes:

1. **SMask not composited** — PDF stores transparency as a separate Soft Mask stream. `extract_image(xref)` returns raw bytes without compositing the SMask, so alpha is lost.
2. **RGBA passthrough** — `_ensure_rgb()` passes RGBA through unchanged, but the alpha channel may be inverted or absent, causing black fill.
3. **Word transparency** — DOCX renderers handle PNG alpha inconsistently.

### Fix

| Step | Change | Effect |
|------|--------|--------|
| 1 | Check for SMask xref after `extract_image()` | Detect separate transparency mask |
| 2 | Read mask pixmap → create proper RGBA image | Correct alpha compositing |
| 3 | Composite onto white background (`Image.alpha_composite`) | Eliminate transparency artifacts |
| 4 | Save all images as PNG | Lossless, consistent format |

### Tasks

- [x] **6a** — Fix transparent image extraction (SMask + alpha compositing) ✅
- [ ] **6b** — Handle palette-mode (P) transparency
- [ ] **6c** — Improve CMYK → RGB color accuracy
- [ ] **6d** — Test with various logo types (PNG, JPEG, vector-based)

---

## Phase 7 — Future Improvements

### Layout & Geometry

- **Per-page margins** — Currently only page 0's margins are applied to all sections
- **Per-page geometry** — Different page sizes across the document not fully supported
- **Cross-page table continuity** — Tables spanning pages produce two independent tables
- **Headers/footers inheritance** — Section breaks don't carry over headers/footers
- **Page/section background colors** — Not yet extracted or applied

### Formatting Fidelity

- **Table border detection** — Thick/thin/colored/none border styles
- **Floating/anchored images** — Exact PDF coordinate positioning via anchored placement
- **Text highlight colors** — Annotation and highlight color extraction
- **AI text override collisions** — Same text snippet on different pages overwrites formatting

### Image Quality

- **CMYK → RGB accuracy** — Color profile-aware conversion
- **SVG/vector logo quality** — Vector graphics rasterized at low DPI lose detail
- **Image DPI metadata** — Missing DPI info causes size miscalculation

### Performance

- **Render caching** — Cache page renders to avoid redundant work
- **Parallel page processing** — Process pages concurrently
- **Batch processing** — Multi-file conversion improvements

### Edge Cases

- **RTL text** — Right-to-left language support
- **CJK fonts** — Chinese/Japanese/Korean font handling
- **Embedded fonts** — Font subsetting and fallback
- **Vector graphics** — Preserve as vector where possible
- **Form fields, annotations, digital signatures** — Currently ignored

---

## Config (`config/settings.json`)

```jsonc
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

AI and visual diff settings are configured via CLI flags and environment variables:

- `GEMINI_API_KEY` — API key for Gemini vision (enables AI strategies)
- `--visual-validate` — Enable SSIM scoring after conversion
- `--ai-compare` — Enable AI comparison/analysis
- `--ai-strategy A|B` — Select correction strategy

---

## Dependencies

```
PyMuPDF>=1.23.0          # PDF rendering & extraction
python-docx>=1.1.0       # DOCX generation
pypdf>=4.0.0             # PDF metadata
Pillow>=10.0.0           # Image processing
tqdm>=4.66.0             # Progress bars
flask>=3.0.0             # Web UI
scikit-image>=0.24.0     # SSIM computation
google-genai>=1.0.0      # Gemini API (google-genai SDK)
```

---

## Usage

```bash
# Basic conversion
python pdf2docx.py input.pdf

# With visual validation (SSIM scoring)
python pdf2docx.py input.pdf --visual-validate

# AI-guided build (Strategy B, recommended)
python pdf2docx.py input.pdf --visual-validate --ai-compare --ai-strategy B

# Post-build correction loop (Strategy A)
python pdf2docx.py input.pdf --visual-validate --ai-compare --ai-strategy A

# Web UI
python -m web.app

# Run tests
python tests/test_ai_comparison.py
```
