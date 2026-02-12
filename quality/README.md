# Quality Module - AI-Powered Document Comparison

This module provides AI-powered visual comparison and automatic correction for PDF to DOCX conversions.

## Overview

The quality module consists of three main components:

1. **VisualDiff** - Renders PDF and DOCX to images and computes structural similarity (SSIM)
2. **AIComparator** - Uses Google Gemini AI to detect visual differences between documents
3. **CorrectionEngine** - Automatically applies fixes based on detected differences

## Architecture

```
quality/
├── visual_diff.py       - Visual comparison using SSIM
├── ai_comparator.py     - AI-powered difference detection
└── correction_engine.py - Automatic correction application
```

## Components

### 1. VisualDiff

Provides pixel-perfect visual comparison between PDF and DOCX documents.

**Features:**
- Renders PDF pages using PyMuPDF at configurable DPI
- Converts DOCX to PDF using LibreOffice, then renders to images
- Computes Structural Similarity Index (SSIM) per page
- Generates diff overlay images highlighting differences
- Handles page count mismatches with overflow penalties

**Usage:**

```python
from quality.visual_diff import VisualDiff

vdiff = VisualDiff(dpi=150, libreoffice_path="auto")
result = vdiff.compare("input.pdf", "output.docx", "/tmp/output_dir")

print(f"Overall SSIM: {result['overall_score']:.1%}")
print(f"Quality Level: {result['quality_level']}")  # green/yellow/red
for i, score in enumerate(result['page_scores']):
    print(f"  Page {i+1}: {score:.1%}")
```

**Requirements:**
- PyMuPDF (for PDF rendering)
- scikit-image (for SSIM computation)
- Pillow (for image manipulation)
- LibreOffice (for DOCX → PDF conversion)

**Quality Thresholds:**
- **Green**: SSIM ≥ 95% (excellent match)
- **Yellow**: SSIM ≥ 85% (good match)
- **Red**: SSIM < 85% (needs improvement)

### 2. AIComparator

Uses Google Gemini vision AI to intelligently detect differences between document pages.

**Features:**
- Sends rendered page images to Gemini API for analysis
- Returns structured differences with type, severity, area, and issue description
- Tracks token usage for cost monitoring
- Gracefully degrades when API key not available
- Supports configurable models and parameters

**Usage:**

```python
from quality.ai_comparator import AIComparator

# Initialize with API key
comparator = AIComparator(
    api_key="your_gemini_api_key",
    model="gemini-2.0-flash",
    max_tokens=8192
)

# Compare page images
all_diffs = comparator.compare_pages(
    pdf_images=["pdf_page_0.png", "pdf_page_1.png"],
    docx_images=["docx_page_0.png", "docx_page_1.png"]
)

# Process results
for page_num, page_diffs in enumerate(all_diffs):
    print(f"\nPage {page_num + 1}: {len(page_diffs)} differences")
    for diff in page_diffs:
        print(f"  [{diff['severity']}] {diff['area']}: {diff['issue']}")
```

**Difference Types:**
- `font_size` - Font size mismatch
- `font_family` - Font family/typeface difference
- `font_color` - Font color difference
- `bold` - Bold formatting difference
- `italic` - Italic formatting difference
- `underline` - Underline formatting difference
- `alignment` - Text/cell alignment difference
- `spacing` - Paragraph spacing difference
- `border` - Table border difference
- `shading` - Cell shading/background difference
- `image` - Image position/size difference
- `layout` - Overall layout difference
- `missing_content` - Content present in PDF but missing in DOCX
- `extra_content` - Content in DOCX not in PDF

**Configuration:**

```yaml
# config.yaml
ai_comparison:
  enabled: false             # Requires API key
  provider: gemini
  model: gemini-2.0-flash
  max_rounds: 3              # Max correction iterations
  max_tokens_per_conversion: 50000
  api_key: ${GEMINI_API_KEY}  # Or set via environment variable
```

**Environment Variable:**

```bash
export GEMINI_API_KEY="your_api_key_here"
```

### 3. CorrectionEngine

Automatically applies fixes based on AI-detected differences.

**Features:**
- Maps difference types to correction handlers
- Modifies DOCX using python-docx
- Applies fixes for font, alignment, spacing, borders, and more
- Validates fixes and saves modified document
- Logs all applied corrections

**Usage:**

```python
from quality.correction_engine import CorrectionEngine

engine = CorrectionEngine("output.docx")

# Apply fixes based on detected differences
differences = [
    {
        "type": "font_size",
        "issue": "font size too small",
        "area": "header row",
        "severity": "high"
    },
    {
        "type": "border",
        "issue": "missing cell borders",
        "area": "table",
        "severity": "medium"
    }
]

num_fixed = engine.apply_fixes(differences)
print(f"Applied {num_fixed} fixes")
```

**Supported Fix Types:**
- ✅ Font size adjustment (increase/decrease)
- ✅ Alignment changes (left/center/right/justify)
- ✅ Paragraph spacing adjustment
- ✅ Table border addition
- ✅ Bold formatting toggle
- ✅ Italic formatting toggle
- ⚠️ Cell shading (requires color info)
- ⚠️ Font color (requires RGB values)
- ❌ Font family (complex - requires font matching)
- ❌ Image repositioning (requires re-extraction)

## Complete Workflow

The complete AI comparison workflow is integrated into the main conversion pipeline:

```python
# Stage 1-5: Normal PDF → DOCX conversion
# ...

# Stage 6: Visual Diff
vdiff = VisualDiff(dpi=150)
vd_result = vdiff.compare(pdf_path, docx_path, output_dir)

if vd_result['overall_score'] < 0.95:
    # Stage 7: AI Comparison + Correction Loop
    comparator = AIComparator(api_key=os.environ['GEMINI_API_KEY'])
    
    for round_num in range(1, max_rounds + 1):
        # Detect differences
        all_diffs = comparator.compare_pages(
            vd_result['pdf_images'],
            vd_result['docx_images']
        )
        
        # Apply corrections
        flat_diffs = [d for page_diffs in all_diffs for d in page_diffs]
        engine = CorrectionEngine(docx_path)
        fixes = engine.apply_fixes(flat_diffs)
        
        if fixes == 0:
            break
        
        # Re-render and check if target reached
        vd_result = vdiff.compare(pdf_path, docx_path, output_dir)
        if vd_result['overall_score'] >= 0.95:
            break
```

## CLI Usage

```bash
# Basic conversion with visual validation
python pdf2docx.py input.pdf --visual-validate

# With AI comparison and auto-correction (requires API key)
export GEMINI_API_KEY="your_key"
python pdf2docx.py input.pdf --visual-validate --ai-compare

# Output includes:
# - output.docx (converted document)
# - output_visual_diff/ (rendered pages and diff images)
# - Quality scores and applied corrections in console output
```

## Testing

Comprehensive test suite included:

```bash
# Run all quality module tests
python tests/test_ai_comparison.py

# Test coverage:
# - AIComparator: 10 tests (initialization, parsing, API key handling)
# - CorrectionEngine: 13 tests (all fix types, error handling)
# - Integration: 2 tests (end-to-end flow, graceful degradation)
```

## Dependencies

```txt
# Core (required)
PyMuPDF>=1.23.0         # PDF rendering
python-docx>=1.1.0      # DOCX manipulation
Pillow>=10.0.0          # Image processing

# Visual diff (required for --visual-validate)
scikit-image>=0.24.0    # SSIM computation

# AI comparison (optional, required for --ai-compare)
google-generativeai>=0.8.0  # Gemini API
```

## Cost Considerations

AI comparison uses the Gemini API which has usage costs:

**Token Usage Guidelines:**
- ~500-1000 tokens per page comparison (with gemini-2.0-flash)
- Typical 10-page document: ~5,000-10,000 tokens
- Max 3 rounds per conversion (configurable)

**Cost Guards:**
- Token usage is logged for each API call
- Configurable `max_tokens_per_conversion` limit
- Automatic termination when SSIM > 95% or no fixes applied

**Optimization Tips:**
1. Use `gemini-2.0-flash` for faster, cheaper comparisons
2. Set `max_rounds` to 2-3 for most documents
3. Only enable AI comparison for high-value conversions
4. Use visual diff alone for quick quality checks

## Error Handling

The module is designed for graceful degradation:

**Without API key:**
- AI comparison is automatically skipped
- Visual diff still works (SSIM-only scoring)
- Warning logged, no errors thrown

**Without LibreOffice:**
- DOCX rendering fails gracefully
- Warning logged, returns empty image list
- Install instructions provided in log

**API failures:**
- Individual page failures don't stop processing
- Errors logged, empty results returned for failed pages
- Remaining pages continue processing

**Invalid differences:**
- Non-dict entries skipped with warning
- Missing 'type' field skipped with warning
- Unknown types logged but don't crash

## Limitations

Current limitations and future improvements:

**Correction Engine Limitations:**
1. Cannot determine exact colors without AI providing RGB values
2. Font family changes require complex font matching
3. Image repositioning requires re-extraction from PDF
4. Complex layout changes not supported

**Visual Diff Limitations:**
1. Requires LibreOffice for DOCX rendering
2. SSIM doesn't capture semantic differences
3. Anti-aliasing differences can affect scores
4. Page size differences need special handling

**AI Comparison Limitations:**
1. Requires paid Gemini API key
2. Token costs scale with document size
3. Model may miss very subtle differences
4. Response quality depends on image clarity

## Performance

**Typical Performance:**
- Visual diff: ~2-5 seconds per page (rendering + SSIM)
- AI comparison: ~1-2 seconds per page (API latency)
- Correction: ~0.1 seconds per fix
- End-to-end (10 pages, 3 rounds): ~1-2 minutes

**Optimization Opportunities:**
- [ ] Cache rendered images between rounds
- [ ] Parallelize page processing
- [ ] Batch API calls for multiple pages
- [ ] Optimize DOCX save (only save if changed)

## Contributing

To add a new correction handler:

1. Add the handler method to `CorrectionEngine`:
   ```python
   def _fix_new_type(self, diff: dict[str, Any]) -> None:
       """Fix description."""
       # Implementation
       self._fixes_applied += 1
   ```

2. Register it in `_handlers`:
   ```python
   _handlers = {
       # ...
       "new_type": _fix_new_type,
   }
   ```

3. Add test in `tests/test_ai_comparison.py`:
   ```python
   def test_fix_new_type(self):
       """Test new type fix."""
       # Test implementation
   ```

## License

Part of the PDFtoWord project. See main README for license information.
