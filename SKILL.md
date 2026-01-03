---
name: pptx_simple
description: "Lightweight PowerPoint presentation creation using python-pptx library. Ideal for: (1) Quick test presentations, (2) Automated report generation from data, (3) Structured content with simple formatting, (4) Python-only environments. Alternative to html2pptx when HTML rendering complexity is unnecessary."
license: Proprietary. LICENSE.txt has complete terms
---

# PPTX Simple - Python-pptx Guide

## Overview

Create .pptx presentations using python-pptx library. Ideal for:
- Quick generation without complex HTML rendering
- Automated report generation from data
- Simple text and shape layouts
- Python-only environments
- Performance-critical scenarios

## When to Use This Approach

**Use python-pptx when**:
- Quick generation from data
- Simple text and shape layouts
- Python-only environment constraints
- Performance is critical
- Report automation

## Quick Start

### Installation

```bash
pip install python-pptx
```

### Basic Example

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# Create presentation
prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(5.625)  # 16:9

# Add slide
blank_layout = prs.slide_layouts[6]
slide = prs.slides.add_slide(blank_layout)

# Add text box
text_box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(1))
text_frame = text_box.text_frame
text_frame.text = "Hello, World!"

# Add paragraph with formatting
p = text_frame.paragraphs[0]
p.font.size = Pt(36)
p.font.bold = True
p.font.color.rgb = RGBColor(68, 114, 196)
p.alignment = PP_ALIGN.CENTER

# Save
prs.save('presentation.pptx')
```

## Detailed Guide

For complete API reference, patterns, limitations, and troubleshooting, see: [python-pptx-guide.md](python-pptx-guide.md)

## Key Rules

**ALWAYS follow these**:
- ✅ Use `RGBColor(r, g, b)` for ALL colors, never tuples `(r, g, b)`
- ✅ Use `Inches()` or `Pt()` for ALL measurements, never raw numbers
- ✅ Use web-safe fonts: Arial, Helvetica, Times New Roman, Georgia, Courier New, Verdana
- ✅ Use blank layout (index 6) for maximum control
- ✅ Check text overflow: Long text will be cut off without warning

## Code Style

- Write concise code
- Use descriptive function names
- Group related operations into reusable functions
- Test text overflow before final save