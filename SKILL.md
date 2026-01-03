---
name: simple-pptx-skill
description: Create PowerPoint presentations using python-pptx for automated reports, data-driven slides, and simple layouts
dependencies: python-pptx
---

## Overview

Automated .pptx creation using python-pptx library. Ideal for data-driven reports, simple layouts, and Python-only environments.

## When to Apply

**Use when**:
- Creating automated reports from data
- Simple text/shape layouts required
- Python-only environment constraints
- Performance-critical scenarios
- Programmatic control over elements needed

**Do NOT use when**:
- Complex HTML/CSS styling needed (use html2pptx)
- Rich media/graphics are primary focus

## Constraints & Limitations

- Text overflow is silently truncated
- Only web-safe fonts supported
- No HTML/CSS rendering
- Basic shapes only (no complex graphics)
- Requires python-pptx dependency

## Key Rules

✅ Use `RGBColor(r,g,b)` never tuples
✅ Use `Inches()`/`Pt()` never raw numbers
✅ Use blank layout `prs.slide_layouts[6]`
✅ Check box dimensions for overflow

## Additional Files

- [PATTERNS.md](PATTERNS.md) - Core code patterns and functions
- [EXAMPLES.md](EXAMPLES.md) - Complete working examples
- [REFERENCE.md](REFERENCE.md) - Colors, fonts, quick reference
- [python-pptx-guide.md](python-pptx-guide.md) - Full API documentation

If you are already familiar with python-pptx, ask the user for permission to skip reading the full documentation to reduce token consumption.

Read PATTERNS.md for implementation patterns, EXAMPLES.md for usage demonstrations, and REFERENCE.md for styling references.