# Quick Reference - Python-pptx

## Color Palette

### Primary Colors
```python
from pptx.dml.color import RGBColor

BLUE = RGBColor(68, 114, 196)
DARK_BLUE = RGBColor(31, 78, 121)
LIGHT_BLUE = RGBColor(173, 216, 230)
```

### Status Colors
```python
GREEN = RGBColor(79, 98, 40)
LIGHT_GREEN = RGBColor(146, 208, 80)
RED = RGBColor(192, 0, 0)
LIGHT_RED = RGBColor(255, 102, 102)
YELLOW = RGBColor(255, 192, 0)
ORANGE = RGBColor(255, 153, 51)
```

### Neutral Colors
```python
BLACK = RGBColor(0, 0, 0)
WHITE = RGBColor(255, 255, 255)
GRAY = RGBColor(89, 89, 89)
LIGHT_GRAY = RGBColor(179, 179, 179)
DARK_GRAY = RGBColor(51, 51, 51)
```

### Accent Colors
```python
PURPLE = RGBColor(112, 48, 160)
PINK = RGBColor(216, 0, 115)
TEAL = RGBColor(0, 176, 240)
```

## Font Sizes

```python
from pptx.util import Pt

TITLE = Pt(44)
SUBTITLE = Pt(32)
HEADING = Pt(24)
BODY = Pt(18)
SMALL = Pt(14)
CAPTION = Pt(10)
```

## Web-Safe Fonts

```python
# Primary
ARIAL = "Arial"
HELVETICA = "Helvetica"

# Serif
TIMES = "Times New Roman"
GEORGIA = "Georgia"

# Monospace
COURIER = "Courier New"

# Sans-serif alternatives
VERDANA = "Verdana"
```

## Standard Slide Layouts (16:9)

```python
from pptx.util import Inches

# 16:9 Aspect Ratio
SLIDE_WIDTH = Inches(10)
SLIDE_HEIGHT = Inches(5.625)

# Margins
MARGIN_LEFT = Inches(0.5)
MARGIN_RIGHT = Inches(9.5)
MARGIN_TOP = Inches(0.5)
MARGIN_BOTTOM = Inches(5.125)

# Common Positions
TITLE_Y = Inches(0.5)
CONTENT_Y = Inches(1.8)
CENTER_X = Inches(5)
CENTER_Y = Inches(2.8125)
```

## Text Alignment

```python
from pptx.enum.text import PP_ALIGN

LEFT = PP_ALIGN.LEFT
CENTER = PP_ALIGN.CENTER
RIGHT = PP_ALIGN.RIGHT
JUSTIFY = PP_ALIGN.JUSTIFY
DISTRIBUTE = PP_ALIGN.DISTRIBUTE
```

## Common Measurements

```python
from pptx.util import Inches, Pt, Cm, Mm

# Common widths
FULL_WIDTH = Inches(9)
HALF_WIDTH = Inches(4.5)
THIRD_WIDTH = Inches(3)
QUARTER_WIDTH = Inches(2.25)

# Common heights
TITLE_HEIGHT = Inches(1.5)
CONTENT_HEIGHT = Inches(3.5)
ROW_HEIGHT = Inches(0.6)

# Spacing
GAP_SMALL = Inches(0.1)
GAP_MEDIUM = Inches(0.3)
GAP_LARGE = Inches(0.7)
```

## Font Styles

```python
# Font size
font.size = Pt(24)

# Bold/Italic
font.bold = True
font.italic = True

# Underline
font.underline = True

# Color
font.color.rgb = RGBColor(0, 0, 0)

# Font name (use web-safe only)
font.name = "Arial"
```

## Shape Types (Common)

```python
from pptx.enum.shapes import MSO_SHAPE

RECTANGLE = MSO_SHAPE.RECTANGLE  # 1
ROUNDED_RECTANGLE = MSO_SHAPE.ROUNDED_RECTANGLE  # 2
OVAL = MSO_SHAPE.OVAL  # 9
DIAMOND = MSO_SHAPE.DIAMOND  # 4
TRIANGLE = MSO_SHAPE.ISOSCELES_TRIANGLE  # 12
```

## Quick Templates

### Title Box
```python
title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
title_box.text_frame.text = "Title Here"
title_box.text_frame.paragraphs[0].font.size = Pt(32)
title_box.text_frame.paragraphs[0].font.bold = True
title_box.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
```

### Bullet List
```python
y = Inches(1.8)
for item in items:
    box = slide.shapes.add_textbox(Inches(1), y, Inches(8), Inches(0.5))
    box.text_frame.text = f"• {item}"
    box.text_frame.paragraphs[0].font.size = Pt(18)
    y = y + Inches(0.6)
```

### Two Columns
```python
# Left column
left_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.8), Inches(4.5), Inches(3.5))
left_box.text_frame.text = "\n".join(f"• {item}" for item in left_items)
left_box.text_frame.paragraphs[0].font.size = Pt(16)

# Right column
right_box = slide.shapes.add_textbox(Inches(5), Inches(1.8), Inches(4.5), Inches(3.5))
right_box.text_frame.text = "\n".join(f"• {item}" for item in right_items)
right_box.text_frame.paragraphs[0].font.size = Pt(16)
```

## Conversion Reference

```
1 inch = 2.54 cm = 25.4 mm
1 cm = 0.394 inches
1 point (pt) = 1/72 inch ≈ 0.0139 inches

Font sizes in points:
12 pt ≈ small body text
14 pt ≈ large body text
18 pt ≈ standard presentation text
24 pt ≈ headers
32 pt ≅ subtitles
44 pt ≅ titles
```

## Common Pitfalls

✅ **DO** - Use `RGBColor(r, g, b)` 
❌ **DON'T** - Use tuples `(r, g, b)` or hex strings

✅ **DO** - Use `Inches()`, `Pt()`, `Cm()`
❌ **DON'T** - Use raw numbers

✅ **DO** - Check box dimensions for text overflow
❌ **DON'T** - Assume text will fit

✅ **DO** - Use web-safe fonts
❌ **DON'T** - Use custom fonts

✅ **DO** - Use blank layout `[6]` for control
❌ **DON'T** - Rely on default layouts
