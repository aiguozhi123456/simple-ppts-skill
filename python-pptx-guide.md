# PPTX Simple - Python-pptx Guide
## Overview

Create .pptx presentations using python-pptx library. Ideal for:
- Quick generation without complex HTML rendering
- Automated report generation from data
- Simple text and shape layouts
- Python-only environments
- Performance-critical scenarios
- **Editing existing presentations** (modify content, replace images, etc.)

## Dependencies

Required: `pip install python-pptx`

## Critical Rules

**MUST follow**:
- ✅ Use `RGBColor(r, g, b)` for ALL colors, never tuples `(r, g, b)`
- ✅ Use `Inches()` or `Pt()` for ALL measurements, never raw numbers
- ✅ Use web-safe fonts: Arial, Helvetica, Times New Roman, Georgia, Courier New, Verdana
- ✅ Use blank layout (index 6) for maximum control
- ✅ Check text overflow: Long text will be cut off without warning

## Core API
### Setup

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE_TYPE

# Create new presentation
prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(5.625)  # 16:9

# OR Load existing presentation
prs = Presentation('existing.pptx')
```

### Common Slide Sizes

- 16:9: `Inches(10) × Inches(5.625)`
- 4:3: `Inches(10) × Inches(7.5)`
- 16:10: `Inches(10) × Inches(6.25)`
### Add Slide

```python
blank_layout = prs.slide_layouts[6]  # Index 6 = blank
slide = prs.slides.add_slide(blank_layout)
```

### Delete Slide

```python
# Remove slide by index
slide_idx = 2  # Zero-based index
rid = prs.slides._sldIdLst[slide_idx].rId
prs.part.drop_rel(rid)
del prs.slides._sldIdLst[slide_idx]
```

### Text Box

```python
text_box = slide.shapes.add_textbox(left, top, width, height)
text_frame = text_box.text_frame
text_frame.text = "Text"

# Access paragraph
p = text_frame.paragraphs[0]

# Add paragraph
p2 = text_frame.add_paragraph()
p2.text = "More text"
```

### Text Formatting

```python
p.font.name = "Arial"
p.font.size = Pt(24)
p.font.bold = True
p.font.color.rgb = RGBColor(68, 114, 196)
p.alignment = PP_ALIGN.LEFT  # LEFT, CENTER, RIGHT
p.space_before = Pt(12)
p.space_after = Pt(6)
p.line_spacing = 1.5
```

### Bullet List

```python
text_frame = text_box.text_frame

p = text_frame.add_paragraph()
p.text = "Item 1"
p.level = 0

p2 = text_frame.add_paragraph()
p2.text = "Sub-item"
p2.level = 1

# Add bullet manually if needed: p.text = "• Item"
```

### Shape

```python
# Shape IDs: 1=Rectangle, 9=Oval, 5=Rounded Rectangle, 7=Triangle, 10=Diamond
shape = slide.shapes.add_shape(1, Inches(1), Inches(2), Inches(3), Inches(2))

# Fill
fill = shape.fill
fill.solid()
fill.fore_color.rgb = RGBColor(68, 114, 196)

# Border
line = shape.line
line.color.rgb = RGBColor(0, 0, 0)
line.width = Pt(2)

# Layer order
shape.z_order = 0  # Lower = background
```

### Table

```python
table = slide.shapes.add_table(rows=3, cols=3, left=Inches(1), top=Inches(1), width=Inches(8), height=Inches(2)).table

# Cell content
table.cell(0, 0).text = "Header"
table.cell(1, 0).text = "Data"

# Cell style
cell = table.cell(0, 0)
fill = cell.fill
fill.solid()
fill.fore_color.rgb = RGBColor(68, 114, 196)
cell.text_frame.paragraphs[0].font.bold = True
cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)

# Column widths
table.columns[0].width = Inches(2)
table.columns[1].width = Inches(3)
```

### Image

```python
# Maintain aspect ratio
slide.shapes.add_picture('image.png', Inches(1), Inches(1), width=Inches(8))

# Fixed dimensions (may distort)
slide.shapes.add_picture('image.png', Inches(1), Inches(1), width=Inches(8), height=Inches(4))
```

### Background

```python
background = slide.background
fill = background.fill
fill.solid()
fill.fore_color.rgb = RGBColor(28, 40, 51)
```
## Editing Existing Presentations

### Read and Iterate

```python
prs = Presentation('file.pptx')

# Iterate through slides
for slide_idx, slide in enumerate(prs.slides):
    print(f"Slide {slide_idx + 1}:")
    
    # Iterate through shapes
    for shape_idx, shape in enumerate(slide.shapes):
        print(f"  Shape {shape_idx}: {shape.shape_type}")
        
        # Read text
        if hasattr(shape, 'text'):
            print(f"    Text: {shape.text}")
        
        # Read detailed text properties
        if hasattr(shape, 'text_frame'):
            for para_idx, p in enumerate(shape.text_frame.paragraphs):
                print(f"    Para {para_idx}: {p.text}")
                print(f"      Size: {p.font.size}, Bold: {p.font.bold}")
        
        # Identify image
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            print(f"    Image: {len(shape.blob)} bytes")
```

### Modify Existing Text

```python
prs = Presentation('file.pptx')

# Method 1: Replace all matching text
for slide in prs.slides:
    for shape in slide.shapes:
        if hasattr(shape, 'text_frame'):
            for paragraph in shape.text_frame.paragraphs:
                # Replace text
                if "Old Text" in paragraph.text:
                    paragraph.text = paragraph.text.replace("Old Text", "New Text")

# Method 2: Modify specific shape by index
slide = prs.slides[0]
shape = slide.shapes[0]  # First shape
if hasattr(shape, 'text_frame'):
    shape.text_frame.text = "Completely new text"

# Method 3: Find and modify by text content
for slide in prs.slides:
    for shape in slide.shapes:
        if hasattr(shape, 'text') and "Product" in shape.text:
            shape.text_frame.paragraphs[0].font.bold = True
            shape.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 0, 0)
```

### Delete Shape

```python
prs = Presentation('file.pptx')
slide = prs.slides[0]

# Delete shape by index
shape = slide.shapes[2]  # Third shape
sp = shape._element
sp.getparent().remove(sp)

# Delete by text content matching
for slide in prs.slides:
    shapes_to_remove = []
    for shape in slide.shapes:
        if hasattr(shape, 'text') and "Delete This" in shape.text:
            shapes_to_remove.append(shape)
    
    for shape in shapes_to_remove:
        sp = shape._element
        sp.getparent().remove(sp)
```

### Replace Image

```python
prs = Presentation('file.pptx')
slide = prs.slides[0]

# Find and replace image
for shape in slide.shapes:
    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        # Delete old image
        sp = shape._element
        sp.getparent().remove(sp)
        
        # Add new image at same position
        slide.shapes.add_picture('new_image.png', 
                                shape.left, shape.top, 
                                width=shape.width, 
                                height=shape.height)
        break  # Replace first image only
```

### Add Content to Existing Slide

```python
prs = Presentation('file.pptx')
slide = prs.slides[0]

# Add new text box to existing slide
new_box = slide.shapes.add_textbox(Inches(1), Inches(4), Inches(8), Inches(1))
new_frame = new_box.text_frame
new_frame.text = "Added content"
new_frame.paragraphs[0].font.size = Pt(20)

# Add new image to existing slide
slide.shapes.add_picture('logo.png', Inches(8), Inches(0.5), width=Inches(1.5))
```

### Modify Table Content

```python
prs = Presentation('file.pptx')

for slide in prs.slides:
    for shape in slide.shapes:
        if shape.has_table:
            table = shape.table
            # Modify cell content
            table.cell(0, 0).text = "New Header"
            table.cell(1, 1).text = "New Data"
            
            # Add new row
            row = table.rows.add_row()
            row.cells[0].text = "New Row Cell 1"
            row.cells[1].text = "New Row Cell 2"
```

### Find Shapes by Type

```python
# Find all pictures
pictures = [s for s in slide.shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE]

# Find all text boxes
text_boxes = [s for s in slide.shapes if s.shape_type == MSO_SHAPE_TYPE.TEXT_BOX]

# Find shapes in specific area
shapes_in_area = []
for shape in slide.shapes:
    if (shape.left >= Inches(1) and shape.left <= Inches(5) and
        shape.top >= Inches(1) and shape.top <= Inches(4)):
        shapes_in_area.append(shape)
```

## Common Patterns

### Title Slide

```python
slide = prs.slides.add_slide(blank_layout)

# Background
background = slide.background
fill = background.fill
fill.solid()
fill.fore_color.rgb = RGBColor(28, 40, 51)

# Title
title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2), Inches(9), Inches(1))
title_frame = title_box.text_frame
title_frame.text = "Title"
p = title_frame.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
p.font.size = Pt(48)
p.font.bold = True
p.font.color.rgb = RGBColor(68, 114, 196)

# Subtitle
subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(3.5), Inches(9), Inches(0.8))
subtitle_frame = subtitle_box.text_frame
subtitle_frame.text = "Subtitle"
p2 = subtitle_frame.paragraphs[0]
p2.alignment = PP_ALIGN.CENTER
p2.font.size = Pt(24)
p2.font.color.rgb = RGBColor(255, 255, 255)
```

### Content Slide with List

```python
slide = prs.slides.add_slide(blank_layout)

# Background
background = slide.background
fill = background.fill
fill.solid()
fill.fore_color.rgb = RGBColor(255, 255, 255)

# Title
title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
title_frame = title_box.text_frame
title_frame.text = "Section Title"
title_frame.paragraphs[0].font.size = Pt(36)
title_frame.paragraphs[0].font.bold = True
title_frame.paragraphs[0].font.color.rgb = RGBColor(28, 40, 51)

# List
text_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(8), Inches(2.5))
text_frame = text_box.text_frame
for item in ["Point 1", "Point 2", "Point 3"]:
    p = text_frame.add_paragraph()
    p.text = f"• {item}"
    p.font.size = Pt(24)
    p.font.color.rgb = RGBColor(44, 62, 80)
    p.space_after = Pt(12)
```
### Two-Column Layout

```python
slide = prs.slides.add_slide(blank_layout)

# Left
left_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(4), Inches(3))
left_box.text_frame.text = "Left content"

# Right
right_box = slide.shapes.add_textbox(Inches(5), Inches(1.5), Inches(4.5), Inches(3))
right_box.text_frame.text = "Right content"
```

### Update Report Template

```python
def update_template(template_path, data, output_path):
    """Load template, replace placeholders, save as new file"""
    prs = Presentation(template_path)
    
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, 'text_frame'):
                for paragraph in shape.text_frame.paragraphs:
                    # Replace {{key}} with value
                    for key, value in data.items():
                        placeholder = f"{{{{{key}}}}}"
                        if placeholder in paragraph.text:
                            paragraph.text = paragraph.text.replace(placeholder, value)
    
    prs.save(output_path)

# Usage
data = {
    'date': '2024-01-15',
    'title': 'Quarterly Report',
    'revenue': '$1,234,567'
}
update_template('template.pptx', data, 'report.pptx')
```

### Batch Image Replacement

```python
def replace_images_in_ppt(pptx_path, image_mapping, output_path):
    """Replace images based on position or index"""
    prs = Presentation(pptx_path)
    
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                # Match by position (example: replace image at top-left)
                if shape.left < Inches(2) and shape.top < Inches(2):
                    if 'logo' in image_mapping:
                        sp = shape._element
                        sp.getparent().remove(sp)
                        slide.shapes.add_picture(image_mapping['logo'],
                                                shape.left, shape.top,
                                                width=shape.width,
                                                height=shape.height)
    
    prs.save(output_path)

# Usage
image_mapping = {'logo': 'new_logo.png'}
replace_images_in_ppt('presentation.pptx', image_mapping, 'updated.pptx')
```

## Design Guidelines

### Colors (RGB)

**Blue**: Dark(28,40,51), Medium(44,62,80), Light(68,114,196)
**Green**: Dark(30,80,44), Medium(64,105,91), Light(93,173,226)
**Orange**: Dark(193,57,43), Medium(231,76,60), Light(243,156,18)
**White**: (255,255,255), **Gray**: (128,128,128)

### Font Sizes

Title: 36-48pt, Subtitle: 24-30pt, Body: 18-24pt, Footer: 12-14pt

### Spacing

Margin: Inches(0.5-1.0), Paragraph: Pt(6-12), Line: 1.2-1.5x
## Limitations

**Unsupported**:
- Gradient fills (use image workaround)
- SmartArt (create manually or use images)
- Native charts (use matplotlib + image)
- Complex shape effects
- Master slide editing (read-only)
- Direct slide reordering (workaround: delete and re-add)

**Workaround for gradients**:
```python
from PIL import Image, ImageDraw
img = Image.new('RGB', (width, height))
draw = ImageDraw.Draw(img)
for y in range(height):
    r = int(c1[0] + (c2[0]-c1[0]) * y / height)
    g = int(c1[1] + (c2[1]-c1[1]) * y / height)
    b = int(c1[2] + (c2[2]-c1[2]) * y / height)
    draw.line([(0, y), (width, y)], fill=(r, g, b))
img.save('gradient.png')
slide.shapes.add_picture('gradient.png', 0, 0, width=prs.slide_width)
```

**Workaround for charts**:
```python
import matplotlib.pyplot as plt
fig, ax = plt.subplots()
ax.bar(['A', 'B', 'C'], [10, 20, 15])
plt.savefig('chart.png', dpi=150, bbox_inches='tight')
slide.shapes.add_picture('chart.png', Inches(1), Inches(1), width=Inches(8))
```
## Troubleshooting

**Text cut off**: Increase height or use `text_frame.auto_size = True`
**Wrong colors**: Always use `RGBColor()` object, not tuple
**Overlapping shapes**: Adjust positions with `Inches()` or use `z_order`
**Table distorted**: Set explicit `table.columns[i].width` and `table.rows[j].height`
**File corrupted**: Always save to a new path when editing existing files
**Shape not found**: Check shape_type and use has_table/has_chart properties
**Text replacement not working**: Check if text is in text_frame or directly in shape.text

## Best Practices

1. **Always save to new path**: `prs.save('new_file.pptx')` when editing
2. **Check shape type before access**: Use `hasattr()` or `shape_type`
3. **Use constants**: `MSO_SHAPE_TYPE`, `MSO_SHAPE`, `PP_ALIGN`, etc.
4. **Test text length**: Use `auto_size = True` for variable content
5. **Back up originals**: Keep copies before bulk modifications
6. **Iterate carefully**: When deleting shapes, collect first then delete

## Complete Example: Edit Existing Presentation

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE

def update_presentation(input_path, output_path):
    # Load existing
    prs = Presentation(input_path)
    
    # Update title on first slide
    slide = prs.slides[0]
    for shape in slide.shapes:
        if hasattr(shape, 'text') and 'PRODUCT' in shape.text.upper():
            shape.text = shape.text.replace('PRODUCT', 'NEW PRODUCT')
            shape.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 0, 0)
    
    # Replace image on slide 3
    if len(prs.slides) > 2:
        slide = prs.slides[2]
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                pos = (shape.left, shape.top)
                size = (shape.width, shape.height)
                sp = shape._element
                sp.getparent().remove(sp)
                slide.shapes.add_picture('new_image.jpg', 
                                        pos[0], pos[1],
                                        width=size[0], height=size[1])
                break
    
    # Add new slide at end
    new_slide = prs.slides.add_slide(prs.slide_layouts[6])
    title_box = new_slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(1))
    title_box.text_frame.text = "Added Slide"
    title_box.text_frame.paragraphs[0].font.size = Pt(36)
    
    # Save
    prs.save(output_path)
    print(f"Updated: {output_path}")

update_presentation('original.pptx', 'updated.pptx')
```

## Read Existing Presentation

```python
prs = Presentation('file.pptx')
for slide_idx, slide in enumerate(prs.slides):
    for shape in slide.shapes:
        if hasattr(shape, 'text'):
            print(shape.text)
        if hasattr(shape, 'text_frame'):
            for p in shape.text_frame.paragraphs:
                print(f"  {p.text} (size: {p.font.size})")
```