# Core Patterns - Python-pptx

## 1. Basic Presentation Structure

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# Initialize presentation (16:9 aspect ratio)
prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(5.625)

# Add blank slide for maximum control
slide = prs.slides.add_slide(prs.slide_layouts[6])
```

## 2. Adding Text Boxes

```python
def add_text_box(slide, text, left, top, width, height, font_size=Pt(18), 
                 bold=False, color=RGBColor(0, 0, 0), alignment=None):
    """Add formatted text box to slide"""
    text_box = slide.shapes.add_textbox(left, top, width, height)
    text_frame = text_box.text_frame
    text_frame.text = text
    
    p = text_frame.paragraphs[0]
    p.font.size = font_size
    p.font.bold = bold
    p.font.color.rgb = color
    if alignment:
        p.alignment = alignment
    
    return text_box
```

## 3. Title Slide Pattern

```python
def create_title_slide(prs, title, subtitle=None):
    """Create title slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(1.5))
    title_frame = title_box.text_frame
    title_frame.text = title
    p = title_frame.paragraphs[0]
    p.font.size = Pt(44)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER
    
    # Subtitle (optional)
    if subtitle:
        sub_box = slide.shapes.add_textbox(Inches(1), Inches(3.5), Inches(8), Inches(1))
        sub_frame = sub_box.text_frame
        sub_frame.text = subtitle
        sub_frame.paragraphs[0].font.size = Pt(24)
        sub_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    return slide
```

## 4. Data-Driven Slide Generation

```python
def create_data_slide(prs, title, data_points, start_y=Inches(1.8), 
                     title_color=RGBColor(31, 78, 121), bullet_color=RGBColor(89, 89, 89)):
    """Create slide with bullet points from data"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = title
    title_frame.paragraphs[0].font.size = Pt(32)
    title_frame.paragraphs[0].font.bold = True
    title_frame.paragraphs[0].font.color.rgb = title_color
    
    # Data points
    y_position = start_y
    for point in data_points:
        box = slide.shapes.add_textbox(Inches(1), y_position, Inches(8), Inches(0.5))
        box.text_frame.text = f"• {point}"
        box.text_frame.paragraphs[0].font.size = Pt(18)
        box.text_frame.paragraphs[0].font.color.rgb = bullet_color
        y_position = y_position + Inches(0.7)
    
    return slide
```

## 5. Multi-Column Layout

```python
def create_two_column_slide(prs, title, left_content, right_content):
    """Create slide with two columns"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_box.text_frame.text = title
    title_box.text_frame.paragraphs[0].font.size = Pt(32)
    title_box.text_frame.paragraphs[0].font.bold = True
    
    # Left column
    left_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.8), Inches(4.5), Inches(3.5))
    left_box.text_frame.text = "\n".join(f"• {item}" for item in left_content)
    left_box.text_frame.paragraphs[0].font.size = Pt(16)
    
    # Right column
    right_box = slide.shapes.add_textbox(Inches(5), Inches(1.8), Inches(4.5), Inches(3.5))
    right_box.text_frame.text = "\n".join(f"• {item}" for item in right_content)
    right_box.text_frame.paragraphs[0].font.size = Pt(16)
    
    return slide
```

## 6. Section Divider Slide

```python
def create_section_slide(prs, section_title, background_color=RGBColor(68, 114, 196)):
    """Create section divider slide with colored background"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Background shape
    background = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(10), Inches(5.625))
    background.fill.solid()
    background.fill.fore_color.rgb = background_color
    
    # Section title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.3), Inches(9), Inches(1))
    title_frame = title_box.text_frame
    title_frame.text = section_title
    p = title_frame.paragraphs[0]
    p.font.size = Pt(48)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.CENTER
    
    return slide
```

## 7. Image Slide Pattern

```python
def create_image_slide(prs, image_path, caption=None):
    """Create slide with centered image"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Add image (centered)
    pic = slide.shapes.add_picture(image_path, Inches(2.5), Inches(1.5), width=Inches(5))
    
    # Caption (optional)
    if caption:
        caption_box = slide.shapes.add_textbox(Inches(1), Inches(4.5), Inches(8), Inches(0.5))
        caption_box.text_frame.text = caption
        caption_box.text_frame.paragraphs[0].font.size = Pt(14)
        caption_box.text_frame.paragraphs[0].font.italic = True
        caption_box.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    return slide
```

## 8. Footer Pattern

```python
def add_footer(slide, text, position="bottom"):
    """Add footer text to slide"""
    if position == "bottom":
        footer = slide.shapes.add_textbox(Inches(0.5), Inches(5.2), Inches(9), Inches(0.3))
    else:
        footer = slide.shapes.add_textbox(Inches(0.5), Inches(0.1), Inches(9), Inches(0.3))
    
    footer.text_frame.text = text
    footer.text_frame.paragraphs[0].font.size = Pt(10)
    footer.text_frame.paragraphs[0].font.color.rgb = RGBColor(128, 128, 128)
    
    return footer
```
