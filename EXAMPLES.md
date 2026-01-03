# Complete Examples - Python-pptx

## Example 1: Basic Presentation

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(5.625)

# Title slide
slide1 = prs.slides.add_slide(prs.slide_layouts[6])
title_box = slide1.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(1.5))
title_frame = title_box.text_frame
title_frame.text = "My Presentation"
title_frame.paragraphs[0].font.size = Pt(44)
title_frame.paragraphs[0].font.bold = True
title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

# Content slide
slide2 = prs.slides.add_slide(prs.slide_layouts[6])
content_box = slide2.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(3.5))
content_box.text_frame.text = "This is the content of the second slide."
content_box.text_frame.paragraphs[0].font.size = Pt(24)

prs.save('basic.pptx')
```

## Example 2: Quarterly Report

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(5.625)

# Slide 1: Title
slide1 = prs.slides.add_slide(prs.slide_layouts[6])
title_box = slide1.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(1.5))
title_frame = title_box.text_frame
title_frame.text = "Q3 2024 Quarterly Report"
title_frame.paragraphs[0].font.size = Pt(44)
title_frame.paragraphs[0].font.bold = True
title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
# Slide 2: Key Metrics
slide2 = prs.slides.add_slide(prs.slide_layouts[6])
title2 = slide2.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
title2.text_frame.text = "Key Metrics"
title2.text_frame.paragraphs[0].font.size = Pt(32)
title2.text_frame.paragraphs[0].font.bold = True
title2.text_frame.paragraphs[0].font.color.rgb = RGBColor(31, 78, 121)  # Title color

metrics = [
    "Revenue: $2.5M (+15% YoY)",
    "New Customers: 500",
    "Retention Rate: 85%",
    "Net Promoter Score: 72"
]

y = Inches(1.8)
for metric in metrics:
    box = slide2.shapes.add_textbox(Inches(1), y, Inches(8), Inches(0.5))
    box.text_frame.text = f"• {metric}"
    box.text_frame.paragraphs[0].font.size = Pt(20)
    y = y + Inches(0.7)
# Slide 3: Challenges
slide3 = prs.slides.add_slide(prs.slide_layouts[6])
title3 = slide3.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
title3.text_frame.text = "Challenges & Opportunities"
title3.text_frame.paragraphs[0].font.size = Pt(32)
title3.text_frame.paragraphs[0].font.bold = True
title3.text_frame.paragraphs[0].font.color.rgb = RGBColor(31, 78, 121)  # Title color

challenges = [
    "Supply chain delays affecting production",
    "Increasing competition in core markets",
    "Opportunity: New product line launch",
    "Opportunity: International expansion"
]

y = Inches(1.8)
for challenge in challenges:
    box = slide3.shapes.add_textbox(Inches(1), y, Inches(8), Inches(0.6))
    box.text_frame.text = f"• {challenge}"
    box.text_frame.paragraphs[0].font.size = Pt(18)
    y = y + Inches(0.7)

prs.save('quarterly_report.pptx')
```

## Example 3: Data from Dictionary

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

data = {
    "Sales Overview": [
        "Total Sales: $4.2M",
        "Growth Rate: 12%",
        "Top Product: Widget Pro",
        "Best Region: North America"
    ],
    "Customer Analysis": [
        "Total Customers: 15,000",
        "New Acquisitions: 2,500",
        "Churn Rate: 3.2%",
        "Customer Satisfaction: 4.5/5"
    ],
    "Financial Health": [
        "Profit Margin: 18%",
        "Operating Expenses: $800K",
        "Cash Flow: Positive",
        "Debt-to-Equity: 0.4"
    ]
}

prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(5.625)

# Title slide
slide1 = prs.slides.add_slide(prs.slide_layouts[6])
title1 = slide1.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(1.5))
title1.text_frame.text = "Business Dashboard"
title1.text_frame.paragraphs[0].font.size = Pt(44)
title1.text_frame.paragraphs[0].font.bold = True

# Data slides
for category, metrics in data.items():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_box.text_frame.text = category
    title_box.text_frame.paragraphs[0].font.size = Pt(32)
    title_box.text_frame.paragraphs[0].font.bold = True
    title_box.text_frame.paragraphs[0].font.color.rgb = RGBColor(31, 78, 121)
    # Metrics
    y = Inches(1.8)
    for metric in metrics:
        box = slide.shapes.add_textbox(Inches(1), y, Inches(8), Inches(0.5))
        box.text_frame.text = f"• {metric}"
        box.text_frame.paragraphs[0].font.size = Pt(20)
        box.text_frame.paragraphs[0].font.color.rgb = RGBColor(89, 89, 89)  # Bullet color
        y = y + Inches(0.6)

prs.save('dashboard.pptx')
```

## Example 4: Presentation with Sections

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(5.625)

# Title slide
slide1 = prs.slides.add_slide(prs.slide_layouts[6])
title1 = slide1.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(1.5))
title1.text_frame.text = "Annual Training Program"
title1.text_frame.paragraphs[0].font.size = Pt(44)
title1.text_frame.paragraphs[0].font.bold = True
title1.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

# Section 1: Introduction
# Section divider
divider1 = prs.slides.add_slide(prs.slide_layouts[6])
bg1 = divider1.shapes.add_shape(1, Inches(0), Inches(0), Inches(10), Inches(5.625))
bg1.fill.solid()
bg1.fill.fore_color.rgb = RGBColor(68, 114, 196)
div_title1 = divider1.shapes.add_textbox(Inches(0.5), Inches(2.3), Inches(9), Inches(1))
div_title1.text_frame.text = "Introduction"
div_title1.text_frame.paragraphs[0].font.size = Pt(48)
div_title1.text_frame.paragraphs[0].font.bold = True
div_title1.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
div_title1.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

# Content
slide2 = prs.slides.add_slide(prs.slide_layouts[6])
content2 = slide2.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(3.5))
content2.text_frame.text = """Welcome to the annual training program!

This program covers:
• Company culture and values
• Product knowledge
• Customer service best practices
• Safety protocols

Please take notes and ask questions."""
content2.text_frame.paragraphs[0].font.size = Pt(20)
# Section 2: Product Knowledge
divider2 = prs.slides.add_slide(prs.slide_layouts[6])
bg2 = divider2.shapes.add_shape(1, Inches(0), Inches(0), Inches(10), Inches(5.625))
bg2.fill.solid()
bg2.fill.fore_color.rgb = RGBColor(79, 98, 40)
div_title2 = divider2.shapes.add_textbox(Inches(0.5), Inches(2.3), Inches(9), Inches(1))
div_title2.text_frame.text = "Product Knowledge"
div_title2.text_frame.paragraphs[0].font.size = Pt(48)
div_title2.text_frame.paragraphs[0].font.bold = True
div_title2.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
div_title2.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

prs.save('training_program.pptx')
```

## Example 5: Using Helper Functions

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

def add_text_box(slide, text, left, top, width, height, font_size=Pt(18), 
                 bold=False, color=RGBColor(0, 0, 0), alignment=None):
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

def create_title_slide(prs, title, subtitle=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title_box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(1.5))
    title_frame = title_box.text_frame
    title_frame.text = title
    title_frame.paragraphs[0].font.size = Pt(44)
    title_frame.paragraphs[0].font.bold = True
    title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    if subtitle:
        sub_box = slide.shapes.add_textbox(Inches(1), Inches(3.5), Inches(8), Inches(1))
        sub_box.text_frame.text = subtitle
        sub_box.text_frame.paragraphs[0].font.size = Pt(24)
        sub_box.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    return slide

prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(5.625)

create_title_slide(prs, "Modern Presentation", "Created with python-pptx")

slide2 = prs.slides.add_slide(prs.slide_layouts[6])
add_text_box(slide2, "Main heading", Inches(1), Inches(1), Inches(8), Inches(0.8), 
             font_size=Pt(28), bold=True, color=RGBColor(31, 78, 121))
add_text_box(slide2, "Subheading with supporting text", Inches(1), Inches(2), Inches(8), Inches(2),
             font_size=Pt(18), color=RGBColor(89, 89, 89))

prs.save('modern_presentation.pptx')
```