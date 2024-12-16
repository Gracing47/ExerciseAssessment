from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import os

def create_title_slide(prs, title, subtitle):
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_shape = slide.shapes.title
    subtitle_shape = slide.placeholders[1]
    
    title_shape.text = title
    subtitle_shape.text = subtitle
    
    # Formatting
    title_shape.text_frame.paragraphs[0].font.size = Pt(44)
    title_shape.text_frame.paragraphs[0].font.bold = True
    subtitle_shape.text_frame.paragraphs[0].font.size = Pt(24)
    
    return slide

def create_content_slide(prs, title, content_list):
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    shapes = slide.shapes
    
    title_shape = shapes.title
    body_shape = shapes.placeholders[1]
    
    title_shape.text = title
    
    tf = body_shape.text_frame
    
    for idx, item in enumerate(content_list):
        if idx == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = item
        p.level = 0
        p.font.size = Pt(18)
    
    return slide

def add_image_slide(prs, title, image_path, subtitle=None):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    shapes = slide.shapes
    
    title_shape = shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.size = Pt(32)
    
    # Add image
    if os.path.exists(image_path):
        left = Inches(1)
        top = Inches(1.5)
        width = Inches(8)
        height = Inches(5)
        slide.shapes.add_picture(image_path, left, top, width, height)
    
    # Add subtitle if provided
    if subtitle:
        left = Inches(1)
        top = Inches(6.5)
        width = Inches(8)
        height = Inches(1)
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        p = tf.add_paragraph()
        p.text = subtitle
        p.font.size = Pt(14)
        p.alignment = PP_ALIGN.CENTER
    
    return slide

def create_presentation():
    prs = Presentation()
    
    # Title slide
    create_title_slide(prs, 
                      "Team Availability Analysis",
                      "August 2022 Overview")
    
    # Service Line Requirements slide
    requirements = [
        "Service Line Requirements:",
        "• Service Line 1: 2.5 agents",
        "• Service Line 2: 4 agents",
        "• Service Line 3: 18 agents",
        "",
        "Total required: 24.5 agents for smooth operations"
    ]
    create_content_slide(prs, "Service Line Requirements", requirements)
    
    # Daily Availability slide
    daily_stats = [
        "Key Statistics:",
        "• Total number of agents: 29",
        "• Average daily available agents: 24.03",
        "• Maximum available agents: 29.00",
        "• Days meeting requirements: 16 days",
        "• Days below requirements: 8 days"
    ]
    create_content_slide(prs, "Daily Availability Overview", daily_stats)
    
    # Daily Availability Chart
    add_image_slide(prs, 
                   "Daily Team Availability", 
                   "../output/daily_availability.png",
                   "Daily availability trend showing actual vs required staffing levels")
    
    # Staffing Gaps Chart
    add_image_slide(prs, 
                   "Staffing Gaps Analysis", 
                   "../output/staffing_gaps.png",
                   "Red bars indicate understaffing, green bars indicate surplus staffing")
    
    # Recommendations slide
    recommendations = [
        "Action Items:",
        "• Implement better leave management to prevent understaffing",
        "• Consider additional backup staff for high-absence days",
        "• Review partial availability patterns to optimize scheduling",
        "• Develop contingency plans for days with known staffing gaps"
    ]
    create_content_slide(prs, "Recommendations", recommendations)
    
    # Save the presentation
    output_dir = "../output/Presentation"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    prs.save(os.path.join(output_dir, "team_availability_analysis.pptx"))

if __name__ == "__main__":
    create_presentation()
