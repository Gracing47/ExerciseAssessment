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
    
    # Formatierung
    title_shape.text_frame.paragraphs[0].font.size = Pt(44)
    title_shape.text_frame.paragraphs[0].font.bold = True
    subtitle_shape.text_frame.paragraphs[0].font.size = Pt(24)
    
    return slide

def create_content_slide(prs, title, content, has_image=False, image_path=None):
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    
    # Titel
    title_shape = slide.shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.size = Pt(36)
    
    # Content
    if has_image and image_path and os.path.exists(image_path):
        # Bild links, Text rechts
        img_left = Inches(1)
        img_top = Inches(2)
        img_width = Inches(4)
        img_height = Inches(4)
        
        slide.shapes.add_picture(image_path, img_left, img_top, img_width, img_height)
        
        # Text rechts vom Bild
        text_left = Inches(6)
        text_top = Inches(2)
        text_width = Inches(4)
        text_height = Inches(4)
        
        textbox = slide.shapes.add_textbox(text_left, text_top, text_width, text_height)
        text_frame = textbox.text_frame
        
        for line in content.split('\n'):
            p = text_frame.add_paragraph()
            p.text = line
            p.font.size = Pt(18)
    else:
        # Nur Text
        body_shape = slide.placeholders[1]
        text_frame = body_shape.text_frame
        
        for line in content.split('\n'):
            p = text_frame.add_paragraph()
            p.text = line
            p.font.size = Pt(18)
    
    return slide

def create_presentation():
    prs = Presentation()
    prs.slide_width = Inches(16)  # 16:9 Format
    prs.slide_height = Inches(9)
    
    # Titelfolie
    create_title_slide(prs, 
                      "Team Availability Analysis 📊",
                      "August 2022 Overview")
    
    # Service Line Requirements
    requirements_content = """
Minimum Staffing Requirements:
• Service Line 1: 2.5 agents
• Service Line 2: 4.0 agents
• Service Line 3: 18.0 agents

Total required: 24.5 agents for smooth operations
"""
    create_content_slide(prs, "Service Line Requirements 📋", requirements_content)
    
    # Daily Availability
    daily_content = """
Key Findings:
• Average daily availability: 2.0 agents
• Maximum availability: 4 agents
• Minimum availability: 0 agents
• Consistent staffing gaps in Service Line 3
"""
    create_content_slide(prs, "Daily Availability Overview 📈", daily_content, True, 
                        "output/daily_availability.png")
    
    # Weekly Patterns
    weekly_content = """
Weekly Analysis:
• Highest availability: Friday (2.2 agents)
• Lowest availability: Wednesday (1.8 agents)
• Weekend coverage requires attention
• Mid-week peaks in attendance
"""
    create_content_slide(prs, "Weekly Patterns 📅", weekly_content, True,
                        "output/weekly_pattern.png")
    
    # Staffing Gaps
    gaps_content = """
Service Line Analysis:
• Service Line 1: Generally manageable
• Service Line 2: Occasional shortages
• Service Line 3: Significant understaffing
"""
    create_content_slide(prs, "Staffing Gaps Analysis ⚠️", gaps_content, True,
                        "output/staffing_gaps.png")
    
    # Monthly Statistics
    stats_content = """
August 2022 Overview:
• Total working days: 31
• Days meeting requirements: 0 (0%)
• Days below requirements: 31 (100%)
• Average daily shortage: 22.5 agents
"""
    create_content_slide(prs, "Monthly Statistics 📊", stats_content)
    
    # Critical Findings
    findings_content = """
Key Issues Identified:
• Consistent understaffing in Service Line 3
• Friday staffing levels particularly concerning
• Weekend coverage gaps
• No days meet minimum requirements
"""
    create_content_slide(prs, "Critical Findings 🔍", findings_content)
    
    # Recommendations
    recommendations_content = """
Immediate Actions Needed:
1. Prioritize recruitment to meet Service Line 3 requirements
2. Implement temporary staff augmentation
3. Review and optimize leave management
4. Develop flexible scheduling
"""
    create_content_slide(prs, "Recommendations ✨", recommendations_content)
    
    # Next Steps
    next_steps_content = """
Action Plan:
1. Short-term: Optimize current staff distribution
2. Medium-term: Recruit additional staff
3. Long-term: Implement workforce management system
4. Regular monitoring and adjustment
"""
    create_content_slide(prs, "Next Steps 🎯", next_steps_content)
    
    # Speichern
    output_dir = "output/Presentation"
    os.makedirs(output_dir, exist_ok=True)
    prs.save(f"{output_dir}/team_availability_analysis.pptx")

if __name__ == "__main__":
    create_presentation()
