#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script to convert PowerPoint presentations to PDF
Supports multiple conversion methods depending on available libraries
"""

import os
import sys
import subprocess
from pathlib import Path


def convert_with_libreoffice(pptx_file, output_dir=None):
    """
    Convert PowerPoint to PDF using LibreOffice (headless mode)
    This is the most reliable method on Linux systems
    """
    if output_dir is None:
        output_dir = os.path.dirname(pptx_file) or '.'
    
    try:
        # Check if LibreOffice is installed
        result = subprocess.run(['which', 'libreoffice'], 
                              capture_output=True, 
                              text=True)
        
        if result.returncode != 0:
            print("❌ LibreOffice not found. Please install it:")
            print("   Ubuntu/Debian: sudo apt-get install libreoffice")
            print("   Fedora: sudo dnf install libreoffice")
            print("   Arch: sudo pacman -S libreoffice-fresh")
            return False
        
        # Convert using LibreOffice
        cmd = [
            'libreoffice',
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', output_dir,
            pptx_file
        ]
        
        print(f"🔄 Converting {pptx_file} to PDF...")
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        if result.returncode == 0:
            pdf_name = Path(pptx_file).stem + '.pdf'
            pdf_path = os.path.join(output_dir, pdf_name)
            print(f"✅ Successfully converted to PDF!")
            print(f"📁 Output file: {pdf_path}")
            return True
        else:
            print(f"❌ Conversion failed: {result.stderr}")
            return False
            
    except Exception as e:
        print(f"❌ Error during conversion: {e}")
        return False


def convert_with_unoconv(pptx_file, output_dir=None):
    """
    Convert PowerPoint to PDF using unoconv
    Alternative method using LibreOffice's UNO bridge
    """
    if output_dir is None:
        output_dir = os.path.dirname(pptx_file) or '.'
    
    try:
        # Check if unoconv is installed
        result = subprocess.run(['which', 'unoconv'], 
                              capture_output=True, 
                              text=True)
        
        if result.returncode != 0:
            print("❌ unoconv not found. Please install it:")
            print("   Ubuntu/Debian: sudo apt-get install unoconv")
            print("   Fedora: sudo dnf install unoconv")
            return False
        
        # Convert using unoconv
        cmd = [
            'unoconv',
            '-f', 'pdf',
            '-o', output_dir,
            pptx_file
        ]
        
        print(f"🔄 Converting {pptx_file} to PDF using unoconv...")
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        if result.returncode == 0:
            pdf_name = Path(pptx_file).stem + '.pdf'
            pdf_path = os.path.join(output_dir, pdf_name)
            print(f"✅ Successfully converted to PDF!")
            print(f"📁 Output file: {pdf_path}")
            return True
        else:
            print(f"❌ Conversion failed: {result.stderr}")
            return False
            
    except Exception as e:
        print(f"❌ Error during conversion: {e}")
        return False


def create_info_pdf():
    """
    Create a simple PDF with information about PowerPoint
    Uses reportlab library
    """
    try:
        from reportlab.lib.pagesizes import letter, A4
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import inch
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
        from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
        from reportlab.lib.colors import HexColor
        
        # Create PDF
        pdf_file = "PowerPoint_Guide.pdf"
        doc = SimpleDocTemplate(pdf_file, pagesize=letter)
        story = []
        styles = getSampleStyleSheet()
        
        # Custom styles
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=28,
            textColor=HexColor('#2980b9'),
            spaceAfter=30,
            alignment=TA_CENTER,
            fontName='Helvetica-Bold'
        )
        
        heading_style = ParagraphStyle(
            'CustomHeading',
            parent=styles['Heading2'],
            fontSize=18,
            textColor=HexColor('#2c3e50'),
            spaceAfter=12,
            spaceBefore=12,
            fontName='Helvetica-Bold'
        )
        
        body_style = ParagraphStyle(
            'CustomBody',
            parent=styles['BodyText'],
            fontSize=12,
            alignment=TA_JUSTIFY,
            spaceAfter=12,
            leading=16
        )
        
        # Title Page
        story.append(Spacer(1, 2*inch))
        story.append(Paragraph("Microsoft PowerPoint", title_style))
        story.append(Paragraph("Complete Guide and Overview", styles['Heading3']))
        story.append(Spacer(1, 0.5*inch))
        story.append(Paragraph("Created by Blackbox AI", styles['Normal']))
        story.append(PageBreak())
        
        # Introduction
        story.append(Paragraph("What is Microsoft PowerPoint?", heading_style))
        story.append(Paragraph(
            "Microsoft PowerPoint is a powerful presentation software developed by Microsoft. "
            "It is part of the Microsoft Office suite and is widely used for creating professional "
            "presentations, slideshows, and visual aids for business, education, and personal use.",
            body_style
        ))
        story.append(Spacer(1, 0.2*inch))
        
        # Key Features
        story.append(Paragraph("Key Features", heading_style))
        features = [
            "<b>Slide Creation:</b> Create unlimited slides with various layouts and designs",
            "<b>Templates:</b> Access thousands of pre-designed templates for different purposes",
            "<b>Animations:</b> Add dynamic animations and transitions between slides",
            "<b>Multimedia Support:</b> Insert images, videos, audio, and charts",
            "<b>Collaboration:</b> Work together with team members in real-time",
            "<b>Presenter View:</b> View notes and upcoming slides while presenting",
            "<b>Export Options:</b> Save presentations as PDF, video, or images",
            "<b>Cloud Integration:</b> Store and access presentations via OneDrive"
        ]
        
        for feature in features:
            story.append(Paragraph(f"• {feature}", body_style))
        
        story.append(Spacer(1, 0.3*inch))
        
        # Common Uses
        story.append(Paragraph("Common Uses", heading_style))
        uses = [
            "<b>Business Presentations:</b> Sales pitches, quarterly reports, project updates",
            "<b>Education:</b> Lectures, student projects, training materials",
            "<b>Conferences:</b> Keynote speeches, workshop materials, seminars",
            "<b>Marketing:</b> Product launches, marketing campaigns, brand presentations",
            "<b>Personal:</b> Photo slideshows, event invitations, portfolios"
        ]
        
        for use in uses:
            story.append(Paragraph(f"• {use}", body_style))
        
        story.append(PageBreak())
        
        # Best Practices
        story.append(Paragraph("Best Practices for Creating Presentations", heading_style))
        
        story.append(Paragraph("<b>1. Design Principles</b>", body_style))
        story.append(Paragraph(
            "• Keep slides simple and uncluttered<br/>"
            "• Use consistent fonts and colors throughout<br/>"
            "• Follow the 6x6 rule: maximum 6 bullets per slide, 6 words per bullet<br/>"
            "• Use high-quality images and graphics<br/>"
            "• Maintain adequate white space",
            body_style
        ))
        story.append(Spacer(1, 0.2*inch))
        
        story.append(Paragraph("<b>2. Content Guidelines</b>", body_style))
        story.append(Paragraph(
            "• Start with a clear outline<br/>"
            "• Use bullet points instead of paragraphs<br/>"
            "• Include one main idea per slide<br/>"
            "• Support claims with data and visuals<br/>"
            "• End with a strong conclusion or call-to-action",
            body_style
        ))
        story.append(Spacer(1, 0.2*inch))
        
        story.append(Paragraph("<b>3. Presentation Tips</b>", body_style))
        story.append(Paragraph(
            "• Practice your presentation multiple times<br/>"
            "• Use presenter notes for reference<br/>"
            "• Maintain eye contact with audience<br/>"
            "• Speak clearly and at a moderate pace<br/>"
            "• Be prepared to answer questions",
            body_style
        ))
        
        story.append(PageBreak())
        
        # PowerPoint Versions
        story.append(Paragraph("PowerPoint Versions and Platforms", heading_style))
        story.append(Paragraph(
            "<b>Desktop Applications:</b><br/>"
            "• PowerPoint for Windows (Microsoft 365, Office 2021, 2019, 2016)<br/>"
            "• PowerPoint for Mac (Microsoft 365, Office 2021, 2019)<br/><br/>"
            "<b>Web and Mobile:</b><br/>"
            "• PowerPoint Online (web browser)<br/>"
            "• PowerPoint Mobile (iOS and Android)<br/>"
            "• PowerPoint for iPad and tablets",
            body_style
        ))
        story.append(Spacer(1, 0.3*inch))
        
        # File Formats
        story.append(Paragraph("File Formats", heading_style))
        formats = [
            "<b>.PPTX:</b> Default PowerPoint format (Office 2007 and later)",
            "<b>.PPT:</b> Legacy PowerPoint format (Office 97-2003)",
            "<b>.PDF:</b> Portable Document Format for sharing",
            "<b>.PPSX:</b> PowerPoint Show (opens directly in presentation mode)",
            "<b>.ODP:</b> OpenDocument Presentation format",
            "<b>.MP4:</b> Video export format"
        ]
        
        for fmt in formats:
            story.append(Paragraph(f"• {fmt}", body_style))
        
        story.append(Spacer(1, 0.3*inch))
        
        # Alternatives
        story.append(Paragraph("PowerPoint Alternatives", heading_style))
        alternatives = [
            "<b>Google Slides:</b> Free, cloud-based presentation software",
            "<b>Apple Keynote:</b> Presentation software for Mac and iOS",
            "<b>LibreOffice Impress:</b> Free, open-source alternative",
            "<b>Prezi:</b> Non-linear, zoom-based presentations",
            "<b>Canva:</b> Design-focused presentation tool",
            "<b>Slidesgo:</b> Template-based presentation creator"
        ]
        
        for alt in alternatives:
            story.append(Paragraph(f"• {alt}", body_style))
        
        story.append(PageBreak())
        
        # Conclusion
        story.append(Paragraph("Conclusion", heading_style))
        story.append(Paragraph(
            "Microsoft PowerPoint remains the industry standard for creating professional presentations. "
            "Its comprehensive feature set, ease of use, and wide compatibility make it an essential tool "
            "for professionals, educators, and students worldwide. Whether you're creating a simple slideshow "
            "or a complex multimedia presentation, PowerPoint provides the tools and flexibility needed to "
            "communicate your ideas effectively.",
            body_style
        ))
        story.append(Spacer(1, 0.3*inch))
        
        story.append(Paragraph(
            "By following best practices and leveraging PowerPoint's powerful features, you can create "
            "engaging, memorable presentations that captivate your audience and deliver your message with impact.",
            body_style
        ))
        
        # Build PDF
        doc.build(story)
        print(f"✅ Successfully created PDF guide!")
        print(f"📁 Output file: {pdf_file}")
        return True
        
    except ImportError:
        print("❌ reportlab library not found. Install it with:")
        print("   pip install reportlab")
        return False
    except Exception as e:
        print(f"❌ Error creating PDF: {e}")
        return False


def main():
    print("=" * 60)
    print("PowerPoint to PDF Converter")
    print("=" * 60)
    print()
    
    # Check if a file was provided
    if len(sys.argv) > 1:
        pptx_file = sys.argv[1]
        
        if not os.path.exists(pptx_file):
            print(f"❌ File not found: {pptx_file}")
            return
        
        if not pptx_file.lower().endswith(('.pptx', '.ppt')):
            print("❌ File must be a PowerPoint file (.pptx or .ppt)")
            return
        
        # Try conversion methods in order of preference
        print("Attempting conversion with LibreOffice...")
        if convert_with_libreoffice(pptx_file):
            return
        
        print("\nAttempting conversion with unoconv...")
        if convert_with_unoconv(pptx_file):
            return
        
        print("\n❌ All conversion methods failed.")
        print("Please install LibreOffice or unoconv to convert PowerPoint files.")
    
    else:
        # No file provided, create an informational PDF about PowerPoint
        print("No PowerPoint file provided.")
        print("Creating an informational PDF about PowerPoint instead...\n")
        create_info_pdf()


if __name__ == "__main__":
    main()
