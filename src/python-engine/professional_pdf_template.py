"""
Professional PDF Template Generator
Creates beautifully formatted PDFs matching the reference report style with:
- Green borders and frames
- Company branding headers/footers
- Trebuchet MS font (or Arial fallback)
- Colored tables and sections
- Professional layout
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, mm
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle,
    Frame, PageTemplate, Image, KeepTogether
)
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from pathlib import Path
import os

# Professional color scheme matching reference
COLORS = {
    'primary_green': colors.HexColor('#C5E0B4'),      # Light green for borders/backgrounds
    'dark_green': colors.HexColor('#70AD47'),          # Dark green for accents
    'header_gray': colors.HexColor('#D9D9D9'),         # Gray for table headers
    'text_black': colors.HexColor('#000000'),          # Pure black for text
    'red_accent': colors.HexColor('#FF0000'),          # Red for important items
    'blue_accent': colors.HexColor('#0000FF'),         # Blue for links
    'yellow_highlight': colors.HexColor('#FFFF00'),    # Yellow for highlights
    'table_border': colors.HexColor('#548235'),        # Dark green for table borders
}


class ProfessionalTemplate:
    """Professional PDF template with consistent formatting"""
    
    def __init__(self, company_name="FINVOIS OPEN BUSINESS SOLUTIONS LLP", 
                 contact="9618221011", tagline="MSME & DPR Consultants"):
        self.company_name = company_name
        self.contact = contact
        self.tagline = tagline
        
        # Try to register Trebuchet MS (Windows font)
        self.primary_font = 'Trebuchet-MS'
        self.font_bold = 'Trebuchet-MS-Bold'
        
        try:
            # Common Windows font paths
            font_paths = [
                'C:/Windows/Fonts/trebuc.ttf',
                'C:/Windows/Fonts/trebucbd.ttf',
            ]
            
            if os.path.exists(font_paths[0]):
                pdfmetrics.registerFont(TTFont('Trebuchet-MS', font_paths[0]))
            if os.path.exists(font_paths[1]):
                pdfmetrics.registerFont(TTFont('Trebuchet-MS-Bold', font_paths[1]))
        except:
            # Fallback to Helvetica if Trebuchet not available
            self.primary_font = 'Helvetica'
            self.font_bold = 'Helvetica-Bold'
    
    def create_header_footer(self, canvas_obj, doc):
        """Draw header and footer on each page"""
        canvas_obj.saveState()
        
        width, height = A4
        
        # Draw green border/frame around page
        canvas_obj.setStrokeColor(COLORS['primary_green'])
        canvas_obj.setLineWidth(3)
        canvas_obj.rect(15*mm, 15*mm, width - 30*mm, height - 30*mm, stroke=1, fill=0)
        
        # Inner decorative border
        canvas_obj.setStrokeColor(COLORS['dark_green'])
        canvas_obj.setLineWidth(1)
        canvas_obj.rect(18*mm, 18*mm, width - 36*mm, height - 36*mm, stroke=1, fill=0)
        
        # Header - Company info
        canvas_obj.setFont(self.font_bold, 9)
        canvas_obj.setFillColor(COLORS['dark_green'])
        
        # Top header
        header_y = height - 25*mm
        canvas_obj.drawString(25*mm, header_y, "Detailed Project Report on Manufacturing of Building Materials")
        
        # Location (right aligned)
        canvas_obj.setFont(self.primary_font, 8)
        canvas_obj.drawRightString(width - 25*mm, header_y - 10, "NTR District, Andhrapradesh")
        
        # Footer - Consultant info
        footer_y = 20*mm
        
        # Company name (centered)
        canvas_obj.setFont(self.font_bold, 8)
        canvas_obj.setFillColor(COLORS['text_black'])
        canvas_obj.drawCentredString(width/2, footer_y + 10, self.company_name)
        
        # Tagline
        canvas_obj.setFont(self.primary_font, 7)
        canvas_obj.setFillColor(colors.gray)
        canvas_obj.drawCentredString(width/2, footer_y + 3, self.tagline)
        
        # Contact
        canvas_obj.setFont(self.primary_font, 7)
        canvas_obj.drawCentredString(width/2, footer_y - 3, f"Contact No: {self.contact}")
        
        # Page number (right bottom)
        canvas_obj.setFont(self.primary_font, 8)
        canvas_obj.setFillColor(COLORS['text_black'])
        page_num = canvas_obj.getPageNumber()
        canvas_obj.drawRightString(width - 25*mm, footer_y, f"Page {page_num}")
        
        canvas_obj.restoreState()
    
    def get_styles(self):
        """Get custom paragraph styles matching reference PDF"""
        styles = getSampleStyleSheet()
        
        # Main title style (cover page)
        styles.add(ParagraphStyle(
            name='CoverTitle',
            parent=styles['Heading1'],
            fontName=self.font_bold,
            fontSize=24,
            textColor=COLORS['dark_green'],
            alignment=TA_CENTER,
            spaceAfter=20,
            leading=30,
            fontWeight='BOLD'
        ))
        
        # Section heading style
        styles.add(ParagraphStyle(
            name='SectionHeading',
            parent=styles['Heading1'],
            fontName=self.font_bold,
            fontSize=14,
            textColor=COLORS['text_black'],
            alignment=TA_LEFT,
            spaceAfter=12,
            spaceBefore=20,
            leading=18,
            borderColor=COLORS['dark_green'],
            borderWidth=0,
            borderPadding=5,
        ))
        
        # Subsection heading
        styles.add(ParagraphStyle(
            name='SubHeading',
            parent=styles['Heading2'],
            fontName=self.font_bold,
            fontSize=12,
            textColor=COLORS['text_black'],
            alignment=TA_LEFT,
            spaceAfter=10,
            spaceBefore=12,
            leading=16,
        ))
        
        # Body text style
        styles.add(ParagraphStyle(
            name='ProfessionalBody',
            parent=styles['BodyText'],
            fontName=self.primary_font,
            fontSize=11,
            textColor=COLORS['text_black'],
            alignment=TA_JUSTIFY,
            spaceAfter=8,
            leading=14,
            leftIndent=0,
            rightIndent=0,
        ))
        
        # Bullet points
        styles.add(ParagraphStyle(
            name='BulletPoint',
            parent=styles['BodyText'],
            fontName=self.primary_font,
            fontSize=10,
            textColor=COLORS['text_black'],
            alignment=TA_LEFT,
            spaceAfter=6,
            leading=13,
            leftIndent=20,
            bulletIndent=10,
        ))
        
        # Highlighted text
        styles.add(ParagraphStyle(
            name='Highlighted',
            parent=styles['BodyText'],
            fontName=self.font_bold,
            fontSize=11,
            textColor=COLORS['red_accent'],
            alignment=TA_CENTER,
            spaceAfter=10,
            leading=14,
        ))
        
        # Index/TOC style
        styles.add(ParagraphStyle(
            name='TOCEntry',
            parent=styles['BodyText'],
            fontName=self.primary_font,
            fontSize=10,
            textColor=COLORS['text_black'],
            alignment=TA_LEFT,
            spaceAfter=6,
            leading=14,
            leftIndent=15,
        ))
        
        # Firm info style
        styles.add(ParagraphStyle(
            name='FirmInfo',
            parent=styles['BodyText'],
            fontName=self.primary_font,
            fontSize=10,
            textColor=COLORS['text_black'],
            alignment=TA_LEFT,
            spaceAfter=5,
            leading=13,
        ))
        
        return styles
    
    def create_professional_table(self, data, col_widths=None, header_bg=None, 
                                 has_header=True, stripe=False):
        """Create a professionally styled table matching reference PDF"""
        
        if not data:
            return None
        
        # Create table with auto-adjusting row heights
        table = Table(data, colWidths=col_widths, repeatRows=1 if has_header else 0)
        
        # Base style
        style_commands = [
            # Overall table style
            ('BACKGROUND', (0, 0), (-1, -1), colors.white),
            ('TEXTCOLOR', (0, 0), (-1, -1), COLORS['text_black']),
            ('FONTNAME', (0, 0), (-1, -1), self.primary_font),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # Changed to TOP for better overflow handling
            
            # Grid
            ('GRID', (0, 0), (-1, -1), 0.5, COLORS['table_border']),
            ('BOX', (0, 0), (-1, -1), 1.5, COLORS['table_border']),
            
            # Padding - reduced for more compact tables
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            
            # Word wrap for long content
            ('WORDWRAP', (0, 0), (-1, -1), True),
        ]
        
        # Header row styling
        if has_header:
            header_color = header_bg if header_bg else COLORS['primary_green']
            style_commands.extend([
                ('BACKGROUND', (0, 0), (-1, 0), header_color),
                ('TEXTCOLOR', (0, 0), (-1, 0), COLORS['text_black']),
                ('FONTNAME', (0, 0), (-1, 0), self.font_bold),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
                ('TOPPADDING', (0, 0), (-1, 0), 10),
                ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),  # Center header vertically
            ])
        
        # Alternating row colors (striped)
        if stripe:
            for i in range(1 if has_header else 0, len(data), 2):
                style_commands.append(
                    ('BACKGROUND', (0, i), (-1, i), colors.Color(0.95, 0.95, 0.95))
                )
        
        table.setStyle(TableStyle(style_commands))
        
        return table
    
    def create_info_box(self, title, content, bg_color=None):
        """Create an information box with border"""
        bg = bg_color if bg_color else COLORS['primary_green']
        
        data = [[title], [content]]
        table = Table(data, colWidths=[6*inch])
        
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), bg),
            ('BACKGROUND', (0, 1), (-1, 1), colors.white),
            ('TEXTCOLOR', (0, 0), (-1, -1), COLORS['text_black']),
            ('FONTNAME', (0, 0), (-1, 0), self.font_bold),
            ('FONTNAME', (0, 1), (-1, 1), self.primary_font),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('FONTSIZE', (0, 1), (-1, 1), 10),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('ALIGN', (0, 1), (-1, 1), 'LEFT'),
            ('BOX', (0, 0), (-1, -1), 2, COLORS['dark_green']),
            ('LINEBELOW', (0, 0), (-1, 0), 1, COLORS['dark_green']),
            ('TOPPADDING', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
            ('TOPPADDING', (0, 1), (-1, 1), 15),
            ('BOTTOMPADDING', (0, 1), (-1, 1), 15),
            ('LEFTPADDING', (0, 0), (-1, -1), 15),
            ('RIGHTPADDING', (0, 0), (-1, -1), 15),
        ]))
        
        return table


if __name__ == "__main__":
    # Test the template
    from reportlab.platypus import SimpleDocTemplate
    
    template = ProfessionalTemplate()
    
    # Create a test PDF
    doc = SimpleDocTemplate(
        "test_professional_template.pdf",
        pagesize=A4,
        topMargin=30*mm,
        bottomMargin=30*mm,
        leftMargin=25*mm,
        rightMargin=25*mm
    )
    
    styles = template.get_styles()
    story = []
    
    # Add test content
    story.append(Paragraph("DETAILED PROJECT REPORT", styles['CoverTitle']))
    story.append(Spacer(1, 20))
    story.append(Paragraph("Test Section Heading", styles['SectionHeading']))
    story.append(Paragraph("This is a test paragraph with professional formatting. " * 10, 
                          styles['ProfessionalBody']))
    
    # Test table
    table_data = [
        ['Header 1', 'Header 2', 'Header 3'],
        ['Data 1', 'Data 2', 'Data 3'],
        ['Data 4', 'Data 5', 'Data 6'],
    ]
    story.append(template.create_professional_table(table_data, stripe=True))
    
    # Build PDF
    doc.build(story, onFirstPage=template.create_header_footer, 
              onLaterPages=template.create_header_footer)
    
    print("✅ Test PDF created: test_professional_template.pdf")
