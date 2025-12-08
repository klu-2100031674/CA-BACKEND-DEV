"""
PDF Report Generator with AI-Generated Content
Combines Excel-derived PDFs with AI-generated contextual content using Grok AI API.
Uses ONLY the two resource PDFs as knowledge source - NO external data allowed.
"""

import os
import sys
import json
import re
from pathlib import Path
from typing import Dict, List, Any, Optional
import openai
import google.generativeai as genai
from PyPDF2 import PdfMerger, PdfReader
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle, KeepTogether
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY, TA_RIGHT
from ai_resource_parser import AIResourceParser


def get_final_sheet_name(template_name: str) -> str:
    """
    Get the correct 'Final workings' sheet name based on template type.
    
    Args:
        template_name: The template file name or identifier
        
    Returns:
        The exact sheet name to use for this template
    """
    template_upper = template_name.upper()
    
    # CC1 -> FinalWorkings
    if 'CC1' in template_upper or 'FORMAT CC1' in template_upper:
        return 'FinalWorkings'
    # CC2 -> FinalWorkings
    elif 'CC2' in template_upper or 'FORMAT CC2' in template_upper:
        return 'FinalWorkings'
    # CC3 -> FinalWorkings
    elif 'CC3' in template_upper or 'FORMAT CC3' in template_upper:
        return 'FinalWorkings'
    # CC4 -> Finalworkings
    elif 'CC4' in template_upper or 'FORMAT CC4' in template_upper:
        return 'Finalworkings'
    # CC5 -> FinalWorkings
    elif 'CC5' in template_upper or 'FORMAT CC5' in template_upper:
        return 'FinalWorkings'
    # CC6 -> Final workings (with space)
    elif 'CC6' in template_upper or 'FORMAT CC6' in template_upper:
        return 'Final workings'
    # Term Loan -> Final workings (with space)
    elif 'TERM LOAN' in template_upper or 'TERM_LOAN' in template_upper:
        return 'Final workings'
    # Default fallback
    else:
        return 'Finalworkings'


# Sheet name normalization utilities
def normalize_sheet_name(sheet_name: str) -> str:
    """Normalize sheet name by stripping whitespace and converting to lowercase."""
    return sheet_name.strip().lower()


def normalize_identifier(value: str) -> str:
    """Create a comparable identifier by removing non-alphanumeric characters."""
    if not value:
        return ''
    return re.sub(r'[^a-z0-9]', '', value.lower())


def find_sheet_match(expected_sheet: str, available_sheets: List[str]) -> str:
    """
    Find the best matching sheet name from available sheets, ignoring case and spaces.
    
    Args:
        expected_sheet: The expected sheet name
        available_sheets: List of actual sheet names in the workbook
        
    Returns:
        The matching sheet name from available_sheets, or None if no match found
    """
    normalized_expected = normalize_sheet_name(expected_sheet)
    
    # First try exact match after normalization
    for sheet in available_sheets:
        if normalize_sheet_name(sheet) == normalized_expected:
            return sheet
    
    # If no exact match, try partial matches (in case of slight variations)
    for sheet in available_sheets:
        if normalized_expected in normalize_sheet_name(sheet) or normalize_sheet_name(sheet) in normalized_expected:
            return sheet
    
    return None

def extract_sheet_title_from_filename(filename: str) -> str:
    """Infer the best-guess sheet title from a generated PDF filename."""
    stem = Path(filename).stem
    match = re.match(r'^sheet_\d+_(.+)$', stem, re.IGNORECASE)
    sheet_title = match.group(1) if match else stem
    sheet_title = sheet_title.replace('_', ' ').strip()
    return sheet_title or filename
from professional_pdf_template import ProfessionalTemplate, COLORS

class AIReportGenerator:
    """Generate comprehensive reports with AI-enhanced content."""
    
    def __init__(self, api_key: str, provider: str = "perplexity"):
        """
        Initialize the AI Report Generator.
        
        Args:
            api_key: AI API key (Perplexity or Grok)
            provider: AI provider to use ("perplexity" or "grok")
        """
        self.api_key = api_key
        self.provider = provider.lower()
        self.ai_parser = AIResourceParser()
        self.knowledge_base = None
        
        # Configure AI client based on provider
        if self.provider == "grok":
            # Configure Grok AI (xAI)
            self.client = openai.OpenAI(
                api_key=api_key,
                base_url="https://api.x.ai/v1"
            )
            self.model = "grok-code-fast-1"  # Using the specified Grok model
            print("🤖 AI Report Generator initialized with Grok (xAI)", file=sys.stderr)
        elif self.provider == "perplexity":
            # Configure Perplexity AI
            self.client = openai.OpenAI(
                api_key=api_key,
                base_url="https://api.perplexity.ai/"
            )
            self.model = "sonar"
            print("🤖 AI Report Generator initialized with Perplexity", file=sys.stderr)
        elif self.provider == "gemini":
            # Configure Gemini AI (Google)
            genai.configure(api_key=api_key)
            self.client = genai.GenerativeModel('gemini-1.5-flash')
            self.model = "gemini-1.5-flash"
            print("🤖 AI Report Generator initialized with Gemini (Google AI)", file=sys.stderr)
        else:
            raise ValueError(f"Unsupported AI provider: {provider}. Use 'perplexity', 'grok', or 'gemini'")
        
        print(f"🤖 Using model: {self.model}", file=sys.stderr)
    
    def load_knowledge_base(self):
        """Load or create the AI knowledge base from resource PDFs."""
        kb_file = Path(__file__).parent / "ai_knowledge_base.json"
        
        if kb_file.exists():
            print("📚 Loading existing knowledge base...", file=sys.stderr)
            self.ai_parser.load_knowledge_base(str(kb_file))
        else:
            print("📚 Creating new knowledge base from resource PDFs...", file=sys.stderr)
            self.ai_parser.parse_all_resources()
            self.ai_parser.save_knowledge_base(str(kb_file))
        
        self.knowledge_base = self.ai_parser.knowledge_base
        print(f"✅ Knowledge base loaded: {self.knowledge_base['total_chunks']} chunks from {self.knowledge_base['total_pages']} pages", file=sys.stderr)
    
    def generate_ai_content(self, section_type: str, excel_data: Dict[str, Any], reference_context: str = "") -> str:
        """
        Generate AI content for a specific section using ONLY the resource PDFs as context.
        
        Args:
            section_type: Type of section (e.g., "executive_summary", "project_description", etc.)
            excel_data: Computed data from Excel sheets
            reference_context: Additional context from reference report
            
        Returns:
            Generated content as string
        """
        if not self.knowledge_base:
            self.load_knowledge_base()
        
        # Search knowledge base for relevant content
        search_queries = {
            "executive_summary": "project summary financial assistance manufacturing",
            "project_profile": "project profile overview business details",
            "firm_constitution": "firm constitution partnership proprietorship company",
            "product_characteristics": "product characteristics market analysis demand",
            "swot_analysis": "SWOT analysis strengths weaknesses opportunities threats",
            "project_description": "project description manufacturing business",
            "manufacturing_process": "manufacturing process production flowchart operations",
            "plant_machinery": "plant machinery equipment technical specifications",
            "inventory_details": "inventory stock raw materials working capital",
            "transportation": "transportation logistics distribution",
            "land_requirements": "land building requirements infrastructure",
            "financial_analysis": "financial analysis profitability balance sheet",
            "ratio_interpretation": "ratio analysis DSCR current ratio financial ratios banking norms",
            "mpbf_calculation": "MPBF calculation working capital turnover method",
            "cash_flow_projection": "cash flow projection statements operating investing financing",
            "funds_flow_analysis": "funds flow statement sources applications capital",
            "loan_eligibility": "loan eligibility criteria financial assistance",
            "recommendations": "recommendations project viability assessment"
        }
        
        query = search_queries.get(section_type, section_type)
        
        # Skip knowledge base search - let Grok use its own knowledge
        relevant_chunks = []
        context_text = ""
        
        # Create prompt for Perplexity
        prompt = self._create_prompt(section_type, excel_data, context_text, reference_context)
        
        print(f"\n🤖 Generating AI content for: {section_type}", file=sys.stderr)
        print(f"   Using {self.provider.title()}'s general knowledge (no knowledge chunks)", file=sys.stderr)
        
        try:
            # Generate content using the configured AI provider
            print(f"   🔑 Using {self.provider.title()} API key: {self.api_key[:10]}... (length: {len(self.api_key)})", file=sys.stderr)
            print(f"   🤖 Model: {self.model}", file=sys.stderr)
            print(f"   📝 Prompt length: {len(prompt)} characters", file=sys.stderr)
            
            # Add delay to avoid rate limiting
            import time
            time.sleep(2)
            
            if self.provider == "gemini":
                # Use Gemini API
                response = self.client.generate_content(prompt[:8000])
                generated_text = response.text
            else:
                # Use OpenAI-style API for Grok and Perplexity
                response = self.client.chat.completions.create(
                    model=self.model,
                    messages=[{"role": "user", "content": prompt[:8000]}]  # Limit prompt length
                )
                generated_text = response.choices[0].message.content
            
            print(f"   ✅ Generated {len(generated_text)} characters", file=sys.stderr)
            return generated_text
            
        except Exception as e:
            print(f"   ❌ Error generating AI content: {str(e)}", file=sys.stderr)
            print(f"   🔍 Error type: {type(e).__name__}", file=sys.stderr)
            
            # Check for different types of errors based on provider
            error_str = str(e).lower()
            if "401" in str(e) or "authorization" in error_str or "invalid api key" in error_str:
                if self.provider == "grok":
                    print(f"   🔐 GROK AUTHENTICATION ERROR: Check if Grok API key is valid", file=sys.stderr)
                    print(f"   💡 Try regenerating your API key at https://console.x.ai/", file=sys.stderr)
                elif self.provider == "gemini":
                    print(f"   🔐 GEMINI AUTHENTICATION ERROR: Check if Gemini API key is valid", file=sys.stderr)
                    print(f"   💡 Try regenerating your API key at https://makersuite.google.com/app/apikey", file=sys.stderr)
                else:
                    print(f"   🔐 PERPLEXITY AUTHENTICATION ERROR: Check if Perplexity API key is valid and active", file=sys.stderr)
                    print(f"   💡 Try regenerating your API key at https://www.perplexity.ai/settings/api", file=sys.stderr)
            elif "403" in str(e) or "credits" in error_str or "permission" in error_str or "quota" in error_str:
                if self.provider == "grok":
                    print(f"   💰 GROK CREDITS ERROR: Your Grok account needs credits", file=sys.stderr)
                    print(f"   💡 Add credits at https://console.x.ai/team/7ca0b680-16db-4157-a272-9379e32ba4ce", file=sys.stderr)
                elif self.provider == "gemini":
                    print(f"   💰 GEMINI QUOTA ERROR: Check your Gemini API quota and billing", file=sys.stderr)
                    print(f"   💡 Check usage at https://makersuite.google.com/app/apikey", file=sys.stderr)
                else:
                    print(f"   🚫 PERPLEXITY PERMISSION ERROR: Check your account permissions", file=sys.stderr)
            
            return f"[AI Content Generation Failed: {str(e)}]"
    
    def _create_prompt(self, section_type: str, excel_data: Dict, context: str, reference_context: str) -> str:
        """Create efficient, targeted prompts for specific report sections."""

        # Base instructions (Strict professional guidelines)
        base = f"""You are a senior financial analyst writing a formal Detailed Project Report (DPR) for a bank loan.
Use the provided Excel data for context: {json.dumps(excel_data, separators=(',', ':'))}

STRICT WRITING RULES:
1. **NO PLACEHOLDERS**: Never use brackets like [Name], [Date], [CIN], [Address]. If specific data is missing, use generic professional terms (e.g., "The Company", "The Promoters", "The Unit", "Relevant Authorities") or omit the detail entirely.
2. **NO TEMPLATE REFERENCES**: Do NOT mention "template", "filename", "timestamp", "provided Excel data", "JSON", or "AI".
3. **PROFESSIONAL TONE**: Write as if YOU are the author. Be concise, direct, and business-like.
4. **FORMATTING**:
   - **Tables**: Use LaTeX table format. Example:
     \\begin{{table}}
     \\caption{{Title}}
     \\begin{{tabular}}{{|l|l|}}
     \\hline
     Header1 & Header2 \\\\
     \\hline
     Row1 & Row2 \\\\
     \\hline
     \\end{{tabular}}
     \\end{{table}}
   - **Headings**: Use LaTeX section commands for headings: \\section{{Title}}, \\subsection{{Subtitle}}.
   - **Lists**: Use LaTeX itemize/enumerate environments for bullet points and numbered lists:
     \\begin{{itemize}}
     \\item Point 1
     \\item Point 2
     \\end{{itemize}}
   - **Emphasis**: Use \\textbf{{bold}} and \\textit{{italic}} for emphasis.
   - Do NOT use Markdown formatting (like #, ##, *, -).
   - Do NOT use Markdown tables or [TABLE] markers.
5. **ASSUMPTIONS**: If data is missing, make reasonable professional assumptions for a standard manufacturing/service unit (e.g., "The unit is located in an industrial area with adequate infrastructure"), but do not explicitly state "assuming".

Task: Generate content for "{section_type}" section."""

        # Section-specific concise prompts
        prompts = {
            "constitution": base + """
Describe the Constitution of the Firm.
- State the legal entity type (Proprietorship/Partnership/Company) based on available data or imply a standard structure.
- Describe the management team generally if names are missing (e.g., "The unit is managed by experienced promoters with significant industry expertise").
- Mention the business activity and operational focus.
- Focus on stability, experience, and business continuity.""",

            "product_characteristics": base + """
Describe the Key Characteristics of the Product/Service.
- Explain the product's utility, application, and target market.
- Highlight unique selling propositions (USPs) such as quality, cost-effectiveness, or service delivery.
- Mention adherence to quality standards (e.g., "The unit follows strict quality control measures").
- Focus on marketability and demand dynamics.""",

            "swot_analysis": base + """
Perform a SWOT Analysis (Strengths, Weaknesses, Opportunities, Threats).
- Strengths: Strategic location, experienced management, established market network.
- Weaknesses: Working capital intensity, competition from unorganized sector.
- Opportunities: Growing domestic demand, potential for export, government support schemes.
- Threats: Raw material price volatility, changing regulatory norms.
Format as a structured analysis.""",

            "competitors": base + """
Provide an Overview of Competitors.
- Describe the competitive landscape (e.g., "The market is characterized by a mix of organized and unorganized players").
- Mention generic competitor types (e.g., "Large established brands", "Local regional players").
- Explain the unit's competitive strategy (e.g., "Focus on quality, timely delivery, and customer relationship management").""",

            "power_approvals": base + """
Detail Power & Statutory Approvals.
- Mention power requirements are met through the State Electricity Board with backup generators.
- State that "The unit has applied for/obtained necessary statutory approvals such as GST Registration, Udyam Registration, and Pollution Control Board clearances."
- Emphasize compliance with local regulations.""",

            "raw_materials": base + """
Describe Raw Materials.
- List typical raw materials required for this type of industry.
- Mention procurement from local and inter-state suppliers to ensure competitive pricing.
- Comment on the stability of the supply chain.
- Mention inventory policy (e.g., "The unit maintains adequate stock levels to ensure uninterrupted production").""",

            "production_process": base + """
Explain the Production/Service Process.
- Describe a standard workflow from raw material intake to finished goods.
- Mention the use of modern/semi-automatic technology to ensure efficiency.
- Highlight quality control checks at various stages (raw material inspection, in-process checks, final testing).""",

            "manpower": base + """
Detail Manpower Requirements.
- Mention the team composition: Administrative, Skilled, Semi-skilled, and Unskilled staff.
- Highlight the employment generation potential.
- Mention that "Skilled labor is available in the vicinity" and that "Training programs are conducted for skill enhancement."
- Mention welfare measures (PF, ESI) as per norms.""",

            "land_building": base + """
Describe Land & Building.
- Describe the location generally (e.g., "Located in a designated industrial area with good connectivity").
- Mention infrastructure availability: Water, Power, and Road access.
- State that the premises are suitable for the proposed manufacturing/service activities.""",

            "timeline": base + """
Provide an Implementation Timeline.
- Outline key phases: Project initiation, Civil works/Renovation, Machinery installation, Trial runs, and Commercial production.
- Provide a realistic schedule (e.g., "The project is expected to be fully operational within a standard implementation period").
- Focus on the readiness for commercial operations.""",

            "index_page": base + """
Create a Table of Contents with page numbers:
1. Cover Page
2. Trading, Profit & Loss Account
3. Balance Sheet
4. Administrative & Selling Expenses
5. Ratio Analysis - Part I
6. Ratio Analysis - Part II
7. Ratio Analysis - Part III
8. MPBF Calculation - Methods 1 & 2
9. MPBF Calculation - Turnover Method
10. Depreciation Calculation
11. Executive Summary

Format as numbered list with "Page X" for each item.""",

            "trading_pl_account": base + """
Generate Trading, Profit & Loss Account analysis. Focus on:
- Revenue trends and growth
- Cost structure analysis
- Profit margins (Gross, Operating, Net)
- Key ratios and efficiency metrics

Include 1-2 key insights about profitability trends.""",

            "balance_sheet_analysis": base + """
Analyze Balance Sheet structure:
- Asset composition and quality
- Liabilities and capital structure
- Working capital position
- Debt-equity ratio trends

Highlight financial stability and liquidity position.""",

            "admin_selling_expenses": base + """
Analyze Administrative & Selling Expenses:
- Expense breakdown and trends
- Cost control effectiveness
- Operating efficiency ratios
- Recommendations for cost optimization

Focus on expense management and efficiency.""",

            "ratio_analysis_part1": base + """
Analyze Current Ratio, Debtors Turnover, and Gross Profit Ratio:
- Liquidity assessment
- Receivables management efficiency
- Basic profitability indicators
- Industry comparisons and benchmarks

Provide brief interpretation of each ratio.""",

            "ratio_analysis_part2": base + """
Analyze Net Profit Ratio, Interest Coverage, Working Capital Turnover, and Stock Turnover:
- Overall profitability
- Debt servicing capacity
- Asset utilization efficiency
- Inventory management effectiveness

Include specific ratio calculations and interpretations.""",

            "ratio_analysis_part3": base + """
Analyze TOL/TNW Ratio and Return on Capital Employed:
- Financial leverage assessment
- Overall capital efficiency
- Risk-return profile
- Investment viability indicators

Provide banking norms comparison.""",

            "mpbf_methods_1_2": base + """
Explain MPBF calculation using Methods 1 (25% of Working Capital Gap) and 2 (25% of Current Assets):
- Methodology differences
- Calculation steps
- Recommended approach
- Banking guidelines compliance

Focus on working capital assessment.""",

            "mpbf_turnover_method": base + """
Explain MPBF Turnover Method calculation:
- Sales-based working capital assessment
- Four operating cycles assumption
- Drawing power limitations
- Permissible finance determination

Include Nayak Committee guidelines reference.""",

            "depreciation_calculation": base + """
Analyze depreciation calculation as per Income Tax Act:
- WDV method application
- Plant and machinery depreciation
- Tax implications
- Asset utilization over time

Show 5-year depreciation schedule analysis.""",

            "executive_summary": base + """
Create concise Executive Summary covering:
- Project overview and objectives
- Financial viability assessment
- Key financial indicators
- Risk factors and mitigation
- Recommendations for loan approval

Keep under 500 words, focus on critical decision points."""
        }

        return prompts.get(section_type, base + f" Generate professional content for {section_type} section using Excel data.")
    
    def create_text_pdf(self, content_sections: List[Dict[str, str]], output_path: str) -> bool:
        """
        Create a professionally formatted PDF from text content sections.
        Uses professional template matching reference PDF style.
        Automatically parses and renders tables from AI-generated content.
        
        Args:
            content_sections: List of {"title": "...", "content": "..."} dictionaries
            output_path: Path to save the generated PDF
            
        Returns:
            True if successful, False otherwise
        """
        try:
            print(f"\n📄 Creating professional PDF with {len(content_sections)} sections", file=sys.stderr)
            
            # Initialize professional template and table parser
            prof_template = ProfessionalTemplate()
            from table_parser_enhanced import TableParser
            table_parser = TableParser(prof_template)
            
            # Create PDF document with professional margins
            doc = SimpleDocTemplate(
                output_path, 
                pagesize=A4,
                topMargin=30*mm,
                bottomMargin=30*mm,
                leftMargin=25*mm,
                rightMargin=25*mm
            )
            
            # Get professional styles
            styles = prof_template.get_styles()
            
            # Build PDF content
            story = []
            
            # Add sections with table parsing
            for section in content_sections:
                # Add section title
                story.append(Paragraph(section['title'], styles['SectionHeading']))
                story.append(Spacer(1, 0.15*inch))
                
                # Parse content and extract tables
                content = section['content']
                elements = table_parser.parse_and_render_content(content, styles)
                
                # Add all parsed elements (paragraphs and tables)
                for element in elements:
                    story.append(element)
                    story.append(Spacer(1, 0.1*inch))
                
                story.append(Spacer(1, 0.2*inch))
            
            # Build PDF with professional template (adds headers/footers with green borders)
            doc.build(story, onFirstPage=prof_template.create_header_footer, 
                     onLaterPages=prof_template.create_header_footer)
            
            print(f"   ✅ Professional PDF created: {output_path}", file=sys.stderr)
            return True
            
        except Exception as e:
            print(f"   ❌ Error creating text PDF: {str(e)}", file=sys.stderr)
            import traceback
            traceback.print_exc(file=sys.stderr)
            return False
    
    def merge_pdfs(self, pdf_files: List[str], output_path: str, remove_blank_pages: bool = True) -> bool:
        """
        Merge multiple PDF files into one, optionally removing blank pages.
        
        Args:
            pdf_files: List of PDF file paths to merge
            output_path: Path for the merged PDF
            remove_blank_pages: If True, skip pages with no text content
            
        Returns:
            True if successful, False otherwise
        """
        try:
            print(f"\n📑 Merging {len(pdf_files)} PDF files", file=sys.stderr)
            if remove_blank_pages:
                print(f"   (removing blank pages)", file=sys.stderr)
            
            merger = PdfMerger()
            blank_pages_removed = 0
            
            for pdf_file in pdf_files:
                if not os.path.exists(pdf_file):
                    print(f"   ⚠️  Skipping missing file: {pdf_file}", file=sys.stderr)
                    continue
                    
                print(f"   Adding: {Path(pdf_file).name}", file=sys.stderr)
                
                if remove_blank_pages:
                    # Check each page for content before adding
                    # Use PyPDF2 instead of pdfplumber for speed
                    
                    # First, identify which pages have content
                    pages_to_add = []
                    try:
                        reader = PdfReader(pdf_file)
                        for page_num, page in enumerate(reader.pages):
                            text = page.extract_text() or ""
                            if text.strip():  # Page has content
                                pages_to_add.append(page_num)
                            else:
                                blank_pages_removed += 1
                                print(f"      ⏭️  Skipping blank page {page_num + 1}", file=sys.stderr)
                        
                        # Now add only non-blank pages
                        if pages_to_add:
                            for page_num in pages_to_add:
                                merger.append(pdf_file, pages=(page_num, page_num + 1))
                    except Exception as e:
                        print(f"      ⚠️  Error checking blank pages: {e}. Adding all pages.", file=sys.stderr)
                        merger.append(pdf_file)
                else:
                    # Add all pages without checking
                    merger.append(pdf_file)
            
            merger.write(output_path)
            merger.close()
            
            file_size = os.path.getsize(output_path)
            if blank_pages_removed > 0:
                print(f"   ✅ Merged PDF created: {output_path} ({file_size:,} bytes)", file=sys.stderr)
                print(f"   🗑️  Removed {blank_pages_removed} blank pages", file=sys.stderr)
            else:
                print(f"   ✅ Merged PDF created: {output_path} ({file_size:,} bytes)", file=sys.stderr)
            return True
            
        except Exception as e:
            print(f"   ❌ Error merging PDFs: {str(e)}", file=sys.stderr)
            import traceback
            traceback.print_exc(file=sys.stderr)
            return False
            
            merger.write(output_path)
            merger.close()
            
            file_size = os.path.getsize(output_path)
            if blank_pages_removed > 0:
                print(f"   ✅ Merged PDF created: {output_path} ({file_size:,} bytes)", file=sys.stderr)
                print(f"   🗑️  Removed {blank_pages_removed} blank pages", file=sys.stderr)
            else:
                print(f"   ✅ Merged PDF created: {output_path} ({file_size:,} bytes)", file=sys.stderr)
            return True
            
        except Exception as e:
            print(f"   ❌ Error merging PDFs: {str(e)}", file=sys.stderr)
            import traceback
            traceback.print_exc(file=sys.stderr)
            return False
    
    def generate_full_report(self, excel_pdfs_dir: str, excel_data: Dict[str, Any], 
                           output_path: str, reference_analysis: Dict = None, template_name: str = 'CC1') -> Dict[str, Any]:
        """
        Generate the complete report by combining Excel PDFs and AI-generated content.
        Follows reference report structure with proper sheet ordering and interspersed AI analysis.
        
        Args:
            excel_pdfs_dir: Directory containing individual Excel sheet PDFs
            excel_data: Computed data from Excel calculations
            output_path: Path for the final merged PDF report
            reference_analysis: Optional reference report analysis for context
            
        Returns:
            Dictionary with generation results and metadata
        """
        print(f"\n{'='*80}", file=sys.stderr)
        print(f"🚀 GENERATING FULL AI-ENHANCED REPORT", file=sys.stderr)
        print(f"{'='*80}\n", file=sys.stderr)
        
        result = {
            "success": False,
            "output_path": output_path,
            "ai_sections_generated": [],
            "excel_pdfs_included": [],
            "errors": []
        }
        
        try:
            # Load knowledge base
            self.load_knowledge_base()
            
            # Define proper sheet order based on reference PDF structure
            # Sheet names map: (sheet_index_prefix, sheet_name_in_file, display_order)
            # Determine sheet name based on template format
            final_sheet_name = get_final_sheet_name(template_name)
            
            is_term_loan = 'TERM LOAN' in template_name.upper() or 'TERM_LOAN' in template_name.upper()
            
            if is_term_loan:
                print("ℹ️  Using Term Loan specific sheet order", file=sys.stderr)
                SHEET_ORDER = [
                    ('sheet_x', 'Cover page', 1),
                    ('sheet_x', 'Index', 2),      # Excel Sheet
                    ('sheet_x', 'profile', 3),    # Excel Sheet
                    ('sheet_x', 'Descriptive', 4),# AI Narrative Sections Placeholder
                    ('sheet_x', 'project cost', 5),
                    ('sheet_x', 'PL BS', 6),
                    ('sheet_x', 'Graph', 7),
                    ('sheet_x', 'RATIO', 8),
                    ('sheet_x', 'FA Sch', 9),
                    ('sheet_x', 'Dep IT act', 10),
                    ('sheet_x', 'Loan sch', 11),
                    ('sheet_x', 'Repayment', 12),
                    ('sheet_x', 'CFs', 13),
                    ('sheet_x', 'IRR', 14),
                    ('sheet_x', 'MIRR', 15),
                    ('sheet_x', 'NPV', 16),
                    ('sheet_x', 'PI Index', 17),
                    ('sheet_x', 'WACC', 18),
                    ('sheet_x', 'Payback period I', 19),
                    ('sheet_x', 'Payback period II', 20),
                    ('sheet_x', 'Altman Z', 21),
                    ('sheet_x', 'Sensitivity Analysis', 22),
                    ('sheet_x', 'workings for sensittivity1', 23),
                    ('sheet_x', 'Workings for Sensitivity2', 24),
                    ('sheet_x', 'CF workings', 25),
                    ('sheet_x', 'Final workings', 26),
                    ('sheet_x', 'MPBF ', 27),
                    ('sheet_x', 'workings for sensitivity1', 28),
                    ('sheet_x', 'Gaurantors', 29),
                    ('sheet_x', 'BEP analysis', 30),
                    ('sheet_x', 'Sales', 31),
                ]
            else:
                SHEET_ORDER = [
                    ('sheet_4', 'coverpage', 1),           # ALWAYS FIRST
                    ('sheet_3', final_sheet_name, 2),       # Project Cost & Summary
                    ('sheet_5', 'PLBS', 3),                # Balance Sheet (P&L)
                    ('sheet_6', 'RATIO', 4),               # Ratio Analysis
                    ('sheet_9', 'Depsch', 5),              # Depreciation Schedule
                    ('sheet_7', 'MPBF ', 6),                # MPBF Method 1
                    ('sheet_8', 'nayak', 7),               # Nayak Committee (WC Assessment)
                    ('sheet_2', 'wp', 8),                  # Working Capital
                ]
            
            # Define AI sections with their placement in the report
            # Position is relative to Excel sheets in SHEET_ORDER
            AI_SECTIONS_CONFIG = [
                {
                    "type": "index_page",
                    "title": "Index / Table of Contents",
                    "after_sheet": "coverpage",  # After Coverpage sheet
                    "pages": 1
                },
                {
                    "type": "project_profile",
                    "title": "1. Project Profile Overview",
                    "after_ai": "index_page",
                    "pages": 1
                },
                {
                    "type": "constitution",
                    "title": "2. Constitution of the Firm",
                    "after_ai": "project_profile",
                    "pages": 1
                },
                {
                    "type": "product_characteristics",
                    "title": "3. Key Characteristics of Product",
                    "after_ai": "constitution",
                    "pages": 1
                },
                {
                    "type": "swot_analysis",
                    "title": "4. SWOT Analysis",
                    "after_ai": "product_characteristics",
                    "pages": 1
                },
                {
                    "type": "competitors",
                    "title": "5. Overview of Competitors",
                    "after_ai": "swot_analysis",
                    "pages": 1
                },
                {
                    "type": "power_approvals",
                    "title": "6. Power & Statutory Approvals",
                    "after_ai": "competitors",
                    "pages": 1
                },
                {
                    "type": "raw_materials",
                    "title": "7. Raw Materials",
                    "after_ai": "power_approvals",
                    "pages": 1
                },
                {
                    "type": "production_process",
                    "title": "8. Production Process",
                    "after_ai": "raw_materials",
                    "pages": 1
                },
                {
                    "type": "manpower",
                    "title": "9. Manpower Requirements",
                    "after_ai": "production_process",
                    "pages": 1
                },
                {
                    "type": "land_building",
                    "title": "10. Land & Building",
                    "after_ai": "manpower",
                    "pages": 1
                },
                {
                    "type": "timeline",
                    "title": "11. Implementation Timeline",
                    "after_ai": "land_building",
                    "pages": 1
                }
            ]
            
            # Collect all Excel sheet PDFs
            excel_pdf_map = {}
            if os.path.exists(excel_pdfs_dir):
                for filename in os.listdir(excel_pdfs_dir):
                    if filename.endswith('.pdf'):
                        pdf_path = os.path.join(excel_pdfs_dir, filename)
                        excel_pdf_map[filename] = pdf_path
            
            print(f"📊 Found {len(excel_pdf_map)} Excel sheet PDFs", file=sys.stderr)
            if not excel_pdf_map:
                error_msg = (
                    "No Excel sheet PDFs were found to merge into the final report. "
                    "Verify that Excel export succeeded before invoking the AI report generator."
                )
                print(f"⚠️  {error_msg}", file=sys.stderr)
                result["errors"].append(error_msg)
                return result
            
            # Generate all AI sections upfront
            print(f"\n🤖 Generating AI content sections...", file=sys.stderr)
            ai_content_pdfs = {}
            
            if is_term_loan:
                import concurrent.futures
                
                # Filter sections to generate
                sections_to_generate = []
                for section_config in AI_SECTIONS_CONFIG:
                    section_type = section_config["type"]
                    if section_type not in ['index_page', 'project_profile']:
                        sections_to_generate.append(section_config)
                
                print(f"🚀 Starting parallel generation for {len(sections_to_generate)} sections...", file=sys.stderr)
                
                def process_section(section_config):
                    section_type = section_config["type"]
                    section_title = section_config["title"]
                    
                    print(f"   ⏳ Starting: {section_title}", file=sys.stderr)
                    try:
                        content = self.generate_ai_content(section_type, excel_data)
                        
                        # Create individual PDF for this AI section
                        ai_pdf_path = output_path.replace('.pdf', f'_ai_{section_type}.pdf')
                        if self.create_text_pdf([{
                            "title": section_title,
                            "content": content
                        }], ai_pdf_path):
                            print(f"   ✅ Completed: {section_title}", file=sys.stderr)
                            return section_type, ai_pdf_path
                    except Exception as e:
                        print(f"   ❌ Failed: {section_title} - {str(e)}", file=sys.stderr)
                        return None
                    return None

                # Use ThreadPoolExecutor for parallel execution
                # Limit max_workers to avoid hitting API rate limits too hard (e.g., 5 concurrent requests)
                with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
                    future_to_section = {executor.submit(process_section, config): config for config in sections_to_generate}
                    
                    for future in concurrent.futures.as_completed(future_to_section):
                        result_tuple = future.result()
                        if result_tuple:
                            sec_type, pdf_path = result_tuple
                            ai_content_pdfs[sec_type] = pdf_path
                            result["ai_sections_generated"].append(sec_type)
            else:
                print("ℹ️  Skipping AI content generation for non-Term Loan report (CC/OD)", file=sys.stderr)
            
            # Build final PDF sequence according to proper order
            print(f"\n📑 Assembling final report in correct order...", file=sys.stderr)
            final_pdf_sequence = []
            position_counter = 1
            
            # Create a map of sheet names to their PDF paths for easy lookup
            sheet_pdf_map = {}
            
            # Pre-process all available PDFs to extract their sheet names
            available_pdf_info = []
            for filename, path in excel_pdf_map.items():
                extracted_name = extract_sheet_title_from_filename(filename)
                normalized_extracted = normalize_identifier(extracted_name)
                available_pdf_info.append({
                    'filename': filename,
                    'path': path,
                    'extracted_name': extracted_name,
                    'normalized_name': normalized_extracted
                })

            for prefix, sheet_name, order in SHEET_ORDER:
                normalized_sheet = normalize_identifier(sheet_name)
                
                # Strategy 1: Exact match on extracted name (Best for "Index" vs "PI Index")
                best_match = None
                for info in available_pdf_info:
                    if info['normalized_name'] == normalized_sheet:
                        best_match = info['path']
                        break
                
                # Strategy 2: Contained match (Fallback for "profile" vs "Project Profile")
                if not best_match:
                    for info in available_pdf_info:
                        if normalized_sheet in info['normalized_name']:
                            best_match = info['path']
                            break
                
                if best_match:
                    sheet_pdf_map[sheet_name] = best_match
            
            # Track which AI sections have been added
            added_ai_sections = set()
            
            def add_ai_section(section_type):
                """Add a specific AI section if it exists and hasn't been added"""
                if section_type not in added_ai_sections and section_type in ai_content_pdfs:
                    # Find the section config for the title
                    section_title = section_type
                    for config in AI_SECTIONS_CONFIG:
                        if config["type"] == section_type:
                            section_title = config["title"]
                            break
                    
                    final_pdf_sequence.append(ai_content_pdfs[section_type])
                    added_ai_sections.add(section_type)
                    nonlocal position_counter
                    print(f"   [{position_counter}] AI: {section_title}", file=sys.stderr)
                    position_counter += 1
                    return True
                return False
            
            def add_ai_sections_after_sheet(sheet_name):
                """Add all AI sections configured to appear after a specific sheet"""
                normalized_sheet = normalize_sheet_name(sheet_name)
                for config in AI_SECTIONS_CONFIG:
                    config_sheet = config.get("after_sheet", "")
                    if normalize_sheet_name(config_sheet) == normalized_sheet:
                        add_ai_section(config["type"])
            
            def add_ai_sections_after_ai(ai_type):
                """Add all AI sections configured to appear after another AI section"""
                for config in AI_SECTIONS_CONFIG:
                    if config.get("after_ai") == ai_type:
                        if add_ai_section(config["type"]):
                            # Recursively add any sections after this one
                            add_ai_sections_after_ai(config["type"])
            
            # 1. Add Coverpage first (case-insensitive search)
            coverpage_key = None
            for key in sheet_pdf_map.keys():
                norm_key = normalize_sheet_name(key)
                if norm_key == 'coverpage' or norm_key == 'cover page':
                    coverpage_key = key
                    break
            
            if coverpage_key:
                final_pdf_sequence.append(sheet_pdf_map[coverpage_key])
                result["excel_pdfs_included"].append(Path(sheet_pdf_map[coverpage_key]).name)
                print(f"   [{position_counter}] Coverpage: {Path(sheet_pdf_map[coverpage_key]).name}", file=sys.stderr)
                position_counter += 1
                
                # Add AI sections that come after coverpage
                add_ai_sections_after_sheet("Coverpage")
                # And add chain of AI sections
                for section_type in added_ai_sections.copy():
                    add_ai_sections_after_ai(section_type)
            else:
                # Fallback: If coverpage not found in map but exists in SHEET_ORDER as order 1
                # Try to find it by order
                for prefix, sheet_name, order in SHEET_ORDER:
                    if order == 1 and sheet_name in sheet_pdf_map:
                        final_pdf_sequence.append(sheet_pdf_map[sheet_name])
                        result["excel_pdfs_included"].append(Path(sheet_pdf_map[sheet_name]).name)
                        print(f"   [{position_counter}] Coverpage (by order): {Path(sheet_pdf_map[sheet_name]).name}", file=sys.stderr)
                        position_counter += 1
                        break
            
            # 2. Process remaining Excel sheets in order
            for prefix, sheet_name, order in SHEET_ORDER:
                if order == 1:  # Skip coverpage (already added)
                    continue
                
                # Skip excluded sheets
                if sheet_name.lower() in ['workings for pl1', 'wokings for pl2', 'workings for pl2']:
                    continue

                # Check if this is an AI section placeholder (Term Loan specific)
                if is_term_loan and sheet_name == "Descriptive":
                    # Add all descriptive sections in order
                    descriptive_sections = [
                        "constitution", "product_characteristics", "swot_analysis", 
                        "competitors", "power_approvals", "raw_materials", 
                        "production_process", "manpower", "land_building", "timeline"
                    ]
                    for section in descriptive_sections:
                        add_ai_section(section)
                    continue

                # Add the Excel sheet
                if sheet_name in sheet_pdf_map:
                    final_pdf_sequence.append(sheet_pdf_map[sheet_name])
                    if Path(sheet_pdf_map[sheet_name]).name not in result["excel_pdfs_included"]:
                        result["excel_pdfs_included"].append(Path(sheet_pdf_map[sheet_name]).name)
                    print(f"   [{position_counter}] Excel: {sheet_name} ({Path(sheet_pdf_map[sheet_name]).name})", file=sys.stderr)
                    position_counter += 1
                    
                    # Add AI sections that should appear after this sheet
                    add_ai_sections_after_sheet(sheet_name)
                    # And add chain of AI sections
                    for section_type in added_ai_sections.copy():
                        add_ai_sections_after_ai(section_type)
            
            # 3. Add sections marked as "end"
            for config in AI_SECTIONS_CONFIG:
                if config.get("position") == "end":
                    add_ai_section(config["type"])
            
            # 4. Add any remaining AI sections that weren't positioned (fallback)
            remaining_sections = set(ai_content_pdfs.keys()) - added_ai_sections
            if remaining_sections:
                print(f"\n   ⚠️  Adding {len(remaining_sections)} unpositioned sections:", file=sys.stderr)
                for section_type in remaining_sections:
                    for config in AI_SECTIONS_CONFIG:
                        if config["type"] == section_type:
                            final_pdf_sequence.append(ai_content_pdfs[section_type])
                            print(f"   [{position_counter}] AI: {config['title']} (fallback)", file=sys.stderr)
                            position_counter += 1
                            break
            
            missing_excel_files = sorted(
                set(excel_pdf_map.keys()) - set(result["excel_pdfs_included"])
            )
            if missing_excel_files:
                print("[Final Report] ℹ️  Adding remaining sheet PDFs to honour the requested list:", file=sys.stderr)
                for missing in missing_excel_files:
                    pdf_path = excel_pdf_map[missing]
                    final_pdf_sequence.append(pdf_path)
                    if missing not in result["excel_pdfs_included"]:
                        result["excel_pdfs_included"].append(missing)
                    sheet_label = extract_sheet_title_from_filename(missing)
                    print(
                        f"   [{position_counter}] Excel: {sheet_label} ({missing}) [appended]",
                        file=sys.stderr
                    )
                    position_counter += 1
            else:
                print("[Final Report] ✅ All generated sheet PDFs are already scheduled for merge", file=sys.stderr)

            total_excel_available = len(excel_pdf_map)
            total_excel_included = len(result["excel_pdfs_included"])
            print(f"\n📊 Total sections in final report: {len(final_pdf_sequence)}", file=sys.stderr)
            print(
                f"[Final Report] Excel sheet PDFs included: {total_excel_included}/{total_excel_available}",
                file=sys.stderr
            )
            
            # Merge all PDFs in the correct sequence
            if self.merge_pdfs(final_pdf_sequence, output_path):
                result["success"] = True
                result["total_sections"] = len(final_pdf_sequence)
                
                # Clean up temporary AI content PDFs
                for ai_pdf in ai_content_pdfs.values():
                    if os.path.exists(ai_pdf):
                        os.unlink(ai_pdf)
            
            print(f"\n{'='*80}", file=sys.stderr)
            print(f"✅ REPORT GENERATION COMPLETE", file=sys.stderr)
            print(f"   Output: {output_path}", file=sys.stderr)
            print(f"   AI Sections: {len(ai_content_pdfs)}", file=sys.stderr)
            print(f"   Excel Sheets: {len([x for x in result['excel_pdfs_included']])}", file=sys.stderr)
            print(f"   Total Sections: {result.get('total_sections', 0)}", file=sys.stderr)
            print(f"{'='*80}\n", file=sys.stderr)
            
        except Exception as e:
            print(f"\n❌ Error generating full report: {str(e)}", file=sys.stderr)
            import traceback
            traceback.print_exc(file=sys.stderr)
            result["errors"].append(str(e))
        
        return result


if __name__ == "__main__":
    # Test the AI Report Generator
    import argparse
    
    parser = argparse.ArgumentParser(description='Generate AI-enhanced PDF report')
    parser.add_argument('--api-key', required=True, help='AI API key (Grok, Perplexity, or Gemini)')
    parser.add_argument('--provider', default='gemini', choices=['grok', 'perplexity', 'gemini'], help='AI provider to use')
    parser.add_argument('--excel-pdfs-dir', required=True, help='Directory with Excel sheet PDFs')
    parser.add_argument('--output', required=True, help='Output PDF path')
    parser.add_argument('--excel-data', help='JSON file with Excel computed data')
    parser.add_argument('--template-name', default='CC1', help='Template name (e.g., CC6)')
    
    args = parser.parse_args()
    
    # Load Excel data if provided
    excel_data = {}
    if args.excel_data and os.path.exists(args.excel_data):
        with open(args.excel_data, 'r') as f:
            excel_data = json.load(f)
    
    # Generate report
    generator = AIReportGenerator(args.api_key, provider=args.provider)
    result = generator.generate_full_report(
        excel_pdfs_dir=args.excel_pdfs_dir,
        excel_data=excel_data,
        output_path=args.output,
        template_name=args.template_name
    )
    
    print(json.dumps(result, indent=2))
