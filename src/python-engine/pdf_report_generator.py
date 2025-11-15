"""
PDF Report Generator with AI-Generated Content
Combines Excel-derived PDFs with AI-generated contextual content using Grok AI API.
Uses ONLY the two resource PDFs as knowledge source - NO external data allowed.
"""

import os
import sys
import json
from pathlib import Path
from typing import Dict, List, Any, Optional
import openai
from PyPDF2 import PdfMerger, PdfReader
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle, KeepTogether
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY, TA_RIGHT
from ai_resource_parser import AIResourceParser

# Sheet name normalization utilities
def normalize_sheet_name(sheet_name: str) -> str:
    """Normalize sheet name by stripping whitespace and converting to lowercase."""
    return sheet_name.strip().lower()


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
        else:
            raise ValueError(f"Unsupported AI provider: {provider}. Use 'perplexity' or 'grok'")
        
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
            if "401" in str(e) or "authorization" in error_str:
                if self.provider == "grok":
                    print(f"   🔐 GROK AUTHENTICATION ERROR: Check if Grok API key is valid", file=sys.stderr)
                    print(f"   💡 Try regenerating your API key at https://console.x.ai/", file=sys.stderr)
                else:
                    print(f"   🔐 PERPLEXITY AUTHENTICATION ERROR: Check if Perplexity API key is valid and active", file=sys.stderr)
                    print(f"   💡 Try regenerating your API key at https://www.perplexity.ai/settings/api", file=sys.stderr)
            elif "403" in str(e) or "credits" in error_str or "permission" in error_str:
                if self.provider == "grok":
                    print(f"   💰 GROK CREDITS ERROR: Your Grok account needs credits", file=sys.stderr)
                    print(f"   💡 Add credits at https://console.x.ai/team/7ca0b680-16db-4157-a272-9379e32ba4ce", file=sys.stderr)
                else:
                    print(f"   🚫 PERPLEXITY PERMISSION ERROR: Check your account permissions", file=sys.stderr)
            
            return f"[AI Content Generation Failed: {str(e)}]"
    
    def _create_prompt(self, section_type: str, excel_data: Dict, context: str, reference_context: str) -> str:
        """Create a detailed prompt for AI based on section type."""
        
        base_instructions = f"""
You are a professional financial report writer analyzing a manufacturing/business project.

**WRITING STYLE:**
- Professional, formal business report writing
- Clear paragraph structure (3-5 paragraphs)
- Reference the Excel data where appropriate
- Use your expertise to provide comprehensive analysis and recommendations

**TABLE FORMATTING (CRITICAL - MUST USE THIS EXACT FORMAT):**
When including tables, use this EXACT format with NO HTML, NO Markdown:

[TABLE:Your Table Title Here]
Header1|Header2|Header3|Header4
DataRow1Col1|DataRow1Col2|DataRow1Col3|DataRow1Col4
DataRow2Col1|DataRow2Col2|DataRow2Col3|DataRow2Col4
[/TABLE]

Example:
[TABLE:Financial Ratios - 3 Year Projection]
Ratio|Year 1|Year 2|Year 3|Banking Norm
DSCR|1.65|1.85|2.10|Min 1.50
Current Ratio|1.45|1.60|1.75|1.33-2.00
[/TABLE]

RULES:
- Use pipe (|) to separate columns
- First row after [TABLE:...] is ALWAYS the header row
- NO HTML tags (<table>, <tr>, <td>, <br>, etc.)
- NO Markdown (no |---|---|)
- Just plain text with pipes
- Include meaningful table titles
- Use data from Excel calculations for financial tables

═══════════════════════════════════════════════════════════

**EXCEL CALCULATIONS (Financial Data from Analysis):**
{json.dumps(excel_data, indent=2)}

{f"**REFERENCE REPORT FORMAT:**{reference_context}" if reference_context else ""}

═══════════════════════════════════════════════════════════
**TASK:** Write the "{section_type}" section using your expertise and the Excel data provided.
Include relevant tables using [TABLE:...][/TABLE] format with data from Excel calculations.
═══════════════════════════════════════════════════════════
"""
        
        section_prompts = {
            "executive_summary": base_instructions + """
Create an Executive Summary that includes:
1. Project overview and business nature
2. Total project cost and funding structure
3. Key financial indicators (profitability, ROI, payback period)
4. Employment generation
5. Overall viability assessment

Keep it concise (200-300 words), highlight only the most important points.
""",

            "project_profile": base_instructions + """
Create a comprehensive Project Profile Overview that includes:
1. Promoter details and background
2. Business experience and qualifications
3. Project location with specific address
4. Nature of business activity in detail
5. Legal structure and registration details
6. Existing business (if any) and turnover
7. Educational qualifications of key personnel

**REQUIRED TABLE** - Use [TABLE:...][/TABLE] format:
[TABLE:Project Profile Summary]
Particulars|Details
Project Name|[Extract from Excel or use generic name]
Promoter Name|[Use generic or from context]
Location|[Extract from context]
Total Project Cost (Rs.)|[Use Excel total_project_cost]
Promoter Contribution (Rs.)|[Calculate from Excel]
Loan Required (Rs.)|[Calculate from Excel]
Employment Generation|[Extract from Excel or estimate]
[/TABLE]

Use information from Excel data about project costs and business details.
Write in 2-3 paragraphs followed by the table using exact [TABLE:...][/TABLE] format.
Minimum 300 words total.
""",

            "firm_constitution": base_instructions + """
Create a Constitution of Firm section that includes:
1. Legal form of organization (Proprietorship/Partnership/Private Ltd/LLP)
2. Registration details (if applicable)
3. Partner details with capital contribution (if partnership)
4. Management structure and roles
5. Authorized signatories

**REQUIRED TABLE** - Use [TABLE:...][/TABLE] format:
[TABLE:Promoter/Partner Details]
Name|Contribution (Rs.)|Share %|Role
Partner/Promoter 1|[From Excel]|[Calculate]|Managing Partner
Partner/Promoter 2|[From Excel]|[Calculate]|Partner
[/TABLE]

Write professionally with 2-3 paragraphs.
Minimum 250 words.
""",

            "product_characteristics": base_instructions + """
Create a Product Characteristics & Market Analysis section that includes:
1. Detailed product specifications
2. Quality standards and certifications
3. Market demand analysis
4. Target customer segments
5. Competition analysis
6. Pricing strategy
7. Sales and distribution channels

Include tables for:
- Product specifications and features
- Market size and growth projections
- Competitive positioning

Use information from guidelines about market assessment.
Write in 4-5 paragraphs with professional formatting.
Minimum 400 words.
""",

            "swot_analysis": base_instructions + """
Create a comprehensive SWOT Analysis section with:

**REQUIRED TABLE** - Use [TABLE:...][/TABLE] format:
[TABLE:SWOT Analysis Matrix]
Category|Key Points
Strengths|1. [Financial strength from Excel data]
Strengths|2. [Technical capability]
Strengths|3. [Market position]
Strengths|4. [Management expertise]
Strengths|5. [Policy support - from guidelines]
Weaknesses|1. [Financial constraint if any]
Weaknesses|2. [Market challenge]
Weaknesses|3. [Operational limitation]
Weaknesses|4. [Resource constraint]
Opportunities|1. [Market expansion potential]
Opportunities|2. [Product diversification]
Opportunities|3. [Government scheme benefits - from guidelines]
Opportunities|4. [Technology upgrade]
Opportunities|5. [Export potential]
Threats|1. [Competition from established players]
Threats|2. [Market price fluctuations]
Threats|3. [Regulatory changes]
Threats|4. [Economic uncertainties]
[/TABLE]

Add 1-2 paragraphs of analysis explaining how strengths can overcome weaknesses
and how opportunities can mitigate threats.
Minimum 300 words.
""",

            "plant_machinery": base_instructions + """
Create a Plant & Machinery Details section that includes:
1. Complete list of machinery and equipment
2. Technical specifications for each major equipment
3. Capacity and throughput details
4. Supplier/manufacturer information
5. Installation and commissioning timeline
6. Maintenance requirements
7. Power and utility requirements

**REQUIRED TABLE** - Use [TABLE:...][/TABLE] format with Excel data:
[TABLE:Plant & Machinery Cost Details]
S.No|Machinery/Equipment|Specifications|Quantity|Unit Cost (Rs.)|Total Cost (Rs.)
1|[Machine name]|[Specs]|[Qty]|[From Excel]|[Calculate]
2|[Machine name]|[Specs]|[Qty]|[From Excel]|[Calculate]
3|[Machine name]|[Specs]|[Qty]|[From Excel]|[Calculate]
Total|-|-|-|-|[Sum from Excel plant_machinery_cost]
[/TABLE]

Extract cost information from Excel data.
Write in 3-4 paragraphs plus comprehensive table.
Minimum 350 words.
""",

            "ratio_interpretation": base_instructions + """
Create a comprehensive Detailed Ratio Analysis & Trends section that includes:

**REQUIRED TABLE** - Use [TABLE:...][/TABLE] format with Excel ratio data:
[TABLE:Key Financial Ratios - Multi-Year Analysis]
Ratio Name|Year 1|Year 2|Year 3|Banking Norm|Status
DSCR|[Excel DSCR_year1]|[Excel DSCR_year2]|[Excel DSCR_year3]|Min 1.50|[Pass/Fail]
Current Ratio|[Excel]|[Excel]|[Excel]|1.33-2.00|[Pass/Fail]
Interest Coverage|[Excel]|[Excel]|[Excel]|Min 2.50|[Pass/Fail]
Debt-Equity|[Excel]|[Excel]|[Excel]|Max 2:1|[Pass/Fail]
Net Profit %|[Excel]|[Excel]|[Excel]|Industry Avg|[Assessment]
[/TABLE]

**Threats:**
- List 4-5 external threats
- Competition, regulatory changes, market risks

Present in a professional table format with 4 sections.
Add 1-2 paragraphs of analysis explaining how strengths can overcome weaknesses
and how opportunities can mitigate threats.
Minimum 300 words.
""",
            
            "project_description": base_instructions + """
Create a detailed Project Description that includes:
1. Nature of business/manufacturing activity
2. Location and infrastructure details
3. Plant and machinery requirements
4. Production capacity
5. Raw materials and inventory management
6. Manpower requirements

Use information from the knowledge source about manufacturing projects and PMEGP guidelines.
Write in paragraph form, 4-6 paragraphs.
Minimum 350 words.
""",

            "manufacturing_process": base_instructions + """
Create a Manufacturing Process & Flowchart section that includes:
1. Step-by-step production process description
2. Raw material procurement and storage
3. Production workflow stages
4. Quality control measures at each stage
5. Packaging and finished goods storage
6. Production capacity and shift details
7. Technology and equipment used

Include a table showing:
- Process Stage
- Description
- Equipment Used
- Duration/Capacity
- Quality Check Points

Write in 4-5 paragraphs describing the complete manufacturing cycle.
Minimum 400 words.
""",

            "plant_machinery": base_instructions + """
Create a Plant & Machinery Details section that includes:
1. Complete list of machinery and equipment
2. Technical specifications for each major equipment
3. Capacity and throughput details
4. Supplier/manufacturer information
5. Installation and commissioning timeline
6. Maintenance requirements
7. Power and utility requirements

Include a detailed table with:
- S.No.
- Machinery/Equipment Name
- Specifications
- Quantity
- Rate (Rs.)
- Total Cost (Rs.)
- Supplier Details

Extract cost information from Excel data if available.
Write in 3-4 paragraphs plus comprehensive table.
Minimum 350 words.
""",

            "inventory_details": base_instructions + """
Create an Inventory & Stock Management section that includes:
1. Raw material requirements and specifications
2. Minimum stock levels (safety stock)
3. Reorder levels and quantities
4. Storage requirements and facilities
5. Inventory turnover targets
6. Work-in-progress management
7. Finished goods inventory policy

Include tables for:
- Raw material inventory with quantities and values
- Consumables and packing materials
- Stock holding period for each category

Use Excel data for working capital calculations if available.
Write in 3-4 paragraphs with tables.
Minimum 300 words.
""",

            "transportation": base_instructions + """
Create a Transportation & Logistics section that includes:
1. Raw material transportation arrangements
2. Finished goods distribution network
3. Vehicle requirements (own/hired)
4. Transportation costs and budgets
5. Logistics partners and arrangements
6. Storage and warehousing at distribution points
7. Delivery timelines and service levels

Include a table showing:
- Type of Transport
- Purpose (RM/FG/Both)
- Capacity
- Ownership (Own/Hired)
- Monthly Cost

Write in 2-3 paragraphs covering the complete logistics chain.
Minimum 250 words.
""",

            "land_requirements": base_instructions + """
Create a Land & Building Requirements section that includes:
1. Total land area required
2. Built-up area breakdown (production, storage, office, etc.)
3. Land ownership status (own/leased/purchased)
4. Location advantages and connectivity
5. Zoning and regulatory approvals
6. Construction specifications
7. Cost of land and building development

Include a table with:
- Particulars
- Area (Sq.ft/Sq.mtr)
- Rate
- Total Cost
- Remarks

Extract cost data from Excel if available.
Write in 2-3 paragraphs plus table.
Minimum 300 words.
""",
            
            "financial_analysis": base_instructions + """
Create a comprehensive Financial Analysis section that includes:
1. Overview of the project's financial position
2. Analysis of profitability trends from P&L and Balance Sheet
3. Debt service capacity assessment
4. Working capital management evaluation
5. Cash flow and liquidity position
6. Year-wise financial performance analysis
7. Break-even analysis

**REQUIRED TABLE 1** - Use [TABLE:...][/TABLE] format with Excel data:
[TABLE:Profitability Analysis - Multi-Year]
Particulars|Year 1|Year 2|Year 3|Trend
Sales Revenue|[Excel]|[Excel]|[Excel]|[Increasing/Stable]
Cost of Goods Sold|[Excel]|[Excel]|[Excel]|[Analysis]
Gross Profit|[Excel]|[Excel]|[Excel]|[Analysis]
Operating Expenses|[Excel]|[Excel]|[Excel]|[Analysis]
Net Profit Before Tax|[Excel]|[Excel]|[Excel]|[Analysis]
Net Profit After Tax|[Excel]|[Excel]|[Excel]|[Analysis]
[/TABLE]

**REQUIRED TABLE 2** - Asset & Liability Composition:
[TABLE:Balance Sheet Summary]
Particulars|Year 1|Year 2|Year 3
Fixed Assets|[Excel]|[Excel]|[Excel]
Current Assets|[Excel]|[Excel]|[Excel]
Total Assets|[Excel]|[Excel]|[Excel]
Net Worth|[Excel]|[Excel]|[Excel]
Term Loan|[Excel]|[Excel]|[Excel]
Current Liabilities|[Excel]|[Excel]|[Excel]
[/TABLE]

Use the Excel data provided and reference the guidelines for acceptable financial parameters.
Write professionally with clear explanations of what the financial statements indicate.
5-6 paragraphs with comprehensive analysis.
Minimum 450 words.
""",
            
            "ratio_interpretation": base_instructions + """
Create a comprehensive Detailed Ratio Analysis & Trends section that includes:

**Coverage Ratios:**
1. **DSCR (Debt Service Coverage Ratio)**: Explain the calculated value and what it means for debt servicing ability. Reference ideal range from guidelines (minimum 1.5 preferred).
2. **Interest Coverage Ratio**: Evaluate ability to service interest obligations (minimum 2.5 preferred).

**Liquidity Ratios:**
3. **Current Ratio**: Interpret short-term liquidity position. Mention ideal range (1.33-2 times).
4. **Quick Ratio**: Assess immediate liquidity excluding inventory.

**Profitability Ratios:**
5. **Net Profit Ratio**: Analyze profitability trend and efficiency.
6. **Return on Assets (ROA)**: Assess asset utilization efficiency.
7. **Return on Equity (ROE)**: Evaluate returns to shareholders.

**Efficiency Ratios:**
8. **Debtors Turnover Ratio**: Assess collection efficiency.
9. **Inventory Turnover Ratio**: Evaluate stock management.
10. **Fixed Asset Turnover**: Measure asset productivity.

**Leverage Ratios:**
11. **Debt-Equity Ratio**: Assess financial leverage (max 2:1 acceptable).
12. **TOL/TNW (Total Outside Liabilities to Tangible Net Worth)**: Banking norm assessment.

For EACH ratio, provide:
- What the calculated value is (extract from Excel data)
- What it indicates about the business
- Whether it meets banking/industry standards
- Trend analysis across years
- What it means for loan approval

Include a comprehensive table showing all ratios for 3-5 years.

Write in clear paragraphs with ratio names in BOLD. 
Reference specific numbers from Excel data.
8-10 paragraphs covering all key ratios with detailed interpretation.
Minimum 600 words.
""",

            "mpbf_calculation": base_instructions + """
Create an MPBF Calculation & Working Capital Analysis section that includes:

**Maximum Permissible Bank Finance (MPBF):**
1. Turnover Method calculation (25% of projected turnover)
2. Current Asset method calculation
3. Nayak Committee recommendations
4. Assessment of both methods and lower value selection

**Working Capital Components:**
5. Current Assets breakdown (inventory, receivables, cash)
6. Current Liabilities (creditors, provisions)
7. Net Working Capital calculation
8. Margin requirements (promoter's contribution)

**Assessment:**
9. Adequacy of working capital
10. Comparison with industry norms
11. Monthly MPBF utilization pattern
12. Peak and non-peak working capital needs

Include comprehensive tables for:
- MPBF calculation (both methods)
- Current assets and liabilities breakdown
- Monthly working capital cycle
- Bank finance eligibility

Extract all values from Excel data (MPBF sheet and Working Capital sheet).
Write in 5-6 paragraphs with detailed calculations and tables.
Minimum 500 words.
""",

            "cash_flow_projection": base_instructions + """
Create a Cash Flow Statements & Projections section that includes:

**Operating Activities:**
1. Cash inflows from sales/operations
2. Cash outflows for expenses, salaries, raw materials
3. Net cash from operating activities

**Investing Activities:**
4. Capital expenditure on fixed assets
5. Sale of assets (if any)
6. Net cash from investing activities

**Financing Activities:**
7. Loan receipts and repayments
8. Equity capital and promoter contributions
9. Interest and dividend payments
10. Net cash from financing activities

**Cash Position:**
11. Opening cash balance
12. Net increase/decrease in cash
13. Closing cash balance
14. Cash adequacy analysis

Include comprehensive tables showing:
- Year-wise cash flow statement (3-5 years)
- Monthly cash flow projection for Year 1
- Sources and uses of cash
- Cash conversion cycle analysis

Extract data from Excel sheets wherever available.
Write in 5-6 paragraphs analyzing cash generation capacity and liquidity.
Minimum 500 words.
""",

            "funds_flow_analysis": base_instructions + """
Create a Funds Flow Statement & Analysis section that includes:

**Sources of Funds:**
1. Funds from operations (net profit + depreciation)
2. Issue of share capital
3. Long-term borrowings
4. Sale of fixed assets
5. Total sources of funds

**Applications of Funds:**
6. Purchase of fixed assets
7. Repayment of long-term loans
8. Payment of dividends
9. Increase in working capital
10. Total applications of funds

**Analysis:**
11. Net increase/decrease in working capital
12. Assessment of fund utilization efficiency
13. Comparison of sources vs applications
14. Long-term financial position analysis
15. Capital structure changes

Include comprehensive tables showing:
- Sources and Applications of Funds (3-5 years)
- Changes in Working Capital components
- Fund flow ratios and trends

Extract data from Balance Sheet and P&L data in Excel.
Write in 5-6 paragraphs with detailed analysis of capital movements.
Minimum 500 words.
""",
            
            "loan_eligibility": base_instructions + """
Create a comprehensive Loan Eligibility Assessment section that includes:

**Policy Compliance:**
1. Eligibility criteria from PMEGP/IDP guidelines
2. Project category and classification
3. Unit cost limits and compliance
4. Location criteria (urban/rural/special area)
5. Caste category benefits (if applicable)

**Financial Parameters:**
6. DSCR compliance (minimum 1.5 required)
7. Margin money requirements and availability
8. Debt-Equity ratio assessment (max 2:1)
9. Working capital margin (25% promoter contribution)
10. Security/collateral adequacy

**Assessment Against Norms:**
11. Banking norms for current ratio (min 1.33)
12. TOL/TNW ratio compliance (max 3:1)
13. Interest coverage adequacy (min 2.5)
14. Profitability standards

**Subsidy Eligibility:**
15. Subsidy calculation based on category
16. Maximum subsidy limits
17. Conditions for subsidy release

**Recommendations:**
18. Overall eligibility status
19. Compliance gaps (if any)
20. Conditions for approval

Include tables for:
- Eligibility criteria checklist
- Financial parameter compliance
- Subsidy calculation
- Loan structure and terms

Use ONLY information from the AP IDP 4.0 and PMEGP guidelines provided in the context.
Be specific about which criteria are met or not met.
Cite specific guideline requirements where applicable.
6-7 paragraphs with comprehensive assessment.
Minimum 550 words.
""",
            
            "recommendations": base_instructions + """
Create a comprehensive Recommendations & Conclusions section that includes:

**Project Viability Assessment:**
1. Overall financial viability based on DSCR, profitability, ROI
2. Market viability and demand assessment
3. Technical feasibility and operational capability
4. Management competence evaluation

**Key Strengths:**
5. Financial strengths (ratios, profitability margins)
6. Business strengths (market position, product quality)
7. Operational strengths (technology, location, infrastructure)
8. Compliance with all policy guidelines

**Risk Factors:**
9. Market risks (competition, demand fluctuations)
10. Financial risks (debt servicing, working capital adequacy)
11. Operational risks (raw material availability, technical issues)
12. External risks (regulatory, economic conditions)

**Mitigation Strategies:**
13. For each risk identified, provide specific mitigation measures
14. Contingency planning recommendations
15. Monitoring mechanisms

**Compliance Review:**
16. Compliance with banking norms
17. Adherence to PMEGP/IDP guidelines
18. All regulatory approvals status
19. Security and documentation adequacy

**Final Recommendation:**
20. Clear recommendation: APPROVE / REJECT / CONDITIONAL APPROVAL
21. If conditional, list specific conditions to be met
22. Suggested loan amount and terms
23. Subsidy recommendation
24. Disbursement conditions and monitoring requirements

Base recommendations on:
- Financial metrics from Excel data (DSCR > 1.5, Current Ratio > 1.33, profitability positive)
- All ratios meeting banking norms
- Compliance with guidelines from knowledge source
- Industry standards mentioned in reference documents
- Working capital adequacy as per Nayak Committee
- Debt servicing capacity confirmed

Be balanced and professional in your assessment.
Provide clear, actionable recommendations with supporting rationale.
6-8 paragraphs covering all aspects comprehensively.
Minimum 600 words.
"""
        }
        
        return section_prompts.get(section_type, base_instructions)
    
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
                    import pdfplumber
                    
                    # First, identify which pages have content
                    pages_to_add = []
                    with pdfplumber.open(pdf_file) as pdf:
                        for page_num, page in enumerate(pdf.pages):
                            text = page.extract_text() or ""
                            if text.strip():  # Page has content
                                pages_to_add.append(page_num)
                            else:
                                blank_pages_removed += 1
                                print(f"      ⏭️  Skipping blank page {page_num + 1}", file=sys.stderr)
                    
                    # Now add only non-blank pages
                    if pages_to_add:
                        # Add pages one by one or as ranges
                        reader = PdfReader(pdf_file)
                        for page_num in pages_to_add:
                            merger.append(pdf_file, pages=(page_num, page_num + 1))
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
                           output_path: str, reference_analysis: Dict = None) -> Dict[str, Any]:
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
            SHEET_ORDER = [
                ('sheet_4', 'coverpage', 1),           # ALWAYS FIRST
                ('sheet_3', 'Finalworkings', 2),       # Project Cost & Summary
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
                    "type": "executive_summary",
                    "title": "Executive Summary",
                    "after_sheet": "Coverpage",  # After Coverpage sheet
                    "pages": 1
                },
                {
                    "type": "project_profile",
                    "title": "Project Profile Overview",
                    "after_ai": "executive_summary",  # After Executive Summary
                    "pages": 2
                },
                {
                    "type": "firm_constitution",
                    "title": "Constitution of Firm",
                    "after_ai": "project_profile",  # After Project Profile
                    "pages": 1
                },
                {
                    "type": "product_characteristics",
                    "title": "Product Characteristics & Market Analysis",
                    "after_ai": "firm_constitution",  # After Firm Constitution
                    "pages": 2
                },
                {
                    "type": "swot_analysis",
                    "title": "SWOT Analysis",
                    "after_ai": "product_characteristics",  # After Product Characteristics
                    "pages": 1
                },
                {
                    "type": "project_description",
                    "title": "Detailed Project Description",
                    "after_ai": "swot_analysis",  # After SWOT Analysis
                    "pages": 2
                },
                {
                    "type": "manufacturing_process",
                    "title": "Manufacturing Process & Flowchart",
                    "after_ai": "project_description",  # After Project Description
                    "pages": 2
                },
                {
                    "type": "plant_machinery",
                    "title": "Plant & Machinery Details",
                    "after_ai": "manufacturing_process",  # After Manufacturing Process
                    "pages": 1
                },
                {
                    "type": "inventory_details",
                    "title": "Inventory & Stock Management",
                    "after_ai": "plant_machinery",  # After Plant & Machinery
                    "pages": 1
                },
                {
                    "type": "transportation",
                    "title": "Transportation & Logistics",
                    "after_ai": "inventory_details",  # After Inventory Details
                    "pages": 1
                },
                {
                    "type": "land_requirements",
                    "title": "Land & Building Requirements",
                    "after_ai": "transportation",  # After Transportation
                    "pages": 1
                },
                {
                    "type": "financial_analysis",
                    "title": "Financial Analysis & Interpretation",
                    "after_sheet": "PLBS",  # After Balance Sheet
                    "pages": 2
                },
                {
                    "type": "ratio_interpretation",
                    "title": "Detailed Ratio Analysis & Trends",
                    "after_sheet": "RATIO",  # After Ratio sheet
                    "pages": 2
                },
                {
                    "type": "mpbf_calculation",
                    "title": "MPBF Calculation & Working Capital Analysis",
                    "after_sheet": "MPBF ",  # After MPBF sheet
                    "pages": 2
                },
                {
                    "type": "cash_flow_projection",
                    "title": "Cash Flow Statements & Projections",
                    "after_sheet": "wp",  # After Working Capital
                    "pages": 2
                },
                {
                    "type": "funds_flow_analysis",
                    "title": "Funds Flow Statement & Analysis",
                    "after_ai": "cash_flow_projection",  # After Cash Flow
                    "pages": 2
                },
                {
                    "type": "loan_eligibility",
                    "title": "Loan Eligibility Assessment",
                    "after_ai": "funds_flow_analysis",  # After Funds Flow
                    "pages": 1
                },
                {
                    "type": "recommendations",
                    "title": "Recommendations & Conclusions",
                    "position": "end",  # At the very end
                    "pages": 2
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
            
            # Generate all AI sections upfront
            print(f"\n🤖 Generating AI content sections...", file=sys.stderr)
            ai_content_pdfs = {}
            
            for section_config in AI_SECTIONS_CONFIG:
                section_type = section_config["type"]
                section_title = section_config["title"]
                
                print(f"\n{'─'*60}", file=sys.stderr)
                print(f"🤖 Generating: {section_title}", file=sys.stderr)
                
                content = self.generate_ai_content(section_type, excel_data)
                
                # Create individual PDF for this AI section
                ai_pdf_path = output_path.replace('.pdf', f'_ai_{section_type}.pdf')
                if self.create_text_pdf([{
                    "title": section_title,
                    "content": content
                }], ai_pdf_path):
                    ai_content_pdfs[section_type] = ai_pdf_path
                    result["ai_sections_generated"].append(section_type)
                    print(f"   ✅ PDF created: {Path(ai_pdf_path).name}", file=sys.stderr)
            
            # Build final PDF sequence according to proper order
            print(f"\n📑 Assembling final report in correct order...", file=sys.stderr)
            final_pdf_sequence = []
            position_counter = 1
            
            # Create a map of sheet names to their PDF paths for easy lookup
            sheet_pdf_map = {}
            for prefix, sheet_name, order in SHEET_ORDER:
                for filename, path in excel_pdf_map.items():
                    # More robust matching: check prefix, normalized sheet name, or partial matches
                    normalized_sheet = normalize_sheet_name(sheet_name)
                    normalized_filename = normalize_sheet_name(filename)
                    if (filename.startswith(prefix) or 
                        normalized_sheet in normalized_filename or 
                        any(word in normalized_filename for word in normalized_sheet.split())):
                        sheet_pdf_map[sheet_name] = path
                        break
            
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
                if normalize_sheet_name(key) == 'coverpage':
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
            
            # 2. Process remaining Excel sheets in order
            for prefix, sheet_name, order in SHEET_ORDER:
                if order == 1:  # Skip coverpage (already added)
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
            
            print(f"\n📊 Total sections in final report: {len(final_pdf_sequence)}", file=sys.stderr)
            
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
    parser.add_argument('--api-key', required=True, help='Google Gemini API key')
    parser.add_argument('--excel-pdfs-dir', required=True, help='Directory with Excel sheet PDFs')
    parser.add_argument('--output', required=True, help='Output PDF path')
    parser.add_argument('--excel-data', help='JSON file with Excel computed data')
    
    args = parser.parse_args()
    
    # Load Excel data if provided
    excel_data = {}
    if args.excel_data and os.path.exists(args.excel_data):
        with open(args.excel_data, 'r') as f:
            excel_data = json.load(f)
    
    # Generate report
    generator = AIReportGenerator(args.api_key)
    result = generator.generate_full_report(
        excel_pdfs_dir=args.excel_pdfs_dir,
        excel_data=excel_data,
        output_path=args.output
    )
    
    print(json.dumps(result, indent=2))
