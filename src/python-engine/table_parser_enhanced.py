"""
Enhanced Table Parser for AI-Generated Content
Supports both structured markers and Markdown tables
Uses ProfessionalTemplate for consistent formatting
"""

import re
from typing import List, Dict, Tuple, Any
from reportlab.lib.units import inch
from reportlab.platypus import Table, Paragraph, ListFlowable, ListItem
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from professional_pdf_template import ProfessionalTemplate


class TableParser:
    """Parse and render tables from AI-generated content"""
    
    def __init__(self, prof_template: ProfessionalTemplate):
        self.prof_template = prof_template
        self.styles = prof_template.get_styles()
    
    def sanitize_html_content(self, text: str) -> str:
        """
        Remove or escape problematic HTML that ReportLab can't handle.
        Keeps only safe tags: <b>, <i>, <u>, <br/> (self-closing)
        """
        if not text:
            return ""
        
        # Replace self-closing <br> with <br/> (proper XML format)
        text = re.sub(r'<br\s*>', '<br/>', text)
        
        # Remove <para> tags (ReportLab adds these automatically)
        text = re.sub(r'</?para>', '', text)
        
        # Remove problematic table-related HTML
        text = re.sub(r'</?table[^>]*>', '', text, flags=re.IGNORECASE)
        text = re.sub(r'</?tr[^>]*>', '', text, flags=re.IGNORECASE)
        text = re.sub(r'</?td[^>]*>', '', text, flags=re.IGNORECASE)
        text = re.sub(r'</?th[^>]*>', '', text, flags=re.IGNORECASE)
        text = re.sub(r'</?thead[^>]*>', '', text, flags=re.IGNORECASE)
        text = re.sub(r'</?tbody[^>]*>', '', text, flags=re.IGNORECASE)
        
        # Remove any other unclosed tags
        text = re.sub(r'<(?!/?[biu]|br/)([^>]+)>', '', text)
        
        return text.strip()
    
    def parse_structured_table(self, content: str) -> List[Tuple[str, List[List[str]]]]:
        """
        Parse structured table markers from AI content.
        
        Format:
        [TABLE:Title]
        Header1|Header2|Header3
        Row1Col1|Row1Col2|Row1Col3
        Row2Col1|Row2Col2|Row2Col3
        [/TABLE]
        
        Returns:
            List of (table_title, table_data) tuples
        """
        tables = []
        
        # Pattern to match table blocks
        pattern = r'\[TABLE:([^\]]+)\](.*?)\[/TABLE\]'
        matches = re.finditer(pattern, content, re.DOTALL | re.IGNORECASE)
        
        for match in matches:
            title = match.group(1).strip()
            table_content = match.group(2).strip()
            
            # Split into rows
            rows = [line.strip() for line in table_content.split('\n') if line.strip()]
            
            # Parse each row (pipe-separated)
            table_data = []
            for row in rows:
                cells = [cell.strip() for cell in row.split('|')]
                if cells:  # Only add non-empty rows
                    table_data.append(cells)
            
            if table_data:
                tables.append((title, table_data))
        
        return tables
    
    def parse_markdown_table(self, content: str) -> List[Tuple[str, List[List[str]]]]:
        """
        Parse Markdown tables from AI content.
        
        Format:
        | Header1 | Header2 | Header3 |
        |---------|---------|---------|
        | Row1C1  | Row1C2  | Row1C3  |
        | Row2C1  | Row2C2  | Row2C3  |
        
        Returns:
            List of (table_title, table_data) tuples
        """
        tables = []
        
        # Pattern to match markdown tables
        # Looks for lines starting with | and containing |
        lines = content.split('\n')
        
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            
            # Check if line starts a markdown table
            if line.startswith('|') and line.endswith('|'):
                table_data = []
                table_title = "Data Table"  # Default title
                
                # Check if there's a title in the previous line
                if i > 0:
                    prev_line = lines[i-1].strip()
                    if prev_line and not prev_line.startswith('|') and len(prev_line) < 100:
                        table_title = prev_line.strip('*#-_ ')
                
                # Collect all table rows
                while i < len(lines):
                    line = lines[i].strip()
                    if not line.startswith('|') or not line.endswith('|'):
                        break
                    
                    # Skip separator rows (like |---|---|)
                    if re.match(r'\|[\s\-:|]+\|', line):
                        i += 1
                        continue
                    
                    # Parse cells
                    cells = [cell.strip() for cell in line.split('|')[1:-1]]  # Skip first/last empty
                    if cells and any(cell for cell in cells):  # At least one non-empty cell
                        table_data.append(cells)
                    
                    i += 1
                
                if table_data:
                    tables.append((table_title, table_data))
                
                continue
            
            i += 1
        
        return tables
    
    def parse_latex_table(self, content: str) -> List[Tuple[str, List[List[str]]]]:
        """
        Parse LaTeX tables from AI content.
        
        Format:
        \\begin{table}
        \\caption{Title}
        \\begin{tabular}{|c|c|}
        Header1 & Header2 \\\\
        Row1 & Row2 \\\\
        \\end{tabular}
        \\end{table}
        
        Returns:
            List of (table_title, table_data) tuples
        """
        tables = []
        
        # Pattern to match LaTeX table blocks (capturing title and content)
        # We look for the tabular environment, and optionally a surrounding table env with caption
        
        # Strategy: Find all tabular environments
        tabular_pattern = r'\\begin\{tabular\}(?:\{.*?\})?(.*?)\\end\{tabular\}'
        matches = re.finditer(tabular_pattern, content, re.DOTALL | re.IGNORECASE)
        
        for match in matches:
            tabular_content = match.group(1).strip()
            
            # Try to find a caption/title nearby (before the tabular)
            start_pos = match.start()
            preceding_text = content[:start_pos]
            
            # Look for \caption{...} in the preceding 300 chars (inside a table env)
            caption_match = re.search(r'\\caption\{([^}]+)\}', preceding_text[-300:])
            title = caption_match.group(1) if caption_match else "Data Table"
            
            # Parse rows
            # Split by \\ (row terminator)
            rows = tabular_content.split(r'\\')
            
            table_data = []
            for row in rows:
                row = row.strip()
                if not row:
                    continue
                
                # Remove \hline and other common commands
                row = row.replace(r'\hline', '')
                row = row.replace(r'\toprule', '')
                row = row.replace(r'\midrule', '')
                row = row.replace(r'\bottomrule', '')
                
                if not row.strip():
                    continue
                
                # Split by & (column separator)
                # Handle escaped & (\&) - simplistic approach: replace \& with placeholder, split, then restore
                row_safe = row.replace(r'\%', 'PERCENT_PLACEHOLDER').replace(r'\&', 'AMP_PLACEHOLDER').replace(r'\$', 'DOLLAR_PLACEHOLDER')
                
                cells = [cell.strip() for cell in row_safe.split('&')]
                
                # Clean up cells and restore special chars
                clean_cells = []
                for cell in cells:
                    cell = cell.replace('AMP_PLACEHOLDER', '&').replace('PERCENT_PLACEHOLDER', '%').replace('DOLLAR_PLACEHOLDER', '$')
                    
                    # Remove common LaTeX formatting
                    cell = re.sub(r'\\textbf\{([^}]+)\}', r'\1', cell)
                    cell = re.sub(r'\\textit\{([^}]+)\}', r'\1', cell)
                    cell = re.sub(r'\\underline\{([^}]+)\}', r'\1', cell)
                    
                    clean_cells.append(cell.strip())
                
                if clean_cells and any(c for c in clean_cells):
                    table_data.append(clean_cells)
            
            if table_data:
                tables.append((title, table_data))
                
        return tables

    def extract_tables_from_content(self, content: str) -> Tuple[str, List[Tuple[str, List[List[str]]]]]:
        """
        Extract all tables from content and return sanitized text + table data.
        
        Returns:
            (sanitized_text_without_tables, list_of_tables)
        """
        # 1. Try LaTeX tables (Highest Priority)
        latex_tables = self.parse_latex_table(content)
        
        # Remove LaTeX table blocks
        # Remove \begin{table}...\end{table} blocks
        clean_content = re.sub(r'\\begin\{table\}.*?\\end\{table\}', '\n\n[TABLE_PLACEHOLDER]\n\n', 
                              content, flags=re.DOTALL | re.IGNORECASE)
        # Also remove standalone tabulars if any remain (that weren't in table envs)
        clean_content = re.sub(r'\\begin\{tabular\}.*?\\end\{tabular\}', '\n\n[TABLE_PLACEHOLDER]\n\n', 
                              clean_content, flags=re.DOTALL | re.IGNORECASE)

        # 2. Try structured markers
        structured_tables = self.parse_structured_table(clean_content)
        
        # Remove structured table blocks from content
        clean_content = re.sub(r'\[TABLE:[^\]]+\].*?\[/TABLE\]', '\n\n[TABLE_PLACEHOLDER]\n\n', 
                              clean_content, flags=re.DOTALL | re.IGNORECASE)
        
        # 3. Then try markdown tables
        markdown_tables = self.parse_markdown_table(clean_content)
        
        # Remove markdown tables from content (basic pattern)
        clean_content = re.sub(r'(\|[^\n]+\|\n)+', '\n\n[TABLE_PLACEHOLDER]\n\n', clean_content)
        
        # Combine all tables
        all_tables = latex_tables + structured_tables + markdown_tables
        
        # Sanitize remaining content
        clean_content = self.sanitize_html_content(clean_content)
        
        return clean_content, all_tables
    
    def create_table_element(self, title: str, data: List[List[str]], 
                           has_header: bool = True) -> Table:
        """
        Create a formatted table using ProfessionalTemplate.
        
        Args:
            title: Table title
            data: 2D list of table data (first row is header if has_header=True)
            has_header: Whether first row is a header
            
        Returns:
            ReportLab Table object
        """
        if not data:
            return None
        
        # Calculate column widths based on content
        num_cols = len(data[0]) if data else 0
        if num_cols == 0:
            return None
        
        # Smart column width calculation based on content length
        available_width = 6.5  # inches (increased from 6.0)
        
        # Calculate average content length for each column
        col_avg_lengths = []
        for col_idx in range(num_cols):
            total_length = 0
            count = 0
            for row in data:
                if col_idx < len(row):
                    total_length += len(str(row[col_idx]))
                    count += 1
            col_avg_lengths.append(total_length / count if count > 0 else 1)
        
        # Calculate proportional widths
        total_avg = sum(col_avg_lengths)
        if total_avg > 0:
            col_widths = [(length / total_avg) * available_width * inch for length in col_avg_lengths]
        else:
            # Fallback to equal widths
            col_widths = [available_width / num_cols * inch] * num_cols
        
        # Apply minimum and maximum constraints
        min_width = 0.8 * inch  # Minimum 0.8 inch per column
        max_width = 5.0 * inch  # Maximum 5 inch per column
        
        col_widths = [max(min_width, min(max_width, w)) for w in col_widths]
        
        # Adjust if total exceeds available width
        total_width = sum(col_widths)
        if total_width > available_width * inch:
            scale = (available_width * inch) / total_width
            col_widths = [w * scale for w in col_widths]
        
        # Wrap cell content in Paragraph objects for better text wrapping
        from reportlab.platypus import Paragraph
        from reportlab.lib.styles import getSampleStyleSheet
        
        styles = getSampleStyleSheet()
        cell_style = ParagraphStyle(
            'CellStyle',
            parent=styles['Normal'],
            fontName=self.prof_template.primary_font,
            fontSize=8,  # Reduced from 9 to 8 for more content
            leading=10,  # Reduced from 11 to 10
            alignment=TA_LEFT,
            wordWrap='CJK',
        )
        
        header_style = ParagraphStyle(
            'HeaderStyle',
            parent=styles['Normal'],
            fontName=self.prof_template.font_bold,
            fontSize=9,  # Reduced from 10 to 9
            leading=11,  # Reduced from 12 to 11
            alignment=TA_CENTER,
            wordWrap='CJK',
        )
        
        # Convert data to Paragraph objects
        wrapped_data = []
        for row_idx, row in enumerate(data):
            wrapped_row = []
            for col_idx, cell in enumerate(row):
                cell_text = str(cell).strip()
                # Use header style for first row if has_header
                if row_idx == 0 and has_header:
                    wrapped_row.append(Paragraph(cell_text, header_style))
                else:
                    wrapped_row.append(Paragraph(cell_text, cell_style))
            wrapped_data.append(wrapped_row)
        
        # Use ProfessionalTemplate's method for consistent styling
        table = self.prof_template.create_professional_table(
            data=wrapped_data,
            col_widths=col_widths,
            has_header=has_header,
            stripe=True  # Alternating row colors for readability
        )
        
        return table

    def replace_inline_latex(self, text: str) -> str:
        """Convert inline LaTeX formatting to ReportLab XML tags."""
        if not text:
            return ""
        # Bold
        text = re.sub(r'\\textbf\{([^}]+)\}', r'<b>\1</b>', text)
        # Italic
        text = re.sub(r'\\textit\{([^}]+)\}', r'<i>\1</i>', text)
        # Underline
        text = re.sub(r'\\underline\{([^}]+)\}', r'<u>\1</u>', text)
        # Escaped characters
        text = text.replace(r'\%', '%').replace(r'\&', '&').replace(r'\$', '$')
        return text

    def extract_lists_from_content(self, content: str) -> Tuple[str, List[List[str]]]:
        """
        Extract LaTeX lists.
        Returns (clean_content, list_of_list_items)
        """
        lists = []
        
        # Pattern for itemize/enumerate
        pattern = r'\\begin\{(?:itemize|enumerate)\}(.*?)\\end\{(?:itemize|enumerate)\}'
        
        def replace_func(match):
            list_content = match.group(1)
            # Parse items
            items = []
            parts = re.split(r'\\item\s+', list_content)
            for part in parts:
                part = part.strip()
                if part:
                    items.append(part)
            
            if items:
                lists.append(items)
                return '\n\n[LIST_PLACEHOLDER]\n\n'
            return match.group(0)
            
        clean_content = re.sub(pattern, replace_func, content, flags=re.DOTALL | re.IGNORECASE)
        return clean_content, lists

    def create_list_element(self, items: List[str], styles: Dict) -> ListFlowable:
        """Create a ReportLab ListFlowable from list items."""
        list_items = []
        for item_text in items:
            # Handle inline formatting
            item_text = self.replace_inline_latex(item_text)
            # Use a bullet style or body style
            p = Paragraph(item_text, styles['ProfessionalBody'])
            list_items.append(ListItem(p))
        
        return ListFlowable(list_items, bulletType='bullet', start='circle', leftIndent=20)
    
    def parse_and_render_content(self, content: str, styles: Dict) -> List[Any]:
        """
        Parse content, extract tables and lists, and return list of ReportLab flowables.
        
        Args:
            content: AI-generated text content (may contain tables)
            styles: ReportLab ParagraphStyle dictionary
            
        Returns:
            List of ReportLab flowables (Paragraphs and Tables)
        """
        elements = []
        
        # Normalize headings
        content = re.sub(r'(\\section\{[^}]+\})', r'\n\n\1\n\n', content)
        content = re.sub(r'(\\subsection\{[^}]+\})', r'\n\n\1\n\n', content)
        content = re.sub(r'(\\subsubsection\{[^}]+\})', r'\n\n\1\n\n', content)

        # 1. Extract tables
        clean_text, tables = self.extract_tables_from_content(content)
        table_index = 0
        
        # 2. Extract lists
        clean_text, lists = self.extract_lists_from_content(clean_text)
        list_index = 0
        
        # 3. Split text into paragraphs
        paragraphs = clean_text.split('\n\n')
        
        for para_text in paragraphs:
            para_text = para_text.strip()
            if not para_text:
                continue
            
            # Handle Table Placeholders
            if '[TABLE_PLACEHOLDER]' in para_text:
                if table_index < len(tables):
                    title, table_data = tables[table_index]
                    
                    # Add table title
                    if title:
                        try:
                            elements.append(Paragraph(f"<b>{title}</b>", styles['TableHeading']))
                        except:
                            pass
                    
                    # Add table
                    table_elem = self.create_table_element(title, table_data, has_header=True)
                    if table_elem:
                        elements.append(table_elem)
                    
                    table_index += 1
                
                para_text = para_text.replace('[TABLE_PLACEHOLDER]', '').strip()
                if not para_text:
                    continue

            # Handle List Placeholders
            if '[LIST_PLACEHOLDER]' in para_text:
                if list_index < len(lists):
                    list_items = lists[list_index]
                    list_elem = self.create_list_element(list_items, styles)
                    if list_elem:
                        elements.append(list_elem)
                    list_index += 1
                para_text = para_text.replace('[LIST_PLACEHOLDER]', '').strip()
                if not para_text:
                    continue
            
            # Handle Headings
            # \section{Title} -> Heading1
            section_match = re.match(r'\\section\{([^}]+)\}', para_text)
            if section_match:
                title = section_match.group(1)
                elements.append(Paragraph(title, styles['Heading1']))
                continue
                
            # \subsection{Title} -> Heading2
            subsection_match = re.match(r'\\subsection\{([^}]+)\}', para_text)
            if subsection_match:
                title = subsection_match.group(1)
                elements.append(Paragraph(title, styles['Heading2']))
                continue

            # \subsubsection{Title} -> Heading3
            subsubsection_match = re.match(r'\\subsubsection\{([^}]+)\}', para_text)
            if subsubsection_match:
                title = subsubsection_match.group(1)
                elements.append(Paragraph(title, styles['Heading3']))
                continue
            
            # Regular Paragraph
            try:
                # Replace inline LaTeX
                para_text = self.replace_inline_latex(para_text)
                # Additional sanitization before creating Paragraph
                para_text = self.sanitize_html_content(para_text)
                if para_text:
                    elements.append(Paragraph(para_text, styles['ProfessionalBody']))
            except Exception as e:
                # If paragraph fails, try plain text
                print(f"   ⚠️  Paragraph parse error, using plain text: {str(e)[:100]}")
                try:
                    # Strip all tags and create simple paragraph
                    plain_text = re.sub(r'<[^>]+>', '', para_text)
                    if plain_text.strip():
                        elements.append(Paragraph(plain_text, styles['ProfessionalBody']))
                except:
                    pass  # Skip problematic paragraphs
        
        return elements


def test_table_parser():
    """Test the table parser with sample content"""
    
    # Sample AI-generated content with both formats
    sample_content = """
This is a project overview paragraph.

[TABLE:Project Cost Breakdown]
Particulars|Amount (Rs.)|Percentage
Land & Building|500000|25%
Plant & Machinery|1200000|60%
Working Capital|300000|15%
[/TABLE]

Another paragraph explaining the financial structure.

**Key Financial Ratios**

| Ratio Name | Year 1 | Year 2 | Year 3 |
|------------|--------|--------|--------|
| Current Ratio | 1.5 | 1.7 | 1.9 |
| DSCR | 1.8 | 2.0 | 2.2 |
| Debt-Equity | 2.0 | 1.5 | 1.2 |

The ratios show improving financial health over the project period.
"""
    
    prof_template = ProfessionalTemplate()
    parser = TableParser(prof_template)
    
    # Test table extraction
    clean_text, tables = parser.extract_tables_from_content(sample_content)
    
    print("=== EXTRACTED TABLES ===")
    for i, (title, data) in enumerate(tables, 1):
        print(f"\nTable {i}: {title}")
        print(f"  Rows: {len(data)}")
        print(f"  Columns: {len(data[0]) if data else 0}")
        for row in data:
            print(f"  {row}")
    
    print("\n=== CLEAN TEXT ===")
    print(clean_text)
    
    # Test rendering
    elements = parser.parse_and_render_content(sample_content, prof_template.get_styles())
    print(f"\n=== RENDERED ELEMENTS ===")
    print(f"Total elements: {len(elements)}")
    for i, elem in enumerate(elements, 1):
        print(f"  {i}. {type(elem).__name__}")


if __name__ == "__main__":
    test_table_parser()
