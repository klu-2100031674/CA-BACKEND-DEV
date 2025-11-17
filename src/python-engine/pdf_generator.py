
import json
import sys
from fpdf import FPDF

# Sheet name normalization utilities
def normalize_sheet_name(sheet_name: str) -> str:
    """Normalize sheet name by stripping whitespace and converting to lowercase."""
    return sheet_name.strip().lower()


def find_sheet_match(expected_sheet: str, available_sheets: list) -> str:
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

class ExcelReportPDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 16)
        self.cell(0, 10, 'Financial Projections Report', 0, 1, 'C')
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

    def add_sheet_title(self, title):
        self.set_font('Arial', 'B', 14)
        self.cell(0, 10, f"Sheet: {title}", 0, 1, 'L')
        self.ln(5)

    def add_table_from_cells(self, cells_data, max_rows=50):
        """Convert cell data to a structured table format and add to PDF"""
        if not cells_data:
            self.set_font('Arial', 'I', 10)
            self.cell(0, 10, 'No data available for this sheet', 0, 1, 'L')
            return

        # Extract row and column information from cell references
        rows_data = {}
        max_col = 0
        max_row = 0
        
        for cell_ref, value in cells_data.items():
            # Parse cell reference like "A1", "B2", etc.
            col_letters = ""
            row_num = ""
            for char in cell_ref:
                if char.isalpha():
                    col_letters += char
                else:
                    row_num += char
            
            if row_num:
                row_idx = int(row_num)
                col_idx = self.column_letter_to_number(col_letters)
                
                if row_idx not in rows_data:
                    rows_data[row_idx] = {}
                rows_data[row_idx][col_idx] = str(value) if value is not None else ""
                
                max_col = max(max_col, col_idx)
                max_row = max(max_row, row_idx)

        # Limit rows for PDF display
        display_rows = min(max_rows, max_row)
        
        if not rows_data:
            self.set_font('Arial', 'I', 10)
            self.cell(0, 10, 'No valid cell data found', 0, 1, 'L')
            return

        # Calculate column width
        usable_width = self.w - 20  # Account for margins
        col_width = min(usable_width / (max_col + 1), 40)  # Max 40 units per column
        
        # Add table headers (column letters)
        self.set_font('Arial', 'B', 8)
        self.cell(15, 8, 'Row', 1, 0, 'C')  # Row number column
        for col in range(1, max_col + 1):
            col_letter = self.number_to_column_letter(col)
            self.cell(col_width, 8, col_letter, 1, 0, 'C')
        self.ln()

        # Add data rows
        self.set_font('Arial', '', 7)
        for row in range(1, display_rows + 1):
            # Row number
            self.cell(15, 8, str(row), 1, 0, 'C')
            
            # Data cells
            for col in range(1, max_col + 1):
                cell_value = rows_data.get(row, {}).get(col, "")
                # Truncate long values
                if len(cell_value) > 15:
                    cell_value = cell_value[:12] + "..."
                self.cell(col_width, 8, cell_value, 1, 0, 'C')
            self.ln()
            
            # Add page break if getting close to bottom
            if self.get_y() > 250:
                self.add_page()
                self.add_sheet_title(f"Continued...")
                break

    def column_letter_to_number(self, letters):
        """Convert column letters to number (A=1, B=2, etc.)"""
        result = 0
        for letter in letters:
            result = result * 26 + (ord(letter.upper()) - ord('A') + 1)
        return result

    def number_to_column_letter(self, num):
        """Convert column number to letters (1=A, 2=B, etc.)"""
        result = ""
        while num > 0:
            num -= 1
            result = chr(num % 26 + ord('A')) + result
            num //= 26
        return result

    def add_summary_section(self, meta_data):
        """Add a summary section with metadata"""
        self.add_page()
        self.set_font('Arial', 'B', 14)
        self.cell(0, 10, 'Report Summary', 0, 1, 'L')
        self.ln(5)
        
        self.set_font('Arial', '', 10)
        if 'totalSheets' in meta_data:
            self.cell(0, 8, f'Total Sheets: {meta_data["totalSheets"]}', 0, 1, 'L')
        if 'formulaRecalculation' in meta_data:
            self.cell(0, 8, f'Formula Recalculation: {meta_data["formulaRecalculation"]}', 0, 1, 'L')
        if 'approach' in meta_data:
            self.cell(0, 8, f'Calculation Engine: {meta_data["approach"]}', 0, 1, 'L')

def generate_pdf(json_data, output_path, template_name='CC1'):
    """
    Generates a comprehensive PDF from the Excel data JSON.
    """
    try:
        data = json.loads(json_data) if isinstance(json_data, str) else json_data
        
        pdf = ExcelReportPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        
        # Add summary page first
        if 'meta' in data:
            pdf.add_summary_section(data['meta'])

        # Process each sheet
        final_sheet_name = 'Final workings' if 'CC6' in template_name else 'Finalworkings'
        important_sheets = ['Assumptions.1', final_sheet_name, 'PLBS', 'RATIO']
        other_sheets = []
        
        # Track which sheets have been processed
        processed_sheets = set()
        
        # First, add important sheets (case-insensitive matching)
        for expected_sheet in important_sheets:
            actual_sheet = find_sheet_match(expected_sheet, [k for k in data.keys() if k not in ['_appliedUpdates', '_cellValues', '_meta', 'meta']])
            if actual_sheet:
                processed_sheets.add(actual_sheet)
                pdf.add_page()
                pdf.add_sheet_title(actual_sheet)
                
                sheet_data = data[actual_sheet]
                if isinstance(sheet_data, dict) and 'cells' in sheet_data:
                    # New format with cells
                    pdf.add_table_from_cells(sheet_data['cells'])
                elif isinstance(sheet_data, list):
                    # Old format - convert to cells format
                    cells = {}
                    for row_idx, row_data in enumerate(sheet_data):
                        if isinstance(row_data, dict):
                            for col_idx, (key, value) in enumerate(row_data.items()):
                                if value is not None and value != '':
                                    col_letter = pdf.number_to_column_letter(col_idx + 1)
                                    cell_ref = f"{col_letter}{row_idx + 1}"
                                    cells[cell_ref] = value
                    pdf.add_table_from_cells(cells)

        # Then add other sheets
        for sheet_name, sheet_data in data.items():
            if (sheet_name not in processed_sheets and 
                sheet_name not in ['_appliedUpdates', '_cellValues', '_meta', 'meta']):
                pdf.add_page()
                pdf.add_sheet_title(sheet_name)
                
                if isinstance(sheet_data, dict) and 'cells' in sheet_data:
                    pdf.add_table_from_cells(sheet_data['cells'])
                elif isinstance(sheet_data, list):
                    # Convert old format
                    cells = {}
                    for row_idx, row_data in enumerate(sheet_data):
                        if isinstance(row_data, dict):
                            for col_idx, (key, value) in enumerate(row_data.items()):
                                if value is not None and value != '':
                                    col_letter = pdf.number_to_column_letter(col_idx + 1)
                                    cell_ref = f"{col_letter}{row_idx + 1}"
                                    cells[cell_ref] = value
                    pdf.add_table_from_cells(cells)

        pdf.output(output_path)
        return True
        
    except Exception as e:
        print(f"Error generating PDF: {str(e)}")
        return False

if __name__ == '__main__':
    json_input_string = sys.argv[1]
    output_file_path = sys.argv[2]
    template_name = sys.argv[3] if len(sys.argv) > 3 else 'CC1'
    
    success = generate_pdf(json_input_string, output_file_path, template_name)
    if success:
        print("PDF generated successfully")
    else:
        print("Failed to generate PDF")
