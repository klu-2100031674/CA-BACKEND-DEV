"""Excel payload applier.

Updates cells based on the incoming payload, leaves the source template untouched,
and writes an updated copy to the temp directory for verification.
"""

import json
import datetime
import os
import sys
import re
from typing import Any, Dict, List
import base64

import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.utils import get_column_letter
import pandas as pd
import numpy as np
from pathlib import Path

# Windows COM for Excel automation
try:
    import win32com.client
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False
    print("Warning: pywin32 not available. PDF generation will use fallback method.", file=sys.stderr)


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


def generate_pdf_from_excel_sheet(excel_path: str, sheet_name: str, output_path: str) -> bool:
    """Generate PDF directly from Excel sheet using Excel COM automation to preserve all formatting."""
    try:
        print(f"[PDF Generator] Starting PDF generation for sheet: {sheet_name}", file=sys.stderr)
        print(f"[PDF Generator] COM_AVAILABLE: {COM_AVAILABLE}", file=sys.stderr)
        print(f"[PDF Generator] Input Excel: {excel_path}", file=sys.stderr)
        print(f"[PDF Generator] Output PDF: {output_path}", file=sys.stderr)
        
        if COM_AVAILABLE:
            # Use Excel COM automation for exact formatting preservation
            print(f"[PDF Generator] Using Excel COM automation for exact formatting", file=sys.stderr)
            excel = None
            try:
                print(f"[PDF Generator] Initializing Excel COM...", file=sys.stderr)
                excel = win32com.client.Dispatch("Excel.Application")
                try:
                    excel.Visible = False
                except Exception as e:
                    print(f"[PDF Generator] Warning: Could not set Excel.Visible to False: {e}", file=sys.stderr)
                excel.DisplayAlerts = False
                print(f"[PDF Generator] Excel COM initialized successfully", file=sys.stderr)
                
                # Open workbook
                print(f"[PDF Generator] Opening workbook: {os.path.abspath(excel_path)}", file=sys.stderr)
                workbook = excel.Workbooks.Open(os.path.abspath(excel_path))
                print(f"[PDF Generator] Workbook opened, total sheets: {workbook.Sheets.Count}", file=sys.stderr)
                
                # Find and select the sheet (case-insensitive matching)
                sheet_found = False
                actual_sheet_name = None
                for sheet in workbook.Sheets:
                    print(f"[PDF Generator] Found sheet: {sheet.Name}", file=sys.stderr)
                    if normalize_sheet_name(sheet.Name) == normalize_sheet_name(sheet_name):
                        sheet.Select()
                        sheet_found = True
                        actual_sheet_name = sheet.Name
                        print(f"[PDF Generator] Sheet '{actual_sheet_name}' selected (matched from '{sheet_name}')", file=sys.stderr)
                        break
                
                if not sheet_found:
                    print(f"[PDF Generator] ERROR: Sheet '{sheet_name}' not found in workbook (tried case-insensitive matching)", file=sys.stderr)
                    workbook.Close(SaveChanges=False)
                    excel.Quit()
                    return False
                
                # Export as PDF with optimal settings
                print(f"[PDF Generator] Exporting to PDF: {os.path.abspath(output_path)}", file=sys.stderr)
                workbook.ActiveSheet.ExportAsFixedFormat(
                    Type=0,  # xlTypePDF
                    Filename=os.path.abspath(output_path),
                    Quality=0,  # 0 = Standard quality (faster, smaller file)
                    IncludeDocProperties=True,
                    IgnorePrintAreas=False,
                    OpenAfterPublish=False
                )
                
                print(f"[PDF Generator] PDF export completed", file=sys.stderr)
                workbook.Close(SaveChanges=False)
                excel.Quit()
                print(f"[PDF Generator] Excel closed successfully", file=sys.stderr)
                
                print(f"[PDF Generator] PDF generated successfully using Excel COM: {output_path}", file=sys.stderr)
                return True
                
            except Exception as com_error:
                print(f"❌ [PDF Generator] Excel COM error: {str(com_error)}", file=sys.stderr)
                import traceback
                traceback.print_exc(file=sys.stderr)
                if excel:
                    try:
                        excel.Quit()
                    except:
                        pass
                return False
        else:
            # Fallback to pandas method (no formatting preservation)
            print(f"⚠️ [PDF Generator] COM not available, using fallback method", file=sys.stderr)
            return generate_pdf_fallback(excel_path, sheet_name, output_path)
            
    except Exception as e:
        print(f"❌ [PDF Generator] Error generating PDF from Excel sheet: {str(e)}", file=sys.stderr)
        import traceback
        traceback.print_exc(file=sys.stderr)
        return False


def generate_pdf_fallback(excel_path: str, sheet_name: str, output_path: str) -> bool:
    """Fallback PDF generation using pandas (no formatting preservation)."""
    try:
        from fpdf import FPDF
        
        class SimplePDF(FPDF):
            def header(self):
                self.set_font('Arial', 'B', 12)
                self.cell(0, 8, f'{sheet_name} Sheet', 0, 1, 'C')
                self.ln(2)
            
            def footer(self):
                self.set_y(-15)
                self.set_font('Arial', 'I', 8)
                self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')
        
        # Read Excel sheet
        df = pd.read_excel(excel_path, sheet_name=sheet_name, engine='openpyxl', header=None)
        df = df.replace([pd.NA, np.inf, -np.inf], '')
        df = df.fillna('')

        # Create PDF
        pdf = SimplePDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()
        pdf.set_font('Arial', '', 8)

        # Calculate column widths
        max_cols = len(df.columns)
        page_width = pdf.w - 30
        col_width = min(page_width / max_cols, 25)

        # Add data rows
        for row_idx in range(len(df)):
            row_data = df.iloc[row_idx]
            if pdf.get_y() > 250:
                pdf.add_page()
                pdf.set_font('Arial', '', 8)

            for col_idx in range(max_cols):
                value = str(row_data.iloc[col_idx]) if col_idx < len(row_data) else ''
                if len(value) > 15:
                    value = value[:12] + '...'

                if row_idx == 0:
                    pdf.set_fill_color(240, 240, 240)
                    pdf.cell(col_width, 8, value, 1, 0, 'C', 1)
                else:
                    pdf.cell(col_width, 8, value, 1, 0, 'C', 0)
            pdf.ln()

        pdf.output(output_path)
        print(f"PDF generated using fallback method: {output_path}", file=sys.stderr)
        return True

    except Exception as e:
        print(f"Fallback PDF generation failed: {str(e)}", file=sys.stderr)
        return False


def generate_pdfs_for_all_sheets(excel_path: str, output_dir: str) -> Dict[str, Any]:
    """
    Generate individual PDF files for ALL sheets in the Excel workbook (excluding Assumptions sheet).
    Uses Excel COM automation to preserve formatting with better page fitting.
    
    Args:
        excel_path: Path to the Excel file
        output_dir: Directory to save the PDF files
        
    Returns:
        Dictionary with sheet names as keys and PDF file paths as values
    """
    print(f"\n{'='*80}", file=sys.stderr)
    print(f"📄 GENERATING PDFs FOR ALL EXCEL SHEETS", file=sys.stderr)
    print(f"{'='*80}\n", file=sys.stderr)
    
    # Sheets to exclude from PDF generation
    EXCLUDED_SHEETS = ['Assumptions.1', 'Assumptions', 'assumptions', 'ASSUMPTIONS']
    
    pdf_files = {
        "sheets": {},
        "success_count": 0,
        "failed_count": 0,
        "total_sheets": 0,
        "excluded_sheets": []
    }
    
    try:
        if not COM_AVAILABLE:
            print(f"❌ Excel COM not available. Cannot generate PDFs with formatting.", file=sys.stderr)
            return pdf_files
        
        # Ensure output directory exists
        os.makedirs(output_dir, exist_ok=True)
        
        # Open workbook with Excel COM
        print(f"[Multi-PDF Generator] Opening workbook: {excel_path}", file=sys.stderr)
        excel = win32com.client.Dispatch("Excel.Application")
        try:
            excel.Visible = False
        except Exception as e:
            print(f"[Multi-PDF Generator] Warning: Could not set Excel.Visible to False: {e}", file=sys.stderr)
        excel.DisplayAlerts = False
        
        workbook = excel.Workbooks.Open(os.path.abspath(excel_path))
        total_sheets = workbook.Sheets.Count
        pdf_files["total_sheets"] = total_sheets
        
        print(f"[Multi-PDF Generator] Found {total_sheets} sheets", file=sys.stderr)
        print(f"{'─'*80}\n", file=sys.stderr)
        
        # Generate PDF for each sheet
        for sheet_idx in range(1, total_sheets + 1):
            sheet = workbook.Sheets(sheet_idx)
            sheet_name = sheet.Name
            
            # Skip excluded sheets (like Assumptions)
            if sheet_name in EXCLUDED_SHEETS:
                print(f"[{sheet_idx}/{total_sheets}] ⏭️  Skipping sheet: '{sheet_name}' (excluded)", file=sys.stderr)
                pdf_files["excluded_sheets"].append(sheet_name)
                continue
            
            print(f"[{sheet_idx}/{total_sheets}] Processing sheet: '{sheet_name}'", file=sys.stderr)
            
            # Create PDF filename (sanitize sheet name)
            safe_sheet_name = re.sub(r'[<>:"/\\|?*]', '_', sheet_name)
            pdf_filename = f"sheet_{sheet_idx}_{safe_sheet_name}.pdf"
            pdf_path = os.path.join(output_dir, pdf_filename)
            
            try:
                # Select the sheet
                sheet.Select()
                
                # Configure page setup for better fitting
                page_setup = workbook.ActiveSheet.PageSetup
                page_setup.Zoom = False  # Disable fixed zoom
                page_setup.FitToPagesWide = 1  # Fit to 1 page wide
                
                # Special handling for Coverpage - must fit on 1 page
                if sheet_name.lower() == 'coverpage':
                    page_setup.FitToPagesTall = 1  # Force Coverpage to 1 page
                    page_setup.Orientation = 1  # Portrait
                else:
                    page_setup.FitToPagesTall = False  # Allow multiple pages vertically for other sheets
                
                page_setup.Orientation = 1  # xlPortrait (use 2 for xlLandscape if needed)
                page_setup.PaperSize = 9  # A4
                page_setup.LeftMargin = excel.InchesToPoints(0.5)
                page_setup.RightMargin = excel.InchesToPoints(0.5)
                page_setup.TopMargin = excel.InchesToPoints(0.5)
                page_setup.BottomMargin = excel.InchesToPoints(0.5)
                
                # Export as PDF
                workbook.ActiveSheet.ExportAsFixedFormat(
                    Type=0,  # xlTypePDF
                    Filename=os.path.abspath(pdf_path),
                    Quality=0,  # Standard quality
                    IncludeDocProperties=True,
                    IgnorePrintAreas=False,
                    OpenAfterPublish=False
                )
                
                # Check if PDF was created
                if os.path.exists(pdf_path):
                    file_size = os.path.getsize(pdf_path)
                    print(f"   ✅ PDF created: {pdf_filename} ({file_size:,} bytes)", file=sys.stderr)
                    
                    pdf_files["sheets"][sheet_name] = {
                        "pdf_path": pdf_path,
                        "pdf_filename": pdf_filename,
                        "sheet_index": sheet_idx,
                        "file_size": file_size,
                        "status": "success"
                    }
                    pdf_files["success_count"] += 1
                else:
                    print(f"   ❌ PDF file not created", file=sys.stderr)
                    pdf_files["sheets"][sheet_name] = {
                        "status": "failed",
                        "error": "PDF file not created"
                    }
                    pdf_files["failed_count"] += 1
                    
            except Exception as sheet_error:
                print(f"   ❌ Error generating PDF for sheet '{sheet_name}': {str(sheet_error)}", file=sys.stderr)
                pdf_files["sheets"][sheet_name] = {
                    "status": "failed",
                    "error": str(sheet_error)
                }
                pdf_files["failed_count"] += 1
        
        # Close workbook and Excel
        workbook.Close(SaveChanges=False)
        excel.Quit()
        
        print(f"\n{'─'*80}", file=sys.stderr)
        print(f"✅ PDF Generation Complete", file=sys.stderr)
        print(f"   Total Sheets: {pdf_files['total_sheets']}", file=sys.stderr)
        print(f"   Excluded: {len(pdf_files['excluded_sheets'])} ({', '.join(pdf_files['excluded_sheets']) if pdf_files['excluded_sheets'] else 'none'})", file=sys.stderr)
        print(f"   Successful: {pdf_files['success_count']}", file=sys.stderr)
        print(f"   Failed: {pdf_files['failed_count']}", file=sys.stderr)
        print(f"{'='*80}\n", file=sys.stderr)
        
        return pdf_files
        
    except Exception as e:
        print(f"❌ Error in multi-sheet PDF generation: {str(e)}", file=sys.stderr)
        import traceback
        traceback.print_exc(file=sys.stderr)
        return pdf_files

def generate_html_from_excel_com(excel_path: str, sheet_name: str) -> tuple:
    """
    Generate HTML from Excel using COM automation to get calculated values.
    This ensures formulas are evaluated and we get the actual values.
    Returns: (html_content, json_data)
    """
    try:
        print(f"[HTML COM Generator] Starting HTML generation using Excel COM", file=sys.stderr)
        
        excel = win32com.client.Dispatch("Excel.Application")
        try:
            excel.Visible = False
        except Exception as e:
            print(f"[HTML COM Generator] Warning: Could not set Excel.Visible to False: {e}", file=sys.stderr)
        excel.DisplayAlerts = False
        
        wb = excel.Workbooks.Open(os.path.abspath(excel_path))
        
        # Get all sheet names
        available_sheets = [ws.Name for ws in wb.Worksheets]
        
        # Find the matching sheet name (handles case and space differences)
        actual_sheet_name = find_sheet_match(sheet_name, available_sheets)
        if not actual_sheet_name:
            print(f"[HTML COM Generator] Sheet '{sheet_name}' not found (tried case-insensitive matching)", file=sys.stderr)
            print(f"[HTML COM Generator] Available sheets: {available_sheets}", file=sys.stderr)
            wb.Close(False)
            excel.Quit()
            return "", {}
        
        # Find the sheet
        sheet = None
        for ws in wb.Worksheets:
            if ws.Name == actual_sheet_name:
                sheet = ws
                break
        
        print(f"[HTML COM Generator] Processing sheet: {actual_sheet_name} (matched from '{sheet_name}')", file=sys.stderr)
        
        # Get used range
        used_range = sheet.UsedRange
        max_row = used_range.Rows.Count
        max_col = used_range.Columns.Count
        
        print(f"[HTML COM Generator] Processing {max_row} rows x {max_col} columns", file=sys.stderr)
        
        # Extract JSON data structure
        json_data = {
            "sheetName": actual_sheet_name,
            "data": {},
            "timestamp": datetime.datetime.now().isoformat()
        }
        
        # Extract firm details from the data for receipt header
        firm_name = ""
        proprietor = ""
        sector = ""
        nature_of_business = ""
        
        # Try to get firm details from common cell positions
        try:
            if max_row >= 3:
                firm_name_cell = sheet.Cells(3, 2).Value
                if firm_name_cell:
                    firm_name = str(firm_name_cell)
            if max_row >= 4:
                proprietor_cell = sheet.Cells(4, 2).Value
                if proprietor_cell:
                    proprietor = str(proprietor_cell)
            if max_row >= 6:
                sector_cell = sheet.Cells(6, 2).Value
                if sector_cell:
                    sector = str(sector_cell)
            if max_row >= 7:
                nature_cell = sheet.Cells(7, 2).Value
                if nature_cell:
                    nature_of_business = str(nature_cell)
        except:
            pass
        
        # Build HTML with modern professional styling
        html_parts = [
            "<!DOCTYPE html>",
            "<html lang='en'>",
            "<head>",
            "<meta charset='UTF-8'>",
            "<meta name='viewport' content='width=device-width, initial-scale=1.0'>",
            f"<title>Financial Report - {firm_name or sheet_name}</title>",
            "<link href='https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&family=Inter:wght@300;400;500;600;700&display=swap' rel='stylesheet'>",
            "<link rel='stylesheet' href='https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css'>", # For professional icons
            "<style>",
            "  :root {",
            "    --primary-purple: #7c3aed;",
            "    --primary-dark-purple: #6d28d9;",
            "    --primary-black: #1f2937;",
            "    --primary-light-black: #374151;",
            "    --ghost-white: #F8F8FF;",
            "    --success-green: #10b981;",
            "    --text-primary: var(--primary-black);",
            "    --text-secondary: #6b7280;",
            "    --bg-primary: #ffffff;",
            "    --bg-secondary: var(--ghost-white);",
            "    --bg-accent: #e5e7eb;",
            "    --border-color: #e5e7eb;",
            "    --shadow-soft: 0 2px 15px -3px rgba(0, 0, 0, 0.07), 0 10px 20px -2px rgba(0, 0, 0, 0.04);",
            "    --shadow-md: 0 4px 6px -1px rgba(0, 0, 0, 0.1);",
            "    --shadow-lg: 0 10px 15px -3px rgba(0, 0, 0, 0.1);",
            "    --shadow-xl: 0 20px 25px -5px rgba(0, 0, 0, 0.1);",
            "  }",
            "  ",
            "  * {",
            "    margin: 0;",
            "    padding: 0;",
            "    box-sizing: border-box;",
            "  }",
            "  ",
            "  body {",
            "    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;",
            "    background-color: var(--bg-secondary);", # Ghost White background
            "    min-height: 100vh;",
            "    padding: 20px;",
            "    line-height: 1.6;",
            "    color: var(--text-primary);",
            "    -webkit-font-smoothing: antialiased;",
            "    -moz-osx-font-smoothing: grayscale;",
            "  }",
            "  ",
            "  .container {",
            "    max-width: 1400px;",
            "    margin: 0 auto;",
            "  }",
            "  ",
            "  .report-card {",
            "    background: var(--bg-primary);",
            "    border-radius: 12px;", # Rounded corners
            "    box-shadow: var(--shadow-soft);", # Soft shadow
            "    overflow: hidden;",
            "    animation: slideUp 0.6s ease-out;",
            "  }",
            "  ",
            "  @keyframes slideUp {",
            "    from {",
            "      opacity: 0;",
            "      transform: translateY(30px);",
            "    }",
            "    to {",
            "      opacity: 1;",
            "      transform: translateY(0);",
            "    }",
            "  }",
            "  ",
            "  /* Header Section */",
            "  .report-header {"
            "    padding: 32px 24px;",
            "    color: black;",
            "    position: relative;",
            "    overflow: hidden;",
            "    border-bottom: 1px solid rgba(255, 255, 255, 0.1);",
            "  }",
            "  ",
            "  .report-header::before {",
            "    content: '';",
            "    position: absolute;",
            "    top: 0;",
            "    right: 0;",
            "    width: 200px;",
            "    height: 200px;",
            "    background: radial-gradient(circle, rgba(255,255,255,0.15) 0%, transparent 70%);",
            "    border-radius: 50%;",
            "    transform: translate(30%, -30%);",
            "  }",
            "  ",
            "  .header-content {",
            "    position: relative;",
            "    z-index: 1;",
            "  }",
            "  ",
            "  .report-badge {",
            "    display: inline-flex;", # Use flex for icon alignment
            "    align-items: center;",
            "    gap: 8px;",
            "    background: rgba(255, 255, 255, 0.2);",
            "    backdrop-filter: blur(5px);",
            "    padding: 6px 16px;",
            "    border-radius: 50px;",
            "    font-size: 12px;",
            "    font-weight: 600;",
            "    letter-spacing: 0.5px;",
            "    text-transform: uppercase;",
            "    margin-bottom: 16px;",
            "  }",
            "  .report-badge i {",
            "    font-size: 14px;",
            "  }",
            "  ",
            "  .firm-name {",
            "    font-family: 'Manrope', sans-serif;", # Manrope for headings
            "    font-size: 28px;",
            "    font-weight: 700;",
            "    margin-bottom: 12px;",
            "    letter-spacing: -0.5px;",
            "  }",
            "  ",
            "  .firm-meta {",
            "    display: flex;",
            "    flex-wrap: wrap;",
            "    gap: 24px;",
            "    margin-top: 20px;",
            "    padding-top: 20px;",
            "    border-top: 1px solid rgba(255, 255, 255, 0.15);",
            "  }",
            "  ",
            "  .meta-item {",
            "    display: flex;",
            "    flex-direction: column;",
            "    gap: 4px;",
            "  }",
            "  ",
            "  .meta-label {",
            "    font-size: 11px;",
            "    font-weight: 500;",
            "    opacity: 0.9;",
            "    text-transform: uppercase;",
            "    letter-spacing: 0.8px;",
            "  }",
            "  ",
            "  .meta-value {",
            "    font-size: 15px;",
            "    font-weight: 600;",
            "  }",
            "  ",
            "  /* Stats Cards */",
            "  .stats-grid {",
            "    display: grid;",
            "    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));",
            "    gap: 1px;",
            "    background: var(--border-color);",
            "    border-bottom: 1px solid var(--border-color);",
            "  }",
            "  ",
            "  .stat-card {",
            "    background: var(--bg-primary);",
            "    padding: 24px 20px;",
            "    text-align: center;",
            "    transition: all 0.3s ease;",
            "  }",
            "  ",
            "  .stat-card:hover {",
            "    background: var(--bg-secondary);",
            "    transform: translateY(-2px);",
            "  }",
            "  ",
            "  .stat-icon {",
            "    width: 40px;",
            "    height: 40px;",
            "    margin: 0 auto 12px;",
            "    border-radius: 8px;",
            "    display: flex;",
            "    align-items: center;",
            "    justify-content: center;",
            "    font-size: 18px;",
            "    color: black;",
            "  }",
            "  ",
            "  .stat-label {",
            "    font-size: 11px;",
            "    font-weight: 600;",
            "    color: var(--text-secondary);",
            "    text-transform: uppercase;",
            "    letter-spacing: 0.8px;",
            "    margin-bottom: 6px;",
            "  }",
            "  ",
            "  .stat-value {",
            "    font-size: 16px;",
            "    font-weight: 700;",
            "    color: var(--text-primary);",
            "    font-family: 'Manrope', sans-serif;", # Manrope for values
            "  }",
            "  ",
            "  /* Table Section */",
            "  .table-section {",
            "    padding: 32px 24px;",
            "  }",
            "  ",
            "  .section-title {",
            "    font-family: 'Manrope', sans-serif;", # Manrope for titles
            "    font-size: 20px;",
            "    font-weight: 700;",
            "    color: var(--text-primary);",
            "    margin-bottom: 6px;",
            "  }",
            "  ",
            "  .section-subtitle {",
            "    font-size: 13px;",
            "    color: var(--text-secondary);",
            "    margin-bottom: 24px;",
            "  }",
            "  ",
            "  .table-wrapper {",
            "    overflow-x: auto;",
            "    border-radius: 8px;",
            "    border: 1px solid var(--border-color);",
            "  }",
            "  ",
            "  table {",
            "    width: 100%;",
            "    border-collapse: collapse;",
            "    background: var(--bg-primary);",
            "  }",
            "  ",
            "  thead {",
            "    background: var(--bg-accent);",
            "    position: sticky;",
            "    top: 0;",
            "    z-index: 10;",
            "  }",
            "  ",
            "  thead th {",
            "    padding: 14px 20px;",
            "    text-align: left;",
            "    font-size: 11px;",
            "    font-weight: 700;",
            "    color: var(--text-primary);",
            "    text-transform: uppercase;",
            "    letter-spacing: 0.8px;",
            "    border-bottom: 1px solid var(--border-color);",
            "  }",
            "  ",
            "  thead th:last-child {",
            "    text-align: right;",
            "  }",
            "  ",
            "  tbody tr {",
            "    border-bottom: 1px solid var(--border-color);",
            "    transition: all 0.2s ease;",
            "  }",
            "  ",
            "  tbody tr:hover {",
            "    background: var(--bg-secondary);",
            "  }",
            "  ",
            "  tbody tr:last-child {",
            "    border-bottom: none;",
            "  }",
            "  ",
            "  td {",
            "    padding: 16px 20px;",
            "    font-size: 13px;",
            "    color: var(--text-primary);",
            "  }",
            "  ",
            "  .item-name {",
            "    font-weight: 500;",
            "  }",
            "  ",
            "  .item-value {",
            "    text-align: right;",
            "    font-weight: 600;",
            "    font-family: 'Inter', monospace;",
            "    font-size: 14px;",
            "  }",
            "  ",
            "  .currency::before {",
            "    content: '₹ ';",
            "    color: var(--text-secondary);",
            "    margin-right: 2px;",
            "    font-weight: 500;",
            "  }",
            "  ",
            "  /* Special Row Styles */",
            "  .section-header {",
            "    background: var(--bg-accent) !important;",
            "  }",
            "  ",
            "  .section-header td {",
            "    font-family: 'Manrope', sans-serif !important;",
            "    font-weight: 700 !important;",
            "    font-size: 13px !important;",
            "    color: var(--text-primary) !important;",
            "    padding: 12px 20px !important;",
            "    text-transform: uppercase;",
            "    letter-spacing: 0.5px;",
            "  }",
            "  ",
            "  .total-row {",
            "  }",
            "  ",
            "  .total-row td {",
            "    color: black !important;",
            "    font-weight: 700 !important;",
            "    font-size: 15px !important;",
            "    padding: 18px 20px !important;",
            "    border-top: 2px solid var(--primary-dark-purple);", # Purple accent border
            "  }",
            "  ",
            "  .subtotal-row {",
            "    background: var(--bg-secondary) !important;", # Ghost White subtotal row
            "  }",
            "  ",
            "  .subtotal-row td {",
            "    font-weight: 600 !important;",
            "    padding: 14px 20px !important;",
            "    color: var(--text-primary);",
            "  }",
            "  ",
            "  /* Footer Section */",
            "  .report-footer {",
            "    background: var(--bg-primary);",
            "    padding: 32px 24px;",
            "    text-align: center;",
            "    border-top: 1px solid var(--border-color);",
            "  }",
            "  ",
            "  .footer-content {",
            "    max-width: 600px;",
            "    margin: 0 auto;",
            "  }",
            "  ",
            "  .footer-title {",
            "    font-family: 'Manrope', sans-serif;", # Manrope for titles
            "    font-size: 18px;",
            "    font-weight: 600;",
            "    color: var(--text-primary);",
            "    margin-bottom: 10px;",
            "  }",
            "  ",
            "  .footer-text {",
            "    font-size: 12px;",
            "    color: var(--text-secondary);",
            "    line-height: 1.7;",
            "    margin-bottom: 20px;",
            "  }",
            "  ",
            "  .action-buttons {",
            "    display: flex;",
            "    gap: 12px;",
            "    justify-content: center;",
            "    flex-wrap: wrap;",
            "  }",
            "  ",
            "  .btn {",
            "    padding: 10px 24px;",
            "    border-radius: 8px;",
            "    font-weight: 600;",
            "    font-size: 13px;",
            "    cursor: pointer;",
            "    transition: all 0.3s ease;",
            "    border: 2px solid var(--primary-black);", # Outlined black button
            "    display: inline-flex;",
            "    align-items: center;",
            "    gap: 8px;",
            "    text-decoration: none;",
            "    color: var(--primary-black);",
            "    background: transparent;",
            "  }",
            "  ",
            "  .btn:hover {",
            "    background: var(--primary-black);", # Hover to solid black
            "    color: white;",
            "    transform: translateY(-1px);",
            "    box-shadow: 0 4px 8px rgba(0,0,0,0.1);",
            "  }",
            "  ",
            "  .btn-primary {",
            "    background: var(--primary-black);", # Primary is solid black
            "    color: white;",
            "    box-shadow: var(--shadow-sm);",
            "    border-color: var(--primary-black);",
            "  }",
            "  .btn-primary:hover {",
            "    background: var(--primary-light-black);", # Darker black on hover
            "    border-color: var(--primary-light-black);",
            "  }",
            "  ",
            "  /* Secondary button is the default outlined style */",
            "  .btn-secondary {",
            "    background: transparent;",
            "    color: var(--primary-black);",
            "    border-color: var(--primary-black);",
            "  }",
            "  .btn-secondary:hover {",
            "    background: var(--primary-black);",
            "    color: white;",
            "  }",
            "  ",
            "  /* Timestamp Badge */",
            "  .timestamp-badge {",
            "    display: inline-flex;",
            "    align-items: center;",
            "    gap: 6px;",
            "    background: var(--bg-accent);",
            "    padding: 6px 12px;",
            "    border-radius: 6px;",
            "    font-size: 11px;",
            "    color: var(--text-secondary);",
            "    margin-top: 20px;",
            "  }",
            "  .timestamp-badge i {",
            "    font-size: 12px;",
            "  }",
            "  ",
            "  /* Responsive Design */",
            "  @media (max-width: 1024px) {",
            "    .report-header {",
            "      padding: 28px 20px 20px;",
            "    }",
            "    ",
            "    .firm-name {",
            "      font-size: 24px;",
            "    }",
            "    ",
            "    .table-section {",
            "      padding: 28px 20px;",
            "    }",
            "  }",
            "  ",
            "  @media (max-width: 768px) {",
            "    body {",
            "      padding: 16px;",
            "    }",
            "    ",
            "    .report-card {",
            "      border-radius: 10px;",
            "    }",
            "    ",
            "    .report-header {",
            "      padding: 24px 16px 16px;",
            "    }",
            "    ",
            "    .firm-name {",
            "      font-size: 20px;",
            "    }",
            "    ",
            "    .firm-meta {",
            "      flex-direction: column;",
            "      gap: 16px;",
            "    }",
            "    ",
            "    .stats-grid {",
            "      grid-template-columns: 1fr;",
            "    }",
            "    ",
            "    .table-section {",
            "      padding: 20px 12px;",
            "    }",
            "    ",
            "    thead th,",
            "    td {",
            "      padding: 12px 14px;",
            "      font-size: 12px;",
            "    }",
            "    ",
            "    .report-footer {",
            "      padding: 24px 16px;",
            "    }",
            "    ",
            "    .action-buttons {",
            "      flex-direction: column;",
            "    }",
            "    ",
            "    .btn {",
            "      width: 100%;",
            "      justify-content: center;",
            "    }",
            "  }",
            "  ",
            "  /* Print Styles */",
            "  @media print {",
            "    body {",
            "      background: white;",
            "      padding: 0;",
            "    }",
            "    ",
            "    .report-card {",
            "      box-shadow: none;",
            "      border-radius: 0;",
            "    }",
            "    ",
            "    .report-header::before {",
            "      display: none;",
            "    }",
            "    ",
            "    .action-buttons {",
            "      display: none;",
            "    }",
            "    ",
            "    tbody tr:hover {",
            "      background: transparent;",
            "    }",
            "  }",
            "  ",
            "  /* Loading Animation */",
            "  @keyframes shimmer {",
            "    0% { background-position: -1000px 0; }",
            "    100% { background-position: 1000px 0; }",
            "  }",
            "  ",
            "  .loading {",
            "    animation: shimmer 2s infinite;",
            "    background: linear-gradient(to right, #f6f7f8 0%, #edeef1 20%, #f6f7f8 40%, #f6f7f8 100%);",
            "    background-size: 1000px 100%;",
            "  }",
            "</style>",
            "</head>",
            "<body>",
            "<div class='container'>",
            "<div class='report-card'>",
            "",
            "<!-- Header Section -->",
            "<div class='report-header'>",
            "<div class='header-content'>",
            f"<span class='report-badge'><i class='fas fa-chart-line'></i> {sheet_name}</span>", # Professional Icon
            f"<h1 class='firm-name'>{firm_name or 'Financial Report'}</h1>",
            "<div class='firm-meta'>",
        ]
        
        if proprietor:
            html_parts.extend([
                "<div class='meta-item'>",
                "<span class='meta-label'>Proprietor</span>",
                f"<span class='meta-value'>{proprietor}</span>",
                "</div>",
            ])
        
        if sector:
            html_parts.extend([
                "<div class='meta-item'>",
                "<span class='meta-label'>Sector</span>",
                f"<span class='meta-value'>{sector}</span>",
                "</div>",
            ])
        
        if nature_of_business:
            html_parts.extend([
                "<div class='meta-item'>",
                "<span class='meta-label'>Nature of Business</span>",
                f"<span class='meta-value'>{nature_of_business}</span>",
                "</div>",
            ])
        
        html_parts.extend([
            "<div class='meta-item'>",
            "<span class='meta-label'>Generated</span>",
            f"<span class='meta-value'>{datetime.datetime.now().strftime('%b %d, %Y')}</span>",
            "</div>",
            "</div>",
            "</div>",
            "</div>",
            "",
            "<!-- Stats Grid -->",
        ])
        
        html_parts.extend([
            "</div>",
            "",
            "<!-- Table Section -->",
            "<div class='table-section'>",
            "<h2 class='section-title'>Financial Details</h2>",
            "<p class='section-subtitle'>Comprehensive breakdown of financial data and calculations</p>",
            "<div class='table-wrapper'>",
            "<table>",
            "<thead>",
            "<tr>",
            "<th>Particulars</th>",
            "<th>Amount</th>",
            "</tr>",
            "</thead>",
            "<tbody>",
        ])
        
        # Process each row
        for row_idx in range(1, max_row + 1):
            row_data = []
            is_header = False
            is_total = False
            is_empty_row = True
            
            # First pass: collect row data
            for col_idx in range(1, max_col + 1):
                cell = sheet.Cells(row_idx, col_idx)
                cell_value = cell.Value
                
                if cell_value is None:
                    cell_value = ""
                elif isinstance(cell_value, (int, float)):
                    if isinstance(cell_value, float):
                        if cell_value % 1 == 0:
                            cell_value = int(cell_value)
                    # Store in JSON
                    json_data["data"][f"R{row_idx}C{col_idx}"] = cell_value
                else:
                    cell_value = str(cell_value)
                    json_data["data"][f"R{row_idx}C{col_idx}"] = cell_value
                
                if cell_value != "":
                    is_empty_row = False
                
                row_data.append({
                    "value": cell_value,
                    "cell": cell,
                    "col_idx": col_idx
                })
            
            # Skip completely empty rows
            if is_empty_row:
                continue
            
            # Detect row type
            first_value = str(row_data[0]["value"]).lower() if row_data else ""
            if any(keyword in first_value for keyword in ["step", "financials", "ratios", "particulars", "profit", "balance", "sheet", "statement"]):
                is_header = True
            elif any(keyword in first_value for keyword in ["total", "net", "grand"]):
                is_total = True
            
            # Build row HTML
            row_class = ""
            if is_header:
                row_class = " class='section-header'"
            elif is_total:
                row_class = " class='total-row'"
            elif "subtotal" in first_value or "sub-total" in first_value:
                row_class = " class='subtotal-row'"
            
            html_parts.append(f"  <tr{row_class}>")
            
            for cell_data in row_data:
                cell = cell_data["cell"]
                cell_value = cell_data["value"]
                col_idx = cell_data["col_idx"]
                
                # Determine cell class
                cell_classes = []
                if col_idx == 1:
                    cell_classes.append("item-name")
                else:
                    cell_classes.append("item-value")
                
                # Format numeric values as currency
                formatted_value = cell_value
                if isinstance(cell_value, (int, float)) and cell_value != "" and col_idx > 1:
                    cell_classes.append("currency")
                    # Format with commas but without currency symbol (CSS will add it)
                    if isinstance(cell_value, float):
                        formatted_value = f"{cell_value:,.2f}"
                    else:
                        formatted_value = f"{cell_value:,}"
                
                # Basic styling from cell
                style_parts = []
                
                # Background color
                try:
                    interior_color = cell.Interior.Color
                    if interior_color != 16777215:  # Not white
                        r = interior_color & 255
                        g = (interior_color >> 8) & 255
                        b = (interior_color >> 16) & 255
                        style_parts.append(f"background-color: rgb({r},{g},{b})")
                except:
                    pass
                
                # Font color (only if custom style not applied)
                try:
                    if not is_header and not is_total:
                        font_color = cell.Font.Color
                        if font_color != 0:  # Not black
                            r = font_color & 255
                            g = (font_color >> 8) & 255
                            b = (font_color >> 16) & 255
                            style_parts.append(f"color: rgb({r},{g},{b})")
                except:
                    pass
                
                # Font weight
                try:
                    if cell.Font.Bold and not is_header and not is_total:
                        style_parts.append("font-weight: bold")
                except:
                    pass
                
                style_attr = "; ".join(style_parts) if style_parts else ""
                class_attr = " ".join(cell_classes) if cell_classes else ""
                
                # Handle merged cells
                merge_attrs = ""
                try:
                    merge_area = cell.MergeArea
                    if merge_area.Cells.Count > 1:
                        rowspan = merge_area.Rows.Count
                        colspan = merge_area.Columns.Count
                        if cell.Row == merge_area.Row and cell.Column == merge_area.Column:
                            if rowspan > 1:
                                merge_attrs += f" rowspan='{rowspan}'"
                            if colspan > 1:
                                merge_attrs += f" colspan='{colspan}'"
                        else:
                            continue
                except:
                    pass
                
                # Output cell
                attrs = []
                if class_attr:
                    attrs.append(f"class='{class_attr}'")
                if style_attr:
                    attrs.append(f"style='{style_attr}'")
                
                attr_str = " " + " ".join(attrs) if attrs else ""
                html_parts.append(f"    <td{attr_str}{merge_attrs}>{formatted_value}</td>")
            
            html_parts.append("  </tr>")
        
        html_parts.extend([
            "</tbody>",
            "</table>",
            "</div>",
            "</div>",
            "",
            "<!-- Footer Section -->",
            "<div class='report-footer'>",
            "<div class='footer-content'>",
            f"<h3 class='footer-title'><i class='fas fa-check-circle'></i> Report Generated Successfully</h3>", # Professional Icon
            "</div>",
            "<div class='timestamp-badge'>",
            f"<i class='fas fa-clock'></i> Generated on " + datetime.datetime.now().strftime('%B %d, %Y at %I:%M %p'), # Professional Icon
            "</div>",
            "</div>",
            "</div>",
            "",
            "</div>",
            "</div>",
            "",
            "<script>",
            "// Store JSON data for programmatic access",
            f"window.reportData = {json.dumps(json_data, ensure_ascii=False)};",
            "",
            "console.log('%c📊 Financial Report Data Loaded', 'color: #7c3aed; font-weight: bold; font-size: 16px; font-family: Inter, sans-serif;');",
            "console.log('%c━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━', 'color: #7c3aed;');",
            "console.log('%c📄 Sheet Name:', 'color: #6b7280; font-weight: 600;', window.reportData.sheetName);",
            "console.log('%c🔢 Total Cells:', 'color: #6b7280; font-weight: 600;', Object.keys(window.reportData.data).length);",
            "console.log('%c⏰ Timestamp:', 'color: #6b7280; font-weight: 600;', window.reportData.timestamp);",
            "console.log('%c━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━', 'color: #7c3aed;');",
            "console.log('%c💡 Access data: window.reportData.data[\"R1C1\"]', 'color: #10b981; font-style: italic;');",
            "",
            "// Download report as PDF (placeholder function)",
            "function downloadReport() {",
            "  alert('PDF download functionality will be implemented by the backend. (Ctrl+P to print)');", # Updated alert
            "  console.log('Download request initiated for:', window.reportData.sheetName);",
            "}",
            "",
            "// Add smooth scroll behavior",
            "document.querySelectorAll('a[href^=\"#\"]').forEach(anchor => {",
            "  anchor.addEventListener('click', function (e) {",
            "    e.preventDefault();",
            "    const target = document.querySelector(this.getAttribute('href'));",
            "    if (target) {",
            "      target.scrollIntoView({ behavior: 'smooth', block: 'start' });",
            "    }",
            "  });",
            "});",
            "",
            "// Add loading state handler",
            "window.addEventListener('load', () => {",
            "  document.querySelectorAll('.loading').forEach(el => {",
            "    el.classList.remove('loading');",
            "  });",
            "});",
            "",
            "// Add table row highlight on click",
            "document.querySelectorAll('tbody tr').forEach(row => {",
            "  row.addEventListener('click', function() {",
            "    document.querySelectorAll('tbody tr').forEach(r => {",
            "      r.style.outline = 'none';",
            "    });",
            "    this.style.outline = '2px solid var(--primary-purple)';", # Highlight with primary purple
            "    this.style.outlineOffset = '-2px';",
            "  });",
            "});",
            "",
            "// Add keyboard navigation",
            "document.addEventListener('keydown', (e) => {",
            "  if (e.ctrlKey && e.key === 'p') {",
            "    e.preventDefault();",
            "    window.print();",
            "  }",
            "});",
            "",
            "// Performance monitoring",
            "if (window.performance) {",
            "  const perfData = window.performance.timing;",
            "  const pageLoadTime = perfData.loadEventEnd - perfData.navigationStart;",
            "  console.log('%c⚡ Page Load Time:', 'color: #10b981; font-weight: 600;', pageLoadTime + 'ms');",
            "}",
            "",
            "// Add animation observer for elements",
            "const observerOptions = {",
            "  threshold: 0.1,",
            "  rootMargin: '0px 0px -50px 0px'",
            "};",
            "",
            "const observer = new IntersectionObserver((entries) => {",
            "  entries.forEach(entry => {",
            "    if (entry.isIntersecting) {",
            "      entry.target.style.opacity = '1';",
            "      entry.target.style.transform = 'translateY(0)';",
            "    }",
            "  });",
            "}, observerOptions);",
            "",
            "document.querySelectorAll('.stat-card, .table-wrapper').forEach(el => {",
            "  el.style.opacity = '0';",
            "  el.style.transform = 'translateY(20px)';",
            "  el.style.transition = 'opacity 0.6s ease, transform 0.6s ease';",
            "  observer.observe(el);",
            "});",
            "</script>",
            "</body>",
            "</html>"
        ])
        
        html_content = "\n".join(html_parts)
        print(f"[HTML COM Generator] SUCCESS: HTML generated successfully ({len(html_content)} chars)", file=sys.stderr)
        print(f"[HTML COM Generator] SUCCESS: JSON data extracted ({len(json_data['data'])} cells)", file=sys.stderr)
        
        # Clean up
        wb.Close(False)
        excel.Quit()
        
        return html_content, json_data
        
    except Exception as e:
        print(f"[HTML COM Generator] ❌ Error: {str(e)}", file=sys.stderr)
        import traceback
        traceback.print_exc(file=sys.stderr)
        try:
            if 'wb' in locals():
                wb.Close(False)
            if 'excel' in locals():
                excel.Quit()
        except:
            pass
        return ""

def generate_html_from_excel_sheet(excel_path: str, sheet_name: str):
    """
    Convert an Excel sheet to HTML with complete styling preservation.
    Returns tuple: (html_content, json_data) for both COM and fallback methods.
    FALLBACK NOW PROPERLY EVALUATES FORMULAS AND DISPLAYS VALUES.
    """
    try:
        print(f"[HTML Generator] Loading workbook: {excel_path}", file=sys.stderr)
        print(f"[HTML Generator] COM_AVAILABLE: {COM_AVAILABLE}", file=sys.stderr)
        
        # Try using win32com first to get calculated values
        html_content = None
        json_data = {}
        if COM_AVAILABLE:
            try:
                print(f"[HTML Generator] Attempting to use Excel COM method", file=sys.stderr)
                html_content, json_data = generate_html_from_excel_com(excel_path, sheet_name)
                if html_content:
                    print(f"[HTML Generator] Successfully generated HTML using COM", file=sys.stderr)
                    return html_content, json_data
                else:
                    print(f"[HTML Generator] COM method returned empty content", file=sys.stderr)
            except Exception as com_error:
                print(f"[HTML Generator] COM method failed, falling back to openpyxl: {com_error}", file=sys.stderr)
                import traceback
                traceback.print_exc(file=sys.stderr)
        else:
            print(f"[HTML Generator] COM not available, using openpyxl fallback", file=sys.stderr)
        
        # ========================================
        # FALLBACK METHOD WITH PROPER VALUE HANDLING
        # ========================================
        print(f"[HTML Generator] Using openpyxl fallback method WITH PROFESSIONAL STYLING", file=sys.stderr)
        
        # CRITICAL FIX: Use pandas to read the Excel file which properly evaluates formulas
        print(f"[HTML Generator] Reading Excel with pandas to get calculated values...", file=sys.stderr)
        
        # Read the specific sheet with pandas (it reads calculated values)
        df_dict = pd.read_excel(excel_path, sheet_name=None, engine='openpyxl')
        
        # Find matching sheet name
        actual_sheet_name = None
        for sheet_key in df_dict.keys():
            if normalize_sheet_name(sheet_key) == normalize_sheet_name(sheet_name):
                actual_sheet_name = sheet_key
                break
        
        if not actual_sheet_name:
            print(f"[HTML Generator] ERROR: Sheet '{sheet_name}' not found", file=sys.stderr)
            print(f"[HTML Generator] Available sheets: {list(df_dict.keys())}", file=sys.stderr)
            return "", {}
        
        df = df_dict[actual_sheet_name]
        print(f"[HTML Generator] Processing sheet: {actual_sheet_name} (matched from '{sheet_name}')", file=sys.stderr)
        print(f"[HTML Generator] Dataframe shape: {df.shape[0]} rows x {df.shape[1]} columns", file=sys.stderr)
        
        # Also load with openpyxl for styling and structure
        wb = load_workbook(excel_path, data_only=False)
        sheet = wb[actual_sheet_name]
        
        # Extract JSON data structure
        json_data = {
            "sheetName": actual_sheet_name,
            "data": {},
            "timestamp": datetime.datetime.now().isoformat()
        }
        
        # Extract firm details from the data for header
        firm_name = ""
        proprietor = ""
        sector = ""
        nature_of_business = ""
        
        # Try to get firm details from pandas dataframe (row 2, col 1 in 0-indexed)
        try:
            if len(df) >= 3 and len(df.columns) >= 2:
                firm_name_val = df.iloc[2, 1]  # Row 3, Col 2 (0-indexed)
                if pd.notna(firm_name_val):
                    firm_name = str(firm_name_val)
            if len(df) >= 4 and len(df.columns) >= 2:
                proprietor_val = df.iloc[3, 1]
                if pd.notna(proprietor_val):
                    proprietor = str(proprietor_val)
            if len(df) >= 6 and len(df.columns) >= 2:
                sector_val = df.iloc[5, 1]
                if pd.notna(sector_val):
                    sector = str(sector_val)
            if len(df) >= 7 and len(df.columns) >= 2:
                nature_val = df.iloc[6, 1]
                if pd.notna(nature_val):
                    nature_of_business = str(nature_val)
        except Exception as e:
            print(f"[HTML Generator] Warning: Could not extract firm details: {e}", file=sys.stderr)
        
        # Build HTML with EXACT SAME professional styling as COM method
        html_parts = [
            "<!DOCTYPE html>",
            "<html lang='en'>",
            "<head>",
            "<meta charset='UTF-8'>",
            "<meta name='viewport' content='width=device-width, initial-scale=1.0'>",
            f"<title>Financial Report - {firm_name or sheet_name}</title>",
            "<link href='https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap' rel='stylesheet'>",
            "<link rel='stylesheet' href='https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css'>",
            "<style>",
            "  :root {",
            "    --primary-color: #8b5cf6;",
            "    --primary-light: #a78bfa;",
            "    --primary-dark: #7c3aed;",
            "    --success-color: #10b981;",
            "    --text-primary: #1f2937;",
            "    --text-secondary: #6b7280;",
            "    --bg-primary: #ffffff;",
            "    --bg-secondary: #f9fafb;",
            "    --bg-accent: #f3f4f6;",
            "    --border-color: #e5e7eb;",
            "    --shadow-sm: 0 1px 2px 0 rgba(0, 0, 0, 0.05);",
            "    --shadow-md: 0 4px 6px -1px rgba(0, 0, 0, 0.1);",
            "    --shadow-lg: 0 10px 15px -3px rgba(0, 0, 0, 0.1);",
            "    --shadow-xl: 0 20px 25px -5px rgba(0, 0, 0, 0.1);",
            "  }",
            "  * { margin: 0; padding: 0; box-sizing: border-box; }",
            "  body {",
            "    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;",
            "    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);",
            "    min-height: 100vh;",
            "    padding: 40px 20px;",
            "    line-height: 1.6;",
            "    color: var(--text-primary);",
            "    -webkit-font-smoothing: antialiased;",
            "  }",
            "  .container { max-width: 1400px; margin: 0 auto; }",
            "  .report-card {",
            "    background: var(--bg-primary);",
            "    border-radius: 20px;",
            "    box-shadow: var(--shadow-xl);",
            "    overflow: hidden;",
            "    animation: slideUp 0.6s ease-out;",
            "  }",
            "  @keyframes slideUp {",
            "    from { opacity: 0; transform: translateY(30px); }",
            "    to { opacity: 1; transform: translateY(0); }",
            "  }",
            "  .report-header {",
            "    background: linear-gradient(135deg, var(--primary-color) 0%, var(--primary-dark) 100%);",
            "    padding: 48px 48px 32px;",
            "    color: white;",
            "    position: relative;",
            "  }",
            "  .report-header::before {",
            "    content: '';",
            "    position: absolute;",
            "    top: 0; right: 0;",
            "    width: 400px; height: 400px;",
            "    background: radial-gradient(circle, rgba(255,255,255,0.1) 0%, transparent 70%);",
            "    border-radius: 50%;",
            "    transform: translate(30%, -30%);",
            "  }",
            "  .header-content { position: relative; z-index: 1; }",
            "  .report-badge {",
            "    display: inline-block;",
            "    background: rgba(255, 255, 255, 0.2);",
            "    backdrop-filter: blur(10px);",
            "    padding: 8px 20px;",
            "    border-radius: 50px;",
            "    font-size: 13px;",
            "    font-weight: 600;",
            "    letter-spacing: 0.5px;",
            "    text-transform: uppercase;",
            "    margin-bottom: 20px;",
            "  }",
            "  .firm-name {",
            "    font-size: 36px;",
            "    font-weight: 700;",
            "    margin-bottom: 16px;",
            "    letter-spacing: -0.5px;",
            "  }",
            "  .firm-meta {",
            "    display: flex;",
            "    flex-wrap: wrap;",
            "    gap: 32px;",
            "    margin-top: 24px;",
            "    padding-top: 24px;",
            "    border-top: 1px solid rgba(255, 255, 255, 0.2);",
            "  }",
            "  .meta-item { display: flex; flex-direction: column; gap: 6px; }",
            "  .meta-label {",
            "    font-size: 12px;",
            "    font-weight: 500;",
            "    opacity: 0.9;",
            "    text-transform: uppercase;",
            "    letter-spacing: 1px;",
            "  }",
            "  .meta-value { font-size: 16px; font-weight: 600; }",
            "  .stats-grid {",
            "    display: grid;",
            "    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));",
            "    gap: 1px;",
            "    background: var(--border-color);",
            "    border-bottom: 1px solid var(--border-color);",
            "  }",
            "  .stat-card {",
            "    background: var(--bg-primary);",
            "    padding: 28px 32px;",
            "    text-align: center;",
            "    transition: all 0.3s ease;",
            "  }",
            "  .stat-card:hover {",
            "    background: var(--bg-secondary);",
            "    transform: translateY(-2px);",
            "  }",
            "  .stat-icon {",
            "    width: 48px; height: 48px;",
            "    margin: 0 auto 16px;",
            "    background: linear-gradient(135deg, var(--primary-light), var(--primary-color));",
            "    border-radius: 12px;",
            "    display: flex;",
            "    align-items: center;",
            "    justify-content: center;",
            "    font-size: 24px;",
            "  }",
            "  .stat-label {",
            "    font-size: 12px;",
            "    font-weight: 600;",
            "    color: var(--text-secondary);",
            "    text-transform: uppercase;",
            "    letter-spacing: 0.8px;",
            "    margin-bottom: 8px;",
            "  }",
            "  .stat-value {",
            "    font-size: 18px;",
            "    font-weight: 700;",
            "    color: var(--text-primary);",
            "  }",
            "  .table-section { padding: 48px; }",
            "  .section-title {",
            "    font-size: 24px;",
            "    font-weight: 700;",
            "    color: var(--text-primary);",
            "    margin-bottom: 8px;",
            "  }",
            "  .section-subtitle {",
            "    font-size: 14px;",
            "    color: var(--text-secondary);",
            "    margin-bottom: 32px;",
            "  }",
            "  .table-wrapper {",
            "    overflow-x: auto;",
            "    border-radius: 12px;",
            "    border: 1px solid var(--border-color);",
            "  }",
            "  table {",
            "    width: 100%;",
            "    border-collapse: collapse;",
            "    background: var(--bg-primary);",
            "  }",
            "  thead {",
            "    background: var(--bg-accent);",
            "    position: sticky;",
            "    top: 0;",
            "    z-index: 10;",
            "  }",
            "  thead th {",
            "    padding: 18px 24px;",
            "    text-align: left;",
            "    font-size: 12px;",
            "    font-weight: 700;",
            "    color: var(--text-primary);",
            "    text-transform: uppercase;",
            "    letter-spacing: 1px;",
            "    border-bottom: 2px solid var(--border-color);",
            "  }",
            "  thead th:last-child { text-align: right; }",
            "  tbody tr {",
            "    border-bottom: 1px solid var(--border-color);",
            "    transition: all 0.2s ease;",
            "  }",
            "  tbody tr:hover { background: var(--bg-secondary); }",
            "  tbody tr:last-child { border-bottom: none; }",
            "  td {",
            "    padding: 20px 24px;",
            "    font-size: 14px;",
            "    color: var(--text-primary);",
            "  }",
            "  .item-name { font-weight: 500; }",
            "  .item-value {",
            "    text-align: right;",
            "    font-weight: 600;",
            "    font-family: 'SF Mono', 'Monaco', 'Courier New', monospace;",
            "    font-size: 15px;",
            "  }",
            "  .currency::before {",
            "    content: '₹ ';",
            "    color: var(--text-secondary);",
            "    margin-right: 4px;",
            "    font-weight: 500;",
            "  }",
                        "  .section-header {",
            "    background: linear-gradient(135deg, #f3f4f6 0%, #e5e7eb 100%) !important;",
            "  }",
            "  .section-header td {",
            "    font-weight: 700 !important;",
            "    font-size: 14px !important;",
            "    color: var(--text-primary) !important;",
            "    padding: 16px 24px !important;",
            "    text-transform: uppercase;",
            "    letter-spacing: 0.5px;",
            "  }",
            "  .total-row {",
            "    background: linear-gradient(135deg, var(--primary-color), var(--primary-dark)) !important;",
            "  }",
            "  .total-row td {",
            "    color: white !important;",
            "    font-weight: 700 !important;",
            "    font-size: 16px !important;",
            "    padding: 24px !important;",
            "    border-top: 3px solid var(--primary-dark);",
            "  }",
            "  .subtotal-row {",
            "    background: var(--bg-accent) !important;",
            "  }",
            "  .subtotal-row td {",
            "    font-weight: 600 !important;",
            "    padding: 18px 24px !important;",
            "    color: var(--text-primary);",
            "  }",
            "  .report-footer {",
            "    background: var(--bg-secondary);",
            "    padding: 40px 48px;",
            "    text-align: center;",
            "    border-top: 1px solid var(--border-color);",
            "  }",
            "  .footer-content {",
            "    max-width: 600px;",
            "    margin: 0 auto;",
            "  }",
            "  .footer-title {",
            "    font-size: 16px;",
            "    font-weight: 600;",
            "    color: var(--text-primary);",
            "    margin-bottom: 12px;",
            "  }",
            "  .footer-text {",
            "    font-size: 13px;",
            "    color: var(--text-secondary);",
            "    line-height: 1.8;",
            "    margin-bottom: 24px;",
            "  }",
            "  .action-buttons {",
            "    display: flex;",
            "    gap: 16px;",
            "    justify-content: center;",
            "    flex-wrap: wrap;",
            "  }",
            "  .btn {",
            "    padding: 12px 32px;",
            "    border-radius: 10px;",
            "    font-weight: 600;",
            "    font-size: 14px;",
            "    cursor: pointer;",
            "    transition: all 0.3s ease;",
            "    border: none;",
            "    display: inline-flex;",
            "    align-items: center;",
            "    gap: 8px;",
            "    text-decoration: none;",
            "  }",
            "  .btn-primary {",
            "    background: linear-gradient(135deg, var(--primary-color), var(--primary-dark));",
            "    color: white;",
            "    box-shadow: 0 4px 12px rgba(139, 92, 246, 0.3);",
            "  }",
            "  .btn-primary:hover {",
            "    transform: translateY(-2px);",
            "    box-shadow: 0 6px 20px rgba(139, 92, 246, 0.4);",
            "  }",
            "  .btn-secondary {",
            "    background: var(--bg-primary);",
            "    color: var(--text-primary);",
            "    border: 2px solid var(--border-color);",
            "  }",
            "  .btn-secondary:hover {",
            "    background: var(--bg-accent);",
            "    border-color: var(--primary-color);",
            "  }",
            "  .timestamp-badge {",
            "    display: inline-flex;",
            "    align-items: center;",
            "    gap: 8px;",
            "    background: var(--bg-accent);",
            "    padding: 8px 16px;",
            "    border-radius: 8px;",
            "    font-size: 12px;",
            "    color: var(--text-secondary);",
            "    margin-top: 24px;",
            "  }",
            "  @media (max-width: 1024px) {",
            "    .report-header { padding: 40px 32px 24px; }",
            "    .firm-name { font-size: 28px; }",
            "    .table-section { padding: 32px 24px; }",
            "  }",
            "  @media (max-width: 768px) {",
            "    body { padding: 20px 10px; }",
            "    .report-card { border-radius: 16px; }",
            "    .report-header { padding: 32px 24px 20px; }",
            "    .firm-name { font-size: 24px; }",
            "    .firm-meta { gap: 20px; }",
            "    .stats-grid { grid-template-columns: repeat(2, 1fr); }",
            "    .table-section { padding: 24px 16px; }",
            "    .section-title { font-size: 20px; }",
            "    thead th, td { padding: 14px 16px; font-size: 13px; }",
            "    .report-footer { padding: 32px 24px; }",
            "    .action-buttons { flex-direction: column; }",
            "    .btn { width: 100%; justify-content: center; }",
            "  }",
            "  @media (max-width: 480px) {",
            "    .stats-grid { grid-template-columns: 1fr; }",
            "    .firm-meta { flex-direction: column; gap: 16px; }",
            "  }",
            "  @media print {",
            "    body { background: white; padding: 0; }",
            "    .report-card { box-shadow: none; border-radius: 0; }",
            "    .report-header::before { display: none; }",
            "    .action-buttons { display: none; }",
            "    tbody tr:hover { background: transparent; }",
            "  }",
            "</style>",
            "</head>",
            "<body>",
            "<div class='container'>",
            "<div class='report-card'>",
            "",
            "<!-- Header Section -->",
            "<div class='report-header'>",
            "<div class='header-content'>",
            f"<span class='report-badge'>📊 {sheet_name}</span>",
            f"<h1 class='firm-name'>{firm_name or 'Financial Report'}</h1>",
            "<div class='firm-meta'>",
        ]
        
        if proprietor:
            html_parts.extend([
                "<div class='meta-item'>",
                "<span class='meta-label'>Proprietor</span>",
                f"<span class='meta-value'>{proprietor}</span>",
                "</div>",
            ])
        
        if sector:
            html_parts.extend([
                "<div class='meta-item'>",
                "<span class='meta-label'>Sector</span>",
                f"<span class='meta-value'>{sector}</span>",
                "</div>",
            ])
        
        if nature_of_business:
            html_parts.extend([
                "<div class='meta-item'>",
                "<span class='meta-label'>Nature of Business</span>",
                f"<span class='meta-value'>{nature_of_business}</span>",
                "</div>",
            ])
        
        html_parts.extend([
            "<div class='meta-item'>",
            "<span class='meta-label'>Generated</span>",
            f"<span class='meta-value'>{datetime.datetime.now().strftime('%b %d, %Y')}</span>",
            "</div>",
            "</div>",
            "</div>",
            "</div>",
            "",
            "<!-- Stats Grid -->",
            "<div class='stats-grid'>",
            "<div class='stat-card'>",
            "<div class='stat-icon'>📄</div>",
            "<div class='stat-label'>Report Type</div>",
            f"<div class='stat-value'>{sheet_name}</div>",
            "</div>",
            "<div class='stat-card'>",
            "<div class='stat-icon'>📅</div>",
            "<div class='stat-label'>Date</div>",
            f"<div class='stat-value'>{datetime.datetime.now().strftime('%b %d, %Y')}</div>",
            "</div>",
            "<div class='stat-card'>",
            "<div class='stat-icon'>🔢</div>",
            "<div class='stat-label'>Report ID</div>",
            f"<div class='stat-value'>#{datetime.datetime.now().strftime('%Y%m%d%H%M')}</div>",
            "</div>",
        ])
        
        if sector:
            html_parts.extend([
                "<div class='stat-card'>",
                "<div class='stat-icon'>🏢</div>",
                "<div class='stat-label'>Sector</div>",
                f"<div class='stat-value'>{sector}</div>",
                "</div>",
            ])
        
        html_parts.extend([
            "</div>",
            "",
            "<!-- Table Section -->",
            "<div class='table-section'>",
            "<h2 class='section-title'>Financial Details</h2>",
            "<p class='section-subtitle'>Comprehensive breakdown of financial data and calculations</p>",
            "<div class='table-wrapper'>",
            "<table>",
            "<thead>",
            "<tr>",
            "<th>Particulars</th>",
            "<th>Amount</th>",
            "</tr>",
            "</thead>",
            "<tbody>",
        ])
        
        # Process merged cells from openpyxl
        merged_ranges = {}
        for merged_range in sheet.merged_cells.ranges:
            merged_ranges[(merged_range.min_row, merged_range.min_col)] = {
                'rowspan': merged_range.max_row - merged_range.min_row + 1,
                'colspan': merged_range.max_col - merged_range.min_col + 1
            }
        
        print(f"[HTML Generator] Processing {len(df)} rows from dataframe", file=sys.stderr)
        
        # Process each row from pandas dataframe
        for row_idx in range(len(df)):
            row_data = []
            is_empty_row = True
            
            # Get all column values for this row
            for col_idx in range(len(df.columns)):
                cell_value = df.iloc[row_idx, col_idx]
                
                # Handle NaN and None
                if pd.isna(cell_value):
                    cell_value = ""
                elif isinstance(cell_value, (int, float, np.integer, np.floating)):
                    # Convert numpy types to Python types
                    if isinstance(cell_value, np.floating):
                        cell_value = float(cell_value)
                        # Handle infinity
                        if np.isinf(cell_value):
                            cell_value = ""
                    elif isinstance(cell_value, np.integer):
                        cell_value = int(cell_value)
                    
                    # Store in JSON (1-indexed for consistency with COM method)
                    if cell_value != "":
                        json_data["data"][f"R{row_idx+1}C{col_idx+1}"] = cell_value
                else:
                    cell_value = str(cell_value)
                    if cell_value != "":
                        json_data["data"][f"R{row_idx+1}C{col_idx+1}"] = cell_value
                
                if cell_value != "":
                    is_empty_row = False
                
                row_data.append({
                    "value": cell_value,
                    "col_idx": col_idx + 1  # 1-indexed for HTML
                })
            
            # Skip completely empty rows
            if is_empty_row:
                continue
            
            # Detect row type
            first_value = str(row_data[0]["value"]).lower() if row_data else ""
            is_header = any(kw in first_value for kw in ["step", "financials", "ratios", "particulars", "profit", "balance", "sheet", "statement"])
            is_total = any(kw in first_value for kw in ["total", "net", "grand"])
            is_subtotal = "subtotal" in first_value or "sub-total" in first_value
            
            row_class = ""
            if is_header:
                row_class = " class='section-header'"
            elif is_total:
                row_class = " class='total-row'"
            elif is_subtotal:
                row_class = " class='subtotal-row'"
            
            html_parts.append(f"  <tr{row_class}>")
            
            for cell_data in row_data:
                cell_value = cell_data["value"]
                col_idx = cell_data["col_idx"]
                
                # Cell classes
                cell_classes = []
                if col_idx == 1:
                    cell_classes.append("item-name")
                else:
                    cell_classes.append("item-value")
                
                # Format numeric values as currency
                formatted_value = cell_value
                if isinstance(cell_value, (int, float)) and cell_value != "" and col_idx > 1:
                    cell_classes.append("currency")
                    # Format with commas but without currency symbol (CSS will add it)
                    try:
                        if isinstance(cell_value, float):
                            formatted_value = f"{cell_value:,.2f}"
                        else:
                            formatted_value = f"{cell_value:,}"
                    except:
                        formatted_value = str(cell_value)
                
                class_attr = " ".join(cell_classes) if cell_classes else ""
                html_parts.append(f"    <td class='{class_attr}'>{formatted_value}</td>")
            
            html_parts.append("  </tr>")
        
        html_parts.extend([
            "</tbody>",
            "</table>",
            "</div>",
            "</div>",
            "",
            "<!-- Footer Section -->",
            "<div class='report-footer'>",
            "<div class='footer-content'>",
            "<h3 class='footer-title'>🎉 Report Generated Successfully</h3>",
            "<p class='footer-text'>",
            "This financial report has been automatically generated with professional formatting. ",
            "All calculations are based on the provided data and formulas.",
            "</p>",
            "<div class='action-buttons'>",
            "<button class='btn btn-primary' onclick='window.print()'>",
            "🖨️ Print Report",
            "</button>",
            "<button class='btn btn-secondary' onclick='downloadReport()'>",
            "📥 Download PDF",
            "</button>",
            "</div>",
            "<div class='timestamp-badge'>",
            "⏰ Generated on " + datetime.datetime.now().strftime('%B %d, %Y at %I:%M %p'),
            "</div>",
            "</div>",
                        "</div>",
            "",
            "</div>",
            "</div>",
            "",
            "<script>",
            "// Store JSON data for programmatic access",
            f"window.reportData = {json.dumps(json_data, ensure_ascii=False)};",
            "",
            "console.log('%c📊 Financial Report Data Loaded', 'color: #8b5cf6; font-weight: bold; font-size: 16px; font-family: Inter, sans-serif;');",
            "console.log('%c━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━', 'color: #8b5cf6;');",
            "console.log('%c📄 Sheet Name:', 'color: #6b7280; font-weight: 600;', window.reportData.sheetName);",
            "console.log('%c🔢 Total Cells:', 'color: #6b7280; font-weight: 600;', Object.keys(window.reportData.data).length);",
            "console.log('%c⏰ Timestamp:', 'color: #6b7280; font-weight: 600;', window.reportData.timestamp);",
            "console.log('%c━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━', 'color: #8b5cf6;');",
            "console.log('%c💡 Access data: window.reportData.data[\"R1C1\"]', 'color: #10b981; font-style: italic;');",
            "",
            "// Download report as PDF (placeholder function)",
            "function downloadReport() {",
            "  alert('PDF download functionality will be implemented by the backend.');",
            "  console.log('Download request initiated for:', window.reportData.sheetName);",
            "}",
            "",
            "// Add smooth scroll behavior",
            "document.querySelectorAll('a[href^=\"#\"]').forEach(anchor => {",
            "  anchor.addEventListener('click', function (e) {",
            "    e.preventDefault();",
            "    const target = document.querySelector(this.getAttribute('href'));",
            "    if (target) {",
            "      target.scrollIntoView({ behavior: 'smooth', block: 'start' });",
            "    }",
            "  });",
            "});",
            "",
            "// Add loading state handler",
            "window.addEventListener('load', () => {",
            "  document.querySelectorAll('.loading').forEach(el => {",
            "    el.classList.remove('loading');",
            "  });",
            "});",
            "",
            "// Add table row highlight on click",
            "document.querySelectorAll('tbody tr').forEach(row => {",
            "  row.addEventListener('click', function() {",
            "    // Remove previous highlights",
            "    document.querySelectorAll('tbody tr').forEach(r => {",
            "      r.style.outline = 'none';",
            "    });",
            "    // Add highlight to clicked row",
            "    this.style.outline = '2px solid #8b5cf6';",
            "    this.style.outlineOffset = '-2px';",
            "  });",
            "});",
            "",
            "// Add keyboard navigation",
            "document.addEventListener('keydown', (e) => {",
            "  if (e.ctrlKey && e.key === 'p') {",
            "    e.preventDefault();",
            "    window.print();",
            "  }",
            "});",
            "",
            "// Performance monitoring",
            "if (window.performance) {",
            "  const perfData = window.performance.timing;",
            "  const pageLoadTime = perfData.loadEventEnd - perfData.navigationStart;",
            "  console.log('%c⚡ Page Load Time:', 'color: #10b981; font-weight: 600;', pageLoadTime + 'ms');",
            "}",
            "",
            "// Add animation observer for elements",
            "const observerOptions = {",
            "  threshold: 0.1,",
            "  rootMargin: '0px 0px -50px 0px'",
            "};",
            "",
            "const observer = new IntersectionObserver((entries) => {",
            "  entries.forEach(entry => {",
            "    if (entry.isIntersecting) {",
            "      entry.target.style.opacity = '1';",
            "      entry.target.style.transform = 'translateY(0)';",
            "    }",
            "  });",
            "}, observerOptions);",
            "",
            "document.querySelectorAll('.stat-card, .table-wrapper').forEach(el => {",
            "  el.style.opacity = '0';",
            "  el.style.transform = 'translateY(20px)';",
            "  el.style.transition = 'opacity 0.6s ease, transform 0.6s ease';",
            "  observer.observe(el);",
            "});",
            "</script>",
            "</body>",
            "</html>"
        ])
        
        html_content = "\n".join(html_parts)
        print(f"[HTML Generator] SUCCESS: HTML generated using FALLBACK with professional styling ({len(html_content)} chars)", file=sys.stderr)
        print(f"[HTML Generator] SUCCESS: JSON data extracted ({len(json_data['data'])} cells)", file=sys.stderr)
        
        # Close workbook
        wb.close()
        
        return html_content, json_data
        
    except Exception as e:
        print(f"[HTML Generator] ❌ Error generating HTML: {str(e)}", file=sys.stderr)
        import traceback
        traceback.print_exc(file=sys.stderr)
        return ""




def _abs_path(path: str) -> str:
    return os.path.abspath(path)


def _r1c1_to_a1(r1c1_ref: str) -> str:
    """
    Convert R1C1 notation (e.g., 'R3C2') to A1 notation (e.g., 'B3').
    
    Args:
        r1c1_ref: Cell reference in R1C1 format (e.g., 'R3C2')
        
    Returns:
        Cell reference in A1 format (e.g., 'B3')
    """
    import re
    match = re.match(r'R(\d+)C(\d+)', r1c1_ref, re.IGNORECASE)
    if not match:
        # If it's not R1C1 format, assume it's already A1 and return as-is
        return r1c1_ref
    
    row = int(match.group(1))
    col = int(match.group(2))
    
    # Convert column number to letter(s)
    col_letter = ''
    while col > 0:
        col -= 1
        col_letter = chr(65 + (col % 26)) + col_letter
        col //= 26
    
    return f'{col_letter}{row}'


def _collect_updates(workbook, updates: List[Dict[str, Any]]):
    from openpyxl.cell.cell import MergedCell
    
    print(f"[_collect_updates] Processing {len(updates)} updates", file=sys.stderr)
    
    applied = []
    for update in updates:
        sheet_name = update.get('sheet')
        cell_addr = update.get('cell')
        value = update.get('value')

        if not sheet_name or not cell_addr:
            continue

        if sheet_name not in workbook.sheetnames:
            raise ValueError(f'Sheet "{sheet_name}" not found in workbook')

        # Convert R1C1 notation to A1 if needed
        cell_addr_a1 = _r1c1_to_a1(cell_addr)

        sheet = workbook[sheet_name]
        cell = sheet[cell_addr_a1]
        
        # Debug logging for specific cells
        if cell_addr.lower() in ['i34', 'i35', 'h28', 'h30', 'h32', 'h33', 'h13', 'h14', 'h15']:
            print(f"[Update Debug] Cell {cell_addr} -> {cell_addr_a1} = {value}", file=sys.stderr)
        
        # Handle merged cells - write to the top-left cell of the merged range
        if isinstance(cell, MergedCell):
            # Find the merged range that contains this cell
            for merged_range in sheet.merged_cells.ranges:
                if cell.coordinate in merged_range:
                    # Get the top-left cell of the merged range (this is the "master" cell)
                    top_left_cell = sheet.cell(merged_range.min_row, merged_range.min_col)
                    top_left_cell.value = value
                    print(f"[Update] Merged cell {cell_addr_a1} -> writing to master cell {top_left_cell.coordinate}", file=sys.stderr)
                    break
        else:
            # Normal cell - write directly
            cell.value = value

        applied.append({
            'sheet': sheet_name,
            'cell': cell_addr,  # Keep original format in response
            'value': value
        })

    print(f"[_collect_updates] Applied {len(applied)} updates successfully", file=sys.stderr)
    return applied


def calculate_excel(input_data: Dict[str, Any], excel_path: str) -> str:
    meta: Dict[str, Any] = {
        'templatePath': _abs_path(excel_path),
        'autoCalculation': 'enabled'  # Excel will auto-calculate formulas
    }

    try:
        workbook = openpyxl.load_workbook(excel_path)
        applied_updates = _collect_updates(workbook, input_data.get('updates', []))

        # Use TEMP_DIR environment variable, fallback to system temp directory
        import tempfile
        output_dir = os.getenv('TEMP_DIR', tempfile.gettempdir())
        os.makedirs(output_dir, exist_ok=True)

        timestamp = datetime.datetime.utcnow().strftime('%Y%m%dT%H%M%SZ')
        template_name = os.path.splitext(os.path.basename(excel_path))[0]
        output_path = _abs_path(
            os.path.join(output_dir, f'{template_name}-updated-{timestamp}.xlsx')
        )

        workbook.save(output_path)
        
        # CRITICAL: Force Excel to recalculate all formulas using COM
        print(f"[Excel Calculator] Forcing formula recalculation via COM...", file=sys.stderr)
        try:
            if COM_AVAILABLE:
                import win32com.client
                excel_app = win32com.client.Dispatch("Excel.Application")
                excel_app.Visible = False
                excel_app.DisplayAlerts = False
                
                wb_com = excel_app.Workbooks.Open(output_path)
                wb_com.Application.Calculation = -4105  # xlCalculationAutomatic
                wb_com.Application.CalculateFullRebuild()  # Force full recalculation
                wb_com.Save()
                wb_com.Close(SaveChanges=True)
                excel_app.Quit()
                
                print(f"[Excel Calculator] ✓ Formulas recalculated successfully", file=sys.stderr)
            else:
                print(f"[Excel Calculator] ⚠ COM not available, formulas NOT recalculated", file=sys.stderr)
        except Exception as calc_error:
            print(f"[Excel Calculator] ⚠ Formula recalculation failed: {calc_error}", file=sys.stderr)

        # Read the Excel file as bytes and encode to base64
        with open(output_path, 'rb') as f:
            excel_bytes = f.read()
        excel_base64 = base64.b64encode(excel_bytes).decode('utf-8')

        # Also extract JSON data for browser display in Luckysheet format
        try:
            import pandas as pd
            all_sheets = pd.read_excel(output_path, sheet_name=None, engine='openpyxl')
            json_output = []
            for sheet_name, df in all_sheets.items():
                try:
                    df_cleaned = df.replace([pd.NA, np.inf, -np.inf], None)
                    df_cleaned = df_cleaned.where(pd.notna(df_cleaned), None)

                    # Convert to Luckysheet format
                    sheet_data = []
                    max_rows = len(df_cleaned)
                    max_cols = len(df_cleaned.columns) if max_rows > 0 else 0
                    
                    for row_idx in range(max_rows):
                        row_data = []
                        for col_idx in range(max_cols):
                            try:
                                value = df_cleaned.iloc[row_idx, col_idx] if row_idx < len(df_cleaned) else None
                                if value is not None and not pd.isna(value):
                                    cell_data = {
                                        'v': value,
                                        'm': str(value) if value is not None else ''
                                    }
                                else:
                                    cell_data = None
                                row_data.append(cell_data)
                            except Exception as cell_error:
                                print(f"Error processing cell {row_idx},{col_idx}: {cell_error}", file=sys.stderr)
                                row_data.append(None)
                        sheet_data.append(row_data)
                    
                    sheet_obj = {
                        'name': sheet_name,
                        'data': sheet_data,
                        'config': {
                            'merge': {},
                            'borderInfo': [],
                            'rowlen': {},
                            'columnlen': {}
                        },
                        'index': len(json_output)  # sheet index
                    }
                    json_output.append(sheet_obj)
                except Exception as sheet_error:
                    print(f"Error processing sheet {sheet_name}: {sheet_error}", file=sys.stderr)
                    # Skip this sheet
                    continue
        except Exception as json_error:
            print(f"Error generating JSON data: {json_error}", file=sys.stderr)
            json_output = []  # Fallback to empty array

        # Determine sheet name based on template
        final_sheet_name = 'Final workings' if 'CC6' in template_name else 'Finalworkings'

        # Generate PDF for Final workings sheet directly from Excel
        pdf_base64 = None
        pdf_file_name = None
        try:
            pdf_output_path = os.path.join(output_dir, f'{template_name}-{final_sheet_name}-{timestamp}.pdf')
            if generate_pdf_from_excel_sheet(output_path, final_sheet_name, pdf_output_path):
                with open(pdf_output_path, 'rb') as f:
                    pdf_bytes = f.read()
                pdf_base64 = base64.b64encode(pdf_bytes).decode('utf-8')
                pdf_file_name = f'{template_name}-{final_sheet_name}-{timestamp}.pdf'
                print(f"PDF generated successfully: {pdf_file_name}", file=sys.stderr)
                # Clean up PDF file after encoding
                os.unlink(pdf_output_path)
            else:
                print("PDF generation failed", file=sys.stderr)
        except Exception as pdf_error:
            print(f"Error generating PDF: {pdf_error}", file=sys.stderr)

        # Generate HTML for Final workings sheet with exact formatting
        html_content = None
        html_json_data = {}
        try:
            result_tuple = generate_html_from_excel_sheet(output_path, final_sheet_name)
            # Handle both old return (str) and new return (tuple) for backward compatibility
            if isinstance(result_tuple, tuple):
                html_content, html_json_data = result_tuple
            else:
                html_content = result_tuple
                html_json_data = {}
            
            if html_content:
                print(f"HTML generated successfully ({len(html_content)} chars)", file=sys.stderr)
                if html_json_data:
                    print(f"HTML JSON data extracted ({len(html_json_data.get('data', {}))} cells)", file=sys.stderr)
            else:
                print("HTML generation returned empty content", file=sys.stderr)
        except Exception as html_error:
            print(f"Error generating HTML: {html_error}", file=sys.stderr)

        # Generate full AI-enhanced report if requested
        full_report_base64 = None
        full_report_filename = None
        if input_data.get('generateFullReport', False) and (input_data.get('grokApiKey') or input_data.get('perplexityApiKey')):
            try:
                print(f"\n{'='*80}", file=sys.stderr)
                print(f"🚀 GENERATING FULL AI-ENHANCED REPORT", file=sys.stderr)
                print(f"{'='*80}\n", file=sys.stderr)
                
                # Import AI report generator
                from pdf_report_generator import AIReportGenerator
                
                # Create PDFs directory
                pdfs_dir = os.path.join(output_dir, f'pdfs_{timestamp}')
                os.makedirs(pdfs_dir, exist_ok=True)
                
                # Generate PDFs for all sheets
                print("[Full Report] Step 1: Generating PDFs for all Excel sheets...", file=sys.stderr)
                sheet_pdfs = generate_pdfs_for_all_sheets(output_path, pdfs_dir)
                print(f"[Full Report] Generated {sheet_pdfs['success_count']} sheet PDFs", file=sys.stderr)
                
                # Prepare Excel data for AI
                excel_data = {
                    'json_data': json_output,
                    'html_data': html_json_data,
                    'template_name': template_name,
                    'timestamp': timestamp
                }
                
                # Initialize AI generator - Grok is the default/preferred AI provider
                if input_data.get('grokApiKey'):
                    ai_generator = AIReportGenerator(input_data['grokApiKey'], provider="grok")
                    print("[Full Report] Using Grok AI for report generation", file=sys.stderr)
                elif input_data.get('perplexityApiKey'):
                    ai_generator = AIReportGenerator(input_data['perplexityApiKey'], provider="perplexity")
                    print("[Full Report] Using Perplexity AI for report generation (fallback)", file=sys.stderr)
                else:
                    # Default to Grok if no specific API key provided
                    grok_key = input_data.get('grokApiKey') or os.environ.get('GROK_API_KEY')
                    if grok_key:
                        ai_generator = AIReportGenerator(grok_key, provider="grok")
                        print("[Full Report] Using Grok AI for report generation (default)", file=sys.stderr)
                    else:
                        # Final fallback to Perplexity
                        perplexity_key = input_data.get('perplexityApiKey') or os.environ.get('PERPLEXITY_API_KEY')
                        if perplexity_key:
                            ai_generator = AIReportGenerator(perplexity_key, provider="perplexity")
                            print("[Full Report] Using Perplexity AI for report generation (final fallback)", file=sys.stderr)
                        else:
                            raise ValueError("No AI API key provided. Set GROK_API_KEY or PERPLEXITY_API_KEY environment variable, or provide apiKey in request.")
                
                # Generate full report
                full_report_path = os.path.join(output_dir, f'{template_name}-full-report-{timestamp}.pdf')
                print("[Full Report] Step 2: Generating AI content and merging...", file=sys.stderr)
                
                report_result = ai_generator.generate_full_report(
                    excel_pdfs_dir=pdfs_dir,
                    excel_data=excel_data,
                    output_path=full_report_path,
                    template_name=template_name
                )
                
                if report_result['success']:
                    # Read and encode the full report
                    with open(full_report_path, 'rb') as f:
                        full_report_bytes = f.read()
                    full_report_base64 = base64.b64encode(full_report_bytes).decode('utf-8')
                    full_report_filename = os.path.basename(full_report_path)
                    
                    print(f"[Full Report] ✅ Full report generated: {full_report_filename}", file=sys.stderr)
                    print(f"[Full Report]    AI Sections: {len(report_result.get('ai_sections_generated', []))}", file=sys.stderr)
                    print(f"[Full Report]    Excel PDFs: {len(report_result.get('excel_pdfs_included', []))}", file=sys.stderr)
                    
                    # Clean up individual Excel sheet PDFs (only keep final report)
                    try:
                        import shutil
                        if os.path.exists(pdfs_dir):
                            shutil.rmtree(pdfs_dir)
                            print(f"[Full Report] 🗑️  Cleaned up individual sheet PDFs from {pdfs_dir}", file=sys.stderr)
                    except Exception as cleanup_error:
                        print(f"[Full Report] ⚠️  Could not clean up temp PDFs: {str(cleanup_error)}", file=sys.stderr)
                else:
                    print(f"[Full Report] ❌ Report generation failed", file=sys.stderr)
                    
            except Exception as full_report_error:
                print(f"[Full Report] Error generating full report: {str(full_report_error)}", file=sys.stderr)
                import traceback
                traceback.print_exc(file=sys.stderr)

        meta['verificationCopy'] = output_path

        result = {
            'success': True,
            'message': 'Workbook updated, encoded, PDF and HTML generated',
            '_appliedUpdates': applied_updates,
            '_meta': meta,
            'excelData': excel_base64,
            'jsonData': json_output,
            'pdfData': pdf_base64,
            'pdfFileName': pdf_file_name,
            'htmlContent': html_content,
            'htmlJsonData': html_json_data,  # Add extracted JSON data from HTML
            'fileName': f'{template_name}-updated-{timestamp}.xlsx',
            'fullReportData': full_report_base64,  # AI-enhanced full report
            'fullReportFileName': full_report_filename
        }

        return json.dumps(result, ensure_ascii=False, default=str)
    except Exception as exc:  # pragma: no cover - operational guard
        return json.dumps({'success': False, 'error': str(exc)})


if __name__ == '__main__':
    import sys
    import io
    
    # Force UTF-8 encoding for stdout to handle emojis and special characters
    if sys.stdout.encoding != 'utf-8':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    
    args = sys.argv[1:]
    excel_file_path = args[0]
    json_input_string = args[1]

    payload = json.loads(json_input_string)
    outcome = calculate_excel(payload, excel_file_path)
    print(outcome)
