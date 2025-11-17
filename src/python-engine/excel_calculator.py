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
        
        # Build HTML with modern receipt styling (similar to reference)
        html_parts = [
            "<!DOCTYPE html>",
            "<html>",
            "<head>",
            "<meta charset='UTF-8'>",
            f"<title>Financial Report - {firm_name or sheet_name}</title>",
            "<style>",
            "  * { margin: 0; padding: 0; box-sizing: border-box; }",
            "  body { margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif; background: #f7f8fc; }",
            "  ",
            "  .receipt-wrapper { min-height: 100vh; padding: 40px 20px; display: flex; justify-content: center; align-items: flex-start; }",
            "  .receipt-container { max-width: 1400px; width: 100%; background: white; border-radius: 16px; box-shadow: 0 4px 24px rgba(0,0,0,0.08); overflow: scroll; }",
            "  ",
            "  /* Top Info Cards */",
            "  .top-info { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 1px; background: #e8eaf6; padding: 1px; }",
            "  .info-card { background: white; padding: 20px 24px; text-align: center; }",
            "  .info-label { font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.8px; color: #000; margin-bottom: 8px; }",
            "  .info-value { font-size: 15px; font-weight: 600; color: #000; }",
            "  .info-value.primary { color: #000; font-size: 16px; }",
            "  ",
            "  /* Firm Details Section */",
            "  .firm-section { padding: 32px 40px; background: #00000017; color: #000; }",
            "  .firm-name { font-size: 28px; font-weight: 700; margin-bottom: 16px; letter-spacing: -0.5px; }",
            "  .firm-details { display: flex; flex-wrap: wrap; gap: 24px; margin-top: 12px; }",
            "  .firm-detail-item { display: flex; align-items: center; gap: 8px; }",
            "  .firm-detail-label { font-size: 12px; font-weight: 500; }",
            "  .firm-detail-value { font-size: 14px; font-weight: 600; }",
            "  .report-meta { margin-top: 16px; padding-top: 16px; border-top: 1px solid rgba(0,0,0,0.1); font-size: 13px; }",
            "  ",
            "  /* Table Section */",
            "  .table-section { padding: 40px; }",
            "  table { width: 100%; border-collapse: separate; border-spacing: 0; }",
            "  ",
            "  thead th { background: #00000017; color: #000; font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; padding: 16px 20px; text-align: left; border-bottom: 2px solid #e5e7eb; }",
            "  thead th:last-child { text-align: right; }",
            "  ",
            "  tbody tr { border-bottom: 1px solid #e5e7eb; transition: all 0.2s; }",
            "  tbody tr:hover { background: #fafafa; }",
            "  tbody tr:last-child { border-bottom: none; }",
            "  ",
            "  td { padding: 18px 20px; font-size: 14px; color: #000; }",
            "  .item-name { font-weight: 500; color: #000; }",
            "  .item-value { text-align: right; font-weight: 600; font-family: 'Courier New', monospace; color: #000; }",
            "  ",
            "  /* Special Rows */",
            "  .section-header { background: #00000017 !important; }",
            "  .section-header td { color: #000 !important; font-weight: 700; font-size: 13px; padding: 14px 20px; letter-spacing: 0.5px; }",
            "  ",
            "  .total-row { background: #00000017 !important; }",
            "  .total-row td { font-weight: 700; font-size: 15px; color: #000; padding: 20px; border-top: 2px solid #e5e7eb; border-bottom: 2px solid #e5e7eb; }",
            "  ",
            "  .subtotal-row { background: #fafafa !important; }",
            "  .subtotal-row td { font-weight: 600; padding: 16px 20px; color: #000; }",
            "  ",
            "  /* Currency */",
            "  .currency::before { content: '₹ '; color: #000; margin-right: 4px; }",
            "  ",
            "  /* Footer */",
            "  .receipt-footer { background: white; padding: 28px 40px; text-align: center; border-top: 1px solid #e5e7eb; }",
            "  .footer-note { font-size: 12px; color: #000; margin-bottom: 6px; }",
            "  .footer-main { font-size: 13px; color: #000; font-weight: 500; margin-bottom: 18px; }",
            "  ",
            "  .action-btn { display: inline-block; padding: 12px 32px; background: #00000017; color: #000; border: 1px solid #e5e7eb; border-radius: 8px; font-weight: 600; font-size: 14px; cursor: pointer; transition: all 0.3s; }",
            "  .action-btn:hover { background: rgba(254, 249, 195, 1); }",
            "  ",
            "  /* Responsive */",
            "  @media (max-width: 768px) {",
            "    .receipt-wrapper { padding: 20px 10px; }",
            "    .top-info { grid-template-columns: repeat(2, 1fr); }",
            "    .firm-section { padding: 24px 20px; }",
            "    .firm-name { font-size: 22px; }",
            "    .table-section { padding: 24px 20px; }",
            "    td, th { padding: 12px 16px; font-size: 13px; }",
            "  }",
            "  ",
            "  @media print {",
            "    body { background: white; }",
            "    .receipt-wrapper { padding: 0; }",
            "    .receipt-container { box-shadow: none; border-radius: 0; }",
            "    .action-btn { display: none; }",
            "  }",
            "</style>",
            "</head>",
            "<body>",
            "<div class='receipt-wrapper'>",
            "<div class='receipt-container'>",
            "",
            "<!-- Top Info Cards -->",
            "<div class='top-info'>",
            "  <div class='info-card'>",
            "    <div class='info-label'>Report Type</div>",
            f"    <div class='info-value'>{sheet_name}</div>",
            "  </div>",
            "  <div class='info-card'>",
            "    <div class='info-label'>Generated On</div>",
            f"    <div class='info-value'>{datetime.datetime.now().strftime('%b %d, %Y')}</div>",
            "  </div>",
            "  <div class='info-card'>",
            "    <div class='info-label'>Report ID</div>",
            f"    <div class='info-value primary'>#{datetime.datetime.now().strftime('%Y%m%d%H%M')}</div>",
            "  </div>",
        ]
        
        if sector:
            html_parts.append("  <div class='info-card'>")
            html_parts.append("    <div class='info-label'>Sector</div>")
            html_parts.append(f"    <div class='info-value'>{sector}</div>")
            html_parts.append("  </div>")
        
        html_parts.extend([
            "</div>",
            "",
            "<!-- Firm Details Section -->",
            "<div class='firm-section'>",
            f"  <div class='firm-name'>{firm_name or 'Financial Report'}</div>",
            "  <div class='firm-details'>",
        ])
        
        if proprietor:
            html_parts.append("    <div class='firm-detail-item'>")
            html_parts.append("      <span class='firm-detail-label'>Proprietor:</span>")
            html_parts.append(f"      <span class='firm-detail-value'>{proprietor}</span>")
            html_parts.append("    </div>")
        
        if nature_of_business:
            html_parts.append("    <div class='firm-detail-item'>")
            html_parts.append("      <span class='firm-detail-label'>Nature of Business:</span>")
            html_parts.append(f"      <span class='firm-detail-value'>{nature_of_business}</span>")
            html_parts.append("    </div>")
        
        html_parts.extend([
            "  </div>",
            f"  <div class='report-meta'>Generated on {datetime.datetime.now().strftime('%B %d, %Y at %I:%M %p')}</div>",
            "</div>",
            "",
            "<!-- Table Section -->",
            "<div class='table-section'>",
            "<table>",
            "<thead>",
            "<tr>",
            "<th>Particulars</th>",
            "<th>Amount</th>",
            "</tr>",
            "</thead>",
            "<tbody>"
        ])
        
        # Process each row
        current_section = None
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
            if any(keyword in first_value for keyword in ["step", "financials", "ratios", "particulars", "profit", "balance", "sheet"]):
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
                    if not is_header:
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
            "",
            "<!-- Receipt Footer -->",
            "<div class='receipt-footer'>",
            "  <div class='footer-main'>This is a computer-generated financial report</div>",
            "  <div class='footer-note'>For any queries or clarifications, please contact your financial advisor</div>",
            "  <button class='action-btn' onclick='window.print(); return false;'>Download / Print Report</button>",
            "</div>",
            "",
            "</div>",
            "</div>",
            "",
            "<script>",
            "// Store JSON data for programmatic access",
            f"window.reportData = {json.dumps(json_data, ensure_ascii=False)};",
            "console.log('%c Financial Report Data Loaded', 'color: #667eea; font-weight: bold; font-size: 14px;');",
            "console.log('Sheet Name:', window.reportData.sheetName);",
            "console.log('Total Cells:', Object.keys(window.reportData.data).length);",
            "console.log('Timestamp:', window.reportData.timestamp);",
            "console.log('%cAccess data: window.reportData.data[\"R1C1\"]', 'color: #999; font-style: italic;');",
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
        return "", {}


def generate_html_from_excel_sheet(excel_path: str, sheet_name: str):
    """
    Convert an Excel sheet to HTML with complete styling preservation.
    Returns tuple: (html_content, json_data) when using COM, or just html_content for fallback.
    Uses win32com to get calculated values instead of formulas.
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
        
        # Fallback to openpyxl method
        print(f"[HTML Generator] Using openpyxl fallback method", file=sys.stderr)
        wb = load_workbook(excel_path, data_only=True)  # data_only=True to get calculated values
        
        # Find the matching sheet name (handles case and space differences)
        actual_sheet_name = find_sheet_match(sheet_name, wb.sheetnames)
        if not actual_sheet_name:
            print(f"[HTML Generator] ERROR: Sheet '{sheet_name}' not found (tried case-insensitive matching)", file=sys.stderr)
            print(f"[HTML Generator] Available sheets: {wb.sheetnames}", file=sys.stderr)
            return ""
        
        sheet = wb[actual_sheet_name]
        print(f"[HTML Generator] Processing sheet: {actual_sheet_name} (matched from '{sheet_name}')", file=sys.stderr)
        
        # Helper function to convert Excel color to hex
        def get_hex_color(color):
            if color is None:
                return None
            if isinstance(color, str):
                return f"#{color}" if not color.startswith('#') else color
            
            # Handle openpyxl Color objects
            if hasattr(color, 'rgb'):
                rgb = color.rgb
                if rgb is not None:
                    # rgb could be a string or RGB object
                    if isinstance(rgb, str):
                        if len(rgb) >= 6:
                            # Handle ARGB format (first 2 chars are alpha)
                            if len(rgb) == 8:
                                return f"#{rgb[2:]}"
                            return f"#{rgb}"
                    else:
                        # RGB object - convert to hex
                        try:
                            return f"#{rgb:06x}"
                        except:
                            pass
            
            # Handle RGB tuple (r, g, b)
            if hasattr(color, 'r') and hasattr(color, 'g') and hasattr(color, 'b'):
                try:
                    r = int(color.r) if color.r else 0
                    g = int(color.g) if color.g else 0
                    b = int(color.b) if color.b else 0
                    return f"#{r:02x}{g:02x}{b:02x}"
                except:
                    pass
            
            return None
        
        # Helper function to get cell style
        def get_cell_style(cell):
            styles = []
            
            # Background color
            if cell.fill and cell.fill.start_color:
                bg_color = get_hex_color(cell.fill.start_color)
                if bg_color and bg_color != "#000000":
                    styles.append(f"background-color: {bg_color}")
            
            # Font styles
            if cell.font:
                if cell.font.color:
                    font_color = get_hex_color(cell.font.color)
                    if font_color:
                        styles.append(f"color: {font_color}")
                
                if cell.font.bold:
                    styles.append("font-weight: bold")
                
                if cell.font.italic:
                    styles.append("font-style: italic")
                
                if cell.font.size:
                    styles.append(f"font-size: {cell.font.size}pt")
                
                if cell.font.name:
                    styles.append(f"font-family: '{cell.font.name}', sans-serif")
            
            # Alignment
            if cell.alignment:
                if cell.alignment.horizontal:
                    h_align = cell.alignment.horizontal
                    if h_align == 'center':
                        styles.append("text-align: center")
                    elif h_align == 'right':
                        styles.append("text-align: right")
                    elif h_align == 'left':
                        styles.append("text-align: left")
                
                if cell.alignment.vertical:
                    v_align = cell.alignment.vertical
                    if v_align == 'center':
                        styles.append("vertical-align: middle")
                    elif v_align == 'top':
                        styles.append("vertical-align: top")
                    elif v_align == 'bottom':
                        styles.append("vertical-align: bottom")
            
            # Borders
            border_styles = []
            if cell.border:
                if cell.border.top and cell.border.top.style:
                    border_styles.append("border-top: 1px solid #000")
                if cell.border.bottom and cell.border.bottom.style:
                    border_styles.append("border-bottom: 1px solid #000")
                if cell.border.left and cell.border.left.style:
                    border_styles.append("border-left: 1px solid #000")
                if cell.border.right and cell.border.right.style:
                    border_styles.append("border-right: 1px solid #000")
            
            styles.extend(border_styles)
            
            # Padding for better appearance
            styles.append("padding: 4px 8px")
            styles.append("white-space: pre-wrap")
            
            return "; ".join(styles) if styles else ""
        
        # Build HTML
        html_parts = [
            "<!DOCTYPE html>",
            "<html>",
            "<head>",
            "<meta charset='UTF-8'>",
            f"<title>{sheet_name}</title>",
            "<style>",
            "  body { margin: 0; padding: 20px; font-family: 'Calibri', 'Arial', sans-serif; background: #f5f5f5; }",
            "  .excel-container { background: white; padding: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); overflow-x: auto; }",
            "  table { border-collapse: collapse; width: 100%; table-layout: auto; }",
            "  td, th { border: 1px solid #d0d0d0; min-width: 50px; }",
            "  td { overflow: hidden; }",
            "  .sheet-title { font-size: 18px; font-weight: bold; margin-bottom: 15px; color: #333; }",
            "</style>",
            "</head>",
            "<body>",
            "<div class='excel-container'>",
            f"<div class='sheet-title'>{sheet_name}</div>",
            "<table>"
        ]
        
        # Get the actual used range
        max_row = sheet.max_row
        max_col = sheet.max_column
        
        print(f"[HTML Generator] Processing {max_row} rows x {max_col} columns", file=sys.stderr)
        
        # Process merged cells
        merged_ranges = {}
        for merged_range in sheet.merged_cells.ranges:
            min_row, min_col = merged_range.min_row, merged_range.min_col
            max_row_merge, max_col_merge = merged_range.max_row, merged_range.max_col
            merged_ranges[(min_row, min_col)] = {
                'rowspan': max_row_merge - min_row + 1,
                'colspan': max_col_merge - min_col + 1
            }
        
        # Generate table rows
        for row_idx, row in enumerate(sheet.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col), start=1):
            html_parts.append("  <tr>")
            
            for col_idx, cell in enumerate(row, start=1):
                # Skip cells that are part of a merge (except the top-left cell)
                skip_cell = False
                for (merge_row, merge_col), merge_info in merged_ranges.items():
                    if merge_row != row_idx or merge_col != col_idx:
                        if (merge_row <= row_idx < merge_row + merge_info['rowspan'] and
                            merge_col <= col_idx < merge_col + merge_info['colspan']):
                            skip_cell = True
                            break
                
                if skip_cell:
                    continue
                
                # Get cell value - use calculated value for formulas, not the formula itself
                if cell.data_type == 'f':  # Formula cell
                    # Use the cached calculated value instead of the formula
                    cell_value = cell.value if hasattr(cell, '_value') else cell.value
                    # Try to get the calculated value from the cell's internal value
                    try:
                        # Access the internal calculated value
                        if hasattr(cell, 'internal_value'):
                            cell_value = cell.internal_value
                        elif hasattr(cell, '_value'):
                            cell_value = cell._value
                        else:
                            # Fallback: re-read the cell to get calculated value
                            cell_value = sheet.cell(row=cell.row, column=cell.column).value
                    except:
                        cell_value = cell.value
                else:
                    cell_value = cell.value
                    
                if cell_value is None:
                    cell_value = ""
                elif isinstance(cell_value, (int, float)):
                    # Format numbers nicely
                    if isinstance(cell_value, float):
                        cell_value = f"{cell_value:.2f}" if cell_value % 1 else str(int(cell_value))
                    else:
                        cell_value = str(cell_value)
                else:
                    cell_value = str(cell_value)
                
                # Get cell style
                style_attr = get_cell_style(cell)
                
                # Check if this cell is the start of a merged range
                merge_attrs = ""
                if (row_idx, col_idx) in merged_ranges:
                    merge_info = merged_ranges[(row_idx, col_idx)]
                    if merge_info['rowspan'] > 1:
                        merge_attrs += f" rowspan='{merge_info['rowspan']}'"
                    if merge_info['colspan'] > 1:
                        merge_attrs += f" colspan='{merge_info['colspan']}'"
                
                # Add cell to HTML
                if style_attr:
                    html_parts.append(f"    <td style='{style_attr}'{merge_attrs}>{cell_value}</td>")
                else:
                    html_parts.append(f"    <td{merge_attrs}>{cell_value}</td>")
            
            html_parts.append("  </tr>")
        
        html_parts.extend([
            "</table>",
            "</div>",
            "</body>",
            "</html>"
        ])
        
        html_content = "\n".join(html_parts)
        print(f"[HTML Generator] SUCCESS: HTML generated successfully ({len(html_content)} chars)", file=sys.stderr)
        
        return html_content
        
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

        output_dir = os.path.join(os.path.dirname(__file__), '../../temp')
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
