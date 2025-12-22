"""
PDF Regenerator Script
Regenerates PDF from an existing Excel file using the same logic as excel_calculator.py
Uses AIReportGenerator for proper AI content and sheet ordering.
Used by admin to regenerate PDF after uploading a revised Excel file.
"""

import sys
import os
import json
import base64
import re
import io
import tempfile
import shutil
from typing import List, Dict, Any, Optional
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

# Import the AI report generator and helper functions from excel_calculator
try:
    from pdf_report_generator import AIReportGenerator
    from excel_calculator import generate_pdfs_for_all_sheets
    AI_GENERATOR_AVAILABLE = True
except ImportError as e:
    print(f"[PDF Regenerator] Warning: Could not import AIReportGenerator or generate_pdfs_for_all_sheets: {e}", file=sys.stderr)
    AI_GENERATOR_AVAILABLE = False

# Check for Windows COM automation support
COM_AVAILABLE = False
try:
    import win32com.client
    import pythoncom
    COM_AVAILABLE = True
except ImportError:
    print("[PDF Regenerator] Warning: win32com not available. PDF generation will use fallback methods.", file=sys.stderr)


def generate_pdf_from_excel_sheet(excel_path: str, sheet_name: str, output_path: str) -> bool:
    """
    Generate a PDF from a specific sheet in an Excel workbook using Excel COM automation.
    This preserves all formatting, formulas, and styles.
    
    Args:
        excel_path: Path to the Excel file
        sheet_name: Name of the sheet to export
        output_path: Path to save the PDF file
        
    Returns:
        True if successful, False otherwise
    """
    if not COM_AVAILABLE:
        print(f"[PDF Generator] Excel COM automation not available.", file=sys.stderr)
        return False
    
    excel = None
    workbook = None
    co_initialized = False
    
    try:
        pythoncom.CoInitialize()
        co_initialized = True
        
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.ScreenUpdating = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        
        workbook = excel.Workbooks.Open(os.path.abspath(excel_path), ReadOnly=True)
        
        # Find the sheet
        sheet = None
        for i in range(1, workbook.Sheets.Count + 1):
            if workbook.Sheets(i).Name == sheet_name:
                sheet = workbook.Sheets(i)
                break
        
        if not sheet:
            print(f"[PDF Generator] Sheet '{sheet_name}' not found in workbook", file=sys.stderr)
            return False
        
        # Select the sheet
        sheet.Select()
        
        # Auto-fit columns to prevent ######## display for numeric values (skip for index sheet)
        try:
            if sheet.Name.strip().lower() != 'index':
                sheet.Columns.AutoFit()
        except Exception as e:
            print(f"[PDF Generator] Warning: Could not AutoFit columns: {str(e)}", file=sys.stderr)
        
        # Configure page setup
        page_setup = workbook.ActiveSheet.PageSetup
        page_setup.Zoom = False
        page_setup.FitToPagesWide = 1
        
        if sheet_name.lower() == 'coverpage':
            page_setup.FitToPagesTall = 1
            page_setup.Orientation = 1  # Portrait
        else:
            page_setup.FitToPagesTall = False
        
        page_setup.Orientation = 1  # Portrait
        page_setup.PaperSize = 9  # A4
        page_setup.LeftMargin = excel.InchesToPoints(0.5)
        page_setup.RightMargin = excel.InchesToPoints(0.5)
        page_setup.TopMargin = excel.InchesToPoints(0.5)
        page_setup.BottomMargin = excel.InchesToPoints(0.5)
        
        # Export as PDF
        workbook.ActiveSheet.ExportAsFixedFormat(
            Type=0,  # xlTypePDF
            Filename=os.path.abspath(output_path),
            Quality=0,
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False
        )
        
        return os.path.exists(output_path)
        
    except Exception as e:
        print(f"[PDF Generator] Error generating PDF: {str(e)}", file=sys.stderr)
        return False
        
    finally:
        if workbook:
            try:
                workbook.Close(SaveChanges=False)
            except:
                pass
        if excel:
            try:
                excel.Quit()
            except:
                pass
        if co_initialized:
            try:
                pythoncom.CoUninitialize()
            except:
                pass


def generate_pdfs_from_excel(excel_path: str, output_dir: str, include_sheets: Optional[List[str]] = None, excluded_sheets: Optional[List[str]] = None) -> Dict[str, Any]:
    """
    Generate individual PDF files for sheets in the Excel workbook.
    
    Args:
        excel_path: Path to the Excel file
        output_dir: Directory to save the PDF files
        include_sheets: List of sheet names to include (None = all sheets except excluded)
        excluded_sheets: List of sheet names to exclude (overrides default)
        
    Returns:
        Dictionary with generation results
    """
    print(f"\n{'='*80}", file=sys.stderr)
    print(f"📄 REGENERATING PDFs FROM EXCEL", file=sys.stderr)
    print(f"{'='*80}\n", file=sys.stderr)
    
    if excluded_sheets is not None:
        EXCLUDED_SHEETS = excluded_sheets
    else:
        EXCLUDED_SHEETS = ['Assumptions.1', 'Assumptions', 'assumptions', 'ASSUMPTIONS']
    
    pdf_files = {
        "sheets": {},
        "success_count": 0,
        "failed_count": 0,
        "total_sheets": 0,
        "excluded_sheets": [],
        "filtered_out_sheets": [],
        "requested_sheets": include_sheets or [],
        "sheet_status": []
    }
    
    # Build include filter if sheets are specified
    include_filter = None
    if include_sheets:
        def _normalized(value: str) -> str:
            return re.sub(r'[\s_\-]+', '', value.strip().lower())
        include_filter = {_normalized(sheet): sheet.strip() for sheet in include_sheets if sheet.strip()}
    
    if not COM_AVAILABLE:
        print(f"❌ Excel COM not available. Cannot generate PDFs.", file=sys.stderr)
        pdf_files["error"] = "Excel COM automation not available"
        return pdf_files
    
    excel = None
    workbook = None
    co_initialized = False
    
    try:
        pythoncom.CoInitialize()
        co_initialized = True
        
        os.makedirs(output_dir, exist_ok=True)
        
        print(f"[PDF Regenerator] Opening workbook: {excel_path}", file=sys.stderr)
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.ScreenUpdating = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        
        workbook = excel.Workbooks.Open(os.path.abspath(excel_path), ReadOnly=True)
        total_sheets = workbook.Sheets.Count
        pdf_files["total_sheets"] = total_sheets
        
        print(f"[PDF Regenerator] Found {total_sheets} sheets", file=sys.stderr)
        
        for sheet_idx in range(1, total_sheets + 1):
            sheet = workbook.Sheets(sheet_idx)
            sheet_name = sheet.Name
            
            normalized_sheet = sheet_name.strip()
            normalized_key = re.sub(r'[\s_\-]+', '', normalized_sheet.lower())
            
            # Skip excluded sheets
            if sheet_name in EXCLUDED_SHEETS:
                print(f"[{sheet_idx}/{total_sheets}] ⏭️  Skipping: '{sheet_name}' (excluded)", file=sys.stderr)
                pdf_files["excluded_sheets"].append(sheet_name)
                continue
            
            # Skip if not in include filter
            if include_filter and normalized_key not in include_filter:
                print(f"[{sheet_idx}/{total_sheets}] ⏭️  Skipping: '{sheet_name}' (not in requested list)", file=sys.stderr)
                pdf_files["filtered_out_sheets"].append(sheet_name)
                continue
            
            print(f"[{sheet_idx}/{total_sheets}] Processing sheet: '{sheet_name}'", file=sys.stderr)
            
            # Create PDF filename
            safe_sheet_name = re.sub(r'[<>:"/\\|?*]', '_', sheet_name)
            pdf_filename = f"sheet_{sheet_idx}_{safe_sheet_name}.pdf"
            pdf_path = os.path.join(output_dir, pdf_filename)
            
            try:
                sheet.Select()
                
                # Auto-fit columns to prevent ######## display for numeric values (skip for index sheet)
                try:
                    if sheet_name.strip().lower() != 'index':
                        sheet.Columns.AutoFit()
                except Exception as e:
                    print(f"   ⚠️  Warning: Could not AutoFit columns on '{sheet_name}': {str(e)}", file=sys.stderr)
                
                page_setup = workbook.ActiveSheet.PageSetup
                page_setup.Zoom = False
                page_setup.FitToPagesWide = 1
                
                if sheet_name.lower() == 'coverpage':
                    page_setup.FitToPagesTall = 1
                    page_setup.Orientation = 1
                else:
                    page_setup.FitToPagesTall = False
                
                page_setup.Orientation = 1
                page_setup.PaperSize = 9  # A4
                page_setup.LeftMargin = excel.InchesToPoints(0.5)
                page_setup.RightMargin = excel.InchesToPoints(0.5)
                page_setup.TopMargin = excel.InchesToPoints(0.5)
                page_setup.BottomMargin = excel.InchesToPoints(0.5)
                
                workbook.ActiveSheet.ExportAsFixedFormat(
                    Type=0,
                    Filename=os.path.abspath(pdf_path),
                    Quality=0,
                    IncludeDocProperties=True,
                    IgnorePrintAreas=False,
                    OpenAfterPublish=False
                )
                
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
                    pdf_files["sheets"][sheet_name] = {"status": "failed", "error": "PDF file not created"}
                    pdf_files["failed_count"] += 1
                    
            except Exception as sheet_error:
                print(f"   ❌ Error: {str(sheet_error)}", file=sys.stderr)
                pdf_files["sheets"][sheet_name] = {"status": "failed", "error": str(sheet_error)}
                pdf_files["failed_count"] += 1
        
        workbook.Close(SaveChanges=False)
        workbook = None
        excel.Quit()
        excel = None
        
        print(f"\n{'─'*80}", file=sys.stderr)
        print(f"✅ PDF Generation Complete", file=sys.stderr)
        print(f"   Success: {pdf_files['success_count']}, Failed: {pdf_files['failed_count']}", file=sys.stderr)
        
        return pdf_files
        
    except Exception as e:
        print(f"❌ Error in PDF generation: {str(e)}", file=sys.stderr)
        pdf_files["error"] = str(e)
        return pdf_files
        
    finally:
        if workbook:
            try:
                workbook.Close(SaveChanges=False)
            except:
                pass
        if excel:
            try:
                excel.Quit()
            except:
                pass
        if co_initialized:
            try:
                pythoncom.CoUninitialize()
            except:
                pass


def merge_pdfs(pdf_paths: List[str], output_path: str) -> bool:
    """
    Merge multiple PDF files into a single PDF.
    
    Args:
        pdf_paths: List of paths to PDF files to merge
        output_path: Path to save the merged PDF
        
    Returns:
        True if successful, False otherwise
    """
    try:
        merger = PdfMerger()
        
        for pdf_path in pdf_paths:
            if os.path.exists(pdf_path):
                merger.append(pdf_path)
        
        merger.write(output_path)
        merger.close()
        
        return os.path.exists(output_path)
        
    except Exception as e:
        print(f"[PDF Merger] Error merging PDFs: {str(e)}", file=sys.stderr)
        return False


def normalize_sheet_name(name: str) -> str:
    """Normalize sheet name for comparison."""
    return re.sub(r'[\s_\-]+', '', name.strip().lower())


def get_stamped_pdf(input_pdf_path, signature_path):
    """
    Creates a temporary PDF with the signature stamped on all pages.
    """
    try:
        if not signature_path or not os.path.exists(signature_path):
            return input_pdf_path

        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=A4)
        # Draw signature at bottom left (matching professional_pdf_template.py)
        # footer_y = 20mm, so footer_y - 5mm = 15mm
        # Use mask='auto' for transparency support
        can.drawImage(signature_path, 25*mm, 15*mm, width=30*mm, height=12*mm, mask='auto', preserveAspectRatio=True)
        can.save()
        packet.seek(0)
        
        new_pdf = PdfReader(packet)
        overlay_page = new_pdf.pages[0]
        
        existing_pdf = PdfReader(open(input_pdf_path, "rb"))
        output = PdfWriter()
        
        for i in range(len(existing_pdf.pages)):
            page = existing_pdf.pages[i]
            page.merge_page(overlay_page)
            output.add_page(page)
        
        temp_stamped = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        with open(temp_stamped.name, "wb") as outputStream:
            output.write(outputStream)
        
        return temp_stamped.name
    except Exception as e:
        print(f"[PDF Merger] Error stamping signature: {str(e)}", file=sys.stderr)
        return input_pdf_path


def get_sheet_order(template_name: str) -> List[str]:
    """
    Get the proper sheet order based on template type.
    Same as defined in pdf_report_generator.py's generate_full_report.
    """
    is_term_loan = 'TERM LOAN' in template_name.upper() or 'TERM_LOAN' in template_name.upper() or 'TL' in template_name.upper()
    
    if is_term_loan:
        return [
            'Cover page', 'Index', 'profile', 'Descriptive', 'project cost', 'PL BS', 'Graph',
            'RATIO', 'FA Sch', 'Dep IT act', 'Loan sch', 'Repayment', 'CFs', 'IRR', 'MIRR',
            'NPV', 'PI Index', 'WACC', 'Payback period I', 'Payback period II', 'Altman Z',
            'Sensitivity Analysis', 'workings for sensittivity1', 'Workings for Sensitivity2',
            'CF workings', 'Final workings', 'MPBF ', 'workings for sensitivity1', 'Gaurantors',
            'BEP analysis', 'Sales'
        ]
    else:
        # CC/OD template order
        return [
            'coverpage', 'final', 'PLBS', 'RATIO', 'Depsch', 'MPBF ', 'nayak', 'wp'
        ]


def merge_pdfs_with_order(pdfs_dir: str, output_path: str, template_name: str, signature_path: str = None) -> bool:
    """
    Merge PDFs in the proper order based on template type.
    Uses SHEET_ORDER to determine correct sequence.
    
    Args:
        pdfs_dir: Directory containing individual sheet PDFs
        output_path: Path to save the merged PDF
        template_name: Template name for determining sheet order
        signature_path: Path to the signature image to stamp on pages
        
    Returns:
        True if successful, False otherwise
    """
    try:
        print(f"[PDF Merger] Merging PDFs with template order: {template_name}", file=sys.stderr)
        
        # Get all PDF files in the directory
        pdf_files = {}
        if os.path.exists(pdfs_dir):
            for filename in os.listdir(pdfs_dir):
                if filename.endswith('.pdf'):
                    pdf_path = os.path.join(pdfs_dir, filename)
                    pdf_files[filename] = pdf_path
        
        if not pdf_files:
            print("[PDF Merger] No PDF files found to merge", file=sys.stderr)
            return False
        
        print(f"[PDF Merger] Found {len(pdf_files)} PDF files", file=sys.stderr)
        
        # Get proper sheet order
        sheet_order = get_sheet_order(template_name)
        
        # Build ordered list of PDF paths
        ordered_pdfs = []
        used_files = set()
        temp_files_to_cleanup = []
        
        # First, add PDFs in the defined order
        for sheet_name in sheet_order:
            normalized_target = normalize_sheet_name(sheet_name)
            
            for filename, filepath in pdf_files.items():
                if filename in used_files:
                    continue
                
                # Extract sheet name from filename (format: sheet_N_sheetname.pdf)
                match = re.match(r'sheet_\d+_(.+)\.pdf', filename, re.IGNORECASE)
                if match:
                    extracted_name = match.group(1).replace('_', ' ')
                    normalized_extracted = normalize_sheet_name(extracted_name)
                    
                    # Check for match
                    if normalized_target == normalized_extracted or normalized_target in normalized_extracted:
                        # Stamp signature if not cover or index
                        final_path = filepath
                        if signature_path and normalized_target not in ['coverpage', 'index', 'cover', 'indexpage', 'cover page']:
                            print(f"   ✍️ Stamping signature on: {sheet_name}", file=sys.stderr)
                            stamped_path = get_stamped_pdf(filepath, signature_path)
                            if stamped_path != filepath:
                                final_path = stamped_path
                                temp_files_to_cleanup.append(stamped_path)
                        
                        ordered_pdfs.append(final_path)
                        used_files.add(filename)
                        print(f"   ✓ Added: {sheet_name} -> {filename}", file=sys.stderr)
                        break
        
        # Add any remaining PDFs that weren't in the order list
        remaining = set(pdf_files.keys()) - used_files
        if remaining:
            print(f"   ⚠️ Adding {len(remaining)} unordered PDFs", file=sys.stderr)
            for filename in sorted(remaining):
                filepath = pdf_files[filename]
                
                # Extract sheet name for exclusion check
                match = re.match(r'sheet_\d+_(.+)\.pdf', filename, re.IGNORECASE)
                extracted_name = match.group(1).replace('_', ' ') if match else filename
                normalized_extracted = normalize_sheet_name(extracted_name)
                
                final_path = filepath
                if signature_path and normalized_extracted not in ['coverpage', 'index', 'cover', 'indexpage', 'cover page']:
                    print(f"   ✍️ Stamping signature on unordered: {filename}", file=sys.stderr)
                    stamped_path = get_stamped_pdf(filepath, signature_path)
                    if stamped_path != filepath:
                        final_path = stamped_path
                        temp_files_to_cleanup.append(stamped_path)
                
                ordered_pdfs.append(final_path)
                print(f"   + Added: {filename} (appended)", file=sys.stderr)
        
        if not ordered_pdfs:
            print("[PDF Merger] No PDFs to merge after ordering", file=sys.stderr)
            return False
        
        # Merge in order
        success = merge_pdfs(ordered_pdfs, output_path)
        
        # Cleanup temp stamped files
        for temp_file in temp_files_to_cleanup:
            try:
                if os.path.exists(temp_file):
                    os.remove(temp_file)
            except:
                pass
                
        return success
        
    except Exception as e:
        print(f"[PDF Merger] Error in ordered merge: {str(e)}", file=sys.stderr)
        return False


def regenerate_pdf_from_excel(
    excel_path: str, 
    selected_sheets: Optional[List[str]] = None, 
    output_pdf_path: Optional[str] = None,
    grok_api_key: Optional[str] = None,
    json_data: Optional[Dict[str, Any]] = None,
    html_data: Optional[Dict[str, Any]] = None,
    template_name: Optional[str] = None,
    signature_path: Optional[str] = None,
    excluded_sheets: Optional[List[str]] = None
) -> Dict[str, Any]:
    """
    Main function to regenerate PDF from an Excel file using AIReportGenerator.
    Follows the same flow as excel_calculator.py for proper sheet ordering and AI content.
    
    Args:
        excel_path: Path to the Excel file
        selected_sheets: List of sheet names to include in PDF (None = all sheets)
        output_pdf_path: Path to save the final merged PDF (optional)
        grok_api_key: Grok API key for AI content generation
        json_data: JSON data from original report for AI context
        html_data: HTML data from original report
        template_name: Template name (CC1, CC2, TL1, etc.) for proper sheet ordering
        signature_path: Path to the admin signature image
        excluded_sheets: List of sheet names to exclude (overrides default)
        
    Returns:
        Dictionary with:
            - success: bool
            - pdf_base64: Base64 encoded PDF data
            - pdf_filename: Suggested filename for the PDF
            - sheets_processed: Number of sheets processed
            - ai_sections_generated: Number of AI sections generated
            - error: Error message if failed
    """
    print(f"\n{'='*80}", file=sys.stderr)
    print(f"🔄 PDF REGENERATION FROM EXCEL (AIReportGenerator)", file=sys.stderr)
    print(f"   Excel: {excel_path}", file=sys.stderr)
    print(f"   Selected Sheets: {selected_sheets if selected_sheets else 'All sheets'}", file=sys.stderr)
    print(f"   Template: {template_name or 'Not specified'}", file=sys.stderr)
    print(f"   AI Content: {'Yes (Grok)' if grok_api_key else 'No (simple merge)'}", file=sys.stderr)
    print(f"   Signature: {'Yes' if signature_path else 'No'}", file=sys.stderr)
    print(f"{'='*80}\n", file=sys.stderr)
    
    result = {
        "success": False,
        "pdf_base64": None,
        "pdf_filename": None,
        "sheets_processed": 0,
        "ai_sections_generated": 0,
        "error": None
    }
    
    if not os.path.exists(excel_path):
        result["error"] = f"Excel file not found: {excel_path}"
        return result
    
    if not COM_AVAILABLE:
        result["error"] = "Excel COM automation not available"
        return result
    
    # Create temp directory for individual PDFs
    timestamp = __import__('datetime').datetime.now().strftime("%Y%m%d_%H%M%S")
    temp_dir = tempfile.mkdtemp(prefix="pdf_regen_")
    pdfs_dir = os.path.join(temp_dir, f'pdfs_{timestamp}')
    os.makedirs(pdfs_dir, exist_ok=True)
    
    try:
        # Step 1: Generate PDFs for all selected sheets using the same function as excel_calculator
        print("[PDF Regenerator] Step 1: Generating PDFs for selected Excel sheets...", file=sys.stderr)
        
        if AI_GENERATOR_AVAILABLE:
            # Use the same generate_pdfs_for_all_sheets from excel_calculator
            sheet_pdfs = generate_pdfs_for_all_sheets(
                excel_path,
                pdfs_dir,
                selected_sheets,
                excluded_sheets
            )
        else:
            # Fallback to local implementation
            sheet_pdfs = generate_pdfs_for_sheets(excel_path, pdfs_dir, selected_sheets)
        
        if sheet_pdfs.get("error"):
            result["error"] = sheet_pdfs["error"]
            return result
        
        if sheet_pdfs.get("success_count", 0) == 0:
            result["error"] = "No PDFs were generated successfully"
            return result
        
        result["sheets_processed"] = sheet_pdfs.get("success_count", 0)
        print(f"[PDF Regenerator] Generated {result['sheets_processed']} sheet PDFs", file=sys.stderr)
        
        # Step 2: Use AIReportGenerator for proper ordering and AI content
        excel_basename = os.path.splitext(os.path.basename(excel_path))[0]
        merged_pdf_filename = f"{excel_basename}_regenerated_{timestamp}.pdf"
        
        if output_pdf_path:
            final_pdf_path = output_pdf_path
        else:
            final_pdf_path = os.path.join(temp_dir, merged_pdf_filename)
        
        # Determine template name if not provided
        if not template_name:
            template_name = 'CC1'  # Default
            # Try to infer from filename
            filename_lower = excel_basename.lower()
            if 'term' in filename_lower or 'tl' in filename_lower:
                template_name = 'TL1'
            elif 'cc2' in filename_lower:
                template_name = 'CC2'
        
        # Use AIReportGenerator if available and grok_api_key is provided
        if AI_GENERATOR_AVAILABLE and grok_api_key:
            print(f"[PDF Regenerator] Step 2: Generating AI content and merging with proper order...", file=sys.stderr)
            
            try:
                # Prepare excel_data for AI generator
                excel_data = {
                    'json_data': json_data or {},
                    'html_data': html_data or {},
                    'template_name': template_name,
                    'timestamp': timestamp
                }
                
                # Initialize AI generator
                ai_generator = AIReportGenerator(grok_api_key, provider="grok", signature_path=signature_path)
                print("[PDF Regenerator] Using Grok AI for report generation", file=sys.stderr)
                
                # Generate full report with AI content and proper sheet ordering
                report_result = ai_generator.generate_full_report(
                    excel_pdfs_dir=pdfs_dir,
                    excel_data=excel_data,
                    output_path=final_pdf_path,
                    template_name=template_name
                )
                
                if report_result.get('success'):
                    result["ai_sections_generated"] = len(report_result.get('ai_sections_generated', []))
                    print(f"[PDF Regenerator] ✅ Full report generated with AI content", file=sys.stderr)
                    print(f"   AI Sections: {result['ai_sections_generated']}", file=sys.stderr)
                    print(f"   Excel PDFs: {len(report_result.get('excel_pdfs_included', []))}", file=sys.stderr)
                else:
                    # AI generation failed, fall back to simple merge with SHEET_ORDER
                    print(f"[PDF Regenerator] ⚠️ AI report generation failed, using simple merge with proper order", file=sys.stderr)
                    if not merge_pdfs_with_order(pdfs_dir, final_pdf_path, template_name, signature_path):
                        result["error"] = "Failed to merge PDF files"
                        return result
                    
            except Exception as ai_error:
                print(f"[PDF Regenerator] ⚠️ AI generation error: {str(ai_error)}, falling back to simple merge", file=sys.stderr)
                import traceback
                traceback.print_exc(file=sys.stderr)
                
                # Fall back to simple merge with proper order
                if not merge_pdfs_with_order(pdfs_dir, final_pdf_path, template_name, signature_path):
                    result["error"] = "Failed to merge PDF files"
                    return result
        else:
            # No AI key, use simple merge with proper sheet ordering
            print(f"[PDF Regenerator] Step 2: Merging PDFs with proper sheet order (no AI)...", file=sys.stderr)
            if not merge_pdfs_with_order(pdfs_dir, final_pdf_path, template_name, signature_path):
                result["error"] = "Failed to merge PDF files"
                return result
        
        # Read and encode the final PDF
        if os.path.exists(final_pdf_path):
            with open(final_pdf_path, 'rb') as f:
                pdf_data = f.read()
            
            result["success"] = True
            result["pdf_base64"] = base64.b64encode(pdf_data).decode('utf-8')
            result["pdf_filename"] = merged_pdf_filename
            result["pdf_size"] = len(pdf_data)
            
            print(f"\n✅ PDF regeneration complete!", file=sys.stderr)
            print(f"   Sheets processed: {result['sheets_processed']}", file=sys.stderr)
            print(f"   AI sections: {result['ai_sections_generated']}", file=sys.stderr)
            print(f"   PDF size: {result['pdf_size']:,} bytes", file=sys.stderr)
        else:
            result["error"] = "Final PDF file was not created"
        
        return result
        
    except Exception as e:
        print(f"❌ Error in PDF regeneration: {str(e)}", file=sys.stderr)
        import traceback
        traceback.print_exc(file=sys.stderr)
        result["error"] = str(e)
        return result
        
    finally:
        # Cleanup temp directory
        try:
            shutil.rmtree(temp_dir, ignore_errors=True)
        except:
            pass


if __name__ == '__main__':
    # Force UTF-8 encoding for stdout
    if sys.stdout.encoding != 'utf-8':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    
    # Expected arguments: 
    # pdf_regenerator.py <input_json>
    # Where input_json contains: excel_path, selected_sheets, grok_api_key, json_data, html_data, template_name
    if len(sys.argv) < 2:
        print(json.dumps({
            "success": False,
            "error": "Usage: pdf_regenerator.py <input_json>"
        }))
        sys.exit(1)
    
    try:
        input_data = json.loads(sys.argv[1])
    except json.JSONDecodeError as e:
        print(json.dumps({
            "success": False,
            "error": f"Invalid JSON input: {str(e)}"
        }))
        sys.exit(1)
    
    # Extract parameters from input
    excel_path = input_data.get('excel_path') or input_data.get('excelPath')
    selected_sheets = input_data.get('selected_sheets') or input_data.get('selectedSheets')
    grok_api_key = input_data.get('grok_api_key') or input_data.get('grokApiKey')
    json_data = input_data.get('json_data') or input_data.get('jsonData')
    html_data = input_data.get('html_data') or input_data.get('htmlData')
    template_name = input_data.get('template_name') or input_data.get('templateName')
    signature_path = input_data.get('signature_path') or input_data.get('signaturePath')
    excluded_sheets = input_data.get('excluded_sheets') or input_data.get('excludedSheets')
    
    if not excel_path:
        print(json.dumps({
            "success": False,
            "error": "excel_path is required"
        }))
        sys.exit(1)
    
    # Ensure selected_sheets is a list
    if isinstance(selected_sheets, str):
        try:
            selected_sheets = json.loads(selected_sheets)
        except:
            selected_sheets = None
    
    if selected_sheets and not isinstance(selected_sheets, list):
        selected_sheets = None
    
    result = regenerate_pdf_from_excel(
        excel_path=excel_path,
        selected_sheets=selected_sheets,
        grok_api_key=grok_api_key,
        json_data=json_data,
        html_data=html_data,
        template_name=template_name,
        signature_path=signature_path,
        excluded_sheets=excluded_sheets
    )
    
    print(json.dumps(result, ensure_ascii=False, default=str))
