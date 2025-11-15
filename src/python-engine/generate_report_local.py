"""
Local Report Generation Test
============================
This script allows you to test the full AI-enhanced report generation locally
without going through the Node.js API.

Usage:
    python generate_report_local.py

Requirements:
    - Valid GEMINI_API_KEY in environment or hardcoded below
    - Excel template at ../../templates/excel/frcc1.xlsx
    - AI resource PDFs in ../../templates/ai resource/
"""

import sys
import json
import os
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv

# Load environment variables from .env file
dotenv_path = os.path.join(os.path.dirname(__file__), '..', '..', '.env')
load_dotenv(dotenv_path)

# Add current directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from excel_calculator import calculate_excel

def generate_local_report(gemini_api_key=None, use_sample_data=True):
    """
    Generate a full report locally with sample or custom data
    
    Args:
        gemini_api_key: Your Gemini API key (or set GEMINI_API_KEY env var)
        use_sample_data: If True, uses minimal sample data; if False, prompts for custom data
    """
    print("\n" + "="*80)
    print("LOCAL AI-ENHANCED REPORT GENERATION TEST")
    print("="*80 + "\n")
    
    # Get API key
    if not gemini_api_key:
        gemini_api_key = os.getenv('GEMINI_API_KEY')
    
    if not gemini_api_key or gemini_api_key == 'your_gemini_api_key_here':
        print("⚠️  WARNING: No valid Gemini API key found!")
        print("\nOptions:")
        print("1. Set GEMINI_API_KEY environment variable")
        print("2. Edit this script and add your API key to the gemini_api_key parameter")
        print("3. Enter it now (or press Enter to skip AI generation)")
        
        user_key = input("\nEnter Gemini API key (or press Enter to skip): ").strip()
        if user_key:
            gemini_api_key = user_key
        else:
            print("\n⚠️  Continuing without AI generation (Excel PDFs only)...\n")
            gemini_api_key = None
    
    # Prepare sample data
    if use_sample_data:
        print("📝 Using sample form data...\n")
        form_data = {
            "projectName": "Local Test Manufacturing Unit",
            "applicantName": "Test Applicant",
            "projectCost": "5000000",
            "loanAmount": "3500000",
            "ownContribution": "1500000",
            "category": "Manufacturing",
            "sector": "General",
            "numberOfEmployees": "15",
            # Add minimal required fields
            "state": "Andhra Pradesh",
            "district": "Visakhapatnam"
        }
    else:
        print("📝 Enter custom form data (or press Enter for defaults)...\n")
        form_data = {}
        fields = [
            ("projectName", "Local Test Project"),
            ("applicantName", "Test Applicant"),
            ("projectCost", "5000000"),
            ("loanAmount", "3500000"),
            ("category", "Manufacturing")
        ]
        
        for field, default in fields:
            value = input(f"{field} [{default}]: ").strip()
            form_data[field] = value if value else default
    
    # Build input data
    input_data = {
        "templateId": "frcc1",
        "formData": form_data,
        "generateFullReport": True,  # Enable AI report generation
        "includeAllSheets": True,    # Generate PDFs for all sheets
    }
    
    # Add API key if available
    if gemini_api_key:
        input_data["geminiApiKey"] = gemini_api_key
        print("✅ Gemini API key configured\n")
    else:
        print("⚠️  No API key - will generate Excel PDFs only\n")
    
    # Get Excel template path (relative to this script's location)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(script_dir, "..", "..", "templates", "excel", f"{input_data['templateId']}.xlsx")
    template_path = os.path.normpath(template_path)
    
    if not os.path.exists(template_path):
        print(f"❌ ERROR: Template not found at {template_path}")
        return {"success": False, "error": "Template not found"}
    
    print("="*80)
    print("STARTING REPORT GENERATION")
    print("="*80)
    print(f"\n⏱️  Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"📊 Template: {input_data['templateId']}")
    print(f"📂 Template Path: {template_path}")
    print(f"🤖 AI Generation: {'ENABLED' if gemini_api_key else 'DISABLED'}")
    print(f"📄 All Sheets: {'YES' if input_data.get('includeAllSheets') else 'NO'}")
    print("\nThis may take 30-60 seconds...\n")
    print("-"*80 + "\n")
    
    try:
        # Call the main calculation function with correct signature
        result_json = calculate_excel(input_data, template_path)
        result = json.loads(result_json)
        
        print("\n" + "-"*80)
        print("\n✅ REPORT GENERATION COMPLETED!\n")
        print("="*80)
        
        # Display results
        if result.get('success'):
            print("📊 GENERATED FILES:\n")
            
            # Excel file
            if result.get('fileName'):
                excel_path = os.path.join('../../temp', result['fileName'])
                excel_abs = os.path.abspath(excel_path)
                print(f"   📗 Excel: {result['fileName']}")
                print(f"      Path: {excel_abs}")
                if os.path.exists(excel_abs):
                    size_kb = os.path.getsize(excel_abs) / 1024
                    print(f"      Size: {size_kb:.1f} KB")
            
            # PDF report (if full report was generated)
            if result.get('fullReportPdf'):
                pdf_filename = result.get('fullReportFilename', 'consolidated_report.pdf')
                pdf_path = os.path.join('../../temp', pdf_filename)
                pdf_abs = os.path.abspath(pdf_path)
                print(f"\n   📕 Full Report PDF: {pdf_filename}")
                print(f"      Path: {pdf_abs}")
                if os.path.exists(pdf_abs):
                    size_kb = os.path.getsize(pdf_abs) / 1024
                    print(f"      Size: {size_kb:.1f} KB")
                    print(f"\n   🎉 AI-ENHANCED REPORT READY!")
            
            # Individual sheet PDF (if only single sheet generated)
            elif result.get('pdfData'):
                pdf_filename = result.get('pdfFileName', 'sheet_report.pdf')
                pdf_path = os.path.join('../../temp', pdf_filename)
                pdf_abs = os.path.abspath(pdf_path)
                print(f"\n   📄 Sheet PDF: {pdf_filename}")
                print(f"      Path: {pdf_abs}")
                if os.path.exists(pdf_abs):
                    size_kb = os.path.getsize(pdf_abs) / 1024
                    print(f"      Size: {size_kb:.1f} KB")
            
            # HTML preview
            if result.get('excelData'):
                print(f"\n   🌐 HTML Preview: Available ({len(result['excelData'])} characters)")
            
            # JSON data
            if result.get('jsonData'):
                print(f"   📋 JSON Data: {len(result['jsonData'])} cells extracted")
            
            print("\n" + "="*80)
            print(f"⏱️  Completed at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            print("="*80)
            
            # Save result to file for inspection
            script_dir = os.path.dirname(os.path.abspath(__file__))
            temp_dir = os.path.join(script_dir, "..", "..", "temp")
            os.makedirs(temp_dir, exist_ok=True)
            
            result_file = os.path.join(temp_dir, f"local_test_result_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
            
            # Don't save base64 data (too large)
            result_summary = {
                "success": result.get('success'),
                "fileName": result.get('fileName'),
                "fullReportFilename": result.get('fullReportFilename'),
                "pdfFileName": result.get('pdfFileName'),
                "excelDataLength": len(result.get('excelData', '')),
                "jsonDataLength": len(result.get('jsonData', [])),
                "hasFullReportPdf": bool(result.get('fullReportPdf')),
                "hasPdfData": bool(result.get('pdfData')),
                "timestamp": datetime.now().isoformat()
            }
            
            with open(result_file, 'w') as f:
                json.dump(result_summary, f, indent=2)
            
            print(f"\n💾 Result summary saved to:")
            print(f"   {result_file}\n")
            
            return result
            
        else:
            print(f"❌ GENERATION FAILED!")
            print(f"\nError: {result.get('error', 'Unknown error')}")
            if result.get('details'):
                print(f"Details: {result.get('details')}")
            return result
            
    except Exception as e:
        print(f"\n❌ ERROR OCCURRED!")
        print(f"\nException: {str(e)}")
        import traceback
        print("\nFull traceback:")
        print(traceback.format_exc())
        return {"success": False, "error": str(e)}


if __name__ == "__main__":
    print("\n")
    print("╔" + "="*78 + "╗")
    print("║" + " "*78 + "║")
    print("║" + "   AI-ENHANCED REPORT GENERATION - LOCAL TEST RUNNER".center(78) + "║")
    print("║" + " "*78 + "║")
    print("╚" + "="*78 + "╝")
    
    # Check for command line arguments
    import argparse
    parser = argparse.ArgumentParser(description='Generate AI-enhanced reports locally')
    parser.add_argument('--api-key', help='Gemini API key')
    parser.add_argument('--custom-data', action='store_true', help='Enter custom form data')
    parser.add_argument('--no-ai', action='store_true', help='Skip AI generation (Excel PDFs only)')
    
    args = parser.parse_args()
    
    # Override API key if --no-ai flag is set
    api_key = None if args.no_ai else args.api_key
    
    # Run generation
    result = generate_local_report(
        gemini_api_key=api_key,
        use_sample_data=not args.custom_data
    )
    
    # Exit code based on success
    sys.exit(0 if result.get('success') else 1)
