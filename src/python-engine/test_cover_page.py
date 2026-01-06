
import os
import sys
import win32com.client
from pathlib import Path

# Config
TEMP_DIR = r"d:\CA-DEV\CA-BACKEND-DEV\temp"
# Pick the file that seems to correspond to the screenshot (frcc3)
TARGET_FILE = "frcc3_existing_2026-01-06T16-10-52-869Z-updated-20260106T161053Z-534cd83d.xlsx"
FULL_PATH = os.path.join(TEMP_DIR, TARGET_FILE)
OUTPUT_PDF_NATIVE = os.path.join(TEMP_DIR, "debug_cover_native.pdf")
OUTPUT_PDF_CLEARED = os.path.join(TEMP_DIR, "debug_cover_cleared.pdf")

def debug_cover_page():
    print(f"--- Cover Page Debugger ---")
    print(f"Target File: {FULL_PATH}")
    
    if not os.path.exists(FULL_PATH):
        print(f"❌ File not found: {FULL_PATH}")
        return

    excel = None
    wb = None
    
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        print("Opening workbook...")
        wb = excel.Workbooks.Open(FULL_PATH)
        
        # Find cover page
        cover_sheet = None
        for sheet in wb.Sheets:
            print(f"Scanning sheet: '{sheet.Name}'")
            norm_name = sheet.Name.lower().strip()
            if any(x in norm_name for x in ['cover page', 'coverpage', 'cover']):
                cover_sheet = sheet
                print(f"✅ Found Cover Page: '{sheet.Name}'")
                break
        
        if not cover_sheet:
            print("❌ Could not find a sheet named 'Cover Page', 'Cover', etc.")
            return

        # 1. INSPECT NATIVE PROPERTIES
        ps = cover_sheet.PageSetup
        print("\n--- Native PageSetup Properties ---")
        print(f"PrintArea: '{ps.PrintArea}'")
        print(f"Orientation: {ps.Orientation} (1=Portrait, 2=Landscape)")
        print(f"Zoom: {ps.Zoom}")
        print(f"FitToPagesWide: {ps.FitToPagesWide}")
        print(f"FitToPagesTall: {ps.FitToPagesTall}")
        print(f"PaperSize: {ps.PaperSize} (9=A4)")
        print(f"LeftMargin: {ps.LeftMargin}")
        print(f"RightMargin: {ps.RightMargin}")
        print(f"TopMargin: {ps.TopMargin}")
        
        # 2. EXPORT NATIVE
        print(f"\nAttempting Native Export to: {OUTPUT_PDF_NATIVE}")
        # Ensure we are NOT ignoring print areas (matching the fixed logic)
        cover_sheet.ExportAsFixedFormat(
            Type=0, # xlTypePDF
            Filename=OUTPUT_PDF_NATIVE,
            Quality=0,
            IncludeDocProperties=True,
            IgnorePrintAreas=False, 
            OpenAfterPublish=False
        )
        print("✅ Native PDF Exported.")

        # 3. EXPORT WITH CLEARED PRINT AREA (Test hypothesis)
        print(f"\nAttempting Export with CLEARED PrintArea to: {OUTPUT_PDF_CLEARED}")
        original_print_area = ps.PrintArea
        ps.PrintArea = ""
        
        cover_sheet.ExportAsFixedFormat(
            Type=0, 
            Filename=OUTPUT_PDF_CLEARED,
            Quality=0,
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False
        )
        print("✅ Cleared PrintArea PDF Exported.")
        
        # Restore just in case
        ps.PrintArea = original_print_area

        # 4. EXPORT WITH A4 + FIT TO PAGE (The "Hybrid" Fix)
        OUTPUT_PDF_FIXED = os.path.join(TEMP_DIR, "debug_cover_fixed.pdf")
        print(f"\nAttempting Export with A4 + FIT TO PAGE to: {OUTPUT_PDF_FIXED}")
        
        # Apply the proposed "Correct" settings
        ps.PrintArea = original_print_area # Keep the print area! important for borders.
        ps.PaperSize = 9 # xlPaperA4
        ps.Zoom = False # Turn off zoom
        ps.FitToPagesWide = 1
        ps.FitToPagesTall = 1
        
        cover_sheet.ExportAsFixedFormat(
            Type=0, 
            Filename=OUTPUT_PDF_FIXED,
            Quality=0,
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False
        )
        print("✅ Fixed PDF Exported.")


    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        
    finally:
        if wb:
            wb.Close(SaveChanges=False)
        if excel:
            excel.Quit()
            # Clean up COM references to ensure process terminates
            del wb
            del excel
        print("\nDone.")

if __name__ == "__main__":
    debug_cover_page()
