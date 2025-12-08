
import sys
import os
from PyPDF2 import PdfReader

def analyze_pdf(pdf_path):
    if not os.path.exists(pdf_path):
        print(f"File not found: {pdf_path}")
        return

    try:
        reader = PdfReader(pdf_path)
        print(f"Total Pages: {len(reader.pages)}")
        
        for i, page in enumerate(reader.pages):
            text = page.extract_text()
            print(f"\n--- Page {i+1} ---")
            # Print first few lines to identify the section
            lines = text.split('\n')
            for line in lines[:10]:
                if line.strip():
                    print(line.strip())
            
            # Check for specific keywords to identify sections
            content_lower = text.lower()
            if "swot" in content_lower:
                print("[Contains SWOT Analysis]")
            if "balance sheet" in content_lower:
                print("[Contains Balance Sheet]")
            if "profit & loss" in content_lower or "profit and loss" in content_lower:
                print("[Contains P&L]")
            if "ratio" in content_lower:
                print("[Contains Ratio Analysis]")
            if "term loan" in content_lower:
                print("[Contains Term Loan]")
                
    except Exception as e:
        print(f"Error reading PDF: {e}")

if __name__ == "__main__":
    pdf_path = r"d:\github-desktop\UI\DEV\CA-DEV\CA-BACKEND-DEV\templates\excel\refferanceTermLoan.pdf"
    analyze_pdf(pdf_path)
