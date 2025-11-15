"""
AI Resource PDF Parser
Extracts and indexes content from the two reference PDFs for RAG-based AI content generation.
Only these PDFs are allowed as knowledge sources - no external data permitted.
"""

import pdfplumber
import json
import re
from pathlib import Path
from typing import List, Dict, Any
import sys

class AIResourceParser:
    """Parse and index AI resource PDFs for knowledge retrieval."""
    
    def __init__(self):
        self.resources_dir = Path(__file__).parent.parent.parent / "templates" / "ai resource"
        self.pdf_files = {
            "ap_idp": "AP IDP 4.0(2).pdf",
            "pmegp": "PMEGP guidelines.pdf"
        }
        self.knowledge_base = {}
    
    def extract_text_from_pdf(self, pdf_path: str) -> List[Dict[str, Any]]:
        """Extract text content from PDF page by page."""
        print(f"\n📖 Extracting text from: {Path(pdf_path).name}", file=sys.stderr)
        
        pages_content = []
        
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            print(f"   Total pages: {total_pages}", file=sys.stderr)
            
            for page_num, page in enumerate(pdf.pages, 1):
                text = page.extract_text()
                if text:
                    # Clean up text
                    text = self._clean_text(text)
                    
                    pages_content.append({
                        "page": page_num,
                        "text": text,
                        "char_count": len(text)
                    })
                    
                    if page_num % 10 == 0:
                        print(f"   Processed {page_num}/{total_pages} pages...", file=sys.stderr)
        
        print(f"   ✅ Extracted {len(pages_content)} pages", file=sys.stderr)
        return pages_content
    
    def _clean_text(self, text: str) -> str:
        """Clean and normalize extracted text."""
        # Remove excessive whitespace
        text = re.sub(r'\s+', ' ', text)
        # Remove special characters but keep alphanumeric and basic punctuation
        text = re.sub(r'[^\w\s\.,;:\-\(\)\[\]\/\%\&\+\=\@\#\$\₹]', '', text)
        return text.strip()
    
    def chunk_text(self, pages_content: List[Dict[str, Any]], chunk_size: int = 1000, overlap: int = 200) -> List[Dict[str, Any]]:
        """Split text into overlapping chunks for better context retrieval."""
        chunks = []
        chunk_id = 0
        
        for page_data in pages_content:
            text = page_data["text"]
            page_num = page_data["page"]
            
            # Split into chunks
            for i in range(0, len(text), chunk_size - overlap):
                chunk_text = text[i:i + chunk_size]
                if len(chunk_text.strip()) > 50:  # Only keep substantial chunks
                    chunks.append({
                        "chunk_id": chunk_id,
                        "page": page_num,
                        "text": chunk_text,
                        "start_pos": i,
                        "end_pos": i + len(chunk_text)
                    })
                    chunk_id += 1
        
        return chunks
    
    def extract_sections(self, pages_content: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Extract structured sections from PDF content based on headings."""
        sections = []
        current_section = None
        
        for page_data in pages_content:
            lines = page_data["text"].split('.')
            
            for line in lines:
                line_stripped = line.strip()
                
                # Detect headings (various patterns)
                is_heading = (
                    len(line_stripped) < 100 and
                    (line_stripped.isupper() or
                     re.match(r'^\d+\.', line_stripped) or
                     re.match(r'^[A-Z][a-z]+:', line_stripped) or
                     re.match(r'^Chapter', line_stripped, re.IGNORECASE) or
                     re.match(r'^Section', line_stripped, re.IGNORECASE))
                )
                
                if is_heading and len(line_stripped) > 3:
                    # Save previous section
                    if current_section:
                        sections.append(current_section)
                    
                    # Start new section
                    current_section = {
                        "heading": line_stripped,
                        "page": page_data["page"],
                        "content": []
                    }
                elif current_section:
                    current_section["content"].append(line_stripped)
        
        # Add last section
        if current_section:
            sections.append(current_section)
        
        # Join content
        for section in sections:
            section["content"] = " ".join(section["content"]).strip()
        
        return sections
    
    def parse_all_resources(self) -> Dict[str, Any]:
        """Parse all AI resource PDFs and create knowledge base."""
        print("\n" + "="*80, file=sys.stderr)
        print("🤖 PARSING AI RESOURCE PDFs", file=sys.stderr)
        print("="*80, file=sys.stderr)
        
        self.knowledge_base = {
            "resources": {},
            "total_pages": 0,
            "total_chunks": 0,
            "metadata": {}
        }
        
        for resource_key, filename in self.pdf_files.items():
            pdf_path = self.resources_dir / filename
            
            if not pdf_path.exists():
                print(f"❌ ERROR: {filename} not found at {pdf_path}", file=sys.stderr)
                continue
            
            print(f"\n{'─'*80}", file=sys.stderr)
            print(f"Processing: {filename}", file=sys.stderr)
            print(f"{'─'*80}", file=sys.stderr)
            
            # Extract pages
            pages_content = self.extract_text_from_pdf(str(pdf_path))
            
            # Create chunks
            chunks = self.chunk_text(pages_content, chunk_size=1500, overlap=300)
            print(f"   Created {len(chunks)} text chunks", file=sys.stderr)
            
            # Extract sections
            sections = self.extract_sections(pages_content)
            print(f"   Identified {len(sections)} sections", file=sys.stderr)
            
            # Store in knowledge base
            self.knowledge_base["resources"][resource_key] = {
                "filename": filename,
                "pages": pages_content,
                "chunks": chunks,
                "sections": sections,
                "total_pages": len(pages_content),
                "total_chunks": len(chunks),
                "total_sections": len(sections)
            }
            
            self.knowledge_base["total_pages"] += len(pages_content)
            self.knowledge_base["total_chunks"] += len(chunks)
        
        # Add metadata
        self.knowledge_base["metadata"] = {
            "resource_files": list(self.pdf_files.values()),
            "parsing_complete": True,
            "strict_mode": True,  # Only these PDFs allowed, no external data
            "restriction": "AI must use ONLY these resource PDFs, no external sources"
        }
        
        print(f"\n{'='*80}", file=sys.stderr)
        print(f"✅ PARSING COMPLETE", file=sys.stderr)
        print(f"   Total Resources: {len(self.knowledge_base['resources'])}", file=sys.stderr)
        print(f"   Total Pages: {self.knowledge_base['total_pages']}", file=sys.stderr)
        print(f"   Total Chunks: {self.knowledge_base['total_chunks']}", file=sys.stderr)
        print(f"{'='*80}\n", file=sys.stderr)
        
        return self.knowledge_base
    
    def search_knowledge_base(self, query: str, top_k: int = 5) -> List[Dict[str, Any]]:
        """Simple keyword-based search in knowledge base."""
        if not self.knowledge_base or "resources" not in self.knowledge_base:
            print("⚠️  Knowledge base not loaded. Call parse_all_resources() first.", file=sys.stderr)
            return []
        
        query_lower = query.lower()
        query_terms = set(query_lower.split())
        
        results = []
        
        # Search through all chunks
        for resource_key, resource_data in self.knowledge_base["resources"].items():
            for chunk in resource_data["chunks"]:
                chunk_text_lower = chunk["text"].lower()
                
                # Calculate relevance score (simple keyword matching)
                matches = sum(1 for term in query_terms if term in chunk_text_lower)
                
                if matches > 0:
                    results.append({
                        "resource": resource_key,
                        "filename": resource_data["filename"],
                        "chunk_id": chunk["chunk_id"],
                        "page": chunk["page"],
                        "text": chunk["text"],
                        "relevance_score": matches,
                        "match_ratio": matches / len(query_terms) if query_terms else 0
                    })
        
        # Sort by relevance and return top_k
        results.sort(key=lambda x: (-x["relevance_score"], -x["match_ratio"]))
        return results[:top_k]
    
    def get_full_resource_text(self, resource_key: str = None) -> str:
        """Get complete text from one or all resources."""
        if not self.knowledge_base or "resources" not in self.knowledge_base:
            return ""
        
        if resource_key:
            if resource_key in self.knowledge_base["resources"]:
                resource = self.knowledge_base["resources"][resource_key]
                return "\n\n".join([page["text"] for page in resource["pages"]])
            return ""
        else:
            # Return all resources combined
            all_text = []
            for res_key, resource in self.knowledge_base["resources"].items():
                all_text.append(f"=== {resource['filename']} ===\n")
                all_text.append("\n\n".join([page["text"] for page in resource["pages"]]))
            return "\n\n".join(all_text)
    
    def save_knowledge_base(self, output_path: str = None):
        """Save knowledge base to JSON file."""
        if not output_path:
            output_path = Path(__file__).parent / "ai_knowledge_base.json"
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(self.knowledge_base, f, indent=2, ensure_ascii=False)
        
        print(f"💾 Knowledge base saved to: {output_path}", file=sys.stderr)
        return output_path
    
    def load_knowledge_base(self, input_path: str = None):
        """Load knowledge base from JSON file."""
        if not input_path:
            input_path = Path(__file__).parent / "ai_knowledge_base.json"
        
        if Path(input_path).exists():
            with open(input_path, 'r', encoding='utf-8') as f:
                self.knowledge_base = json.load(f)
            print(f"📚 Knowledge base loaded from: {input_path}", file=sys.stderr)
            return True
        else:
            print(f"❌ Knowledge base file not found: {input_path}", file=sys.stderr)
            return False


if __name__ == "__main__":
    # Initialize parser
    parser = AIResourceParser()
    
    # Parse all resource PDFs
    knowledge_base = parser.parse_all_resources()
    
    # Save to JSON
    output_file = parser.save_knowledge_base()
    
    # Test search
    print("\n" + "="*80, file=sys.stderr)
    print("🔍 TESTING KNOWLEDGE BASE SEARCH", file=sys.stderr)
    print("="*80 + "\n", file=sys.stderr)
    
    test_queries = [
        "manufacturing",
        "loan eligibility",
        "project cost",
        "financial assistance"
    ]
    
    for query in test_queries:
        print(f"\nQuery: '{query}'", file=sys.stderr)
        results = parser.search_knowledge_base(query, top_k=3)
        print(f"Found {len(results)} results:", file=sys.stderr)
        for i, result in enumerate(results, 1):
            print(f"\n  Result {i}:", file=sys.stderr)
            print(f"    Source: {result['filename']} (Page {result['page']})", file=sys.stderr)
            print(f"    Relevance: {result['relevance_score']} matches", file=sys.stderr)
            print(f"    Preview: {result['text'][:150]}...", file=sys.stderr)
    
    print(f"\n{'='*80}", file=sys.stderr)
    print("✅ All tests completed", file=sys.stderr)
    print(f"{'='*80}\n", file=sys.stderr)
