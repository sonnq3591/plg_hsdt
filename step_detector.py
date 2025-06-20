#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF Step Counter - Extract text with PyMuPDF and detect 21 vs 23 steps with OpenAI
Separate module - doesn't touch the core replacement function
"""

import os
import fitz  # PyMuPDF
import openai
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

class StepDetector:
    def __init__(self):
        """Initialize the step detector with OpenAI API"""
        self.openai_api_key = os.getenv('OPENAI_API_KEY')
        if not self.openai_api_key:
            raise ValueError("❌ OPENAI_API_KEY not found in .env file!")
        
        openai.api_key = self.openai_api_key
        print("✅ StepDetector initialized")

    def extract_pdf_text_pymupdf(self, pdf_path):
        """
        Extract text from PDF using PyMuPDF (fitz)
        More reliable than PyPDF2 for Vietnamese text
        """
        try:
            print(f"📖 Extracting text from: {pdf_path}")
            
            # Open PDF
            doc = fitz.open(pdf_path)
            text = ""
            page_count = len(doc)  # Get page count before processing
            
            # Extract text from all pages
            for page_num in range(page_count):
                page = doc.load_page(page_num)
                page_text = page.get_text()
                text += f"\n--- PAGE {page_num + 1} ---\n"
                text += page_text
            
            doc.close()
            
            print(f"✅ Extracted {len(text)} characters from {page_count} pages")
            return text
            
        except Exception as e:
            print(f"❌ Error extracting PDF text: {str(e)}")
            return None

    def detect_steps_with_openai(self, pdf_text):
        """
        Ask OpenAI to analyze the PDF content and determine if it's 21 or 23 steps
        Looking for section 3.2 or similar content about implementation steps
        """
        
        # Create a focused prompt based on the images you showed
        prompt = f"""
Analyze this Vietnamese document content from CHUONG_V.pdf and find the section about implementation steps.

Look for sections with these characteristics:
1. Title containing "quy trình chỉnh lý" (implementation process) 
2. Content mentioning "Tuân thủ theo các bước thực hiện chỉnh lý" (follow implementation steps)
3. Reference to "trình tự 21 bước công việc" (21-step process) OR "trình tự 23 bước công việc" (23-step process)
4. A table with "Số TT" (sequential number) and "Nội dung công việc" (work content) columns
5. Section 3.2 or similar numbering about "Yêu cầu về quy trình chỉnh lý" or "Công việc thực hiện của mỗi bước"

Your task:
- Find the section that describes the step-by-step implementation process
- Count if it mentions 21 steps or 23 steps in the process
- Look for phrases like "21 bước công việc", "23 bước công việc", or count the actual steps in any process table

Return ONLY the number: "21" or "23"
If you cannot determine clearly, return "UNKNOWN"

DOCUMENT CONTENT:
{pdf_text}

STEP COUNT:"""

        try:
            print("🤖 Asking OpenAI to analyze step count...")
            
            response = openai.ChatCompletion.create(
                model='gpt-4o',
                messages=[
                    {
                        "role": "system", 
                        "content": "You are an expert at analyzing Vietnamese documents about administrative processes. You specialize in finding implementation step counts in document sections. Be precise and only return the step count number."
                    },
                    {
                        "role": "user", 
                        "content": prompt
                    }
                ],
                max_tokens=50,
                temperature=0.0  # No creativity, just accuracy
            )
            
            step_count = response.choices[0].message.content.strip()
            print(f"🤖 OpenAI response: '{step_count}'")
            
            # Validate response
            if step_count in ["21", "23"]:
                return int(step_count)
            elif step_count == "UNKNOWN":
                print("⚠️ OpenAI couldn't determine step count")
                return None
            else:
                print(f"⚠️ Unexpected response: {step_count}")
                return None
                
        except Exception as e:
            print(f"❌ OpenAI API Error: {str(e)}")
            return None

    def test_step_detection(self, pdf_path):
        """
        Test the complete step detection process
        """
        print("\n🧪 TESTING: Step Detection Process")
        print("=" * 50)
        
        # Check if PDF exists
        if not os.path.exists(pdf_path):
            print(f"❌ PDF file not found: {pdf_path}")
            return None
        
        # Step 1: Extract PDF text
        print("📖 Step 1: Extracting PDF text with PyMuPDF...")
        pdf_text = self.extract_pdf_text_pymupdf(pdf_path)
        
        if not pdf_text:
            print("❌ Failed to extract PDF text")
            return None
        
        # Show a preview of extracted text
        preview = pdf_text[:500]
        print(f"📄 Text preview: {preview}...")
        
        # Step 2: Detect steps with OpenAI
        print("\n🤖 Step 2: Analyzing with OpenAI...")
        step_count = self.detect_steps_with_openai(pdf_text)
        
        if step_count:
            print(f"✅ SUCCESS: Detected {step_count} steps!")
            return step_count
        else:
            print("❌ Failed to detect step count")
            return None

def main():
    """Test the step detection functionality"""
    print("🔍 PDF Step Counter - PyMuPDF + OpenAI")
    print("=" * 50)
    
    try:
        # Initialize detector
        detector = StepDetector()
        
        # Test with CHUONG_V.pdf
        pdf_file = "pdf_inputs/CHUONG_V_stp.pdf"  # Adjust path as needed
        
        result = detector.test_step_detection(pdf_file)
        
        if result:
            print(f"\n🎯 FINAL RESULT: {result} steps detected")
            print(f"📄 This means we should use: {result}_BUOC.docx")
        else:
            print(f"\n❌ FAILED: Could not detect step count")
            
    except ValueError as e:
        print(e)
        print("💡 Please create a .env file with: OPENAI_API_KEY=your_key_here")
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()