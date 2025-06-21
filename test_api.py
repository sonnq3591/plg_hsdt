#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Test Walking Skeleton API
Quick test script to verify the API works before Teams integration
"""

import requests
import time
from pathlib import Path

def test_api_health():
    """Test API health endpoint"""
    try:
        response = requests.get("http://localhost:8000/api/health")
        if response.status_code == 200:
            print("✅ API Health: OK")
            print(f"   Response: {response.json()}")
            return True
        else:
            print(f"❌ API Health: Failed ({response.status_code})")
            return False
    except Exception as e:
        print(f"❌ API Health: Connection failed - {e}")
        return False

def test_api_root():
    """Test API root endpoint"""
    try:
        response = requests.get("http://localhost:8000/")
        if response.status_code == 200:
            print("✅ API Root: OK")
            data = response.json()
            print(f"   Template: {data.get('template')}")
            print(f"   Placeholders: {data.get('placeholders')}")
            return True
        else:
            print(f"❌ API Root: Failed ({response.status_code})")
            return False
    except Exception as e:
        print(f"❌ API Root: Connection failed - {e}")
        return False

def test_api_templates():
    """Test templates endpoint"""
    try:
        response = requests.get("http://localhost:8000/api/templates")
        if response.status_code == 200:
            print("✅ API Templates: OK")
            data = response.json()
            skeleton = data.get('walking_skeleton', {})
            print(f"   Status: {skeleton.get('status')}")
            print(f"   Placeholders: {len(skeleton.get('placeholders', []))}")
            return True
        else:
            print(f"❌ API Templates: Failed ({response.status_code})")
            return False
    except Exception as e:
        print(f"❌ API Templates: Connection failed - {e}")
        return False

def test_document_processing():
    """Test the main document processing endpoint"""
    print("\n🚀 Testing document processing...")
    
    # Check if PDF files exist
    pdf_files = {
        'tbmt_pdf': 'pdf_inputs/TBMT.pdf',
        'bmmt_pdf': 'pdf_inputs/BMMT.pdf',
        'chuong_iii_pdf': 'pdf_inputs/CHUONG_III.pdf',
        'chuong_v_pdf': 'pdf_inputs/CHUONG_V.pdf',
        'hsmt_pdf': 'pdf_inputs/HSMT.pdf'
    }
    
    missing_files = []
    files_to_upload = {}
    
    for param_name, file_path in pdf_files.items():
        if Path(file_path).exists():
            files_to_upload[param_name] = open(file_path, 'rb')
            print(f"✅ Found: {file_path}")
        else:
            missing_files.append(file_path)
            print(f"❌ Missing: {file_path}")
    
    if missing_files:
        print(f"\n❌ Cannot test processing - missing {len(missing_files)} files:")
        for f in missing_files:
            print(f"   - {f}")
        return False
    
    try:
        print("\n📤 Uploading files and processing...")
        start_time = time.time()
        
        response = requests.post(
            "http://localhost:8000/api/process-document",
            files=files_to_upload
        )
        
        processing_time = time.time() - start_time
        
        if response.status_code == 200:
            # Save the result
            output_file = "api_test_result.docx"
            with open(output_file, 'wb') as f:
                f.write(response.content)
            
            file_size = len(response.content)
            print(f"✅ Document Processing: SUCCESS")
            print(f"   Processing time: {processing_time:.1f} seconds")
            print(f"   Output file: {output_file}")
            print(f"   File size: {file_size:,} bytes")
            print(f"   Content-Type: {response.headers.get('content-type')}")
            return True
        else:
            print(f"❌ Document Processing: Failed ({response.status_code})")
            print(f"   Error: {response.text}")
            return False
            
    except Exception as e:
        print(f"❌ Document Processing: Exception - {e}")
        return False
    
    finally:
        # Close file handles
        for f in files_to_upload.values():
            f.close()

def main():
    """Run all API tests"""
    print("🧪 WALKING SKELETON API TESTS")
    print("=" * 50)
    print("🎯 Testing: http://localhost:8000")
    print("📋 Make sure the API server is running!")
    print()
    
    tests = [
        ("API Health Check", test_api_health),
        ("API Root Endpoint", test_api_root),
        ("API Templates Endpoint", test_api_templates),
        ("Document Processing", test_document_processing)
    ]
    
    results = []
    
    for test_name, test_func in tests:
        print(f"🔍 {test_name}...")
        result = test_func()
        results.append((test_name, result))
        print()
    
    # Summary
    print("=" * 50)
    print("📊 TEST RESULTS SUMMARY:")
    
    passed = 0
    for test_name, result in results:
        status = "✅ PASS" if result else "❌ FAIL"
        print(f"   {status}: {test_name}")
        if result:
            passed += 1
    
    print(f"\n🎯 Results: {passed}/{len(results)} tests passed")
    
    if passed == len(results):
        print("🎉 ALL TESTS PASSED! API is ready for Teams integration!")
    else:
        print("❌ Some tests failed. Check the API server and file paths.")
    
    return passed == len(results)

if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)