#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Setup Walking Skeleton Environment
Helps you prepare everything needed for the walking skeleton test
"""

import os
from pathlib import Path
import sys

def check_python_version():
    """Check Python version"""
    version = sys.version_info
    print(f"🐍 Python version: {version.major}.{version.minor}.{version.micro}")
    
    if version.major < 3 or (version.major == 3 and version.minor < 8):
        print("❌ Python 3.8+ required!")
        return False
    
    print("✅ Python version OK")
    return True

def check_env_file():
    """Check .env file exists with OpenAI API key"""
    env_file = Path(".env")
    
    if not env_file.exists():
        print("❌ .env file not found!")
        print("📝 Creating sample .env file...")
        
        sample_content = """# OpenAI API Configuration
OPENAI_API_KEY=your_openai_api_key_here

# Instructions:
# 1. Get your API key from: https://platform.openai.com/api-keys
# 2. Replace 'your_openai_api_key_here' with your actual API key
# 3. Save this file as .env in your project root
"""
        env_file.write_text(sample_content)
        print(f"📄 Created sample .env file at: {env_file.absolute()}")
        print("⚠️ Please edit .env and add your OpenAI API key!")
        return False
    
    # Check if API key is set
    content = env_file.read_text()
    if "your_openai_api_key_here" in content:
        print("❌ Please update .env with your actual OpenAI API key!")
        return False
    
    if "OPENAI_API_KEY=" not in content:
        print("❌ OPENAI_API_KEY not found in .env file!")
        return False
    
    print("✅ .env file exists with API key")
    return True

def check_required_python_files():
    """Check if all required Python files exist"""
    required_files = [
        "processor.py",
        "processor_pham_vi.py", 
        "processor_can_cu.py",
        "processor_muc_dich.py",
        "combined_processor.py"
    ]
    
    print("🔍 Checking Python processors...")
    missing = []
    
    for file_name in required_files:
        if Path(file_name).exists():
            print(f"✅ Found: {file_name}")
        else:
            print(f"❌ Missing: {file_name}")
            missing.append(file_name)
    
    if missing:
        print(f"❌ Missing {len(missing)} processor files!")
        return False
    
    print("✅ All processor files found")
    return True

def check_required_template_files():
    """Check if required template files exist"""
    required_files = [
        "02_MUC_DO_HIEU_BIET_template.docx",
        "21_BUOC.docx",
        "23_BUOC.docx"
    ]
    
    print("📄 Checking template files...")
    missing = []
    
    for file_name in required_files:
        if Path(file_name).exists():
            print(f"✅ Found: {file_name}")
        else:
            print(f"❌ Missing: {file_name}")
            missing.append(file_name)
    
    if missing:
        print(f"❌ Missing {len(missing)} template files!")
        print("📋 Please ensure these files are in your project root:")
        for f in missing:
            print(f"   - {f}")
        return False
    
    print("✅ All template files found")
    return True

def setup_pdf_folder():
    """Setup PDF input folder structure"""
    pdf_folder = Path("pdf_inputs")
    
    if not pdf_folder.exists():
        pdf_folder.mkdir()
        print(f"📁 Created folder: {pdf_folder}")
    else:
        print(f"✅ Folder exists: {pdf_folder}")
    
    # Check for PDF files
    required_pdfs = [
        "TBMT.pdf",
        "BMMT.pdf", 
        "CHUONG_III.pdf",
        "CHUONG_V.pdf",
        "HSMT.pdf"
    ]
    
    print("📎 Checking PDF files...")
    missing_pdfs = []
    
    for pdf_name in required_pdfs:
        pdf_path = pdf_folder / pdf_name
        if pdf_path.exists():
            size = pdf_path.stat().st_size
            print(f"✅ Found: {pdf_name} ({size:,} bytes)")
        else:
            print(f"❌ Missing: {pdf_name}")
            missing_pdfs.append(pdf_name)
    
    if missing_pdfs:
        print(f"\n📋 Please place these PDF files in {pdf_folder}:")
        for pdf in missing_pdfs:
            print(f"   - {pdf}")
        return False
    
    print("✅ All PDF files found")
    return True

def install_requirements():
    """Check if requirements can be installed"""
    print("📦 Checking Python packages...")
    
    required_packages = [
        "fastapi",
        "uvicorn", 
        "python-multipart",
        "openai",
        "python-docx",
        "PyPDF2",
        "PyMuPDF",
        "python-dotenv"
    ]
    
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package.replace("-", "_"))
            print(f"✅ {package}")
        except ImportError:
            print(f"❌ {package}")
            missing_packages.append(package)
    
    if missing_packages:
        print(f"\n📦 Install missing packages with:")
        print(f"pip install {' '.join(missing_packages)}")
        return False
    
    print("✅ All packages installed")
    return True

def show_walking_skeleton_plan():
    """Show the walking skeleton plan"""
    print("\n🚀 WALKING SKELETON PLAN")
    print("=" * 50)
    print("🎯 GOAL: Get ONE template working end-to-end")
    print("📋 Template: 02_MUC_DO_HIEU_BIET.docx")
    print("🔢 Placeholders: 5 total")
    print("\n📊 WORKFLOW:")
    print("1. PDF Input → 5 processors → DOCX Output")
    print("2. Test locally → FastAPI → Teams Bot")
    print("3. Add remaining 14 templates incrementally")
    print("\n🔄 NEXT STEPS:")
    print("1. Run: python test_walking_skeleton.py")
    print("2. If successful → Run: python walking_skeleton_api.py")
    print("3. Test API → Create Teams bot")

def main():
    """Main setup function"""
    print("🏗️ WALKING SKELETON SETUP")
    print("=" * 50)
    
    checks = [
        ("Python Version", check_python_version),
        ("Environment File", check_env_file),
        ("Python Processors", check_required_python_files),
        ("Template Files", check_required_template_files),
        ("PDF Input Folder", setup_pdf_folder),
        ("Python Packages", install_requirements)
    ]
    
    all_passed = True
    
    for check_name, check_func in checks:
        print(f"\n🔍 {check_name}:")
        if not check_func():
            all_passed = False
    
    print("\n" + "=" * 50)
    
    if all_passed:
        print("🎉 SETUP COMPLETE!")
        print("✅ All requirements satisfied")
        print("\n🚀 Ready to run walking skeleton:")
        print("   python test_walking_skeleton.py")
    else:
        print("❌ SETUP INCOMPLETE!")
        print("🛠️ Please fix the issues above before proceeding")
    
    show_walking_skeleton_plan()
    
    return all_passed

if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)