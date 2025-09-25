#!/usr/bin/env python3
"""
Build Script for Contract Analyzer Desktop App
Creates a standalone executable from the GUI application
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path

def build_app():
    """Build the Contract Analyzer into a standalone app."""
    
    print("Building Contract Analyzer Desktop App...")
    print("="*50)
    
    # Check if PyInstaller is installed
    try:
        import PyInstaller
        print("✓ PyInstaller found")
    except ImportError:
        print("Installing PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("✓ PyInstaller installed")
    
    # Check if the GUI file exists
    gui_file = "contract_analyzer_gui.py"
    if not Path(gui_file).exists():
        print(f"❌ Error: {gui_file} not found")
        print("Make sure the GUI file is in the current directory")
        return False
    
    print(f"✓ Found {gui_file}")
    
    # Clean previous builds
    if Path("build").exists():
        shutil.rmtree("build")
        print("✓ Cleaned previous build directory")
    
    if Path("dist").exists():
        shutil.rmtree("dist")
        print("✓ Cleaned previous dist directory")
    
    # Build command for different platforms
    build_cmd = [
        "pyinstaller",
        "--onefile",                    # Single executable
        "--windowed",                   # No console window (GUI only)
        "--name", "Contract Analyzer",  # App name
        # Add icon if available
        # "--icon", "icon.ico",         # Uncomment and add icon file
        gui_file
    ]
    
    # Add README if it exists
    if Path("README.md").exists():
        build_cmd.insert(-1, "--add-data")
        build_cmd.insert(-1, "README.md:.")
    
    print("Building application...")
    try:
        result = subprocess.run(build_cmd, check=True, capture_output=True, text=True)
        print("✓ Build completed successfully")
        
        # Find the executable
        if sys.platform == "darwin":  # macOS
            app_path = Path("dist/Contract Analyzer")
        elif sys.platform == "win32":  # Windows
            app_path = Path("dist/Contract Analyzer.exe")
        else:  # Linux
            app_path = Path("dist/Contract Analyzer")
        
        if app_path.exists():
            print(f"✓ Executable created: {app_path}")
            print(f"✓ File size: {app_path.stat().st_size / 1024 / 1024:.1f} MB")
            
            # Test the executable
            print("\nTesting executable...")
            try:
                if sys.platform == "win32":
                    subprocess.Popen([str(app_path)], creationflags=subprocess.CREATE_NO_WINDOW)
                else:
                    subprocess.Popen([str(app_path)])
                print("✓ App launched successfully (close it to continue)")
            except Exception as e:
                print(f"⚠ Could not test launch: {e}")
            
            return True
        else:
            print("❌ Executable not found after build")
            return False
            
    except subprocess.CalledProcessError as e:
        print(f"❌ Build failed: {e}")
        print("Error output:", e.stderr)
        return False

def create_installer_info():
    """Create additional files for distribution."""
    
    print("\nCreating distribution files...")
    
    # Create README for users
    readme_content = """# Contract Analyzer Desktop App

## Quick Start
1. Double-click "Contract Analyzer" to launch
2. Enter your Claude API key (get one at console.anthropic.com)
3. Select your contract file (PDF, Word, or Text)
4. Choose where to save the Excel analysis
5. Click "Analyze Contract"

## System Requirements
- Windows 10/11, macOS 10.14+, or Linux
- Internet connection for AI analysis
- Excel or compatible spreadsheet software (optional)

## Features
- Three-pass financial data validation
- Comprehensive payment tracking spreadsheets
- Professional CPA-level analysis
- Confidence scoring for all extracted data
- Compliance and budget monitoring

## Support
For issues or questions, contact your system administrator.

## API Key
You'll need a Claude API key from Anthropic:
1. Visit console.anthropic.com
2. Sign up/sign in
3. Go to API Keys section
4. Create a new key
5. Copy and paste into the app
"""
    
    with open("dist/README.txt", "w") as f:
        f.write(readme_content)
    print("✓ Created README.txt")
    
    # Create version info
    version_info = f"""Contract Analyzer Desktop App
Version: 1.0
Build Date: {os.popen('date').read().strip()}
Platform: {sys.platform}

This application provides professional contract financial analysis
with AI-powered three-pass validation for maximum accuracy.
"""
    
    with open("dist/version.txt", "w") as f:
        f.write(version_info)
    print("✓ Created version.txt")

def main():
    """Main build function."""
    print("Contract Analyzer App Builder")
    print("This will create a standalone desktop application")
    print()
    
    # Check Python version
    if sys.version_info < (3, 8):
        print("❌ Python 3.8 or higher required")
        return
    
    print(f"✓ Python {sys.version_info.major}.{sys.version_info.minor} detected")
    
    # Build the app
    if build_app():
        create_installer_info()
        
        print("\n" + "="*50)
        print("BUILD COMPLETE!")
        print("="*50)
        print("Your app is ready in the 'dist' folder:")
        
        if sys.platform == "darwin":
            print("• Contract Analyzer (macOS executable)")
        elif sys.platform == "win32":
            print("• Contract Analyzer.exe (Windows executable)")
        else:
            print("• Contract Analyzer (Linux executable)")
        
        print("• README.txt (user instructions)")
        print("• version.txt (build information)")
        
        print(f"\nTo distribute: Share the entire 'dist' folder")
        print("Users can run the app by double-clicking the executable")
        
    else:
        print("\n❌ Build failed. Check errors above.")

if __name__ == "__main__":
    main()