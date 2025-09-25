# Enhanced Contract Financial Analyzer

A professional desktop application for analyzing contract documents with AI-powered three-pass validation and comprehensive payment tracking spreadsheets.

## Features

- **Three-Pass Validation System**: Combines regex pattern matching, AI extraction, and cross-validation for maximum accuracy
- **Professional Excel Output**: Creates comprehensive tracking spreadsheets with proper number formatting
- **CPA-Level Analysis**: Designed for professional accounting and financial management
- **Confidence Scoring**: Each extracted data point includes confidence levels (HIGH/MEDIUM/LOW)
- **Comprehensive Tracking**: Payment schedules, expense tracking, compliance checklists, and budget monitoring

## Requirements

### For Running the Source Code
- Python 3.8 or higher
- Required packages:
  ```bash
  pip install anthropic openpyxl pandas PyPDF2 python-docx tkinter
  ```

### For Using the Built App
- Windows 10/11, macOS 10.14+, or Linux
- Internet connection for AI analysis
- Claude API key from console.anthropic.com

## Installation

### Option 1: Run from Source
1. Clone or download this repository
2. Install dependencies: `pip install anthropic openpyxl pandas PyPDF2 python-docx`
3. Run the GUI: `python contract_analyzer_gui.py`

### Option 2: Build Standalone App
1. Install PyInstaller: `pip install pyinstaller`
2. Run the build script: `python build_app.py`
3. Find your app in the `dist` folder

### Option 3: Command Line Interface
For batch processing or automation:
```bash
python enhanced_contract_analyzer.py contract.pdf -o output.xlsx -k YOUR_API_KEY
```

## Usage

### GUI Application
1. Launch the application (double-click the executable or run `python contract_analyzer_gui.py`)
2. Enter your Claude API key (get one at console.anthropic.com)
3. Select your contract file (supports PDF, Word, or text files)
4. Choose where to save the Excel analysis
5. Click "Analyze Contract" and wait for results

### API Key Setup
- Get your API key from [console.anthropic.com](https://console.anthropic.com)
- The app will save your key locally for future use
- Keys are stored in `~/.contract_analyzer/config.txt`

## Output

The application generates a comprehensive Excel file with multiple worksheets:

1. **Contract Summary**: Key contract information with confidence scores
2. **Payment Tracking**: Invoice submission and payment timeline
3. **Expense Tracking**: Reimbursable expense management
4. **Compliance Checklist**: Required documentation and deadlines
5. **Budget Monitor**: Hour and budget utilization tracking
6. **Validation Details**: Three-pass validation methodology results
7. **Full Analysis**: Complete AI analysis text

## File Support

- **PDF**: Contracts in PDF format (requires PyPDF2)
- **Word**: .docx and .doc files (requires python-docx)
- **Text**: Plain text files (.txt)

## Three-Pass Validation System

1. **Pass 1**: High-precision regex pattern extraction for financial data
2. **Pass 2**: Direct Claude AI extraction with focused prompts
3. **Pass 3**: Cross-validation and confidence scoring

This system ensures maximum accuracy for critical financial information like contract values, hourly rates, dates, and payment terms.

## Professional Features

### For CPAs and Financial Professionals
- Proper number formatting for Excel calculations
- Confidence scoring suitable for audit trails
- Comprehensive compliance tracking
- Professional documentation standards
- Budget monitoring with automated warnings

### Contract Types Supported
- Time & Materials contracts
- Fixed payment contracts
- Milestone-based contracts
- Mixed payment structures

## Building the Application

To create a standalone executable:

```bash
# Install build tools
pip install pyinstaller

# Run the build script
python build_app.py
```

This creates a single executable file that includes all dependencies.

## Project Structure

```
contract-analyzer/
├── contract_analyzer_gui.py       # Main GUI application
├── enhanced_contract_analyzer.py  # Command-line version
├── build_app.py                   # Build script for standalone app
├── README.md                      # This file
└── dist/                          # Built application (after building)
    ├── Contract Analyzer          # Standalone executable
    ├── README.txt                 # User instructions
    └── version.txt               # Build information
```

## Troubleshooting

### Common Issues
- **Import Errors**: Install required packages with pip
- **PDF Reading Issues**: Install PyPDF2 (`pip install PyPDF2`)
- **Word Document Issues**: Install python-docx (`pip install python-docx`)
- **API Errors**: Check your Claude API key and internet connection

### Performance
- Analysis typically takes 30-60 seconds depending on contract length
- Larger contracts (50+ pages) may take longer
- The application uses threading to maintain responsiveness

## Security

- API keys are stored locally in your home directory
- No contract data is permanently stored by the application
- All AI processing happens through Anthropic's secure API

## License

This project is for professional use. Please ensure compliance with your organization's software policies.

## Support

For technical issues:
1. Check that all dependencies are installed correctly
2. Verify your Claude API key is valid
3. Ensure you have internet connectivity
4. Review the console output for specific error messages

## Version History

- **v1.0**: Initial release with three-pass validation system
- Professional CPA-level analysis and reporting
- Comprehensive Excel output with multiple tracking sheets

## Development

Built with:
- Python 3.8+
- Tkinter for GUI
- Anthropic Claude API for AI analysis
- OpenPyXL for Excel generation
- PyInstaller for standalone app creation