#!/usr/bin/env python3
"""
Enhanced Contract Financial Analyzer - Command Line Version
Now uses shared core library for better architecture and maintainability
"""

import argparse
import os
import sys
from pathlib import Path
from contract_analyzer_core import HighPrecisionContractAnalyzer

class CLIContractAnalyzer:
    """CLI-specific wrapper around shared core functionality"""

    def __init__(self, api_key: str):
        """Initialize CLI wrapper with shared core analyzer"""
        self.analyzer = HighPrecisionContractAnalyzer(api_key)

    def analyze_with_cli_progress(self, contract_text: str) -> dict:
        """Analyze contract with CLI-specific progress indicators"""

        print("Enhanced Contract Financial Analyzer - CLI Version")
        print("=" * 60)

        # Step 1: Three-pass validation with progress
        print("\n🔍 Starting Three-Pass Financial Data Validation...")
        print("  → Running Pass 1: High-precision regex pattern extraction...")

        try:
            # Use shared core three-pass validation
            validated_data = self.analyzer.cross_validate_financial_data(contract_text)

            if validated_data.get('error'):
                print(f"  ❌ Validation failed: {validated_data['error']}")
                return {'success': False, 'error': validated_data['error']}

            print("  ✓ Pass 1: Regex extraction complete")
            print("  ✓ Pass 2: Claude AI extraction complete")
            print("  ✓ Pass 3: Cross-validation complete")

            # Display validation results
            print(f"\n📊 Validation Results:")
            print(f"  • Total Contract Value: {validated_data.get('total_contract_value', 'Not found')}")
            print(f"  • Hourly Rate: {validated_data.get('hourly_rate', 'Not found')}")
            print(f"  • Start Date: {validated_data.get('start_date', 'Not found')}")
            print(f"  • End Date: {validated_data.get('end_date', 'Not found')}")
            print(f"  • Payment Terms: {validated_data.get('payment_terms', 'Not found')}")

            # Step 2: Comprehensive analysis
            print(f"\n📋 Performing Comprehensive Contract Analysis...")
            comprehensive_result = self.analyzer.analyze_contract_comprehensive(contract_text)

            if not comprehensive_result.get('success'):
                print(f"  ❌ Comprehensive analysis failed: {comprehensive_result.get('analysis')}")
                return {'success': False, 'error': comprehensive_result.get('analysis')}

            print("  ✓ CPA-level financial analysis complete")

            # Step 3: Extract structured data
            print(f"  → Extracting comprehensive tracking data...")
            contract_info, payment_schedule, tracking_requirements = \
                self.analyzer.extract_comprehensive_data(comprehensive_result["analysis"])

            print("  ✓ Structured data extraction complete")

            # Override with high-precision financial data
            if validated_data.get('total_contract_value'):
                contract_info['total_value'] = validated_data['total_contract_value']
            if validated_data.get('hourly_rate'):
                contract_info['hourly_rate'] = validated_data['hourly_rate']
            if validated_data.get('start_date'):
                contract_info['start_date'] = validated_data['start_date']
            if validated_data.get('end_date'):
                contract_info['end_date'] = validated_data['end_date']
            if validated_data.get('payment_terms'):
                contract_info['payment_terms'] = validated_data['payment_terms']

            return {
                'success': True,
                'validated_data': validated_data,
                'comprehensive_analysis': comprehensive_result["analysis"],
                'contract_info': contract_info,
                'payment_schedule': payment_schedule,
                'tracking_requirements': tracking_requirements
            }

        except Exception as e:
            print(f"  ❌ Analysis failed: {str(e)}")
            return {'success': False, 'error': str(e)}

    def display_analysis_summary(self, results: dict):
        """Display analysis summary in CLI format"""

        print(f"\n📈 Analysis Summary:")
        print("=" * 40)

        contract_info = results.get('contract_info', {})
        payment_schedule = results.get('payment_schedule', [])
        tracking_requirements = results.get('tracking_requirements', {})

        # Contract basics
        print(f"Client: {contract_info.get('client', 'Not specified')}")
        print(f"Vendor: {contract_info.get('vendor', 'Not specified')}")
        print(f"Contract Type: {contract_info.get('contract_type', 'Not specified')}")
        print(f"Total Value: {contract_info.get('total_value', 'Not specified')}")
        print(f"Duration: {contract_info.get('start_date', 'N/A')} to {contract_info.get('end_date', 'N/A')}")

        # Payment info
        print(f"\n💰 Payment Information:")
        print(f"Payment Terms: {contract_info.get('payment_terms', 'Not specified')}")
        print(f"Invoice Frequency: {contract_info.get('invoice_frequency', 'Not specified')}")

        if payment_schedule:
            print(f"Payment Schedule: {len(payment_schedule)} payments scheduled")

        # Tracking requirements
        if tracking_requirements:
            expense_tracking = tracking_requirements.get('expense_tracking', {})
            compliance = tracking_requirements.get('compliance', {})

            print(f"\n📋 Tracking Requirements:")
            if expense_tracking.get('travel_expenses') == 'yes':
                print(f"• Travel expense tracking required")
            if compliance.get('w9_required') == 'yes':
                print(f"• W-9 form required")
            if compliance.get('time_breakdown_required') == 'yes':
                print(f"• Time breakdown required in invoices")

    def run_analysis(self, contract_file: str, output_file: str) -> bool:
        """Main CLI analysis workflow"""

        try:
            # Read contract file
            print(f"📄 Reading contract file: {contract_file}")
            contract_text = self.analyzer.read_contract_file(contract_file)
            print(f"  ✓ Successfully read {len(contract_text)} characters")

            # Perform analysis
            results = self.analyze_with_cli_progress(contract_text)

            if not results.get('success'):
                print(f"\n❌ Analysis failed: {results.get('error')}")
                return False

            # Display summary
            self.display_analysis_summary(results)

            # Generate Excel file
            print(f"\n📊 Generating Excel spreadsheet...")
            self.analyzer.create_comprehensive_spreadsheet(
                results['contract_info'],
                results['payment_schedule'],
                results['tracking_requirements'],
                results['validated_data'],
                results['comprehensive_analysis'],
                output_file
            )

            print(f"  ✓ Excel file created: {output_file}")
            print(f"\n🎉 Analysis Complete!")
            print(f"Results saved to: {Path(output_file).absolute()}")

            # Show what's in the Excel file
            print(f"\n📋 Excel file contains 7 worksheets:")
            print(f"  1. Contract Summary - Key contract info with confidence scores")
            print(f"  2. Payment Tracking - Invoice and payment timeline")
            print(f"  3. Expense Tracking - Reimbursable expenses management")
            print(f"  4. Compliance Checklist - Required documentation and deadlines")
            print(f"  5. Budget Monitor - Hour and budget utilization tracking")
            print(f"  6. Validation Details - Three-pass validation methodology")
            print(f"  7. Full Analysis - Complete AI analysis text")

            return True

        except FileNotFoundError as e:
            print(f"❌ File not found: {e}")
            return False
        except Exception as e:
            print(f"❌ Unexpected error: {e}")
            return False

def load_api_key() -> str:
    """Load API key from various sources"""

    # Check environment variable first
    api_key = os.getenv('CLAUDE_API_KEY')
    if api_key:
        return api_key

    # Check config file
    config_dir = Path.home() / ".contract_analyzer"
    config_file = config_dir / "config.txt"

    if config_file.exists():
        try:
            with open(config_file, 'r') as f:
                api_key = f.read().strip()
                if api_key:
                    return api_key
        except Exception:
            pass

    return None

def save_api_key(api_key: str):
    """Save API key to config file"""
    config_dir = Path.home() / ".contract_analyzer"
    config_dir.mkdir(exist_ok=True)
    config_file = config_dir / "config.txt"

    try:
        with open(config_file, 'w') as f:
            f.write(api_key)
        print(f"  ✓ API key saved to {config_file}")
    except Exception as e:
        print(f"  ⚠ Could not save API key: {e}")

def main():
    """Main CLI entry point"""

    parser = argparse.ArgumentParser(
        description='Enhanced Contract Financial Analyzer - Professional CPA-level analysis',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python enhanced_contract_analyzer.py contract.pdf -o analysis.xlsx
  python enhanced_contract_analyzer.py contract.docx -o results.xlsx -k your_api_key_here

Environment Variables:
  CLAUDE_API_KEY    Claude API key (alternative to -k flag)
        """
    )

    parser.add_argument('contract_file',
                       help='Path to contract file (PDF, Word .docx, or text file)')
    parser.add_argument('-o', '--output', required=True,
                       help='Output Excel file path (.xlsx)')
    parser.add_argument('-k', '--api-key',
                       help='Claude API key (or set CLAUDE_API_KEY environment variable)')
    parser.add_argument('--save-key', action='store_true',
                       help='Save API key to config file for future use')
    parser.add_argument('-v', '--verbose', action='store_true',
                       help='Enable verbose output')

    args = parser.parse_args()

    # Validate input file
    if not Path(args.contract_file).exists():
        print(f"❌ Contract file not found: {args.contract_file}")
        sys.exit(1)

    # Check file extension
    allowed_extensions = {'.pdf', '.docx', '.doc', '.txt'}
    file_ext = Path(args.contract_file).suffix.lower()
    if file_ext not in allowed_extensions:
        print(f"❌ Unsupported file type: {file_ext}")
        print(f"Supported formats: {', '.join(allowed_extensions)}")
        sys.exit(1)

    # Get API key
    api_key = args.api_key or load_api_key()

    if not api_key:
        print("❌ Claude API key required!")
        print("Options:")
        print("  1. Use -k flag: python enhanced_contract_analyzer.py contract.pdf -o output.xlsx -k YOUR_KEY")
        print("  2. Set environment variable: export CLAUDE_API_KEY=YOUR_KEY")
        print("  3. Get your API key at: https://console.anthropic.com")
        sys.exit(1)

    # Save API key if requested
    if args.save_key and args.api_key:
        save_api_key(args.api_key)

    # Validate output file extension
    if not args.output.endswith('.xlsx'):
        print("❌ Output file must have .xlsx extension")
        sys.exit(1)

    # Create output directory if needed
    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        # Initialize CLI analyzer
        cli_analyzer = CLIContractAnalyzer(api_key)

        # Run analysis
        success = cli_analyzer.run_analysis(args.contract_file, args.output)

        if success:
            print(f"\n✅ Contract analysis completed successfully!")
            sys.exit(0)
        else:
            print(f"\n❌ Contract analysis failed!")
            sys.exit(1)

    except KeyboardInterrupt:
        print(f"\n⚠ Analysis interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()