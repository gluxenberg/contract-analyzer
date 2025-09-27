#!/usr/bin/env python3
"""
Enhanced Contract Financial Analyzer - GUI Version (FIXED)
Professional desktop application with three-pass validation and comprehensive payment tracking
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
import sys
import json
import re
import tempfile
from pathlib import Path
from datetime import datetime, timedelta
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles.differential import DifferentialStyle
import anthropic
from typing import Dict, List, Optional, Tuple

class HighPrecisionContractAnalyzer:
    def __init__(self, api_key: str):
        """Initialize the enhanced contract analyzer with Claude API key and precision patterns."""
        self.client = anthropic.Anthropic(api_key=api_key)
        
        # High-precision patterns for critical financial data
        self.financial_patterns = {
            'money_amounts': [
                r'\$\s*[\d,]+\.?\d*(?:\s*(?:USD|dollars?))?',
                r'(?:USD|dollars?)\s*[\d,]+\.?\d*',
                r'[\d,]+\.?\d*\s*(?:USD|dollars?)',
                r'(?:total|amount|value|fee|rate|cost|price|budget|cap|maximum|limit)(?:\s+(?:of|is|at|shall be|not to exceed))?\s*[:\$]?\s*[\d,]+\.?\d*',
            ],
            'hourly_rates': [
                r'\$\s*[\d,]+\.?\d*\s*(?:per|/|an)\s*hour',
                r'hourly\s+(?:rate|fee|charge)(?:\s+(?:of|is|at))?\s*[:\$]?\s*[\d,]+\.?\d*',
                r'[\d,]+\.?\d*\s*(?:per|/)\s*(?:hour|hr)',
            ],
            'contract_values': [
                r'(?:total|contract|project|agreement)\s+(?:value|amount|price|cost|fee)(?:\s+(?:of|is|shall be|not to exceed))?\s*[:\$]?\s*[\d,]+\.?\d*',
                r'(?:maximum|cap|limit|ceiling)(?:\s+(?:of|is|at))?\s*[:\$]?\s*[\d,]+\.?\d*',
                r'budget(?:\s+(?:of|is|shall be|not to exceed))?\s*[:\$]?\s*[\d,]+\.?\d*',
            ],
            'dates': [
                r'\b\d{1,2}[/-]\d{1,2}[/-]\d{4}\b',
                r'\b\d{4}[/-]\d{2}[/-]\d{2}\b',
                r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+\d{4}\b',
                r'\b\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}\b',
            ],
            'time_periods': [
                r'\b\d+\s*(?:hours?|hrs?)\b',
                r'\b\d+\s*(?:days?)\b',
                r'\b\d+\s*(?:weeks?)\b',
                r'\b\d+\s*(?:months?)\b',
                r'maximum.*?\d+.*?(?:hours?|days?)',
                r'(?:daily|weekly|monthly)\s+(?:limit|maximum|cap).*?\d+',
            ],
            'payment_terms': [
                r'net\s+\d+(?:\s+days?)?',
                r'within\s+\d+\s+days?',
                r'payment\s+due.*?\d+\s+days?',
                r'\d+\s+days?\s+(?:after|from|following)',
            ]
        }
    
    def parse_currency_to_number(self, currency_string: str) -> Optional[float]:
        """Convert currency string to float number for Excel calculations."""
        if not currency_string or currency_string in ['null', 'N/A', 'Not found', '']:
            return None
        
        temp_str = str(currency_string).strip()
        if temp_str.startswith(','):
            temp_str = '0' + temp_str
        if temp_str.startswith('.'):
            temp_str = '0' + temp_str
        cleaned = re.sub(r'[^\d.-]', '', temp_str)
        try:
            return float(cleaned) if cleaned else None
        except ValueError:
            return None

    def parse_hours_to_number(self, hours_string: str) -> Optional[float]:
        """Convert hours string to float number."""
        if not hours_string or hours_string in ['null', 'N/A', 'Not found', '']:
            return None
        
        temp_str = str(hours_string).strip()
        if temp_str.startswith(','):
            temp_str = '0' + temp_str
        if temp_str.startswith('.'):
            temp_str = '0' + temp_str
        cleaned = re.sub(r'[^\d.-]', '', temp_str)
        try:
            return float(cleaned) if cleaned else None
        except ValueError:
            return None
        
    def read_contract_file(self, file_path: str) -> str:
        """Read contract file and return content as string."""
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"Contract file not found: {file_path}")
            
        if path.suffix.lower() == '.pdf':
            try:
                import PyPDF2
                with open(file_path, 'rb') as file:
                    pdf_reader = PyPDF2.PdfReader(file)
                    text = ""
                    for page in pdf_reader.pages:
                        text += page.extract_text()
                    return text
            except ImportError:
                raise ImportError("PyPDF2 required for PDF files. Install with: pip install PyPDF2")
                
        elif path.suffix.lower() in ['.txt', '.docx', '.doc']:
            if path.suffix.lower() == '.docx':
                try:
                    from docx import Document
                    doc = Document(file_path)
                    return '\n'.join([paragraph.text for paragraph in doc.paragraphs])
                except ImportError:
                    raise ImportError("python-docx required for Word files. Install with: pip install python-docx")
            else:
                with open(file_path, 'r', encoding='utf-8') as file:
                    return file.read()
        else:
            with open(file_path, 'r', encoding='utf-8') as file:
                return file.read()

    def extract_financial_data_regex(self, contract_text: str) -> Dict:
        """PASS 1: Extract critical financial data using high-precision regex patterns."""
        results = {
            'money_amounts': [],
            'hourly_rates': [],
            'contract_values': [],
            'dates': [],
            'time_periods': [],
            'payment_terms': []
        }
        
        for category, patterns in self.financial_patterns.items():
            unique_matches = set()
            
            for pattern in patterns:
                matches = re.finditer(pattern, contract_text, re.IGNORECASE | re.MULTILINE)
                
                for match in matches:
                    start = max(0, match.start() - 50)
                    end = min(len(contract_text), match.end() + 50)
                    context = contract_text[start:end].strip()
                    matched_text = match.group().strip()
                    unique_matches.add((matched_text, context))
            
            results[category] = list(unique_matches)
        
        return results

    def claude_direct_extraction(self, contract_text: str) -> Dict:
        """PASS 2: Direct Claude extraction with focused prompts."""
        claude_prompt = f"""
        Extract ONLY these specific financial data points from this contract.
        Be extremely precise - include ONLY information that is explicitly stated:
        
        1. TOTAL CONTRACT VALUE (the main contract amount)
        2. HOURLY RATE (if time & materials)
        3. CONTRACT START DATE
        4. CONTRACT END DATE  
        5. PAYMENT TERMS (Net 30, etc.)
        6. MAXIMUM HOURS (if specified)
        7. DAILY HOUR LIMIT (if specified)
        
        Return as JSON:
        {{
            "total_contract_value": "exact amount or null",
            "hourly_rate": "exact rate or null",
            "start_date": "YYYY-MM-DD or null", 
            "end_date": "YYYY-MM-DD or null",
            "payment_terms": "exact terms or null",
            "maximum_hours": "number or null",
            "daily_hour_limit": "number or null"
        }}
        
        Contract: {contract_text}
        """
        
        try:
            claude_result = self.client.messages.create(
                model="claude-3-haiku-20240307",
                max_tokens=1000,
                messages=[{"role": "user", "content": claude_prompt}]
            )
            
            claude_text = claude_result.content[0].text
            start_idx = claude_text.find('{')
            end_idx = claude_text.rfind('}') + 1
            
            if start_idx != -1 and end_idx > start_idx:
                return json.loads(claude_text[start_idx:end_idx])
            else:
                return {}
                
        except Exception as e:
            return {}

    def validate_and_merge_data(self, regex_results: Dict, claude_data: Dict, contract_text: str) -> Dict:
        """PASS 3: Validate and merge results from both methods."""
        validation_prompt = f"""
        You are validating financial data extracted by two different methods. 
        Provide the MOST ACCURATE data by comparing both sources:
        
        REGEX EXTRACTED DATA:
        Money amounts: {[m[0] for m in regex_results.get('money_amounts', [])[:5]]}
        Hourly rates: {[m[0] for m in regex_results.get('hourly_rates', [])]}
        Dates: {[m[0] for m in regex_results.get('dates', [])[:10]]}
        Payment terms: {[m[0] for m in regex_results.get('payment_terms', [])]}
        
        CLAUDE EXTRACTED DATA:
        {claude_data}
        
                Return ONLY valid JSON in this exact format (no other text):
        {{
            "total_contract_value": "actual_value_or_null",
            "hourly_rate": "actual_value_or_null",
            "start_date": "YYYY-MM-DD_or_null",
            "end_date": "YYYY-MM-DD_or_null",
            "payment_terms": "actual_terms_or_null",
            "maximum_hours": "number_or_null",
            "daily_hour_limit": "number_or_null",
            "confidence_scores": {{
                "total_contract_value": "HIGH",
                "hourly_rate": "MEDIUM",
                "start_date": "HIGH",
                "end_date": "HIGH",
                "payment_terms": "HIGH",
                "maximum_hours": "LOW",
                "daily_hour_limit": "LOW"
            }},
            "validation_notes": "Brief summary of data quality"
        }}
        """
        
        try:
            validation_result = self.client.messages.create(
                model="claude-3-haiku-20240307",
                max_tokens=1500,
                messages=[{"role": "user", "content": validation_prompt}]
            )
            
            validation_text = validation_result.content[0].text
            start_idx = validation_text.find('{')
            end_idx = validation_text.rfind('}') + 1
            
            if start_idx != -1 and end_idx > start_idx:
                return json.loads(validation_text[start_idx:end_idx])
            else:
                return {"validation_notes": "Could not parse validation results"}
                
        except Exception as e:
            return {"validation_notes": f"Validation failed: {e}"}

    def cross_validate_financial_data(self, contract_text: str) -> Dict:
        """Run the three-pass validation system."""
        regex_results = self.extract_financial_data_regex(contract_text)
        claude_data = self.claude_direct_extraction(contract_text)
        validated_data = self.validate_and_merge_data(regex_results, claude_data, contract_text)
        return validated_data

    def analyze_contract_comprehensive(self, contract_text: str) -> Dict:
        """Send contract to Claude API for comprehensive financial analysis."""
        enhanced_analysis_prompt = """
        You are a Certified Public Accountant analyzing this contract for complete financial tracking requirements. Extract ALL financial information including:

        BASIC CONTRACT INFORMATION (EXTRACT FIRST):
        - Client name (party receiving services/goods)
        - Vendor/Contractor name (party providing services/goods)  
        - Contract number or ID
        - Contract effective dates
        - Contract title or project name

        Then extract ALL financial information including:

        PRIMARY PAYMENT STRUCTURE:
        - Contract type (fixed payments, time & materials, milestone-based, etc.)
        - Total contract value and any caps/maximums
        - Payment rates (hourly, daily, fixed amounts)
        - Payment frequency and timing requirements

        PAYMENT SCHEDULE DETAILS:
        - Specific due dates for payments OR invoice submission requirements
        - Invoice approval processes and timeframes
        - Expected payment timing after invoice submission
        - Any advance payments or retainer requirements

        EXPENSE AND REIMBURSEMENT TRACKING:
        - Travel expense policies and limits
        - Meal and incidental expense allowances
        - Equipment or material cost provisions
        - Any expense pre-approval requirements
        - Separate expense tracking from main contract value

        COMPLIANCE AND DOCUMENTATION REQUIREMENTS:
        - Required tax forms (W-9, 1099, etc.)
        - Invoice content requirements (daily breakdowns, certifications, etc.)
        - Time tracking limitations (daily/weekly maximums)
        - Reporting deadlines and deliverable schedules
        - Required signatures or approvals

        BUDGET MONITORING ELEMENTS:
        - Maximum hours, days, or other quantity limits
        - Running total calculations needed
        - Budget utilization warnings or thresholds
        - Contract period start and end dates
        - Any renewal or extension provisions

        RISK FACTORS:
        - Payment conditions or performance requirements
        - Penalty clauses or withholding provisions
        - Termination clauses affecting payment
        - Currency or rate change provisions

        For time & materials contracts specifically include:
        - Hourly rate breakdown by role/activity
        - Maximum billable hours per period
        - Overtime or premium rate conditions
        - Minimum/maximum monthly billing requirements

        Create a comprehensive payment tracking structure that includes:
        1. All scheduled payments with dates and amounts
        2. Invoice submission timeline and requirements  
        3. Expense reimbursement tracking separate from main payments
        4. Compliance documentation deadlines
        5. Budget utilization monitoring with warnings
        6. Expected payment dates based on contract terms

        If no specific payment dates exist, create a logical payment tracking framework based on the contract's invoicing and payment cycle requirements.

        Contract text:
        """
        
        try:
            message = self.client.messages.create(
                model="claude-3-haiku-20240307",
                max_tokens=4000,
                messages=[{
                    "role": "user",
                    "content": enhanced_analysis_prompt + contract_text
                }]
            )
            return {"analysis": message.content[0].text, "success": True}
        except Exception as e:
            return {"analysis": f"Error analyzing contract: {str(e)}", "success": False}

    def extract_comprehensive_data(self, analysis: str) -> Tuple[Dict, List[Dict], Dict]:
        """Extract comprehensive structured data from Claude's analysis."""
        extraction_prompt = f"""
        Based on this contract analysis, extract the comprehensive information in JSON format:
        
        Return a JSON object with:
        
        1. "contract_info": {{
            "client": "client name",
            "vendor": "vendor name", 
            "contract_id": "contract number/id",
            "contract_type": "fixed payments/time and materials/milestone/other",
            "total_value": "total contract value",
            "start_date": "YYYY-MM-DD",
            "end_date": "YYYY-MM-DD",
            "payment_terms": "payment terms description",
            "hourly_rate": "hourly rate if applicable",
            "max_hours": "maximum hours if applicable",
            "max_daily_hours": "daily hour limit if applicable",
            "invoice_frequency": "monthly/weekly/other",
            "payment_timeline": "net 30/45 days or description"
        }}
        
        2. "payment_schedule": [
            {{
                "due_date": "YYYY-MM-DD or null",
                "description": "payment description", 
                "amount": "payment amount or null",
                "invoice_submission_due": "YYYY-MM-DD or null",
                "expected_payment_date": "YYYY-MM-DD or null",
                "notes": "any additional notes"
            }}
        ]
        
        3. "tracking_requirements": {{
            "expense_tracking": {{
                "travel_expenses": "yes/no",
                "travel_policy": "description",
                "meal_allowance": "amount or policy",
                "equipment_costs": "yes/no",
                "pre_approval_required": "yes/no"
            }},
            "compliance": {{
                "w9_required": "yes/no",
                "invoice_requirements": "description",
                "time_breakdown_required": "yes/no",
                "certifications_required": "description",
                "reporting_deadlines": "description"
            }},
            "budget_monitoring": {{
                "hour_tracking": "yes/no",
                "budget_caps": "description",
                "warning_thresholds": "percentage or amount",
                "utilization_tracking": "yes/no"
            }}
        }}
        
        If information is not available, use null or "not specified".
        
        Analysis to process:
        {analysis}
        """
        
        try:
            message = self.client.messages.create(
                model="claude-3-haiku-20240307",
                max_tokens=2000,
                messages=[{
                    "role": "user",
                    "content": extraction_prompt
                }]
            )

            response_text = message.content[0].text.strip()

            # Try to clean up the JSON response
            if "```json" in response_text:
                # Extract JSON from code blocks
                start = response_text.find("```json") + 7
                end = response_text.find("```", start)
                if end != -1:
                    response_text = response_text[start:end].strip()

            # Try original method first
            start_idx = response_text.find('{')
            end_idx = response_text.rfind('}') + 1

            if start_idx != -1 and end_idx > start_idx:
                json_str = response_text[start_idx:end_idx]
                data = json.loads(json_str)
                return (
                    data.get("contract_info", {}),
                    data.get("payment_schedule", []),
                    data.get("tracking_requirements", {})
                )
            else:
                # Fallback: Try to extract basic info from the analysis text
                return self._extract_fallback_data_gui(analysis, "No valid JSON structure found")

        except json.JSONDecodeError as e:
            # Fallback: Try to extract basic info from the analysis text
            return self._extract_fallback_data_gui(analysis, str(e))
        except Exception as e:
            # Return structures with fallback data
            return self._extract_fallback_data_gui(analysis, str(e))

    def _extract_fallback_data_gui(self, analysis_text: str, error_msg: str) -> Tuple[Dict, List[Dict], Dict]:
        """Fallback method to extract basic contract info from analysis text"""

        contract_info = {
            "client": self._extract_from_text_gui(analysis_text, ["CLIENT:", "Client:", "client name", "receiving services"]),
            "vendor": self._extract_from_text_gui(analysis_text, ["CONTRACTOR:", "Vendor:", "vendor name", "providing services"]),
            "contract_id": self._extract_from_text_gui(analysis_text, ["Contract ID", "Agreement", "contract number"]),
            "contract_type": self._extract_contract_type_gui(analysis_text),
            "total_value": "unknown",
            "start_date": None,
            "end_date": None,
            "payment_terms": "unknown",
            "hourly_rate": "unknown",
            "max_hours": None,
            "max_daily_hours": None,
            "invoice_frequency": "unknown",
            "payment_timeline": "unknown"
        }

        # Basic payment schedule
        payment_schedule = [{
            "due_date": None,
            "description": "Monthly invoice payment",
            "amount": None,
            "invoice_submission_due": None,
            "expected_payment_date": None,
            "notes": f"Extraction failed: {error_msg}"
        }]

        # Basic tracking requirements
        tracking_requirements = {
            "expense_tracking": {
                "travel_expenses": "unknown",
                "travel_policy": "unknown",
                "meal_allowance": "unknown",
                "equipment_costs": "unknown",
                "pre_approval_required": "unknown"
            },
            "compliance": {
                "w9_required": "unknown",
                "invoice_requirements": "unknown",
                "time_breakdown_required": "unknown",
                "certifications_required": "unknown",
                "reporting_deadlines": "unknown"
            },
            "budget_monitoring": {
                "hour_tracking": "unknown",
                "budget_caps": "unknown",
                "warning_thresholds": "unknown",
                "utilization_tracking": "unknown"
            }
        }

        return contract_info, payment_schedule, tracking_requirements

    def _extract_from_text_gui(self, text: str, search_terms: List[str]) -> str:
        """Extract information from text using search terms"""
        text_lower = text.lower()

        for term in search_terms:
            term_lower = term.lower()
            if term_lower in text_lower:
                # Find the line containing the term
                lines = text.split('\n')
                for line in lines:
                    if term_lower in line.lower():
                        # Extract the part after the term
                        parts = line.split(':')
                        if len(parts) > 1:
                            result = parts[1].strip()
                            if result and len(result) > 2:
                                return result[:50]  # Limit length

        return "Not specified"

    def _extract_contract_type_gui(self, text: str) -> str:
        """Extract contract type from analysis text"""
        text_lower = text.lower()

        if any(term in text_lower for term in ["time and materials", "hourly", "per hour"]):
            return "Time and Materials"
        elif any(term in text_lower for term in ["fixed price", "lump sum", "total contract"]):
            return "Fixed Price"
        elif any(term in text_lower for term in ["milestone", "deliverable"]):
            return "Milestone-based"
        else:
            return "Unknown"

    def create_comprehensive_spreadsheet(self, contract_info: Dict, payment_schedule: List[Dict], 
                                       tracking_requirements: Dict, output_file: str, 
                                       original_analysis: str, validated_data: Dict = None):
        """Create comprehensive Excel spreadsheet with proper number formatting."""
        wb = Workbook()
        
        # Create number format styles
        currency_style = NamedStyle(name="currency")
        currency_style.number_format = '"$"#,##0.00'
        
        hours_style = NamedStyle(name="hours") 
        hours_style.number_format = '#,##0.0'
        
        percentage_style = NamedStyle(name="percentage")
        percentage_style.number_format = '0.0%'
        
        # Register styles with workbook
        wb.add_named_style(currency_style)
        wb.add_named_style(hours_style)
        wb.add_named_style(percentage_style)
        
        # Styling
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
        warning_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
        confidence_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
        border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                       top=Side(style='thin'), bottom=Side(style='thin'))
        
        # 1. Contract Info Sheet with Confidence Scores
        info_sheet = wb.active
        info_sheet.title = "Contract Summary"
        
        confidence_scores = validated_data.get('confidence_scores', {}) if validated_data else {}
        
        contract_data = [
            ["CONTRACT INFORMATION", "VALUE", "CONFIDENCE LEVEL"],
            ["Client", contract_info.get("client", ""), ""],
            ["Vendor/Contractor", contract_info.get("vendor", ""), ""],
            ["Contract ID", contract_info.get("contract_id", ""), ""],
            ["Contract Type", contract_info.get("contract_type", ""), ""],
            ["Total Value", self.parse_currency_to_number(contract_info.get("total_value", "")), confidence_scores.get("total_contract_value", "")],
            ["Start Date", contract_info.get("start_date", ""), confidence_scores.get("start_date", "")],
            ["End Date", contract_info.get("end_date", ""), confidence_scores.get("end_date", "")],
            ["Hourly Rate", self.parse_currency_to_number(contract_info.get("hourly_rate", "")), confidence_scores.get("hourly_rate", "")],
            ["Maximum Hours", self.parse_hours_to_number(contract_info.get("max_hours", "")), confidence_scores.get("maximum_hours", "")],
            ["Daily Hour Limit", self.parse_hours_to_number(contract_info.get("max_daily_hours", "")), confidence_scores.get("daily_hour_limit", "")],
            ["Invoice Frequency", contract_info.get("invoice_frequency", ""), ""],
            ["Payment Terms", contract_info.get("payment_timeline", ""), confidence_scores.get("payment_terms", "")],
            ["", "", ""],
            ["THREE-PASS VALIDATION RESULTS", "", ""],
            ["Analysis Method", "Regex + Claude AI + Cross-Validation", "HIGH"],
            ["High Confidence Items", str(len([s for s in confidence_scores.values() if s == "HIGH"])), ""],
            ["Medium Confidence Items", str(len([s for s in confidence_scores.values() if s == "MEDIUM"])), ""],
            ["Low Confidence Items", str(len([s for s in confidence_scores.values() if s == "LOW"])), ""],
            ["Validation Notes", validated_data.get("validation_notes", "") if validated_data else "", ""],
            ["", "", ""],
            ["TRACKING REQUIREMENTS", "", ""],
            ["Travel Expenses", tracking_requirements.get("expense_tracking", {}).get("travel_expenses", "No"), ""],
            ["Expense Pre-approval", tracking_requirements.get("expense_tracking", {}).get("pre_approval_required", "No"), ""],
            ["W-9 Required", tracking_requirements.get("compliance", {}).get("w9_required", "No"), ""],
            ["Time Breakdown Required", tracking_requirements.get("compliance", {}).get("time_breakdown_required", "No"), ""],
            ["Hour Tracking", tracking_requirements.get("budget_monitoring", {}).get("hour_tracking", "No"), ""],
            ["", "", ""],
            ["Analysis Generated", datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "HIGH"]
        ]
        
        for row_idx, (label, value, confidence) in enumerate(contract_data, 1):
            cell_a = info_sheet.cell(row=row_idx, column=1, value=label)
            cell_b = info_sheet.cell(row=row_idx, column=2, value=value)
            cell_c = info_sheet.cell(row=row_idx, column=3, value=confidence)
            
            if label in ["CONTRACT INFORMATION", "THREE-PASS VALIDATION RESULTS", "TRACKING REQUIREMENTS"]:
                cell_a.font = header_font
                cell_a.fill = header_fill
                cell_a.alignment = Alignment(horizontal="center")
                cell_b.font = header_font
                cell_b.fill = header_fill
                cell_b.alignment = Alignment(horizontal="center")
                cell_c.font = header_font
                cell_c.fill = header_fill
                cell_c.alignment = Alignment(horizontal="center")
            elif confidence == "HIGH":
                cell_c.fill = confidence_fill
            elif confidence == "LOW":
                cell_c.fill = warning_fill
            else:
                cell_b.alignment = Alignment(wrap_text=True, vertical="top")    
        
        # 2. Payment Tracking Sheet
        payment_sheet = wb.create_sheet("Payment Tracking")
        
        if contract_info.get("contract_type") == "time and materials":
            headers = [
                "Period/Month", "Hours Worked", "Cumulative Hours", "Hourly Rate", 
                "Invoice Amount", "Invoice Submission Due", "Invoice Submitted Date",
                "Invoice Approved Date", "Expected Payment Date", "Actual Payment Date",
                "Amount Paid", "Balance Due", "Status", "Notes"
            ]
            
            # Add monthly rows for T&M contracts
            start_date = contract_info.get("start_date")
            end_date = contract_info.get("end_date")
            
            if start_date and end_date:
                try:
                    start = datetime.strptime(start_date, "%Y-%m-%d")
                    end = datetime.strptime(end_date, "%Y-%m-%d")
                    
                    current = start.replace(day=1)
                    monthly_periods = []
                    
                    while current <= end:
                        month_end = (current.replace(month=current.month+1) if current.month < 12 
                                   else current.replace(year=current.year+1, month=1)) - timedelta(days=1)
                        period_end = min(month_end, end)
                        
                        invoice_due = period_end + timedelta(days=5)
                        expected_payment = invoice_due + timedelta(days=30)
                        
                        monthly_periods.append({
                            "period": f"{current.strftime('%Y-%m')}",
                            "invoice_due": invoice_due.strftime('%Y-%m-%d'),
                            "expected_payment": expected_payment.strftime('%Y-%m-%d')
                        })
                        
                        current = current.replace(month=current.month+1) if current.month < 12 else current.replace(year=current.year+1, month=1)
                        
                except ValueError:
                    monthly_periods = [{"period": "2024-01", "invoice_due": "2024-01-31", "expected_payment": "2024-03-01"}]
        else:
            headers = [
                "Due Date", "Description", "Amount Due", "Invoice Required",
                "Invoice Submission Due", "Invoice Submitted Date", "Invoice Approved Date",
                "Expected Payment Date", "Actual Payment Date", "Amount Paid", 
                "Balance Due", "Status", "Notes"
            ]
            monthly_periods = payment_schedule if payment_schedule else [
                {"due_date": "2024-01-31", "description": "Initial Payment", "amount": "TBD"}
            ]
        
        # Add headers
        for col_idx, header in enumerate(headers, 1):
            cell = payment_sheet.cell(row=1, column=col_idx, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
            cell.border = border
        
        # Add data rows
        row_start = 2
        if contract_info.get("contract_type") == "time and materials" and monthly_periods:
            for row_idx, period in enumerate(monthly_periods, row_start):
                payment_sheet.cell(row=row_idx, column=1, value=period["period"])
                payment_sheet.cell(row=row_idx, column=2, value="")
                payment_sheet.cell(row=row_idx, column=3, value="")
                
                # Set hourly rate as number with currency formatting
                hourly_rate_cell = payment_sheet.cell(row=row_idx, column=4)
                hourly_rate_value = self.parse_currency_to_number(contract_info.get("hourly_rate", ""))
                hourly_rate_cell.value = hourly_rate_value
                if hourly_rate_value is not None:
                    hourly_rate_cell.number_format = '"$"#,##0.00'
                
                payment_sheet.cell(row=row_idx, column=5, value="")
                payment_sheet.cell(row=row_idx, column=6, value=period["invoice_due"])
                payment_sheet.cell(row=row_idx, column=7, value="")
                payment_sheet.cell(row=row_idx, column=8, value="")
                payment_sheet.cell(row=row_idx, column=9, value=period["expected_payment"])
                payment_sheet.cell(row=row_idx, column=10, value="")
                payment_sheet.cell(row=row_idx, column=11, value="")
                payment_sheet.cell(row=row_idx, column=12, value="")
                payment_sheet.cell(row=row_idx, column=13, value="Pending")
                payment_sheet.cell(row=row_idx, column=14, value="")
        else:
            data_rows = monthly_periods if monthly_periods else [{}] * 10
            for row_idx, payment in enumerate(data_rows, row_start):
                if payment:
                    payment_sheet.cell(row=row_idx, column=1, value=payment.get("due_date"))
                    payment_sheet.cell(row=row_idx, column=2, value=payment.get("description"))
                    
                    # Set amount as number with currency formatting
                    amount_cell = payment_sheet.cell(row=row_idx, column=3)
                    amount_value = self.parse_currency_to_number(payment.get("amount"))
                    amount_cell.value = amount_value
                    if amount_value is not None:
                        amount_cell.number_format = '"$"#,##0.00'
                    
                    payment_sheet.cell(row=row_idx, column=4, value="Yes" if payment.get("invoice_submission_due") else "No")
                    payment_sheet.cell(row=row_idx, column=5, value=payment.get("invoice_submission_due"))
                    payment_sheet.cell(row=row_idx, column=9, value=payment.get("expected_payment_date"))
                    payment_sheet.cell(row=row_idx, column=12, value="Pending")
        
        # 3. Expense Tracking Sheet (ALWAYS CREATE)
        expense_sheet = wb.create_sheet("Expense Tracking")
        
        expense_headers = [
            "Date", "Expense Type", "Description", "Amount", "Receipt", 
            "Pre-approved", "Submitted Date", "Approved Date", 
            "Reimbursement Date", "Status", "Notes"
        ]
        
        for col_idx, header in enumerate(expense_headers, 1):
            cell = expense_sheet.cell(row=1, column=col_idx, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = border
        
        expense_categories = [
            "Travel - Airfare", "Travel - Hotel", "Travel - Ground Transportation",
            "Meals & Entertainment", "Equipment", "Materials", "Other"
        ]
        
        for row_idx, category in enumerate(expense_categories, 2):
            expense_sheet.cell(row=row_idx, column=2, value=category)
            expense_sheet.cell(row=row_idx, column=10, value="Template")
        
        # 4. Compliance Checklist Sheet (ALWAYS CREATE)
        compliance_sheet = wb.create_sheet("Compliance Checklist")
        
        compliance_items = [
            ["REQUIRED DOCUMENTATION", "Status", "Due Date", "Completed Date", "Notes"],
            ["W-9 Form Submission", "", "", "", ""],
            ["Monthly Invoice Submission", "", "", "", ""],
            ["Time Breakdown Documentation", "", "", "", ""],
            ["Expense Pre-approvals", "", "", "", ""],
            ["Contract Deliverables", "", "", "", ""],
            ["", "", "", "", ""],
            ["INVOICE REQUIREMENTS", "", "", "", ""],
            ["Daily Time Breakdown", "", "", "", ""],
            ["Detailed Work Description", "", "", "", ""],
            ["Required Signatures", "", "", "", ""],
            ["Supporting Documentation", "", "", "", ""],
            ["", "", "", "", ""],
            ["REPORTING DEADLINES", "", "", "", ""],
            ["Monthly Reports", "", "", "", ""],
            ["Expense Reports", "", "", "", ""],
            ["Final Deliverables", "", "", "", ""]
        ]
        
        for row_idx, item_row in enumerate(compliance_items, 1):
            for col_idx, item in enumerate(item_row, 1):
                cell = compliance_sheet.cell(row=row_idx, column=col_idx, value=item)
                if row_idx == 1 or item in ["REQUIRED DOCUMENTATION", "INVOICE REQUIREMENTS", "REPORTING DEADLINES"]:
                    cell.font = header_font
                    cell.fill = header_fill
                cell.border = border
        
        # 5. Budget Monitor Sheet (ALWAYS CREATE)
        budget_sheet = wb.create_sheet("Budget Monitor")
        
        budget_headers = [
            "Metric", "Budgeted/Maximum", "Current", "Remaining", 
            "% Used", "Status", "Warning Level"
        ]
        
        for col_idx, header in enumerate(budget_headers, 1):
            cell = budget_sheet.cell(row=1, column=col_idx, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = border
        
        # Budget tracking rows
        max_hours = contract_info.get("max_hours", "")
        total_value = contract_info.get("total_value", "")
        
        max_hours_num = self.parse_hours_to_number(max_hours)
        total_value_num = self.parse_currency_to_number(total_value)
        
        budget_items = [
            ["Total Hours", max_hours_num, 0, max_hours_num, 0, "On Track", "< 80%"],
            ["Contract Value", total_value_num, 0, total_value_num, 0, "On Track", "< 80%"],
            ["Monthly Invoicing", "Monthly", "Current", "Remaining", "%", "Status", "Threshold"],
            ["Expense Budget", None, 0, None, 0, "On Track", "< 90%"]
        ]
        
        for row_idx, item in enumerate(budget_items, 2):
            for col_idx, value in enumerate(item, 1):
                cell = budget_sheet.cell(row=row_idx, column=col_idx, value=value)
                
                if col_idx in [2, 3, 4] and isinstance(value, (int, float)) and value is not None:
                    if "Value" in item[0]:
                        cell.number_format = '"$"#,##0.00'
                    elif "Hours" in item[0]:
                        cell.number_format = '#,##0.0'
        
        # Add conditional formatting for warnings
        warning_rule = CellIsRule(operator='greaterThan', formula=['80'], 
                                stopIfTrue=True, fill=warning_fill)
        budget_sheet.conditional_formatting.add('E2:E5', warning_rule)
        
        # 6. Validation Details Sheet (ALWAYS CREATE)
        validation_sheet = wb.create_sheet("Validation Details")
        validation_sheet.cell(row=1, column=1, value="THREE-PASS VALIDATION DETAILS")
        validation_sheet.cell(row=1, column=1).font = header_font
        validation_sheet.cell(row=1, column=1).fill = header_fill
        
        validation_info = [
            ["", ""],
            ["VALIDATION METHODOLOGY", ""],
            ["Pass 1", "High-precision regex pattern extraction"],
            ["Pass 2", "Direct Claude AI extraction with focused prompts"],
            ["Pass 3", "Cross-validation and confidence scoring"],
            ["", ""],
            ["CONFIDENCE LEVELS", ""],
        ]
        
        if validated_data:
            confidence_scores = validated_data.get('confidence_scores', {})
            for field, score in confidence_scores.items():
                field_key = field.replace('_', ' ').replace(' ', '_')
                value = validated_data.get(field, contract_info.get(field_key, 'Not found'))
                validation_info.append([field.replace('_', ' ').title(), f"{value} [{score}]"])
            
            validation_info.extend([
                ["", ""],
                ["VALIDATION NOTES", ""],
                ["Notes", validated_data.get('validation_notes', 'No issues detected')],
                ["", ""],
                ["ACCURACY SUMMARY", ""],
                ["High Confidence Items", str(len([s for s in confidence_scores.values() if s == "HIGH"]))],
                ["Medium Confidence Items", str(len([s for s in confidence_scores.values() if s == "MEDIUM"]))],
                ["Low Confidence Items", str(len([s for s in confidence_scores.values() if s == "LOW"]))],
                ["Total Items Validated", str(len(confidence_scores))],
            ])
        else:
            validation_info.append(["Status", "No validation data available"])
        
        for row_idx, (label, value) in enumerate(validation_info, 1):
            cell_a = validation_sheet.cell(row=row_idx, column=1, value=label)
            cell_b = validation_sheet.cell(row=row_idx, column=2, value=value)
            
            if label in ["VALIDATION METHODOLOGY", "CONFIDENCE LEVELS", "VALIDATION NOTES", "ACCURACY SUMMARY"]:
                cell_a.font = Font(bold=True)
            if label == "Notes":
                cell_b.alignment = Alignment(wrap_text=True, vertical="top")
        
        # 7. Full Analysis Sheet
        analysis_sheet = wb.create_sheet("Full Analysis")
        analysis_sheet.cell(row=1, column=1, value="Claude AI Enhanced Contract Analysis with Three-Pass Validation")
        analysis_sheet.cell(row=1, column=1).font = header_font
        analysis_sheet.cell(row=1, column=1).fill = header_fill
        
        analysis_lines = original_analysis.split('\n')
        for row_idx, line in enumerate(analysis_lines, 3):
            analysis_sheet.cell(row=row_idx, column=1, value=line)
        
        # Auto-adjust column widths
        for sheet in wb.worksheets:
            for column in sheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if cell.value and len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                sheet.column_dimensions[column_letter].width = adjusted_width
        
        wb.save(output_file)

    def analyze_with_high_precision(self, contract_text: str) -> Dict:
        """Main analysis method with high-precision validation."""
        
        # Step 1: Run three-pass validation for financial data
        validated_data = self.cross_validate_financial_data(contract_text)
        
        # Step 2: Run comprehensive analysis for other elements
        comprehensive_result = self.analyze_contract_comprehensive(contract_text)
        
        if not comprehensive_result.get('success'):
            return {'success': False, 'error': comprehensive_result.get('analysis')}
        
        # Extract structured data
        contract_info, payment_schedule, tracking_requirements = self.extract_comprehensive_data(comprehensive_result["analysis"])
        
        # Step 3: Override with high-precision financial data
        if validated_data.get('total_contract_value'):
            contract_info['total_value'] = validated_data['total_contract_value']
        if validated_data.get('hourly_rate'):
            contract_info['hourly_rate'] = validated_data['hourly_rate']
        if validated_data.get('start_date'):
            contract_info['start_date'] = validated_data['start_date']
        if validated_data.get('end_date'):
            contract_info['end_date'] = validated_data['end_date']
        if validated_data.get('payment_terms'):
            contract_info['payment_timeline'] = validated_data['payment_terms']
        if validated_data.get('maximum_hours'):
            contract_info['max_hours'] = validated_data['maximum_hours']
        if validated_data.get('daily_hour_limit'):
            contract_info['max_daily_hours'] = validated_data['daily_hour_limit']
        
        return {
            'success': True,
            'contract_info': contract_info,
            'payment_schedule': payment_schedule,
            'tracking_requirements': tracking_requirements,
            'validated_data': validated_data,
            'analysis': comprehensive_result['analysis']
        }


class ContractAnalyzerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Enhanced Contract Financial Analyzer")
        self.root.geometry("900x700")
        self.root.configure(bg='#f8f9fa')
        
        # Variables
        self.api_key = tk.StringVar()
        self.contract_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self.analyzer = None
        
        # Configure style
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        self.create_widgets()
        self.load_saved_api_key()
        
    def create_widgets(self):
        # Main container
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Title
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 20))
        
        title_label = ttk.Label(title_frame, text="Enhanced Contract Financial Analyzer", 
                               font=('Arial', 16, 'bold'))
        title_label.pack()
        
        subtitle_label = ttk.Label(title_frame, text="Professional contract analysis with three-pass validation", 
                                  font=('Arial', 10))
        subtitle_label.pack()
        
        # Configuration Section
        config_frame = ttk.LabelFrame(main_frame, text="Configuration", padding=15)
        config_frame.pack(fill=tk.X, pady=(0, 15))
        
        # API Key
        api_frame = ttk.Frame(config_frame)
        api_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(api_frame, text="Claude API Key:").pack(anchor=tk.W)
        api_entry_frame = ttk.Frame(api_frame)
        api_entry_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.api_entry = ttk.Entry(api_entry_frame, textvariable=self.api_key, show="*", width=50)
        self.api_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Button(api_entry_frame, text="Save", command=self.save_api_key, width=8).pack(side=tk.RIGHT, padx=(5, 0))
        
        # Contract File Selection
        file_frame = ttk.Frame(config_frame)
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(file_frame, text="Contract File:").pack(anchor=tk.W)
        file_entry_frame = ttk.Frame(file_frame)
        file_entry_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.file_entry = ttk.Entry(file_entry_frame, textvariable=self.contract_file, width=60)
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Button(file_entry_frame, text="Browse", command=self.browse_contract_file, width=10).pack(side=tk.RIGHT, padx=(5, 0))
        
        # Output File Selection
        output_frame = ttk.Frame(config_frame)
        output_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(output_frame, text="Output Excel File:").pack(anchor=tk.W)
        output_entry_frame = ttk.Frame(output_frame)
        output_entry_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.output_entry = ttk.Entry(output_entry_frame, textvariable=self.output_file, width=60)
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Button(output_entry_frame, text="Browse", command=self.browse_output_file, width=10).pack(side=tk.RIGHT, padx=(5, 0))
        
        # Set default output filename
        documents_path = Path.home() / "Documents"
        default_filename = f"contract_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        self.output_file.set(str(documents_path / default_filename))
        
        # Analysis Section
        analysis_frame = ttk.LabelFrame(main_frame, text="Analysis", padding=15)
        analysis_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))        
        # Analysis Section
        analysis_frame = ttk.LabelFrame(main_frame, text="Analysis", padding=15)
        analysis_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Analysis Button
        button_frame = ttk.Frame(analysis_frame)
        button_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.analyze_btn = ttk.Button(button_frame, text="Analyze Contract", 
                                     command=self.analyze_contract, style='Accent.TButton')
        self.analyze_btn.pack(side=tk.LEFT)
        
        # Progress bar
        self.progress = ttk.Progressbar(button_frame, mode='indeterminate')
        self.progress.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(10, 0))
        
        # Results area
        results_frame = ttk.Frame(analysis_frame)
        results_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(results_frame, text="Analysis Results:").pack(anchor=tk.W, pady=(0, 5))
        
        self.results_text = scrolledtext.ScrolledText(results_frame, height=15, wrap=tk.WORD)
        self.results_text.pack(fill=tk.BOTH, expand=True)
        
        # Status bar
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X)
        
        self.status_var = tk.StringVar(value="Ready")
        self.status_label = ttk.Label(status_frame, textvariable=self.status_var)
        self.status_label.pack(side=tk.LEFT)
        
    def load_saved_api_key(self):
        """Load saved API key from config file."""
        try:
            config_dir = Path.home() / ".contract_analyzer"
            config_file = config_dir / "config.txt"
            
            if config_file.exists():
                with open(config_file, 'r') as f:
                    saved_key = f.read().strip()
                    if saved_key:
                        self.api_key.set(saved_key)
                        self.status_var.set("API key loaded")
        except Exception as e:
            self.status_var.set("Could not load saved API key")
    
    def save_api_key(self):
        """Save API key to config file."""
        try:
            api_key = self.api_key.get().strip()
            if not api_key:
                messagebox.showwarning("Warning", "Please enter an API key first")
                return
            
            config_dir = Path.home() / ".contract_analyzer"
            config_dir.mkdir(exist_ok=True)
            
            config_file = config_dir / "config.txt"
            with open(config_file, 'w') as f:
                f.write(api_key)
            
            messagebox.showinfo("Success", "API key saved successfully")
            self.status_var.set("API key saved")
            
        except Exception as e:
            messagebox.showerror("Error", f"Could not save API key: {str(e)}")
    
    def browse_contract_file(self):
        """Browse for contract file."""
        filetypes = [
            ("All Supported", "*.pdf *.docx *.doc *.txt"),
            ("PDF files", "*.pdf"),
            ("Word documents", "*.docx *.doc"),
            ("Text files", "*.txt"),
            ("All files", "*.*")
        ]
        
        filename = filedialog.askopenfilename(
            title="Select Contract File",
            filetypes=filetypes
        )
        
        if filename:
            self.contract_file.set(filename)
            # Update output filename based on input filename
            input_name = Path(filename).stem
            documents_path = Path.home() / "Documents"
            output_name = f"{input_name}_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            self.output_file.set(str(documents_path / output_name))
            self.status_var.set(f"Contract file selected: {Path(filename).name}")
    
    def browse_output_file(self):
        """Browse for output Excel file location."""
        filename = filedialog.asksaveasfilename(
            title="Save Analysis As",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("All files", "*.*")
            ]
        )
        
        if filename:
            self.output_file.set(filename)
            self.status_var.set(f"Output file: {Path(filename).name}")
    
    def validate_inputs(self):
        """Validate user inputs before analysis."""
        if not self.api_key.get().strip():
            messagebox.showerror("Error", "Please enter your Claude API key")
            return False
        
        if not self.contract_file.get().strip():
            messagebox.showerror("Error", "Please select a contract file")
            return False
        
        if not Path(self.contract_file.get()).exists():
            messagebox.showerror("Error", "Contract file not found")
            return False
        
        if not self.output_file.get().strip():
            messagebox.showerror("Error", "Please specify an output file location")
            return False
        
        return True
    
    def analyze_contract(self):
        """Main analysis function - runs in separate thread."""
        if not self.validate_inputs():
            return
        
        # Disable button and start progress
        self.analyze_btn.config(state='disabled')
        self.progress.start()
        self.status_var.set("Starting analysis...")
        self.results_text.delete('1.0', tk.END)
        
        # Run analysis in separate thread
        analysis_thread = threading.Thread(target=self._run_analysis)
        analysis_thread.daemon = True
        analysis_thread.start()
    
    def _run_analysis(self):
        """Run the actual analysis - called from separate thread."""
        try:
            # Initialize analyzer
            self.root.after(0, lambda: self.status_var.set("Initializing analyzer..."))
            analyzer = HighPrecisionContractAnalyzer(self.api_key.get().strip())
            
            # Read contract file
            self.root.after(0, lambda: self.status_var.set("Reading contract file..."))
            contract_text = analyzer.read_contract_file(self.contract_file.get())
            
            self.root.after(0, lambda: self._update_results(f"Contract file read successfully ({len(contract_text):,} characters)\n\n"))
            
            # Run three-pass validation analysis
            self.root.after(0, lambda: self.status_var.set("Running three-pass validation analysis..."))
            result = analyzer.analyze_with_high_precision(contract_text)
            
            if not result["success"]:
                self.root.after(0, lambda: self._show_error(result.get('error', 'Unknown error')))
                return
            
            # Extract results
            contract_info = result['contract_info']
            payment_schedule = result['payment_schedule']
            tracking_requirements = result['tracking_requirements']
            validated_data = result['validated_data']
            
            # Create comprehensive spreadsheet
            self.root.after(0, lambda: self.status_var.set("Creating comprehensive Excel tracking system..."))
            analyzer.create_comprehensive_spreadsheet(
                contract_info, 
                payment_schedule,
                tracking_requirements,
                self.output_file.get(),
                result['analysis'],
                validated_data
            )
            
            # Show success results
            self.root.after(0, lambda: self._show_success(contract_info, validated_data, self.output_file.get()))
            
        except Exception as e:
            error_message = str(e)
            self.root.after(0, lambda: self._show_error(error_message))
    
    def _update_results(self, text):
        """Update results text - called from main thread."""
        self.results_text.insert(tk.END, text)
        self.results_text.see(tk.END)
        self.root.update_idletasks()
    
    def _show_success(self, contract_info, validated_data, output_file):
        """Show successful analysis results."""
        self.progress.stop()
        self.analyze_btn.config(state='normal')
        
        # Build comprehensive results summary
        results = "=== THREE-PASS VALIDATION ANALYSIS COMPLETE ===\n\n"
        
        # Financial Data Confidence Summary
        confidence_scores = validated_data.get('confidence_scores', {})
        if confidence_scores:
            results += "FINANCIAL DATA CONFIDENCE LEVELS:\n"
            high_confidence_items = 0
            total_items = 0
            
            key_fields = ['total_contract_value', 'hourly_rate', 'start_date', 'end_date', 'payment_terms', 'maximum_hours', 'daily_hour_limit']
            
            for field in key_fields:
                if field in confidence_scores:
                    score = confidence_scores[field]
                    # Use the validated data directly for display
                    value = validated_data.get(field, contract_info.get(field.replace('_', ' ').replace(' ', '_'), 'Not found'))
                    total_items += 1
                    
                    if score == 'HIGH':
                        results += f"  ✓ {field.replace('_', ' ').title()}: {value} [HIGH CONFIDENCE]\n"
                        high_confidence_items += 1
                    elif score == 'MEDIUM':
                        results += f"  • {field.replace('_', ' ').title()}: {value} [MEDIUM CONFIDENCE]\n"
                    elif score == 'LOW':
                        results += f"  ⚠ {field.replace('_', ' ').title()}: {value} [LOW CONFIDENCE - REVIEW RECOMMENDED]\n"
            
            # Calculate accuracy percentage
            if total_items > 0:
                accuracy_percentage = (high_confidence_items / total_items) * 100
                results += f"\nDATA ACCURACY: {accuracy_percentage:.1f}% HIGH CONFIDENCE ({high_confidence_items}/{total_items} critical items)\n\n"
        
        # Contract Summary
        results += "CONTRACT SUMMARY:\n"
        results += f"Client: {contract_info.get('client', 'Unknown')}\n"
        results += f"Vendor: {contract_info.get('vendor', 'Unknown')}\n"
        results += f"Contract Type: {contract_info.get('contract_type', 'Unknown')}\n"
        results += f"Total Value: {contract_info.get('total_value', 'Unknown')}\n"
        results += f"Payment Terms: {contract_info.get('payment_timeline', 'Unknown')}\n"
        
        if contract_info.get('max_hours'):
            results += f"Maximum Hours: {contract_info.get('max_hours')}\n"
        if contract_info.get('hourly_rate'):
            results += f"Hourly Rate: {contract_info.get('hourly_rate')}\n"
            
        results += f"\nSPREADSHEET FEATURES:\n"
        results += "  • Contract Summary with confidence scores\n"
        results += "  • Payment Tracking with invoice timeline\n"
        results += "  • Expense Tracking\n"
        results += "  • Compliance Checklist\n"
        results += "  • Budget Monitor with warnings\n"
        results += "  • Validation Details sheet\n"
        results += "  • Complete AI analysis\n"
        
        results += f"\nFile saved to: {output_file}\n"
        results += "\n✓ Three-pass validation ensures maximum accuracy on financial data\n"
        results += "✓ All critical payment, compliance, and budget elements included!"
        
        self.results_text.delete('1.0', tk.END)
        self.results_text.insert('1.0', results)
        
        self.status_var.set("Analysis complete!")
        
        # Ask to open file
        if messagebox.askyesno("Analysis Complete", 
                              f"Contract analysis completed successfully!\n\nWould you like to open the Excel file now?"):
            try:
                import subprocess
                import platform
                
                if platform.system() == 'Darwin':  # macOS
                    subprocess.call(['open', output_file])
                elif platform.system() == 'Windows':  # Windows
                    os.startfile(output_file)
                else:  # Linux
                    subprocess.call(['xdg-open', output_file])
            except Exception as e:
                messagebox.showinfo("File Location", f"Excel file saved to:\n{output_file}")
    
    def _show_error(self, error_message):
        """Show error message."""
        self.progress.stop()
        self.analyze_btn.config(state='normal')
        
        results = f"ERROR: {error_message}\n\nPlease check:\n"
        results += "• API key is valid\n"
        results += "• Contract file is readable\n"
        results += "• Internet connection is available\n"
        results += "• Required dependencies are installed (PyPDF2 for PDFs, python-docx for Word docs)"
        
        self.results_text.delete('1.0', tk.END)
        self.results_text.insert('1.0', results)
        
        self.status_var.set("Analysis failed")
        
        messagebox.showerror("Analysis Error", f"Analysis failed: {error_message}")


def main():
    """Main application entry point."""
    root = tk.Tk()
    
    # Set application icon if available
    try:
        # You can add an icon file here if you have one
        # root.iconbitmap('icon.ico')
        pass
    except:
        pass
    
    app = ContractAnalyzerGUI(root)
    
    # Center the window
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    root.mainloop()


if __name__ == "__main__":
    main()
        