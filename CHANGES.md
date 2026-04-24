## Summary of Changes - Excel Export Feature

### ✅ Files Created
1. **details_report_excel.py** - New Excel generator module
   - Generates Excel reports with 10 columns:
     - IP Address
     - Check Number
     - Level
     - Description
     - Result
     - Info
     - Solution
     - Impact
     - Policy Value
     - Configured Value
   - Professional formatting with header colors and borders
   - Frozen header row for easy navigation
   - Optimized column widths

### ✅ Files Modified
1. **requirements.txt**
   - Added: `openpyxl>=3.0.0` (for Excel file generation)

2. **nessus-report-gen.py** (Main interactive generator)
   - Added import: `from details_report_excel import build_excel`
   - Updated `ask_user()` function:
     - Added Option 3: Detailed (Excel) to report type menu
   - Updated `generate_name()` function:
     - Returns .xlsx extension for report_type == "3"
   - Updated `process_file()` function:
     - Added handler for Excel report generation
   - Updated `main()` function:
     - Added file extension validation for Excel files
     - Excel combined mode creates separate files (since Excel can't combine like DOCX)

3. **README.md**
   - Updated contents section to include `details_report_excel.py`
   - Updated introduction to mention Excel reports
   - Added Section 4b: Detailed compliance report (Excel)
   - Added Excel report features documentation
   - Added Section 10: Complete documentation for interactive main generator

### ✅ Features Implemented
- ✅ Excel export with 10 columns as requested
- ✅ Includes IP Address column
- ✅ Professional formatting with header colors (light blue #B4C6E7)
- ✅ Frozen header row for easier navigation
- ✅ Cell borders for clear data separation
- ✅ Word wrapping for long text values
- ✅ Integrated into main nessus-report-gen.py
- ✅ Support for L1/L2 filtering
- ✅ Support for single file and folder batch processing

### 🚀 How to Use

**Option 1: Direct Excel Generation**
```bash
python3 details_report_excel.py "/path/to/scan.nessus" --output "report.xlsx"
```

**Option 2: Interactive Main Generator (RECOMMENDED)**
```bash
python3 nessus-report-gen.py
# Select: 3 for Excel option
```

### 📋 Column Mapping
| # | Column | Source |
|---|--------|--------|
| 1 | IP Address | Host IP from report |
| 2 | Check Number | compliance-check-name |
| 3 | Level | L1 or L2 from check name |
| 4 | Description | Check description |
| 5 | Result | compliance-result or Unknown |
| 6 | Info | compliance-info |
| 7 | Solution | compliance-solution |
| 8 | Impact | compliance-impact |
| 9 | Policy Value | compliance-policy-value |
| 10 | Configured Value | compliance-actual-value |

### 🔧 Dependencies
- openpyxl >= 3.0.0 (installed)
- python-docx >= 1.1.0 (already installed)
- tabulate >= 0.9.0 (already installed)

All dependencies are listed in requirements.txt
