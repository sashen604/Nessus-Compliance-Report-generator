# Windows 7 Compliance Reports Generator

## Overview

This script automatically detects **Windows 7 compliance checks** from Nessus scan files and generates professional reports in both **Word and Excel** formats. It creates two types of reports:

1. **Summary Reports** - Overview with statistics and key findings
2. **Detailed Reports** - Complete compliance details with solutions and impact

Each type is available in both Word (.docx) and Excel (.xlsx) formats.

---

## Quick Start

### Generate All 4 Reports (Default)
```bash
python3 windows7_compliance_reports.py "/path/to/scan.nessus"
```

**Output:**
- `windows7_compliance_summary.docx` - Summary report in Word
- `windows7_compliance_summary.xlsx` - Summary report in Excel
- `windows7_compliance_details.docx` - Details report in Word
- `windows7_compliance_details.xlsx` - Details report in Excel

---

## Usage Options

### Basic Usage
```bash
python3 windows7_compliance_reports.py "scan.nessus"
```

### Custom Output Path
```bash
python3 windows7_compliance_reports.py "scan.nessus" --output "output/my_report"
```

**Output:**
- `output/my_report_summary.docx`
- `output/my_report_summary.xlsx`
- `output/my_report_details.docx`
- `output/my_report_details.xlsx`

### Word Format Only
```bash
python3 windows7_compliance_reports.py "scan.nessus" --format word
```

**Output:**
- `windows7_compliance_summary.docx`
- `windows7_compliance_details.docx`

### Excel Format Only
```bash
python3 windows7_compliance_reports.py "scan.nessus" --format excel
```

**Output:**
- `windows7_compliance_summary.xlsx`
- `windows7_compliance_details.xlsx`

### Summary Report Only
```bash
python3 windows7_compliance_reports.py "scan.nessus" --type summary
```

**Output:**
- `windows7_compliance_summary.docx`
- `windows7_compliance_summary.xlsx`

### Details Report Only
```bash
python3 windows7_compliance_reports.py "scan.nessus" --type details
```

**Output:**
- `windows7_compliance_details.docx`
- `windows7_compliance_details.xlsx`

### Combine Options
```bash
python3 windows7_compliance_reports.py "scan.nessus" \
  --output "reports/kandy_win7" \
  --format excel \
  --type details
```

**Output:**
- `reports/kandy_win7_details.xlsx`

---

## Report Formats

### Summary Reports

#### Word Format (Summary)
```
Windows 7 Compliance Summary Report

Host: 196.1.14.1

Summary Statistics
- L1 Checks: 5
- L2 Checks: 15
- Total Checks: 20

Compliance Summary Table
[Professional table with Check IDs, Levels, Descriptions, Status, Benchmark]
```

#### Excel Format (Summary)
```
Sheet 1: Summary
- Title: "Windows 7 Compliance Summary Report"
- Statistics displayed

Sheet 2: Compliance Data
[5 columns: Check ID | Level | Description | Status | Benchmark]
[Data rows with formatting and borders]
```

---

### Details Reports

#### Word Format (Details)
```
Windows 7 Compliance Detailed Report

L1 Checks: 5
L2 Checks: 15
Total: 20

[For each check:]
[L1] Check Description
  Result: PASSED / FAILED
  Info: [compliance information]
  Solution: [recommended solution]
  Impact: [business impact]
  Policy Value: [expected value]
  Configured Value: [actual value]
```

#### Excel Format (Details)
```
10 Columns:
1. IP Address
2. Check ID
3. Description
4. Level (L1/L2)
5. Result (PASSED/FAILED)
6. Info
7. Solution
8. Impact
9. Policy Value
10. Configured Value

Professional formatting:
- Blue headers with white text
- Borders on all cells
- Text wrapping enabled
- Optimized column widths
- Frozen header row
```

---

## Command-Line Arguments

### Required
- `scan_path` - Path to the Nessus .nessus file

### Optional

#### `--output`
- **Default:** `windows7_compliance`
- **Example:** `--output "output/kandy_win7"`
- Creates files with suffixes: `_summary.docx`, `_summary.xlsx`, `_details.docx`, `_details.xlsx`

#### `--format`
- **Default:** `both`
- **Options:** `both`, `word`, `excel`
- Controls which output formats to generate

#### `--type`
- **Default:** `both`
- **Options:** `both`, `summary`, `details`
- Controls which report types to generate

---

## Examples

### Example 1: Standard Report Generation
```bash
python3 windows7_compliance_reports.py \
  "/home/user/scans/server01.nessus"
```

Result:
- All 4 files created: 2 Word + 2 Excel
- Summary and Details reports
- Files named with default prefix

### Example 2: Executive Summary Only (Excel)
```bash
python3 windows7_compliance_reports.py \
  "/home/user/scans/server01.nessus" \
  --output "reports/exec_summary" \
  --format excel \
  --type summary
```

Result:
- `reports/exec_summary_summary.xlsx`
- Professional summary in Excel format for executives

### Example 3: Detailed Audit Report (Word)
```bash
python3 windows7_compliance_reports.py \
  "/home/user/scans/server01.nessus" \
  --output "audit_reports/detailed_audit" \
  --format word \
  --type details
```

Result:
- `audit_reports/detailed_audit_details.docx`
- Complete details for compliance audits

### Example 4: Batch Processing Multiple Servers
```bash
for file in /home/user/scans/server*.nessus; do
  python3 windows7_compliance_reports.py "$file" \
    --output "output/$(basename "$file" .nessus)" \
    --format both \
    --type both
done
```

Result:
- Creates 4 files for each server scanned

---

## Output Details

### Summary Report Contents

**Word Summary:**
- Host IP address
- L1/L2/Total check counts
- Compliance summary table with:
  - Check ID
  - Level (L1/L2)
  - Description (first 50 chars)
  - Status (PASSED/FAILED)
  - Benchmark name

**Excel Summary:**
- Summary sheet with statistics
- Data sheet with same table as Word
- Professional formatting and styling
- Cell borders and text wrapping

### Details Report Contents

**Word Details:**
- Host IP address
- Complete check counts
- For each check:
  - Full check description
  - Compliance result
  - Information
  - Recommended solution
  - Business impact
  - Policy and actual values

**Excel Details:**
- 10-column layout
- All compliance data
- Professional color scheme
- Frozen header row
- Text wrapping for readability
- Borders on all cells

---

## Features

✅ **Automatic Windows 7 Detection**
- Detects CIS Windows 7 benchmarks
- Filters only relevant compliance checks
- Skips non-Windows-7 items

✅ **Dual Format Output**
- Word (.docx) for sharing and printing
- Excel (.xlsx) for data analysis

✅ **Multiple Report Types**
- Summary for executives and quick overview
- Details for compliance auditors

✅ **Professional Formatting**
- Color-coded headers
- Consistent styling
- Easy to read and share

✅ **Flexible Configuration**
- Choose format(s) to generate
- Choose report type(s)
- Custom output locations

✅ **Statistics**
- L1 vs L2 check counts
- Total compliance items
- Results summary

---

## Output File Examples

### File Sizes
- Summary Word: ~37 KB
- Summary Excel: ~5.7 KB
- Details Word: ~37 KB
- Details Excel: ~5.3 KB

(Sizes vary based on number of compliance checks)

### File Naming Convention
```
[OUTPUT_PREFIX]_summary.docx
[OUTPUT_PREFIX]_summary.xlsx
[OUTPUT_PREFIX]_details.docx
[OUTPUT_PREFIX]_details.xlsx
```

**Example:**
```
windows7_report_summary.docx
windows7_report_summary.xlsx
windows7_report_details.docx
windows7_report_details.xlsx
```

---

## Statistics & Counts

The reports include:

✓ **L1 Checks** - Critical compliance items
✓ **L2 Checks** - Standard compliance items  
✓ **Total** - Combined count
✓ **Host IP** - Target system address
✓ **Benchmark** - CIS benchmark version used

---

## Troubleshooting

### Q: No Windows 7 checks found
**A:** The file may not contain Windows 7 compliance checks. The script only processes:
- Files with "CIS_MS_Windows_7" in check names
- Files with "Windows_7" in audit file names
- Check names containing "Windows 7"

### Q: Output directory not created
**A:** The script automatically creates directories. Ensure you have write permissions to the path.

### Q: Files are empty
**A:** May indicate no matching compliance checks. Check the console output for the count.

### Q: Excel file won't open
**A:** Ensure openpyxl is properly installed:
```bash
pip install openpyxl
```

---

## Integration with Main Generator

This script complements the main `nessus-report-gen.py`. Use this for:
- **Specific Windows 7 analysis**
- **Executive summaries**
- **Compliance audits**
- **Quick format generation**

---

## Version Information

- **Version:** 1.0
- **Date:** 2024-04-24
- **Status:** Production Ready
- **Tested:** Python 3.8+
- **Requirements:** openpyxl, python-docx

---

## Support

For issues or feature requests, refer to:
1. Check console output for error messages
2. Verify the .nessus file contains compliance checks
3. Ensure all dependencies are installed

---

Made with ❤️ for Windows 7 compliance teams
