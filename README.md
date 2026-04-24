# Nessus Compliance & Summary Reports

Generate terminal summaries, CSV/JSON exports, detailed Word reports, and Excel reports from Nessus `.nessus` / `.xml` files.

## Contents
- `nessus_summary.py` – CLI summary tables + CSV/JSON export
- `compliance_table_docx.py` – simple compliance table Word export
- `details_report_docx.py` – detailed compliance Word report
- `details_report_excel.py` – detailed compliance Excel report ✅ NEW
- `nessus-report-gen.py` – Interactive main report generator
- `generate_compliance_docs.sh` – batch simple compliance docx
- `summary_reports_docx.py` – one Word file with summaries for many scans
- `summary_compliance_folder_docx.py` – one Word file with compliance tables for many scans

## 1) Setup (recommended: virtual environment)

```bash
python3 -m venv .venv

Source .venv/bin/activate #Linux
.venv/Script/Activate.ps1 #windows

pip install -r requirements.txt
```
## 2) run the tool

```bash
python nessus-report-gen.py
```
## 3) CLI summary (terminal + CSV/JSON)

```bash
"./.venv/bin/python" nessus_summary.py "/path/to/scan.nessus"
```

Examples:
```bash
"./.venv/bin/python" nessus_summary.py "/path/to/scan.nessus" --summary host --summary compliance
"./.venv/bin/python" nessus_summary.py "/path/to/scan.nessus" --failed-only
"./.venv/bin/python" nessus_summary.py "/path/to/scan.nessus" --export findings.csv --export-json findings.json
```

## 3) Simple compliance table (Word)

```bash
"./.venv/bin/python" compliance_table_docx.py "/path/to/scan.nessus" --output "/path/to/compliance_table.docx"
```

## 4) Detailed compliance report (Word)

```bash
"./.venv/bin/python" details_report_docx.py "/path/to/scan.nessus" --output "/path/to/detailed_report.docx"
```

## 4b) Detailed compliance report (Excel) ✅ NEW

```bash
"./.venv/bin/python" details_report_excel.py "/path/to/scan.nessus" --output "/path/to/detailed_report.xlsx"
```

**Excel report features:**
- 10 columns: IP Address, Check Number, Level, Description, Result, Info, Solution, Impact, Policy Value, Configured Value
- Professional formatting with headers and borders
- Frozen header row for easy navigation
- Optimized column widths

## 5) Generate One Word document with all .nessus file in side Folder
```bash
python3 summary_compliance_folder_docx.py "path/reports" --output "path/compliance_summary_report.docx"
```

## 6) Genarate One word Document With all .nessus File inside Folder with Details report 
```bash
python3 Detail_report_folder_docx.py ../report/ --output finala_details_report.docx
```

## 7) Batch generation scripts

### Simple compliance tables for many `.nessus` files
```bash
./generate_compliance_docs.sh "/path/to/nessus/folder" "/path/to/output-docs"
```

### Detailed reports for many `.nessus` files
```bash
./Detailsreportgenratetor/generate_detailed_reports.sh "/path/to/nessus/folder" "/path/to/output-docs"
```

## 8) One Word file with summaries for many scans

```bash
"./.venv/bin/python" summary_reports_docx.py "/path/to/nessus/folder" --output "/path/to/summary_report.docx"
```

## 9) One Word file with compliance tables for many scans

```bash
"./.venv/bin/python" summary_compliance_folder_docx.py "/path/to/nessus/folder" --output "/path/to/compliance_summary_report.docx"
```

## 10) Interactive Main Report Generator (All-in-One) ✅ RECOMMENDED

This is the easiest way to generate any report type:

```bash
python3 nessus-report-gen.py
```

**Interactive menu:**
1. Choose input: Single .nessus file or folder
2. Choose report type:
   - **Option 1:** Summary (Table)
   - **Option 2:** Detailed (Professional DOCX)
   - **Option 3:** Detailed (Excel) ✅ NEW
3. Choose output mode:
   - Combined into one file
   - Separate files per report
4. Filter by level: L1 + L2 or L1 only
5. Set output folder and filename

**Example workflow:**
```
Enter .nessus file OR folder path: /path/to/folder
Report Type: 3 (Excel)
Output Mode: 1 (Combined)  ← Now creates ONE Excel file with all scans!
Level Filter: 1 (L1 + L2)
Output folder: output
Output file name: all_compliance_scans
```

**Combined Excel Output:**
- One file contains all .nessus file scans
- Each file's data separated by file name header (green background)
- Blank row between each report section
- All 10 columns for each scan
- Easy to review and compare across files

## Notes
- All compliance fields are parsed, including `cm:compliance-*` tags.
- Output `.docx`, `.csv`, and `.json` files are ignored by `.gitignore`.
- If you need a different layout or styling, update `Detailsreportgenratetor/details_report_docx.py`.


