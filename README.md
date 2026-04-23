# Nessus Compliance & Summary Reports

Generate terminal summaries, CSV/JSON exports, and detailed Word reports from Nessus `.nessus` / `.xml` files.

## Contents
- `nessus_summary.py` – CLI summary tables + CSV/JSON export
- `compliance_table_docx.py` – simple compliance table Word export
- `Detailsreportgenratetor/details_report_docx.py` – detailed compliance Word report
- `generate_compliance_docs.sh` – batch simple compliance docx
- `Detailsreportgenratetor/generate_detailed_reports.sh` – batch detailed docx
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
"./.venv/bin/python" Detailsreportgenratetor/details_report_docx.py "/path/to/scan.nessus" --output "/path/to/detailed_report.docx"
```
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

## Notes
- All compliance fields are parsed, including `cm:compliance-*` tags.
- Output `.docx`, `.csv`, and `.json` files are ignored by `.gitignore`.
- If you need a different layout or styling, update `Detailsreportgenratetor/details_report_docx.py`.


