#!/bin/bash

cat << 'EOF'

╔════════════════════════════════════════════════════════════════╗
║     ✅ WINDOWS 7 COMPLIANCE REPORTS - QUICK REFERENCE        ║
╚════════════════════════════════════════════════════════════════╝

📍 LOCATION:
   windows7_compliance_reports.py

📊 WHAT IT DOES:
   Automatically detects Windows 7 compliance checks
   Generates 4 reports: 2 Word + 2 Excel (Summary + Details)

═══════════════════════════════════════════════════════════════════

🚀 QUICK START:

   Single Command (All 4 Reports):
   python3 windows7_compliance_reports.py "scan.nessus"

   Output:
   ✅ windows7_compliance_summary.docx
   ✅ windows7_compliance_summary.xlsx
   ✅ windows7_compliance_details.docx
   ✅ windows7_compliance_details.xlsx

═══════════════════════════════════════════════════════════════════

📋 COMMON COMMANDS:

1. All reports with custom name:
   python3 windows7_compliance_reports.py "scan.nessus" \
     --output "output/kandy_win7"

2. Excel only, summary only:
   python3 windows7_compliance_reports.py "scan.nessus" \
     --format excel --type summary

3. Word only, details only:
   python3 windows7_compliance_reports.py "scan.nessus" \
     --format word --type details

4. Custom path with all options:
   python3 windows7_compliance_reports.py "scan.nessus" \
     --output "reports/my_report" \
     --format both \
     --type both

═══════════════════════════════════════════════════════════════════

⚙️  COMMAND OPTIONS:

   --output TEXT
      Default: windows7_compliance
      Example: --output "output/server01"
      Result: Creates files with suffixes (_summary, _details)

   --format {both|word|excel}
      Default: both
      • both   → .docx AND .xlsx files
      • word   → .docx files only
      • excel  → .xlsx files only

   --type {both|summary|details}
      Default: both
      • both    → Summary AND Details reports
      • summary → Summary report only
      • details → Details report only

═══════════════════════════════════════════════════════════════════

📊 REPORT TYPES:

SUMMARY REPORTS:
  • Statistics (L1/L2/Total counts)
  • Compliance summary table
  • Quick overview format
  • Executive-friendly
  → Good for: Quick reviews, executive summaries

DETAILS REPORTS:
  • All compliance data
  • Info, solution, impact sections
  • Policy vs configured values
  • Complete information
  → Good for: Audits, detailed reviews, remediation

═══════════════════════════════════════════════════════════════════

📁 OUTPUT FILES:

Created files follow this pattern:
  [PREFIX]_summary.docx    ← Word summary
  [PREFIX]_summary.xlsx    ← Excel summary
  [PREFIX]_details.docx    ← Word details
  [PREFIX]_details.xlsx    ← Excel details

Default prefix: windows7_compliance
Custom prefix: Use --output flag

Example output structure:
  output/
  ├── kandy_win7_summary.docx
  ├── kandy_win7_summary.xlsx
  ├── kandy_win7_details.docx
  └── kandy_win7_details.xlsx

═══════════════════════════════════════════════════════════════════

🎨 FORMATTING:

WORD REPORTS:
  • Professional document formatting
  • Tables with colored headers
  • Section-based layout
  • Easy to print and share
  • File size: ~37 KB

EXCEL REPORTS:
  • Data-friendly format
  • Color-coded headers (blue)
  • Borders on all cells
  • Text wrapping enabled
  • Frozen header row
  • File size: ~5-6 KB

═══════════════════════════════════════════════════════════════════

📋 REPORT CONTENTS:

SUMMARY CONTENT:
  • Host IP address
  • L1 check count
  • L2 check count
  • Total check count
  • Table: Check ID | Level | Description | Status | Benchmark

DETAILS CONTENT:
  • Host IP address
  • All statistics
  • For each check:
    ✓ Check description
    ✓ Compliance result
    ✓ Information
    ✓ Recommended solution
    ✓ Business impact
    ✓ Policy value
    ✓ Configured value

═══════════════════════════════════════════════════════════════════

💼 USE CASES:

1. Executive Report:
   python3 windows7_compliance_reports.py scan.nessus \
     --format excel --type summary

2. Compliance Audit:
   python3 windows7_compliance_reports.py scan.nessus \
     --format word --type details \
     --output "audit/detailed_audit"

3. Both Formats (All Data):
   python3 windows7_compliance_reports.py scan.nessus \
     --format both --type both

4. Batch Multiple Servers:
   for f in scans/*.nessus; do
     python3 windows7_compliance_reports.py "$f" \
       --output "reports/$(basename $f .nessus)"
   done

═══════════════════════════════════════════════════════════════════

✨ FEATURES:

✅ Auto-detect Windows 7 checks
✅ Filter non-Windows-7 items
✅ Generate 4 files in one command
✅ Professional formatting
✅ Customizable output
✅ Multiple format options
✅ Statistics included
✅ Host tracking
✅ Full compliance details
✅ Easy to share and print

═══════════════════════════════════════════════════════════════════

❓ QUICK HELP:

Q: How many files are created?
A: By default 4 files (2 Word + 2 Excel)
   Use --format and --type flags to change

Q: Can I use only Word or only Excel?
A: Yes! Use --format word  or  --format excel

Q: Can I get only summary or only details?
A: Yes! Use --type summary  or  --type details

Q: How do I organize outputs by server?
A: Use --output "path/server_name"
   Example: --output "reports/server01"

Q: Will it work with non-Windows-7 scans?
A: It will process them but may find 0 checks
   It only extracts Windows 7 compliance items

═══════════════════════════════════════════════════════════════════

📞 HELP COMMAND:

   python3 windows7_compliance_reports.py --help

═══════════════════════════════════════════════════════════════════

Made with ❤️ for Windows 7 compliance teams

EOF
