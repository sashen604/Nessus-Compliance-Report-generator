# ✅ Combined Excel Export - Feature Complete

## What's New

### 🎯 Main Feature: One Excel File, Multiple Nessus Scans

You can now process **all .nessus files** from a folder and export them into **one consolidated Excel file** with proper separation and formatting.

---

## Usage

### Interactive Mode (Recommended)
```bash
python3 nessus-report-gen.py
```

**When prompted:**
1. Enter folder path with multiple `.nessus` files
2. Choose report type: `3` (Excel)
3. Choose output mode: `1` (Combined) ← **Creates ONE file**
4. Choose level filter: `1` (L1+L2) or `2` (L1 only)
5. Set output folder and filename

**Result:** One `.xlsx` file with all scans consolidated

---

## Excel Output Format

### File Structure
```
Combined_Report.xlsx
├── Sheet: "Combined Compliance"
│   ├── 📄 File: scan1.nessus (Green Header Row)
│   │   ├── IP | Check # | Level | Description | Result | Info | Solution | Impact | Policy | Value
│   │   ├── Data row 1
│   │   ├── Data row 2
│   │   └── Data row N
│   │
│   ├── (BLANK ROW - SEPARATOR)
│   │
│   ├── 📄 File: scan2.nessus (Green Header Row)
│   │   ├── IP | Check # | Level | Description | Result | Info | Solution | Impact | Policy | Value
│   │   ├── Data row 1
│   │   └── Data row M
│   │
│   └── (BLANK ROW)
```

### Styling
- **Green Headers (#70AD47):** File names and separators
- **Blue Headers (#B4C6E7):** Column headers per file
- **White Data:** Compliance details
- **Borders:** All cells have thin borders
- **Text Wrapping:** Enabled for readability

### Columns (10 Total per File)
1. **IP Address** - Host IP being scanned
2. **Check Number** - Compliance check ID
3. **Level** - L1 or L2 compliance level
4. **Description** - Check description
5. **Result** - Compliance result status
6. **Info** - Additional information
7. **Solution** - Recommended solution
8. **Impact** - Business impact
9. **Policy Value** - Expected value
10. **Configured Value** - Actual value

---

## Examples

### Example 1: Combine 3 Scan Reports
```bash
# Folder structure:
# scans/
#   ├── server-prod.nessus
#   ├── server-staging.nessus
#   └── server-dev.nessus

python3 nessus-report-gen.py
# Select: Excel (3), Combined (1)
# Result: all_scans.xlsx with 3 sections
```

### Example 2: Only L1 Findings
```bash
python3 nessus-report-gen.py
# Select: Excel (3), Combined (1), L1 Only (2)
# Result: Only critical L1 items in the report
```

### Example 3: Separate Files Instead
```bash
python3 nessus-report-gen.py
# Select: Excel (3), Separate (2)
# Result: 
#   ├── server-prod_detailed.xlsx
#   ├── server-staging_detailed.xlsx
#   └── server-dev_detailed.xlsx
```

---

## Features

✅ **Process Multiple Files**
- Scan entire folders with one command
- No need to process each file individually

✅ **Professional Formatting**
- Clear visual separation between files
- Color-coded headers
- Consistent styling throughout

✅ **Easy Navigation**
- Blank rows between sections for clarity
- File names clearly labeled
- Column headers repeated for each file

✅ **Filtering Support**
- L1 + L2 checks or L1 only
- Applied consistently across all files

✅ **Single Output File**
- All data in one convenient Excel file
- Easy to share and review
- Easier comparison across scans

---

## Technical Details

### New Functions Added

**`build_excel_combined(all_files_data, output_path)`**
- Combines multiple file scan data into one worksheet
- Parameters:
  - `all_files_data`: List of tuples `(filename, ip, items)`
  - `output_path`: Output .xlsx file path
- Handles formatting and row spacing automatically

### Modified Files
1. `details_report_excel.py` - Added combined function
2. `nessus-report-gen.py` - Updated main() to use combined Excel
3. `README.md` - Updated documentation

---

## Troubleshooting

**Q: My folder has 50 .nessus files, will it work?**
A: Yes! The tool processes all `.nessus` files in the folder.

**Q: Can I combine DOCX reports too?**
A: Yes! DOCX combined mode creates one document with page breaks between reports.

**Q: How do I filter by L1 only in combined mode?**
A: When prompted, choose `2` for Level Filter.

**Q: Can I still export separate Excel files?**
A: Yes! Choose `2` for Output Mode (Separate file per report).

---

## Quick Reference

| Feature | Command | Result |
|---------|---------|--------|
| Single file → Excel | `python3 nessus-report-gen.py` then `3` | One .xlsx |
| Folder → One Excel | `python3 nessus-report-gen.py` then `3`, `1` | Combined .xlsx |
| Folder → Separate Excels | `python3 nessus-report-gen.py` then `3`, `2` | Multiple .xlsx |
| Direct command | `python3 details_report_excel.py file.nessus` | One .xlsx |

---

## Version History

- **v2.0** - Added combined Excel export feature ✅
- **v1.5** - Initial Excel export support
- **v1.0** - DOCX report generation

---

Made with ❤️ for compliance reporting
