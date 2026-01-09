# Tableau Workbook Error Fixer

A Python-based tool to automatically fix XML schema validation errors in Tableau workbook files (.twbx).

## Quick Start

### For Windows Users

**Option 1: Drag and Drop** (Easiest)
1. Download `fix_workbook.bat` from this repository
2. Drag your `.twbx` file onto the batch file
3. Done! A fixed version will be created with `_fixed` suffix

**Option 2: Command Line**
```powershell
python fix_tableau_workbook.py "C:\path\to\your\workbook.twbx"
```

**Option 3: PowerShell**
```powershell
.\fix_workbook.ps1 "C:\path\to\your\workbook.twbx"
```

### For Mac/Linux Users

```bash
python3 fix_tableau_workbook.py "/path/to/your/workbook.twbx"
```

## What Problems Does This Fix?

This tool resolves the following Tableau errors:

- ✅ **Error Code D2E8DA72** - XML schema validation failures
- ✅ **Invalid UUID formats** - e.g., `{PIE00001-0000-0000-0000-000000000001}`
- ✅ **Undeclared elements** - `parameters`, `mark-type`
- ✅ **Invalid attribute values** - `font-color` enumeration errors
- ✅ **Content model violations** - Elements in wrong locations

## Files in This Repository

- `fix_tableau_workbook.py` - Main Python script that performs the fixes
- `fix_workbook.bat` - Windows batch script for easy drag-and-drop usage
- `fix_workbook.ps1` - PowerShell script with enhanced features
- `TABLEAU_FIX_GUIDE.md` - Detailed documentation and troubleshooting

## Requirements

- Python 3.6 or higher (no additional packages needed)
- Your corrupted Tableau workbook file (.twbx)

## How It Works

1. Extracts the .twbx file (it's a ZIP archive)
2. Parses the XML to identify schema violations
3. Applies targeted fixes for each error type
4. Repackages into a new `_fixed.twbx` file
5. Your original file remains unchanged

## Example

```bash
# Before: Covid Dashboard - Enhanced V2.twbx (broken)
python fix_tableau_workbook.py "Covid Dashboard - Enhanced V2.twbx"
# After: Covid Dashboard - Enhanced V2_fixed.twbx (working!)
```

## Documentation

For detailed usage instructions, troubleshooting, and technical details, see [TABLEAU_FIX_GUIDE.md](TABLEAU_FIX_GUIDE.md).

## License

MIT License