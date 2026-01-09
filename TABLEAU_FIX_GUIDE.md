# Tableau Workbook Error Fixer

This repository contains a tool to fix XML schema validation errors in Tableau workbook files (.twbx).

## Problem

When opening a Tableau workbook, you may encounter errors like:

```
Error Code: D2E8DA72
Error(2741,15): no declaration found for element 'parameters'
Error(3186,66): value '{PIE00001-0000-0000-0000-000000000001}' does not match regular expression facet
Error(3237,36): no declaration found for element 'mark-type'
Error(3960,55): value 'font-color' not in enumeration
Error(4295,12): element 'parameters' is not allowed for content model
```

## Solution

This script automatically fixes:
- ✓ Invalid UUID/GUID formats (e.g., `{PIE00001-...}` → valid UUID)
- ✓ Undeclared `parameters` element in wrong location
- ✓ Invalid `mark-type` elements
- ✓ Invalid `font-color` enumeration values
- ✓ Content model violations

## Usage

### On Windows

1. **Download your workbook** to a known location (e.g., Downloads folder)

2. **Open Command Prompt or PowerShell**

3. **Run the fix script:**

   ```powershell
   python fix_tableau_workbook.py "C:\Users\yoga sundaram\Downloads\Covid Dashboard - Enhanced V2.twbx"
   ```

4. **Open the fixed workbook:**
   - The script creates a new file: `Covid Dashboard - Enhanced V2_fixed.twbx`
   - Open this file in Tableau Desktop

### On Linux/Mac

```bash
python3 fix_tableau_workbook.py "/path/to/your/workbook.twbx"
```

## How It Works

1. **Extracts** the .twbx file (which is a ZIP archive containing XML)
2. **Parses** the XML content to identify schema violations
3. **Fixes** each type of error:
   - Replaces invalid UUIDs with properly formatted ones
   - Removes improperly placed elements
   - Corrects enumeration values
4. **Repackages** the workbook into a new .twbx file
5. **Preserves** all dashboards, data sources, and visualizations

## Requirements

- Python 3.6 or higher (comes with Windows 10/11, macOS, and most Linux distributions)
- No additional packages required (uses only standard library)

## Troubleshooting

### "Python is not recognized"

**Windows:** Install Python from [python.org](https://www.python.org/downloads/)
- Make sure to check "Add Python to PATH" during installation

### "File not found"

Make sure to:
- Use quotes around the file path if it contains spaces
- Provide the full path to the workbook file

### Still getting errors after fix?

The script creates a backup by generating a new file with `_fixed` suffix. If issues persist:
1. Check the console output for any error messages
2. Ensure the original file is a valid .twbx file
3. Try opening the fixed file in the latest version of Tableau Desktop

## For Your Specific Error

Based on your error message, run:

```powershell
python fix_tableau_workbook.py "C:\Users\yoga sundaram\Downloads\Covid Dashboard - Enhanced V2.twbx"
```

This will create:
```
C:\Users\yoga sundaram\Downloads\Covid Dashboard - Enhanced V2_fixed.twbx
```

## What Gets Fixed

### Invalid UUIDs
```xml
<!-- Before -->
<worksheet name='Pie' uuid='{PIE00001-0000-0000-0000-000000000001}'>

<!-- After -->
<worksheet name='Pie' uuid='{A1B2C3D4-1234-5678-9ABC-DEF012345678}'>
```

### Invalid Elements
```xml
<!-- Before -->
<mark-type type='pie'/>

<!-- After -->
(removed - not part of Tableau schema)
```

### Misplaced Parameters
```xml
<!-- Before -->
<workbook>
  ...
  <parameters>...</parameters>  <!-- Wrong location -->
</workbook>

<!-- After -->
<workbook>
  ...
  <!-- parameters moved/removed -->
</workbook>
```

## License

MIT License - Feel free to use and modify as needed.
