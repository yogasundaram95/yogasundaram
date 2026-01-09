#!/usr/bin/env python3
"""
Tableau Workbook XML Validator and Fixer

This script fixes common XML schema validation errors in Tableau workbook (.twbx) files.
It handles:
- Invalid UUID/GUID formats
- Undeclared elements (parameters, mark-type)
- Invalid attribute values
- Content model violations
"""

import zipfile
import xml.etree.ElementTree as ET
import re
import uuid
import sys
import os
from pathlib import Path
import shutil
from typing import Dict, Set


class TableauWorkbookFixer:
    def __init__(self, workbook_path: str):
        self.workbook_path = Path(workbook_path)
        self.temp_dir = None
        self.twb_file = None

    def is_valid_uuid(self, guid_str: str) -> bool:
        """Check if a string is a valid UUID format"""
        pattern = r'^\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}$'
        return bool(re.match(pattern, guid_str))

    def generate_valid_uuid(self) -> str:
        """Generate a valid UUID in Tableau format"""
        return '{' + str(uuid.uuid4()).upper() + '}'

    def fix_invalid_uuids(self, content: str) -> str:
        """Fix invalid UUID formats in the XML content"""
        # Pattern to find invalid UUIDs like {PIE00001-0000-0000-0000-000000000001}
        invalid_uuid_pattern = r'\{[A-Z0-9]+-[0-9]+-[0-9]+-[0-9]+-[0-9]+\}'

        # Keep track of replacements to maintain consistency
        uuid_map: Dict[str, str] = {}

        def replace_uuid(match):
            invalid_uuid = match.group(0)
            if invalid_uuid not in uuid_map:
                uuid_map[invalid_uuid] = self.generate_valid_uuid()
            return uuid_map[invalid_uuid]

        fixed_content = re.sub(invalid_uuid_pattern, replace_uuid, content)
        return fixed_content

    def fix_mark_type_element(self, content: str) -> str:
        """Remove or fix invalid mark-type elements"""
        # Remove mark-type elements that are not part of the schema
        # Pattern to match <mark-type type="..."/> or <mark-type type="...">...</mark-type>
        content = re.sub(r'<mark-type[^>]*/>[\s]*', '', content)
        content = re.sub(r'<mark-type[^>]*>.*?</mark-type>[\s]*', '', content, flags=re.DOTALL)
        return content

    def fix_parameters_element(self, content: str) -> str:
        """Remove parameters element from invalid location"""
        # The parameters element is appearing in the wrong place in the XML hierarchy
        # We need to remove it from where it's not allowed

        # Find and remove parameters element that's in the wrong location
        # (at the root level of workbook instead of inside datasource)
        lines = content.split('\n')
        fixed_lines = []
        skip_until_close = False
        in_workbook_params = False
        depth = 0

        for i, line in enumerate(lines):
            # Check if we're at the root workbook parameters (line 4295 in error)
            if '<parameters' in line and not skip_until_close:
                # Look back to see if we're inside a datasource element
                context_lines = '\n'.join(fixed_lines[-50:]) if len(fixed_lines) > 50 else '\n'.join(fixed_lines)

                # If we're not in a datasource context, this is invalid
                if '<datasource' not in context_lines or '</datasource>' in context_lines.split('<datasource')[-1]:
                    skip_until_close = True
                    in_workbook_params = True
                    if '<parameters/>' in line or '/>' in line:
                        skip_until_close = False
                        continue
                    continue

            if skip_until_close:
                if '</parameters>' in line:
                    skip_until_close = False
                continue

            fixed_lines.append(line)

        return '\n'.join(fixed_lines)

    def fix_font_color_value(self, content: str) -> str:
        """Fix invalid font-color enumeration value"""
        # Replace font-color with a valid value (likely should be 'color' or removed)
        content = re.sub(r"<format[^>]*value='font-color'[^>]*>",
                        lambda m: m.group(0).replace("'font-color'", "'bold'"),
                        content)
        content = re.sub(r'<format[^>]*value="font-color"[^>]*>',
                        lambda m: m.group(0).replace('"font-color"', '"bold"'),
                        content)
        return content

    def extract_workbook(self) -> str:
        """Extract .twbx file and return path to .twb file"""
        if not self.workbook_path.exists():
            raise FileNotFoundError(f"Workbook not found: {self.workbook_path}")

        # Create temp directory
        self.temp_dir = Path(f"{self.workbook_path.stem}_temp")
        if self.temp_dir.exists():
            shutil.rmtree(self.temp_dir)
        self.temp_dir.mkdir()

        # Extract .twbx (it's a zip file)
        with zipfile.ZipFile(self.workbook_path, 'r') as zip_ref:
            zip_ref.extractall(self.temp_dir)

        # Find the .twb file
        twb_files = list(self.temp_dir.glob("*.twb"))
        if not twb_files:
            raise FileNotFoundError("No .twb file found in the workbook")

        self.twb_file = twb_files[0]
        return str(self.twb_file)

    def fix_workbook(self):
        """Main method to fix the workbook"""
        print(f"Processing: {self.workbook_path}")

        # Extract the workbook
        print("Extracting workbook...")
        twb_path = self.extract_workbook()

        # Read the XML content
        print("Reading XML content...")
        with open(twb_path, 'r', encoding='utf-8') as f:
            content = f.read()

        original_size = len(content)

        # Apply fixes
        print("Applying fixes...")
        print("  - Fixing invalid UUIDs...")
        content = self.fix_invalid_uuids(content)

        print("  - Removing invalid mark-type elements...")
        content = self.fix_mark_type_element(content)

        print("  - Fixing parameters element placement...")
        content = self.fix_parameters_element(content)

        print("  - Fixing font-color enumeration...")
        content = self.fix_font_color_value(content)

        # Write the fixed content back
        print("Writing fixed XML...")
        with open(twb_path, 'w', encoding='utf-8') as f:
            f.write(content)

        # Repackage the workbook
        output_path = self.workbook_path.parent / f"{self.workbook_path.stem}_fixed.twbx"
        print(f"Repackaging workbook to: {output_path}")

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path in self.temp_dir.rglob('*'):
                if file_path.is_file():
                    arcname = file_path.relative_to(self.temp_dir)
                    zipf.write(file_path, arcname)

        # Cleanup
        print("Cleaning up temporary files...")
        shutil.rmtree(self.temp_dir)

        print(f"\n✓ Fixed workbook saved to: {output_path}")
        print(f"  Original size: {original_size:,} bytes")
        print(f"  Fixed size: {len(content):,} bytes")

        return output_path


def main():
    if len(sys.argv) != 2:
        print("Usage: python fix_tableau_workbook.py <path_to_workbook.twbx>")
        print("\nExample:")
        print('  python fix_tableau_workbook.py "C:\\Users\\yoga sundaram\\Downloads\\Covid Dashboard - Enhanced V2.twbx"')
        sys.exit(1)

    workbook_path = sys.argv[1]

    try:
        fixer = TableauWorkbookFixer(workbook_path)
        fixed_path = fixer.fix_workbook()
        print(f"\n✓ Success! You can now open: {fixed_path}")
        return 0
    except Exception as e:
        print(f"\n✗ Error: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
