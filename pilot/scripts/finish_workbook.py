"""Set the native hidden-sheet flag not exposed by artifact-tool's documented API.

All cells, formulas, tables and formatting are authored with artifact-tool.
This pass changes only the Tracking sheet start tag in xl/workbook.xml.
"""
import os
import copy
from pathlib import Path
import re
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET


def hide_tracking(path):
    path = Path(path)
    with zipfile.ZipFile(path) as source:
        xml = source.read('xl/workbook.xml').decode('utf-8')
        changed = []

        def hide(match):
            tag = match.group(0)
            if not re.search(r'\bname=([\"\x27])Tracking\1', tag):
                return tag
            tag = re.sub(r'\sstate=([\"\x27]).*?\1', '', tag)
            changed.append(True)
            return re.sub(r'(/?>)$', r' state="hidden"\1', tag)

        updated = re.sub(r'<(?:[\w.-]+:)?sheet\b[^>]*>', hide, xml)
        if len(changed) != 1:
            raise ValueError('Expected exactly one Tracking sheet to hide')
        parsed = ET.fromstring(updated)
        ns = {'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        tracking = [s for s in parsed.find('s:sheets', ns) if s.get('name') == 'Tracking']
        if tracking[0].get('state') != 'hidden':
            raise AssertionError('Tracking visibility did not update')
        handle, temporary = tempfile.mkstemp(prefix='finish-', suffix='.xlsx', dir=path.parent)
        os.close(handle)
        try:
            with zipfile.ZipFile(temporary, 'w') as target:
                for entry in source.infolist():
                    content = updated.encode('utf-8') if entry.filename == 'xl/workbook.xml' else source.read(entry.filename)
                    target.writestr(copy.copy(entry), content)
            with zipfile.ZipFile(temporary) as check:
                for entry in source.infolist():
                    if entry.filename != 'xl/workbook.xml' and check.read(entry.filename) != source.read(entry.filename):
                        raise AssertionError('Unexpected workbook-part change')
        except BaseException:
            os.unlink(temporary)
            raise
    os.replace(temporary, path)


if __name__ == '__main__':
    hide_tracking(sys.argv[1])
    print('Hidden device register verified; all other workbook parts preserved.')
