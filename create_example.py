"""
create_example.py — generates inputs/smart_home.xlsx

Run once to create the example input file:
  py create_example.py

Then convert it:
  py topo2visio.py smart_home.xlsx
"""

import zipfile
from pathlib import Path

OUTPUT_PATH = Path(__file__).parent / "inputs" / "smart_home.xlsx"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _col_letter(n):
    """1 → 'A', 2 → 'B', 26 → 'Z', 27 → 'AA' ..."""
    result = ""
    while n:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result


def _cell(col, row, value):
    """One inline-string cell XML fragment."""
    ref = f"{_col_letter(col)}{row}"
    # Escape XML special characters in cell values
    value = value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return f'<c r="{ref}" t="inlineStr"><is><t>{value}</t></is></c>'


def _sheet(data):
    """
    Build a worksheet XML string from a dict: {row: {col: value}}.
    Rows and cols are 1-based integers.
    """
    rows_xml = ""
    for row_num in sorted(data):
        cols = data[row_num]
        cells = "".join(_cell(col, row_num, cols[col]) for col in sorted(cols))
        rows_xml += f'<row r="{row_num}">{cells}</row>'

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData>{rows_xml}</sheetData>'
        '</worksheet>'
    )


# ---------------------------------------------------------------------------
# Sheet definitions
# ---------------------------------------------------------------------------

# Layout: {row: {col: "NodeName|group"}}
# Grid layout — imagine this as the final diagram shape:
#
#   col1               col3              col5
#   TempSensor ---\
#   MotionSensor --+-- Hub ------------ Lights
#   LightSensor --/    |                Thermostat
#                      |                Alarm
#                      Gateway
#                      |
#                      CloudServer
#                      |
#                      MobileApp
#
LAYOUT = {
    1: {1: "TempSensor|sensor",    3: "Hub|controller",   5: "Lights|output"},
    2: {1: "MotionSensor|sensor",                         5: "Thermostat|output"},
    3: {1: "LightSensor|sensor",                          5: "Alarm|output"},
    5: {                           3: "Gateway|controller"},
    7: {                           3: "CloudServer|cloud"},
    9: {                           3: "MobileApp|cloud"},
}

# Edges: rows of [source, separator, target, label]
# Row 1 = header, data starts at row 2
EDGES = {
    1: {1: "Source",      2: "Sep", 3: "Target",      4: "Label"},
    2: {1: "TempSensor",  2: ">",   3: "Hub",          4: "temp"},
    3: {1: "MotionSensor",2: ">",   3: "Hub",          4: "motion"},
    4: {1: "LightSensor", 2: ">",   3: "Hub",          4: "lux"},
    5: {1: "Hub",         2: ">",   3: "Lights",       4: "on/off"},
    6: {1: "Hub",         2: ">",   3: "Thermostat",   4: "setpoint"},
    7: {1: "Hub",         2: ">",   3: "Alarm",        4: "trigger"},
    8: {1: "Hub",         2: "<>",  3: "Gateway",      4: "sync"},
    9: {1: "Gateway",     2: "<>",  3: "CloudServer",  4: "MQTT"},
   10: {1: "CloudServer", 2: "<>",  3: "MobileApp",    4: "REST"},
}

# Style: rows of [group, width, height, draw.io style string]
# Row 1 = header, data starts at row 2
STYLES = {
    1: {1: "Group",       2: "Width", 3: "Height", 4: "Style"},
    2: {1: "sensor",      2: "80",    3: "35",
        4: "ellipse;whiteSpace=wrap;html=1;fillColor=#dae8fc;strokeColor=#6c8ebf;"},
    3: {1: "controller",  2: "120",   3: "50",
        4: "rounded=1;whiteSpace=wrap;html=1;fillColor=#fff2cc;strokeColor=#d6b656;fontStyle=1;fontSize=11;"},
    4: {1: "output",      2: "90",    3: "40",
        4: "rounded=1;whiteSpace=wrap;html=1;fillColor=#d5e8d4;strokeColor=#82b366;"},
    5: {1: "cloud",       2: "120",   3: "50",
        4: "rounded=1;whiteSpace=wrap;html=1;fillColor=#e1d5e7;strokeColor=#9673a6;fontStyle=1;fontSize=11;"},
}


# ---------------------------------------------------------------------------
# xlsx assembly
# ---------------------------------------------------------------------------

CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml"  ContentType="application/xml"/>'
    '<Override PartName="/xl/workbook.xml"'
    ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    '<Override PartName="/xl/worksheets/sheet1.xml"'
    ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '<Override PartName="/xl/worksheets/sheet2.xml"'
    ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '<Override PartName="/xl/worksheets/sheet3.xml"'
    ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '</Types>'
)

ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1"'
    ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
    ' Target="xl/workbook.xml"/>'
    '</Relationships>'
)

WORKBOOK = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<sheets>'
    '<sheet name="Layout" sheetId="1" r:id="rId1"/>'
    '<sheet name="Edges"  sheetId="2" r:id="rId2"/>'
    '<sheet name="Style"  sheetId="3" r:id="rId3"/>'
    '</sheets>'
    '</workbook>'
)

WORKBOOK_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1"'
    ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"'
    ' Target="worksheets/sheet1.xml"/>'
    '<Relationship Id="rId2"'
    ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"'
    ' Target="worksheets/sheet2.xml"/>'
    '<Relationship Id="rId3"'
    ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"'
    ' Target="worksheets/sheet3.xml"/>'
    '</Relationships>'
)


def build_xlsx(path):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml",          CONTENT_TYPES)
        zf.writestr("_rels/.rels",                  ROOT_RELS)
        zf.writestr("xl/workbook.xml",              WORKBOOK)
        zf.writestr("xl/_rels/workbook.xml.rels",   WORKBOOK_RELS)
        zf.writestr("xl/worksheets/sheet1.xml",     _sheet(LAYOUT))
        zf.writestr("xl/worksheets/sheet2.xml",     _sheet(EDGES))
        zf.writestr("xl/worksheets/sheet3.xml",     _sheet(STYLES))


if __name__ == "__main__":
    OUTPUT_PATH.parent.mkdir(exist_ok=True)
    build_xlsx(OUTPUT_PATH)
    print(f"Created: {OUTPUT_PATH}")
    print("Now run:  py topo2visio.py smart_home.xlsx")
