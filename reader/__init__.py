"""
Reads an .xlsx file using only Python stdlib — no external libraries.

An xlsx file is a zip archive. The files we care about inside it:
  xl/workbook.xml              — lists sheet names and their relationship IDs
  xl/_rels/workbook.xml.rels   — maps relationship IDs to actual file paths
  xl/sharedStrings.xml         — all string values (cells store an index, not the string itself)
  xl/worksheets/sheet1.xml     — cell grid for sheet 1 (one file per sheet)

Returns:
  {
    "Layout": {(row, col): "NodeName"},     row/col are 1-based integers
    "Edges":  [["A", ">", "B", "label"]]   one list per row, columns A-D
  }
"""

import zipfile
import xml.etree.ElementTree as ET

# Every element in an xlsx XML file is wrapped in this namespace.
# We include it in every tag lookup so ElementTree can find them.
_XLSX_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

# The namespace used for relationship IDs inside workbook.xml
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# The namespace used inside the .rels file itself
_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _tag(name):
    """Prefix a tag name with the xlsx namespace so ET.find() works."""
    return f"{{{_XLSX_NS}}}{name}"


def _col_letters_to_num(letters):
    """
    Convert Excel column letters to a 1-based integer.
    A=1, Z=26, AA=27, AZ=52, BA=53, ...
    Works like base-26 with no zero digit.
    """
    num = 0
    for ch in letters.upper():
        num = num * 26 + (ord(ch) - ord('A') + 1)
    return num


def _parse_cell_ref(ref):
    """
    Split a cell reference like 'B3' into (row=3, col=2).
    Both are 1-based. Col letters come first, then row digits.
    """
    col_letters = ""
    row_digits  = ""
    for ch in ref:
        if ch.isalpha():
            col_letters += ch
        else:
            row_digits += ch
    return int(row_digits), _col_letters_to_num(col_letters)


def _read_shared_strings(zf):
    """
    Parse xl/sharedStrings.xml into a list indexed by position.
    When a cell has type 's', its value is an index into this list.
    Rich-text cells (<r><t> runs) are concatenated into a single string.
    """
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []  # some xlsx files have no strings at all

    root = ET.parse(zf.open("xl/sharedStrings.xml")).getroot()
    strings = []

    for si in root.findall(_tag("si")):
        t_direct = si.find(_tag("t"))
        if t_direct is not None:
            # Simple string: <si><t>value</t></si>
            strings.append(t_direct.text or "")
        else:
            # Rich text: <si><r><t>part1</t></r><r><t>part2</t></r></si>
            runs = [r.find(_tag("t")) for r in si.findall(_tag("r"))]
            strings.append("".join(t.text or "" for t in runs if t is not None))

    return strings


def _read_sheet_paths(zf):
    """
    Build a dict: {sheet_name: path_inside_zip}.

    Two-step process:
    1. workbook.xml maps sheet name → relationship ID (rId)
    2. workbook.xml.rels maps rId → file path (relative to xl/)
    """
    # Step 1: name → rId
    wb_root = ET.parse(zf.open("xl/workbook.xml")).getroot()
    sheets_el = wb_root.find(_tag("sheets"))

    rid_to_name = {}
    for sheet in sheets_el.findall(_tag("sheet")):
        rid  = sheet.get(f"{{{_REL_NS}}}id")
        name = sheet.get("name")
        rid_to_name[rid] = name

    # Step 2: rId → file path
    rels_root = ET.parse(zf.open("xl/_rels/workbook.xml.rels")).getroot()

    name_to_path = {}
    for rel in rels_root.findall(f"{{{_RELS_NS}}}Relationship"):
        rid    = rel.get("Id")
        target = rel.get("Target")  # e.g. "worksheets/sheet1.xml"
        if rid in rid_to_name:
            # Target is relative to xl/ — prepend it
            path = f"xl/{target}" if not target.startswith("/") else target.lstrip("/")
            name_to_path[rid_to_name[rid]] = path

    return name_to_path


def _read_cells(zf, sheet_path, shared_strings):
    """
    Parse one worksheet and return all non-empty cells as {(row, col): value_str}.

    Cell types we handle:
      t="s"         — shared string: value is an index into shared_strings
      t="inlineStr" — inline string: value is inside <is><t>
      (default)     — number or formula result: value is inside <v> as text
    """
    root = ET.parse(zf.open(sheet_path)).getroot()
    sheet_data = root.find(_tag("sheetData"))
    if sheet_data is None:
        return {}

    cells = {}
    for row_el in sheet_data.findall(_tag("row")):
        for c_el in row_el.findall(_tag("c")):
            ref  = c_el.get("r")        # cell address, e.g. "A1"
            typ  = c_el.get("t", "n")   # "s", "inlineStr", or "n" (number/default)
            v_el = c_el.find(_tag("v")) # <v> holds the raw value or string index
            is_el = c_el.find(_tag("is")) # <is> holds inline string content

            if typ == "s" and v_el is not None:
                idx   = int(v_el.text)
                value = shared_strings[idx] if idx < len(shared_strings) else ""
            elif typ == "inlineStr" and is_el is not None:
                t = is_el.find(_tag("t"))
                value = (t.text or "") if t is not None else ""
            elif v_el is not None:
                value = v_el.text or ""
            else:
                continue  # truly empty cell

            value = value.strip()
            if value:  # skip whitespace-only cells
                row, col = _parse_cell_ref(ref)
                cells[(row, col)] = value

    return cells


def _cells_to_edge_rows(cells):
    """
    Convert a flat {(row,col): value} dict from the Edges sheet into
    a list of [source, separator, target, label] rows.

    Expected column layout (1-based):
      col 1 = Source
      col 2 = Separator
      col 3 = Target
      col 4 = Label (optional)

    Row 1 is assumed to be a header and is skipped.
    """
    rows_by_index = {}  # row number → {col: value}
    for (row, col), value in cells.items():
        if row == 1:
            continue  # skip header
        rows_by_index.setdefault(row, {})[col] = value

    edge_rows = []
    for row_idx in sorted(rows_by_index):
        r = rows_by_index[row_idx]
        edge_rows.append([
            r.get(1, ""),  # source
            r.get(2, ""),  # separator
            r.get(3, ""),  # target
            r.get(4, ""),  # label (optional, may be absent)
        ])

    return edge_rows


def _cells_to_style_rows(cells):
    """
    Convert a flat {(row,col): value} dict from the Style sheet into
    a list of [group, width, height, style_string] rows.

    Expected column layout (1-based):
      col 1 = Group name
      col 2 = Width  (pixels)
      col 3 = Height (pixels)
      col 4 = draw.io style string

    Row 1 is assumed to be a header and is skipped.
    """
    rows_by_index = {}
    for (row, col), value in cells.items():
        if row == 1:
            continue  # skip header
        rows_by_index.setdefault(row, {})[col] = value

    style_rows = []
    for row_idx in sorted(rows_by_index):
        r = rows_by_index[row_idx]
        style_rows.append([
            r.get(1, ""),   # group name
            r.get(2, ""),   # width
            r.get(3, ""),   # height
            r.get(4, ""),   # style string
        ])

    return style_rows


def read_xlsx(path):
    """
    Read an xlsx file and return structured sheet data.

    Returns:
      {
        "Layout": {(row, col): cell_value},   cell_value may be "NodeName" or "NodeName|Group"
        "Edges":  [[source, sep, target, label], ...],
        "Style":  [[group, width, height, style_string], ...]
      }
    Sheets with other names are ignored.
    Missing sheets produce empty dicts/lists.
    """
    with zipfile.ZipFile(path) as zf:
        shared_strings = _read_shared_strings(zf)
        sheet_paths    = _read_sheet_paths(zf)

        result = {"Layout": {}, "Edges": [], "Style": []}

        for sheet_name, sheet_path in sheet_paths.items():
            cells = _read_cells(zf, sheet_path, shared_strings)

            if sheet_name == "Layout":
                result["Layout"] = cells

            elif sheet_name == "Edges":
                result["Edges"] = _cells_to_edge_rows(cells)

            elif sheet_name == "Style":
                result["Style"] = _cells_to_style_rows(cells)

    return result
