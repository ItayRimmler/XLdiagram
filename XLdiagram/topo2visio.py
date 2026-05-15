"""
topo2visio — convert an Excel diagram spec into draw.io XML.

Input:  one .xlsx file with two sheets:
  Layout  — node name in a cell = that node's position on the canvas
  Edges   — rows of: Source | Separator | Target | Label

Output: one .drawio file, openable in draw.io, exportable to Visio.

Usage:
  py topo2visio.py mydiagram.xlsx
  (input is read from ./inputs/, output goes to ./outputs/)

Separator reference:
  >    forward arrow
  <    backward arrow
  o>   dot at source, arrow at target
  <o   arrow at source, dot at target
  <>   bidirectional
  -    plain wire (no arrows)
"""

import sys
import argparse
import xml.etree.ElementTree as ET
from pathlib import Path

# Anchor all imports to this script's directory so the program works
# regardless of where the user runs it from (any working directory).
sys.path.insert(0, str(Path(__file__).parent))

from reader  import read_xlsx
from parser  import parse_layout, parse_edges
from builder import build_xml

SCRIPT_DIR  = Path(__file__).parent
INPUTS_DIR  = SCRIPT_DIR / "inputs"   # drop .xlsx files here
OUTPUTS_DIR = SCRIPT_DIR / "outputs"  # .drawio files appear here

# Pixels per Excel grid cell.
# col 1 → x=MARGIN, col 2 → x=MARGIN+SCALE, col 3 → x=MARGIN+2*SCALE ...
# Increase to spread nodes further apart. Decrease to pack tighter.
SCALE = 120

# Pixel offset applied to all nodes so the diagram doesn't start at the canvas corner.
# Increase if nodes are getting clipped at the top-left edge.
MARGIN = 100


def main():
    parser = argparse.ArgumentParser(
        description="Convert an Excel diagram spec (.xlsx) into a draw.io file (.drawio)"
    )
    parser.add_argument("input", help="Input .xlsx filename (looked up inside ./inputs/)")
    args = parser.parse_args()

    input_path  = INPUTS_DIR  / args.input
    output_path = OUTPUTS_DIR / Path(args.input).with_suffix(".drawio").name

    if not input_path.exists():
        print(f"Error: {input_path} not found")
        sys.exit(1)

    # --- Step 1: read raw cell data from the xlsx file ---
    sheets = read_xlsx(input_path)

    # --- Step 2: convert raw data into graph structures ---
    # positions: {node_name: (x_pixels, y_pixels)}  — empty if no Layout sheet
    # edges:     list of edge dicts with source, target, label, style
    positions = parse_layout(sheets.get("Layout", {}), scale=SCALE, margin=MARGIN)
    edges     = parse_edges(sheets.get("Edges",   []))

    # --- Step 3: assemble draw.io XML ---
    xml_model = build_xml(positions, edges)

    # --- Step 4: write output file ---
    tree = ET.ElementTree(xml_model)
    ET.indent(tree, space="  ")  # pretty-print with indentation
    tree.write(output_path, encoding="unicode", xml_declaration=False)

    print(f"Written: {output_path}")
    print(f"  {len(positions)} positioned nodes, {len(edges)} edges")


if __name__ == "__main__":
    main()
