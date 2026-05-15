"""
topo2visio — convert an Excel diagram spec into draw.io XML.

Input:  one .xlsx file with up to three sheets:
  Layout  — node name (optionally "Name|Group") in a cell = position on canvas
  Edges   — rows of: Source | Separator | Target | Label
  Style   — rows of: Group | Width | Height | draw.io style string  (optional)

Output: one .drawio file, openable in draw.io, exportable to Visio.

Usage:
  py topo2visio.py mydiagram.xlsx
  py topo2visio.py mydiagram.xlsx --scale 150 --margin 50

Separator reference:
  >    forward arrow
  <    backward arrow
  o>   dot at source, arrow at target
  <o   arrow at source, dot at target
  <>   bidirectional
  -    plain wire (no arrows)
  o-   dot at source, plain line to target
  -o   plain line from source, dot at target

Phantom nodes:
  Prefix a node name with "." in the Layout sheet (e.g. ".p1") to create
  an invisible routing point. Edges can pass through it without a visible box.
"""

import sys
import argparse
import xml.etree.ElementTree as ET
from pathlib import Path

# Anchor all imports to this script's directory so the program works
# regardless of where the user runs it from (any working directory).
sys.path.insert(0, str(Path(__file__).parent))

from reader  import read_xlsx
from parser  import parse_layout, parse_edges, parse_styles
from builder import build_xml, NODE_W, NODE_H, VERTEX_STYLE

SCRIPT_DIR  = Path(__file__).parent
INPUTS_DIR  = SCRIPT_DIR / "inputs"   # drop .xlsx files here
OUTPUTS_DIR = SCRIPT_DIR / "outputs"  # .drawio files appear here

# Default values — overridable via CLI flags
DEFAULT_SCALE  = 120   # pixels per Excel grid cell
DEFAULT_MARGIN = 100   # pixel offset from canvas edge so nodes aren't clipped


def main():
    parser = argparse.ArgumentParser(
        description="Convert an Excel diagram spec (.xlsx) into a draw.io file (.drawio)"
    )
    parser.add_argument("input",
        help="Input .xlsx filename (looked up inside ./inputs/)")
    parser.add_argument("--scale", type=int, default=DEFAULT_SCALE,
        help=f"Pixels per Excel grid cell (default: {DEFAULT_SCALE})")
    parser.add_argument("--margin", type=int, default=DEFAULT_MARGIN,
        help=f"Pixel offset from canvas edge (default: {DEFAULT_MARGIN})")
    args = parser.parse_args()

    input_path  = INPUTS_DIR  / args.input
    output_path = OUTPUTS_DIR / Path(args.input).with_suffix(".drawio").name

    if not input_path.exists():
        print(f"Error: {input_path} not found")
        sys.exit(1)

    # --- Step 1: read raw cell data from the xlsx file ---
    sheets = read_xlsx(input_path)

    # --- Step 2: convert raw data into graph structures ---
    positions, groups = parse_layout(
        sheets.get("Layout", {}),
        scale=args.scale,
        margin=args.margin,
    )
    edges  = parse_edges(sheets.get("Edges", []))
    styles = parse_styles(
        sheets.get("Style", []),
        default_w=NODE_W,
        default_h=NODE_H,
        default_style=VERTEX_STYLE,
    )

    # --- Step 3: assemble draw.io XML ---
    xml_model = build_xml(positions, edges, groups=groups, styles=styles)

    # --- Step 4: write output file ---
    tree = ET.ElementTree(xml_model)
    ET.indent(tree, space="  ")
    tree.write(output_path, encoding="unicode", xml_declaration=False)

    print(f"Written: {output_path}")
    print(f"  {len(positions)} positioned nodes, {len(edges)} edges, {len(styles)} style groups")


if __name__ == "__main__":
    main()
