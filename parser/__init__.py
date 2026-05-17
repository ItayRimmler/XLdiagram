"""
Converts raw xlsx data into structured graph objects ready for the builder.

parse_layout — Layout cells → {node_name: (x_pixels, y_pixels)}, {node_name: group}
parse_edges  — Edge rows   → list of edge dicts
parse_styles — Style rows  → {group_name: {w, h, style}}

No XML, no file I/O here. Pure data transformation.
"""

import re

def _strip_parens(name: str) -> str:
    """Remove a trailing parenthetical from a node name.
    'Node1 (Deez Nuts)' -> 'Node1'
    Parentheticals anywhere else are left alone.
    """
    return re.sub(r'\s*\([^)]*\)\s*$', '', name).strip()


# Each separator is a pictogram of what it renders in the diagram.
# The value is the draw.io arrow style string for that connection type.
#
#   >   arrowhead points right             forward arrow
#   <   arrowhead points left              backward arrow
#   o>  filled dot, then arrowhead         dot at source, arrow at target
#   <o  arrowhead, then filled dot         arrow at source, dot at target
#   <>  arrowheads facing outward          bidirectional
#   -   flat line                          plain wire, no arrows
ARROW_STYLES = {
    ">":  "startArrow=none;endArrow=block;endFill=1;",
    "<":  "startArrow=block;startFill=1;endArrow=none;",
    "o>": "startArrow=oval;startFill=1;endArrow=block;endFill=1;",
    "<o": "startArrow=block;startFill=1;endArrow=oval;endFill=1;",
    "<>": "startArrow=block;startFill=1;endArrow=block;endFill=1;",
    "-":  "startArrow=none;endArrow=none;",
    "o-": "startArrow=oval;startFill=1;endArrow=none;",   # dot at source, plain line to target
    "-o": "startArrow=none;endArrow=oval;endFill=1;",     # plain line from source, dot at target
}

# Fallback style for unrecognized separators — plain wire is the safest default
_DEFAULT_STYLE = ARROW_STYLES["-"]


def parse_layout(layout_cells, scale, margin=100):
    """
    Convert Layout sheet cells into pixel positions and group assignments.

    layout_cells: {(row, col): cell_value}  from reader
                  cell_value is either "NodeName" or "NodeName|GroupName"

    Returns (positions, groups):
      positions: {node_name: (x_pixels, y_pixels)}
      groups:    {node_name: group_name}  — only for nodes that declared a group

    Excel grid mapping:
      col 1 → x = margin,   col 2 → x = margin + scale ...
      row 1 → y = margin,   row 2 → y = margin + scale ...

    If layout_cells is empty, returns ({}, {}) and all nodes stack at (0,0).
    """
    positions = {}
    groups    = {}

    for (row, col), cell_value in layout_cells.items():
        # Split "NodeName|GroupName" — the | separates name from optional group
        if "|" in cell_value:
            name, group = cell_value.split("|", 1)
            name  = _strip_parens(name)
            group = group.strip()
            groups[name] = group
        else:
            name = _strip_parens(cell_value)

        x = (col - 1) * scale + margin
        y = (row - 1) * scale + margin
        positions[name] = (x, y)

    return positions, groups


def parse_styles(style_rows, default_w, default_h, default_style):
    """
    Convert Style sheet rows into a dict keyed by group name.

    style_rows: [[active, group, width, height, style_string], ...]  from reader
      active: "#" = inactive (skip), blank = active

    Returns {group_name: {"w": int, "h": int, "style": str}}
    """
    styles = {}
    for active, group, width, height, style_str in style_rows:
        if str(active).strip() == "#" or not group:
            continue  # skip commented-out or blank rows
        try:
            w = int(width)
        except (ValueError, TypeError):
            w = default_w

        try:
            h = int(height)
        except (ValueError, TypeError):
            h = default_h

        styles[group] = {
            "w":     w,
            "h":     h,
            "style": style_str if style_str else default_style,
        }

    return styles


def parse_edges(edge_rows):
    """
    Convert raw edge rows from the Edges sheet into structured edge dicts.

    edge_rows: [[source, separator, target, label], ...]  from reader

    Each output edge dict:
      source  — name of the source node
      target  — name of the target node
      label   — text displayed on the edge (may be empty string)
      style   — draw.io arrow style string derived from the separator

    Rows with empty source or target are skipped silently.
    Unknown separators fall back to plain wire style.
    """
    edges = []
    for source, sep, target, label in edge_rows:
        if not source or not target:
            continue  # incomplete row — missing a node name, skip

        style = ARROW_STYLES.get(sep, _DEFAULT_STYLE)

        edges.append({
            "source": _strip_parens(source),
            "target": _strip_parens(target),
            "label":  label,
            "style":  style,
        })

    return edges
