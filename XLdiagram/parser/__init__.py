"""
Converts raw xlsx data into structured graph objects ready for the builder.

parse_layout — Layout cells → {node_name: (x_pixels, y_pixels)}
parse_edges  — Edge rows   → list of edge dicts

No XML, no file I/O here. Pure data transformation.
"""

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
    Convert Layout sheet cells into pixel positions for each node.

    layout_cells: {(row, col): node_name}  from reader
    scale:        pixels per Excel grid cell — controls how spread-out the diagram is
    margin:       pixel offset added to every position so the diagram doesn't
                  start at the canvas corner (0,0) and get clipped

    Excel grid mapping:
      col 1 → x = margin,         col 2 → x = margin + scale ...
      row 1 → y = margin,         row 2 → y = margin + scale ...

    If layout_cells is empty (no Layout sheet), returns {} and all nodes
    will be placed at (0, 0) by the builder — they stack, which is intentional.
    """
    positions = {}
    for (row, col), name in layout_cells.items():
        x = (col - 1) * scale + margin  # col 1 → margin, col 2 → margin+scale ...
        y = (row - 1) * scale + margin  # same for rows
        positions[name] = (x, y)
    return positions


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
            "source": source,
            "target": target,
            "label":  label,
            "style":  style,
        })

    return edges
