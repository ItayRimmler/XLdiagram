"""
Assembles the draw.io XML tree from parsed graph data.

draw.io XML structure:
  <mxGraphModel>
    <root>
      <mxCell id="0" />                        required scaffold — always present
      <mxCell id="1" parent="0" />             required scaffold — all user cells live under this
      <mxCell id="2" vertex="1" ...>           a node (box)
        <mxGeometry x="..." y="..." />
      </mxCell>
      <mxCell id="3" edge="1"                  a connection line
             source="2" target="...">
        <mxGeometry relative="1" />
      </mxCell>
    </root>
  </mxGraphModel>

ID rules:
  0 and 1 are reserved for the scaffold — never use them for user content.
  User cell IDs start at 2 and increment by 1.
  All attribute values must be strings — ElementTree crashes on integers.
"""

import xml.etree.ElementTree as ET

# Default node box dimensions — used when no Style sheet is provided or a node has no group
NODE_W = 80
NODE_H = 40

# Default draw.io style for node boxes — rounded rectangle with text wrapping
VERTEX_STYLE = "rounded=1;whiteSpace=wrap;html=1;"

# Phantom nodes: invisible routing points edges can pass through.
# Convention: any node whose name starts with "." is a phantom.
# They exist in the XML (so edges can connect) but render as nothing.
# Same size as real nodes so edge anchor points align perfectly.
PHANTOM_STYLE = "point;fillColor=none;strokeColor=none;opacity=0;"
PHANTOM_W = NODE_W
PHANTOM_H = NODE_H

# Base routing style for all edges — arrow style is appended per edge
EDGE_BASE_STYLE = "edgeStyle=orthogonalEdgeStyle;html=1;rounded=0;"


def _resolve_style(name, groups, styles):
    """
    Look up the visual style for a node using a fallback chain.

    groups: {node_name: group_name}       from parse_layout
    styles: {group_name: {w, h, style}}   from parse_styles

    Fallback chain:
      1. Node's declared group (e.g. "R1|sensor" → look up "sensor")
      2. Group named "default" if it exists in the Style sheet
      3. The only group defined, if exactly 1 group exists
      4. Hardcoded NODE_W / NODE_H / VERTEX_STYLE constants

    Returns {"w": int, "h": int, "style": str}
    """
    fallback = {"w": NODE_W, "h": NODE_H, "style": VERTEX_STYLE}

    if not styles:
        return fallback  # no Style sheet at all

    # Try the node's own group assignment
    group = groups.get(name)
    if group and group in styles:
        return styles[group]

    # Try the group literally named "default"
    if "default" in styles:
        return styles["default"]

    # If exactly one group is defined, every node uses it
    if len(styles) == 1:
        return next(iter(styles.values()))

    return fallback


def _make_model_root():
    """
    Create the top-level mxGraphModel element and its required scaffold cells.

    Returns (model_element, root_element):
      model_element — the <mxGraphModel> tag, written to file as the root
      root_element  — the <root> child, where all user cells are appended
    """
    model = ET.Element("mxGraphModel", {
        "grid":       "1",
        "gridSize":   "10",
        "pageWidth":  "1169",  # A4 landscape in draw.io units
        "pageHeight": "827",
    })
    root = ET.SubElement(model, "root")

    # id=0: the graph's root cell — required by draw.io, has no visual
    ET.SubElement(root, "mxCell", {"id": "0"})

    # id=1: the default parent for all user cells — required by draw.io
    ET.SubElement(root, "mxCell", {"id": "1", "parent": "0"})

    return model, root


def _add_node(root, cell_id, name, x, y, w, h, style):
    """
    Append one node (vertex box) to the XML tree.

    cell_id — unique string ID for this cell
    name    — text label displayed inside the box
    x, y    — top-left corner position in pixels
    w, h    — box dimensions in pixels (from group style or defaults)
    style   — draw.io style string (from group style or defaults)
    """
    cell = ET.SubElement(root, "mxCell", {
        "id":     cell_id,
        "value":  name,
        "style":  style,
        "vertex": "1",   # marks this as a node (not an edge)
        "parent": "1",   # child of the default layer (id=1)
    })
    ET.SubElement(cell, "mxGeometry", {
        "x":      str(x),
        "y":      str(y),
        "width":  str(w),
        "height": str(h),
        "as":     "geometry",  # required — tells draw.io this is geometry data
    })


def _add_edge(root, cell_id, source_id, target_id, label, arrow_style):
    """
    Append one edge (connection line) to the XML tree.

    cell_id     — unique string ID for this edge cell
    source_id   — string ID of the source node cell
    target_id   — string ID of the target node cell
    label       — text shown along the edge (empty string = no label)
    arrow_style — draw.io style string for arrowhead shape/direction
    """
    cell = ET.SubElement(root, "mxCell", {
        "id":     cell_id,
        "value":  label,
        "style":  EDGE_BASE_STYLE + arrow_style,
        "edge":   "1",       # marks this as a connection (not a node)
        "source": source_id,
        "target": target_id,
        "parent": "1",
    })
    # relative="1" means edge label position is relative to the edge midpoint
    ET.SubElement(cell, "mxGeometry", {"relative": "1", "as": "geometry"})


def build_xml(positions, edges, groups=None, styles=None, zero_phantom=False):
    """
    Build and return the complete mxGraphModel XML element.

    positions — {node_name: (x, y)}               from parse_layout
    edges     — list of edge dicts                 from parse_edges
    groups    — {node_name: group_name}            from parse_layout (optional)
    styles    — {group_name: {w, h, style}}        from parse_styles (optional)

    Node discovery:
      Nodes come from two sources — Layout sheet (positions) and edge endpoints.
      Union ensures nodes mentioned only in edges still get created.
      Nodes with no layout entry are placed at (0, 0) and stack visually.

    Returns the root ET.Element (<mxGraphModel>) ready to be written to a file.
    """
    groups = groups or {}
    styles = styles or {}

    model, root = _make_model_root()

    # next_id tracks the next unused integer ID.
    # Starts at 2 because 0 and 1 are scaffold cells.
    next_id = 2

    # name_to_id lets the edge loop reference nodes by name → draw.io ID.
    name_to_id = {}

    # Union of all node names from layout and edge endpoints
    all_names = (
        set(positions)
        | {e["source"] for e in edges}
        | {e["target"] for e in edges}
    )

    for name in sorted(all_names):  # sorted = deterministic output order
        x, y    = positions.get(name, (0, 0))
        cell_id = str(next_id)

        if name.startswith("."):
            pw = 0 if zero_phantom else PHANTOM_W
            ph = 0 if zero_phantom else PHANTOM_H
            cell = ET.SubElement(root, "mxCell", {
                "id":     cell_id,
                "value":  "",
                "style":  PHANTOM_STYLE,
                "vertex": "1",
                "parent": "1",
            })
            ET.SubElement(cell, "mxGeometry", {
                "x":      str(x),
                "y":      str(y),
                "width":  str(pw),
                "height": str(ph),
                "as":     "geometry",
            })
        else:
            s = _resolve_style(name, groups, styles)
            _add_node(root, cell_id, name, x, y, s["w"], s["h"], s["style"])

        name_to_id[name] = cell_id
        next_id += 1

    for edge in edges:
        src_id = name_to_id.get(edge["source"])
        tgt_id = name_to_id.get(edge["target"])

        if not src_id or not tgt_id:
            # Can't happen (union above covers all names) but guard for safety
            print(f"  Warning: skipping edge '{edge['source']} → {edge['target']}' — node not found")
            continue

        _add_edge(root, str(next_id), src_id, tgt_id, edge["label"], edge["style"])
        next_id += 1

    return model
