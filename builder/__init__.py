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
    fallback = {"w": NODE_W, "h": NODE_H, "style": VERTEX_STYLE}

    if not styles:
        return fallback

    group = groups.get(name)
    if group and group in styles:
        return styles[group]

    if "default" in styles:
        return styles["default"]

    if len(styles) == 1:
        return next(iter(styles.values()))

    return fallback


def _make_model_root():
    model = ET.Element("mxGraphModel", {
        "grid":       "1",
        "gridSize":   "10",
        "pageWidth":  "1169",
        "pageHeight": "827",
    })
    root = ET.SubElement(model, "root")
    ET.SubElement(root, "mxCell", {"id": "0"})
    ET.SubElement(root, "mxCell", {"id": "1", "parent": "0"})
    return model, root


def _add_node(root, cell_id, name, x, y, w, h, style):
    cell = ET.SubElement(root, "mxCell", {
        "id":     cell_id,
        "value":  name,
        "style":  style,
        "vertex": "1",
        "parent": "1",
    })
    ET.SubElement(cell, "mxGeometry", {
        "x":      str(x),
        "y":      str(y),
        "width":  str(w),
        "height": str(h),
        "as":     "geometry",
    })


def _add_edge(root, cell_id, source_id, target_id, label, arrow_style):
    cell = ET.SubElement(root, "mxCell", {
        "id":     cell_id,
        "value":  label,
        "style":  EDGE_BASE_STYLE + arrow_style,
        "edge":   "1",
        "source": source_id,
        "target": target_id,
        "parent": "1",
    })
    ET.SubElement(cell, "mxGeometry", {"relative": "1", "as": "geometry"})


def build_xml(positions, edges, groups=None, styles=None, zero_phantom=False):
    groups = groups or {}
    styles = styles or {}

    model, root = _make_model_root()

    next_id = 2
    name_to_id = {}

    all_names = (
        set(positions)
        | {e["source"] for e in edges}
        | {e["target"] for e in edges}
    )

    for name in sorted(all_names):
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
            print(f"  Warning: skipping edge '{edge['source']} -> {edge['target']}' — node not found")
            continue

        _add_edge(root, str(next_id), src_id, tgt_id, edge["label"], edge["style"])
        next_id += 1

    return model
