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

# Visual style for node boxes — rounded rectangle with text wrapping
VERTEX_STYLE = "rounded=1;whiteSpace=wrap;html=1;"

# Phantom nodes: invisible routing points edges can pass through.
# Convention: any node whose name starts with "." is a phantom.
# They exist in the XML (so edges can connect) but render as nothing.
# Phantom nodes are the same size as real nodes so edges connect at identical anchor points.
# Using a different size would shift where edges attach, creating visual misalignment.
PHANTOM_STYLE = "point;fillColor=none;strokeColor=none;opacity=0;"
PHANTOM_W = NODE_W
PHANTOM_H = NODE_H

# Base routing style for all edges — arrow style is appended per edge
EDGE_BASE_STYLE = "edgeStyle=orthogonalEdgeStyle;html=1;rounded=0;"

# Node box dimensions in pixels
NODE_W = 80
NODE_H = 40


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


def _add_node(root, cell_id, name, x, y):
    """
    Append one node (vertex box) to the XML tree.

    cell_id — unique string ID for this cell (must be unique across the whole file)
    name    — text label displayed inside the box
    x, y    — top-left corner position in pixels
              (mxGeometry origin is top-left of the canvas)
    """
    cell = ET.SubElement(root, "mxCell", {
        "id":     cell_id,
        "value":  name,
        "style":  VERTEX_STYLE,
        "vertex": "1",   # marks this as a node (not an edge)
        "parent": "1",   # child of the default layer (id=1)
    })
    ET.SubElement(cell, "mxGeometry", {
        "x":      str(x),
        "y":      str(y),
        "width":  str(NODE_W),
        "height": str(NODE_H),
        "as":     "geometry",  # required attribute — tells draw.io this is geometry data
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


def build_xml(positions, edges):
    """
    Build and return the complete mxGraphModel XML element.

    positions — {node_name: (x, y)}  from parse_layout
    edges     — list of edge dicts    from parse_edges

    Node discovery:
      Nodes come from two sources — the Layout sheet (positions) and edge endpoints.
      We take the union so a node mentioned only in an edge still gets created.
      Nodes with no layout entry are placed at (0, 0) and will stack visually.

    Returns the root ET.Element (<mxGraphModel>) ready to be written to a file.
    """
    model, root = _make_model_root()

    # next_id tracks the next unused integer ID for cells.
    # Starts at 2 because 0 and 1 are the scaffold cells.
    next_id = 2

    # name_to_id lets the edge-building step look up a node's ID by its name.
    # Built during the node loop, consumed during the edge loop.
    name_to_id = {}

    # Collect every node name that appears anywhere — layout OR edge endpoints.
    # Union ensures we never silently drop a node that only appears in an edge.
    all_names = (
        set(positions)
        | {e["source"] for e in edges}
        | {e["target"] for e in edges}
    )

    for name in sorted(all_names):  # sorted = deterministic output order
        x, y = positions.get(name, (0, 0))  # no layout entry → stack at origin
        cell_id = str(next_id)

        if name.startswith("."):
            # Phantom node: invisible routing point, not a real diagram element.
            # Edges can connect to it to force a path through a specific position.
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
                "width":  str(PHANTOM_W),
                "height": str(PHANTOM_H),
                "as":     "geometry",
            })
        else:
            _add_node(root, cell_id, name, x, y)

        name_to_id[name] = cell_id
        next_id += 1

    for edge in edges:
        src_id = name_to_id.get(edge["source"])
        tgt_id = name_to_id.get(edge["target"])

        if not src_id or not tgt_id:
            # This can't happen (union above guarantees all names exist),
            # but guard anyway in case of future refactoring.
            print(f"  Warning: skipping edge '{edge['source']} → {edge['target']}' — node not found")
            continue

        _add_edge(root, str(next_id), src_id, tgt_id, edge["label"], edge["style"])
        next_id += 1

    return model
