# topo2visio

Convert an Excel diagram spec into a draw.io file, exportable to Visio.

You describe what connects to what. The script handles positions and XML.
No dragging. No alignment bugs. No Visio open required until the very end.

---

## Flow

```
diagram.xlsx  →  topo2visio.py  →  diagram.drawio  →  draw.io (export to .vsdx)
```

---

## Setup

No pip installs. Python 3.3+ only.

```
py topo2visio.py mydiagram.xlsx
```

Input is read from `inputs/`. Output goes to `outputs/`.

---

## Excel file format

Two sheets: **Layout** and **Edges**.

### Layout sheet

Place a node name in any cell. Its row and column become its position on the diagram.

```
     A        B        C
1   GND
2            R1       R2
3                     C1
```

Nodes not in the Layout sheet are placed at (0, 0) and stack — fix them in draw.io manually.

**Phantom nodes:** prefix a name with `.` (e.g. `.p1`) to create an invisible routing point.
Edges can connect to it to force a path through a specific position. No box is drawn.

### Edges sheet

| Source | Separator | Target | Label    |
|--------|-----------|--------|----------|
| R1     | >         | R2     |          |
| GND    | -         | R1     | ground   |
| R1     | o>        | C1     | 5V       |

Row 1 is a header and is skipped. Label column is optional.

---

## Separator reference

Each separator is a pictogram of what it renders.

| Separator | Renders as                          |
|-----------|-------------------------------------|
| `>`       | forward arrow                       |
| `<`       | backward arrow                      |
| `<>`      | bidirectional arrow                 |
| `o>`      | dot at source, arrow at target      |
| `<o`      | arrow at source, dot at target      |
| `o-`      | dot at source, plain line to target |
| `-o`      | plain line from source, dot at target |
| `-`       | plain wire (no arrows)              |

---

## Configuration

In `topo2visio.py`:

| Constant | Default | Effect                                      |
|----------|---------|---------------------------------------------|
| `SCALE`  | 120     | Pixels per Excel grid cell — controls spacing |
| `MARGIN` | 100     | Offset from canvas edge so nodes aren't clipped |

---

## File structure

```
topo2visio.py       — entry point, CLI
reader/             — reads xlsx using stdlib only (zipfile + ElementTree)
parser/             — converts raw cells into positions and edge dicts
builder/            — assembles draw.io XML
inputs/             — drop .xlsx files here
outputs/            — .drawio files appear here
PLAN.md             — project roadmap and design decisions
```
