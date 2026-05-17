"""
GUI for topo2visio — pick an .xlsx file, convert, open the result.
"""

import sys
import os
import glob
import subprocess
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path

# PyInstaller: when frozen, BASE_DIR is next to the .exe; otherwise next to gui.py
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent

sys.path.insert(0, str(BASE_DIR))

from reader  import read_xlsx
from parser  import parse_layout, parse_edges, parse_styles
from builder import build_xml, NODE_W, NODE_H, VERTEX_STYLE

OUTPUTS_DIR = BASE_DIR / "outputs"
OUTPUTS_DIR.mkdir(exist_ok=True)

DEFAULT_SCALE  = 120
DEFAULT_MARGIN = 100

DRAWIO_CANDIDATES = [
    Path(r"C:\Program Files\draw26.io\draw.io\draw.io.exe"),
    Path(r"C:\Program Files\draw.io\draw.io.exe"),
    Path(r"C:\Program Files (x86)\draw.io\draw.io.exe"),
    Path(r"C:\Users") / os.environ.get("USERNAME", "") / r"AppData\Local\Programs\draw.io\draw.io.exe",
]

INSTRUCTIONS = (
    "HOW TO USE\n"
    "──────────────────────────────────────────────────\n"
    "1. Open  inputs/template.xlsx  or  inputs/example.xlsx\n"
    "   to understand the format.\n\n"
    "2. Layout sheet  — place node names in cells.\n"
    "   Cell position = diagram position.\n"
    "   Use  Name|Group  to assign a visual style group.\n\n"
    "3. Edges sheet  — one row per connection:\n"
    "   Source | Separator | Target | Label\n"
    "   Separators:  >  <  o>  <o  <>  -  o-  -o\n\n"
    "4. Style sheet  — define groups with width, height,\n"
    "   and a draw.io style string.\n"
    "   Lines starting with # are inactive (templates).\n"
    "   Remove the # to activate a style.\n\n"
    "5. Click Browse, pick your .xlsx, click Convert.\n"
    "──────────────────────────────────────────────────"
)


def _find_drawio():
    # Check known fixed paths first
    for p in DRAWIO_CANDIDATES:
        if p.exists():
            return p
    # Glob for any version-named install: "C:\Program Files\draw*.io\draw.io\draw.io.exe"
    for pattern in [
        r"C:\Program Files\draw*.io\draw.io\draw.io.exe",
        r"C:\Program Files (x86)\draw*.io\draw.io\draw.io.exe",
    ]:
        hits = glob.glob(pattern)
        if hits:
            return Path(hits[0])
    # Registry: look up what opens .drawio files
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r".drawio") as k:
            prog_id = winreg.QueryValue(k, "")
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT,
                            rf"{prog_id}\shell\open\command") as k:
            cmd = winreg.QueryValue(k, "")
        # cmd looks like: "C:\...\draw.io.exe" "%1"
        exe = Path(cmd.split('"')[1])
        if exe.exists():
            return exe
    except Exception:
        pass
    return None


def convert(input_path: Path):
    sheets = read_xlsx(input_path)
    positions, groups = parse_layout(sheets.get("Layout", {}), scale=DEFAULT_SCALE, margin=DEFAULT_MARGIN)
    edges  = parse_edges(sheets.get("Edges", []))
    styles = parse_styles(sheets.get("Style", []), default_w=NODE_W, default_h=NODE_H, default_style=VERTEX_STYLE)

    xml_model = build_xml(positions, edges, groups=groups, styles=styles)
    output_path = OUTPUTS_DIR / input_path.with_suffix(".drawio").name

    tree = ET.ElementTree(xml_model)
    ET.indent(tree, space="  ")
    tree.write(output_path, encoding="unicode", xml_declaration=False)

    return output_path, len(positions), len(edges)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("topo2visio")
        self.resizable(False, False)
        self._output_path = None
        self._drawio = _find_drawio()
        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 16, "pady": 6}

        # Instructions
        instr = tk.Text(self, height=18, width=58, relief="flat",
                        bg="#f7f7f7", font=("Consolas", 9), wrap="word")
        instr.insert("1.0", INSTRUCTIONS)
        instr.config(state="disabled")
        instr.pack(padx=16, pady=(14, 4))

        tk.Frame(self, height=1, bg="#cccccc").pack(fill="x", padx=16)

        # draw.io exe row
        drawio_frame = tk.Frame(self)
        drawio_frame.pack(fill="x", **pad)

        drawio_label = self._drawio.name if self._drawio else "draw.io not found"
        drawio_color = "black" if self._drawio else "#cc0000"
        self._drawio_var = tk.StringVar(value=drawio_label)
        tk.Label(drawio_frame, text="draw.io:", width=8, anchor="w").pack(side="left")
        tk.Label(drawio_frame, textvariable=self._drawio_var, anchor="w",
                 width=34, relief="sunken", bg="white", fg=drawio_color).pack(side="left", fill="x", expand=True)
        tk.Button(drawio_frame, text="Browse…", command=self._browse_drawio).pack(side="left", padx=(8, 0))

        # File row
        file_frame = tk.Frame(self)
        file_frame.pack(fill="x", **pad)

        self._path_var = tk.StringVar(value="No file selected")
        tk.Label(file_frame, text="xlsx file:", width=8, anchor="w").pack(side="left")
        tk.Label(file_frame, textvariable=self._path_var, anchor="w",
                 width=34, relief="sunken", bg="white").pack(side="left", fill="x", expand=True)
        tk.Button(file_frame, text="Browse…", command=self._browse).pack(side="left", padx=(8, 0))

        # Convert
        self._convert_btn = tk.Button(self, text="Convert", width=16,
                                      command=self._convert, state="disabled")
        self._convert_btn.pack(**pad)

        # Status
        self._status_var = tk.StringVar(value="")
        tk.Label(self, textvariable=self._status_var, fg="gray",
                 wraplength=440, justify="left").pack(**pad)

        # Open output
        self._open_btn = tk.Button(self, text="Open in draw.io", width=16,
                                   command=self._open_output, state="disabled")
        self._open_btn.pack(pady=(0, 14))

        if not self._drawio:
            self._open_btn.config(text="Open output file")

    def _browse_drawio(self):
        path = filedialog.askopenfilename(
            title="Locate draw.io executable",
            filetypes=[("Executable", "*.exe"), ("All files", "*.*")],
            initialdir=r"C:\Program Files",
        )
        if not path:
            return
        self._drawio = Path(path)
        self._drawio_var.set(self._drawio.name)
        self._open_btn.config(text="Open in draw.io")

    def _browse(self):
        path = filedialog.askopenfilename(
            title="Select Excel diagram file",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not path:
            return
        self._input_path = Path(path)
        self._path_var.set(self._input_path.name)
        self._convert_btn.config(state="normal")
        self._open_btn.config(state="disabled")
        self._status_var.set("")

    def _convert(self):
        self._status_var.set("Converting…")
        self.update_idletasks()
        try:
            out, n_nodes, n_edges = convert(self._input_path)
            self._output_path = out
            self._status_var.set(f"Done — {n_nodes} nodes, {n_edges} edges  →  {out.name}")
            self._open_btn.config(state="normal")
        except Exception as e:
            self._status_var.set(f"Error: {e}")
            messagebox.showerror("Conversion failed", str(e))

    def _open_output(self):
        if self._drawio:
            subprocess.Popen([str(self._drawio), str(self._output_path)])
        else:
            os.startfile(self._output_path)


if __name__ == "__main__":
    App().mainloop()
