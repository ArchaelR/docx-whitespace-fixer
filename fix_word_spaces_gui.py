"""
fix_word_spaces_gui.py  —  Fixes Word Online phantom spaces in .docx / .odt
Pure Python stdlib, no pip installs needed.

ROOT CAUSE
----------
Word Online/LibreOffice encodes multiple spaces as <text:s text:c="44"/> XML
elements — NOT literal spaces in text nodes. This tool removes those elements
entirely, injecting a single space into the surrounding text/tail, then
collapses any remaining literal multi-space runs.
"""

import re, os, zipfile, shutil, tempfile
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# ── Namespaces ─────────────────────────────────────────────────────────────────
DOCX_W   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ODT_TEXT = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"

DOCX_NS = {
    "wpc":"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "mc":"http://schemas.openxmlformats.org/markup-compatibility/2006",
    "o":"urn:schemas-microsoft-com:office:office",
    "r":"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "m":"http://schemas.openxmlformats.org/officeDocument/2006/math",
    "v":"urn:schemas-microsoft-com:vml",
    "wp":"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "wp14":"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "w10":"urn:schemas-microsoft-com:office:word",
    "w":DOCX_W,
    "w14":"http://schemas.microsoft.com/office/word/2010/wordml",
    "w15":"http://schemas.microsoft.com/office/word/2012/wordml",
    "w16":"http://schemas.microsoft.com/office/word/2018/wordml",
    "w16cid":"http://schemas.microsoft.com/office/word/2016/wordml/cid",
    "w16cex":"http://schemas.microsoft.com/office/word/2018/wordml/cex",
    "w16se":"http://schemas.microsoft.com/office/word/2015/wordml/symex",
    "wpg":"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "wpi":"http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
    "wne":"http://schemas.microsoft.com/office/word/2006/wordml",
    "wps":"http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "xml":"http://www.w3.org/XML/1998/namespace",
}
ODT_NS = {
    "office":"urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "text":ODT_TEXT,
    "style":"urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    "draw":"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
    "fo":"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
    "xlink":"http://www.w3.org/1999/xlink",
    "dc":"http://purl.org/dc/elements/1.1/",
    "meta":"urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
    "number":"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
    "svg":"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
    "table":"urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "loext":"urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0",
    "xml":"http://www.w3.org/XML/1998/namespace",
}
for _p, _u in {**DOCX_NS, **ODT_NS}.items():
    ET.register_namespace(_p, _u)

EXTRA_SPACE = re.compile(r"[ \u00a0]{2,}")

# ── DOCX ──────────────────────────────────────────────────────────────────────
W_T = f"{{{DOCX_W}}}t"
W_P = f"{{{DOCX_W}}}p"

def _fix_docx_xml(xml_bytes):
    root = ET.fromstring(xml_bytes)
    groups = chars = 0
    for wt in root.iter(W_T):
        for attr in ("text", "tail"):
            val = getattr(wt, attr)
            if not val: continue
            fixed, n = EXTRA_SPACE.subn(" ", val)
            if n:
                groups += n; chars += len(val) - len(fixed)
                setattr(wt, attr, fixed)
                if attr == "text" and fixed != fixed.strip():
                    wt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return ET.tostring(root, encoding="unicode").encode("utf-8"), groups, chars

def _plain_docx(xml_bytes):
    root = ET.fromstring(xml_bytes)
    lines = []
    for para in root.iter(W_P):
        lines.append("".join(wt.text for wt in para.iter(W_T) if wt.text))
    return "\n".join(lines)

def fix_docx(inp, out):
    tmp_fd, tmp = tempfile.mkstemp(suffix=".docx"); os.close(tmp_fd)
    try:
        tg = tc = 0; before = after = ""
        with zipfile.ZipFile(inp, "r") as zin:
            before = _plain_docx(zin.read("word/document.xml"))
            with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == "word/document.xml":
                        data, g, c = _fix_docx_xml(data); tg += g; tc += c
                        after = _plain_docx(data)
                    zout.writestr(item, data)
        shutil.move(tmp, out)
        return {"groups": tg, "chars": tc, "before": before, "after": after}
    except Exception:
        if os.path.exists(tmp): os.unlink(tmp)
        raise

# ── ODT ───────────────────────────────────────────────────────────────────────
# Multiple spaces in ODT = <text:s text:c="N"/>
# Fix: remove the element entirely, inject a single space into surrounding text/tail,
# then collapse any remaining literal multi-space runs.

T_S      = f"{{{ODT_TEXT}}}s"
T_P      = f"{{{ODT_TEXT}}}p"
T_C_ATTR = f"{{{ODT_TEXT}}}c"

def _fix_odt_xml(xml_bytes):
    root = ET.fromstring(xml_bytes)
    groups = [0]; chars = [0]

    def collapse(s):
        """Collapse 2+ spaces/nbsp to single space, tracking stats."""
        if not s: return s
        fixed, n = EXTRA_SPACE.subn(" ", s)
        if n: groups[0] += n; chars[0] += len(s) - len(fixed)
        return fixed

    def fix_element(parent):
        parent.text = collapse(parent.text)

        to_remove = []  # indices of text:s children with c > 1
        for i, child in enumerate(parent):
            fix_element(child)  # recurse first
            if child.tag == T_S:
                raw_c = child.get(T_C_ATTR)
                count = int(raw_c) if raw_c is not None else 1
                if count > 1:
                    # Will be removed; record index and its tail
                    groups[0] += 1; chars[0] += count - 1
                    to_remove.append((i, child.tail or ""))
                else:
                    child.tail = collapse(child.tail)
            else:
                child.tail = collapse(child.tail)

        # Remove in reverse order so indices stay valid
        for i, tail in reversed(to_remove):
            # Merge: inject one space + collapsed tail into prev sibling's tail
            # or into parent.text if this was the first child
            space_plus_tail = collapse(" " + tail)
            if i == 0:
                parent.text = collapse((parent.text or "") + space_plus_tail)
            else:
                prev = parent[i - 1]
                prev.tail = collapse((prev.tail or "") + space_plus_tail)
            parent.remove(parent[i])

    fix_element(root)
    return ET.tostring(root, encoding="unicode").encode("utf-8"), groups[0], chars[0]

def _plain_odt(xml_bytes):
    root = ET.fromstring(xml_bytes)
    lines = []
    for para in root.iter(T_P):
        parts = []
        def _collect(el):
            if el.tag == T_S:
                parts.append(" " * int(el.get(T_C_ATTR, "1")))
                if el.tail: parts.append(el.tail)
                return
            if el.text: parts.append(el.text)
            for ch in el: _collect(ch)
            if el.tail: parts.append(el.tail)
        if para.text: parts.append(para.text)
        for ch in para: _collect(ch)
        lines.append("".join(parts))
    return "\n".join(lines)

def fix_odt(inp, out):
    tmp_fd, tmp = tempfile.mkstemp(suffix=".odt"); os.close(tmp_fd)
    try:
        tg = tc = 0; before = after = ""
        with zipfile.ZipFile(inp, "r") as zin:
            before = _plain_odt(zin.read("content.xml"))
            with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == "content.xml":
                        data, g, c = _fix_odt_xml(data); tg += g; tc += c
                        after = _plain_odt(data)
                    zout.writestr(item, data)
        shutil.move(tmp, out)
        return {"groups": tg, "chars": tc, "before": before, "after": after}
    except Exception:
        if os.path.exists(tmp): os.unlink(tmp)
        raise

def fix_file(inp, out):
    ext = os.path.splitext(inp)[1].lower()
    if ext == ".docx": return fix_docx(inp, out)
    if ext == ".odt":  return fix_odt(inp, out)
    raise ValueError(f"Unsupported type: {ext}  (use .docx or .odt)")

# ── Self-test ─────────────────────────────────────────────────────────────────
def _selftest():
    cases = [
        ("text:s c=44",    b'<r xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><text:p>cheers<text:s text:c="44"/>erupting</text:p></r>', "cheers erupting"),
        ("literal spaces", b'<r xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><text:p>cheers                    erupting</text:p></r>', "cheers erupting"),
        ("nbsp x5",        '<r xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><text:p>cheers\u00a0\u00a0\u00a0\u00a0\u00a0erupting</text:p></r>'.encode("utf-8"), "cheers erupting"),
        ("mix s+tail",     b'<r xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><text:p>cheers<text:s text:c="5"/>   erupting</text:p></r>', "cheers erupting"),
        ("span tail",      b'<r xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><text:p><text:span>cheers</text:span>                    erupting</text:p></r>', "cheers erupting"),
    ]
    for name, xml, expected in cases:
        out_bytes, g, c = _fix_odt_xml(xml)
        text = "".join(ET.fromstring(out_bytes.decode()).itertext())
        assert expected in text, f"FAIL [{name}]: got {text!r}"
    print(f"Self-test PASSED: all {len(cases)} cases produce single space")

# ── GUI ───────────────────────────────────────────────────────────────────────
BG="#0f0f0f"; PANEL="#161616"; BORDER="#242424"
ACCENT="#e8ff47"; FG="#efefef"; DIM="#555555"
GREEN="#4eff91"; RED_HL="#ff6b6b"
MONO=("Courier New", 10)

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Space Fixer — .docx / .odt")
        self.configure(bg=BG)
        self.resizable(True, True)
        self.minsize(780, 600)
        self._inp    = tk.StringVar()
        self._out    = tk.StringVar()
        self._status = tk.StringVar(value="Open a .docx or .odt file to get started.")
        self._build()
        w, h = 980, 720
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _build(self):
        hdr = tk.Frame(self, bg=BG, pady=16); hdr.pack(fill="x", padx=24)
        tk.Label(hdr, text="SPACE FIXER",
                 font=("Courier New", 16, "bold"), fg=ACCENT, bg=BG).pack(side="left")
        tk.Label(hdr, text="  removes phantom spaces from .docx / .odt — bold/italic/colour preserved",
                 font=("Georgia", 9, "italic"), fg=DIM, bg=BG).pack(side="left")

        note = tk.Frame(self, bg="#131300", pady=8); note.pack(fill="x", padx=20, pady=(0,10))
        tk.Label(note,
                 text='ℹ  Word Online stores multiple spaces as <text:s text:c="N"/> XML elements.\n'
                      '   This tool removes those elements and all literal multi-space runs.',
                 font=("Courier New", 9), fg="#aaaa44", bg="#131300",
                 justify="left").pack(side="left", padx=14)

        fp = tk.Frame(self, bg=PANEL, pady=12); fp.pack(fill="x", padx=20, pady=(0,8))
        fp.columnconfigure(1, weight=1)
        self._file_row(fp, "INPUT  (.docx / .odt)", self._inp, self._browse_inp, 0)
        self._file_row(fp, "OUTPUT (.docx / .odt)", self._out, self._browse_out, 1)

        bf = tk.Frame(self, bg=BG, pady=4); bf.pack(fill="x", padx=20, pady=(0,10))
        self._btn(bf, "⚡  FIX & SAVE", self._do_fix, accent=True).pack(side="left")

        style = ttk.Style(); style.theme_use("default")
        style.configure("D.TNotebook", background=BG, borderwidth=0)
        style.configure("D.TNotebook.Tab", background=BORDER, foreground=DIM,
                        font=("Courier New", 9, "bold"), padding=[12, 5])
        style.map("D.TNotebook.Tab",
                  background=[("selected", PANEL)], foreground=[("selected", ACCENT)])
        nb = ttk.Notebook(self, style="D.TNotebook")
        nb.pack(fill="both", expand=True, padx=20, pady=(0,4))
        self._before_box = self._tab(nb, "BEFORE")
        self._after_box  = self._tab(nb, "AFTER")
        self._diff_box   = self._tab(nb, "CHANGES")

        sb = tk.Frame(self, bg=PANEL, height=30); sb.pack(fill="x", padx=20, pady=(4,14))
        tk.Label(sb, textvariable=self._status, font=MONO, fg=DIM,
                 bg=PANEL, anchor="w").pack(side="left", padx=10, pady=6)

    def _file_row(self, parent, label, var, cmd, row):
        tk.Label(parent, text=label, font=("Courier New", 9, "bold"), fg=DIM, bg=PANEL,
                 width=22, anchor="e").grid(row=row, column=0, padx=(14,8), pady=6, sticky="e")
        tk.Entry(parent, textvariable=var, font=MONO, bg="#1c1c1c", fg=FG,
                 insertbackground=ACCENT, relief="flat", bd=0).grid(
            row=row, column=1, sticky="ew", padx=(0,8), ipady=5)
        self._btn(parent, "Browse…", cmd).grid(row=row, column=2, padx=(0,14))

    def _tab(self, nb, title):
        frame = tk.Frame(nb, bg=PANEL); nb.add(frame, text=title)
        txt = tk.Text(frame, font=MONO, bg=PANEL, fg=FG, relief="flat", bd=0,
                      wrap="word", padx=12, pady=10, state="disabled",
                      selectbackground="#333300", selectforeground=ACCENT)
        sc = tk.Scrollbar(frame, orient="vertical", command=txt.yview,
                          bg=PANEL, troughcolor=BG, width=8)
        txt.configure(yscrollcommand=sc.set)
        sc.pack(side="right", fill="y"); txt.pack(side="left", fill="both", expand=True)
        return txt

    def _btn(self, parent, label, cmd, accent=False):
        fg  = BG if accent else FG
        bg  = ACCENT if accent else "#252525"
        abg = "#c8dd30" if accent else "#333333"
        b = tk.Button(parent, text=label, command=cmd,
                      font=("Courier New", 10, "bold"),
                      fg=fg, bg=bg, activeforeground=fg, activebackground=abg,
                      relief="flat", bd=0, padx=16, pady=8, cursor="hand2")
        b.bind("<Enter>", lambda e: b.config(bg=abg))
        b.bind("<Leave>", lambda e: b.config(bg=bg))
        return b

    def _write(self, box, text):
        box.config(state="normal"); box.delete("1.0", "end")
        box.insert("1.0", text);   box.config(state="disabled")

    def _browse_inp(self):
        p = filedialog.askopenfilename(
            title="Select input file",
            filetypes=[("Word / ODT", "*.docx *.odt"), ("Word (.docx)", "*.docx"),
                       ("OpenDocument (.odt)", "*.odt"), ("All files", "*.*")])
        if not p: return
        self._inp.set(p)
        base, ext = os.path.splitext(p)
        self._out.set(f"{base}_fixed{ext}")
        self._status.set(f"Loaded: {os.path.basename(p)}")

    def _browse_out(self):
        p = filedialog.asksaveasfilename(
            title="Save fixed file as…", defaultextension=".docx",
            filetypes=[("Word (.docx)", "*.docx"), ("OpenDocument (.odt)", "*.odt"),
                       ("All files", "*.*")])
        if p: self._out.set(p)

    def _do_fix(self):
        inp = self._inp.get().strip()
        out = self._out.get().strip()
        if not inp:
            messagebox.showwarning("No input", "Please select an input file."); return
        if not os.path.isfile(inp):
            messagebox.showerror("Not found", f"File not found:\n{inp}"); return
        if not out:
            messagebox.showwarning("No output", "Please choose a save location."); return
        try:
            stats = fix_file(inp, out)
        except Exception as exc:
            messagebox.showerror("Error", f"Failed:\n\n{exc}"); return

        self._write(self._before_box, stats["before"])
        self._write(self._after_box,  stats["after"])
        self._show_diff(stats["before"], stats["after"])

        if stats["groups"] == 0:
            msg = "✔  No phantom spaces found — file was already clean."
        else:
            msg = (f"✔  {stats['groups']} space group(s) collapsed, "
                   f"{stats['chars']} extra character(s) removed.  "
                   f"Saved → {os.path.basename(out)}")
        self._status.set(msg)
        messagebox.showinfo("Done!",
            f"Saved: {out}\n\n"
            f"Space groups collapsed : {stats['groups']}\n"
            f"Extra characters removed: {stats['chars']}\n\n"
            "All bold / italic / colour / font formatting is intact.")

    def _show_diff(self, before, after):
        box = self._diff_box
        box.config(state="normal"); box.delete("1.0", "end")
        box.tag_config("rem",   foreground=RED_HL, background="#2a1010")
        box.tag_config("add",   foreground=GREEN,  background="#0a2a15")
        box.tag_config("same",  foreground=DIM)
        box.tag_config("label", foreground=ACCENT, font=("Courier New", 9, "bold"))
        changed = 0
        for bl, al in zip(before.splitlines(), after.splitlines()):
            if bl == al: box.insert("end", f"  {al}\n", "same")
            else:
                changed += 1
                box.insert("end", f"- {bl}\n", "rem")
                box.insert("end", f"+ {al}\n", "add")
        if changed == 0:
            box.insert("end", "\n  (no differences — file was already clean)", "label")
        box.config(state="disabled")

if __name__ == "__main__":
    _selftest()
    App().mainloop()
