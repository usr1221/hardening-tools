#!/usr/bin/env python3
"""DOCX Tree Editor -- browse, search and edit the parts inside a
.docx/.dotx/.docm/.dotm (or any zip) without unzipping by hand.

  * Left  : folder tree of every part in the package.
  * Right : text editor for the selected part (XML/text). Binary parts (images,
            vbaProject.bin, embedded fonts) are shown read-only as a hex preview.
  * Find bar (Ctrl+F): search the current part, or "All parts" to search every
            part at once (incl. binary as latin-1) with a jump-to-match results list.
  * Toolbar: Open, Save, Save As, Format XML, Add/Delete/Rename Part.

Edits are held in memory; Save repackages the zip (original part order preserved,
deflate compression). Editing one part never touches the others' bytes.

No third-party deps -- pure standard library (tkinter).

Usage:
    python docx_tree_editor.py [file.docx]
    python docx_tree_editor.py --selftest      # headless round-trip + search test
"""
import os
import re
import sys
import zipfile

# ---------------------------------------------------------------------------
# Pure package I/O + search (no GUI) -- unit-tested by --selftest
# ---------------------------------------------------------------------------
BINARY_EXTS = {".bin", ".png", ".jpg", ".jpeg", ".jpe", ".gif", ".bmp", ".emf",
               ".wmf", ".tif", ".tiff", ".ico", ".odttf", ".ttf", ".otf"}


def read_package(path):
    """Return (order, parts): order = part names in file order, parts = name->bytes."""
    order, parts = [], {}
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            order.append(name)
            parts[name] = z.read(name)
    return order, parts


def write_package(path, order, parts):
    """Write parts back to a zip at path, preserving the given order."""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name in order:
            if name in parts:
                z.writestr(name, parts[name])


def is_binary(name, data):
    ext = os.path.splitext(name)[1].lower()
    if ext in BINARY_EXTS:
        return True
    try:
        data.decode("utf-8")
        return b"\x00" in data
    except UnicodeDecodeError:
        return True


def decode_for_search(name, data):
    """Text for searching. utf-8 for text parts, latin-1 for binary so byte
    offsets line up 1:1 with characters."""
    if is_binary(name, data):
        return data.decode("latin-1"), True
    try:
        return data.decode("utf-8"), False
    except UnicodeDecodeError:
        return data.decode("latin-1"), True


def compile_pattern(pattern, *, regex=False, case=False):
    flags = 0 if case else re.IGNORECASE
    return re.compile(pattern if regex else re.escape(pattern), flags)


def search_all(order, parts, rx, cap=5000):
    """Yield dicts: {name, binary, start, length, line, preview} across all parts."""
    hits = []
    for name in order:
        text, binary = decode_for_search(name, parts[name])
        for m in rx.finditer(text):
            start = m.start()
            line = text.count("\n", 0, start) + 1
            ls = text.rfind("\n", 0, start) + 1
            le = text.find("\n", start)
            le = len(text) if le < 0 else le
            hits.append({
                "name": name, "binary": binary, "start": start,
                "length": max(1, m.end() - start), "line": line,
                "preview": text[ls:le].strip()[:140],
            })
            if len(hits) >= cap:
                return hits
    return hits


def hex_preview(data, limit=4096):
    out, n = [], min(len(data), limit)
    for off in range(0, n, 16):
        chunk = data[off:off + 16]
        hexs = " ".join(f"{b:02x}" for b in chunk)
        text = "".join(chr(b) if 32 <= b < 127 else "." for b in chunk)
        out.append(f"{off:08x}  {hexs:<47}  {text}")
    header = f"[binary part: {len(data)} bytes -- read only]\n\n"
    if len(data) > limit:
        out.append(f"\n... {len(data) - limit} more bytes not shown ...")
    return header + "\n".join(out)


HEX_HEADER_LINES = 2  # hex_preview: header line + blank line before row 0


def selftest():
    import tempfile
    here = os.path.abspath(os.path.dirname(__file__))
    cand = [f for f in os.listdir(here)
            if f.lower().endswith((".docx", ".dotx", ".docm", ".dotm"))]
    if not cand:
        print("selftest: no .docx/.dotx in folder to test with"); return 1
    src = os.path.join(here, cand[0])
    order, parts = read_package(src)

    tmp = os.path.join(tempfile.gettempdir(), "_dte_selftest.docx")
    write_package(tmp, order, parts)
    order2, parts2 = read_package(tmp)
    os.remove(tmp)
    assert order == order2, "part order changed"
    assert all(parts[n] == parts2[n] for n in order), "content changed"

    rx = compile_pattern("xml", regex=False, case=False)
    hits = search_all(order, parts, rx)
    assert hits, "search_all found nothing for 'xml'"
    # offset accuracy: the matched slice must equal the pattern (case-insensitive)
    for h in hits[:20]:
        text, _ = decode_for_search(h["name"], parts[h["name"]])
        assert text[h["start"]:h["start"] + h["length"]].lower() == "xml"
    print(f"selftest OK: {os.path.basename(src)} -> {len(order)} parts round-tripped; "
          f"search found {len(hits)} 'xml' matches with correct offsets")
    return 0


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------
def launch(initial_path=None):
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, simpledialog
    import xml.dom.minidom as minidom

    class Editor(tk.Tk):
        def __init__(self):
            super().__init__()
            self.geometry("1200x760")
            self.zip_path = None
            self.order = []
            self.parts = {}
            self.dirty = set()
            self.current = None
            self.part_to_node = {}
            self._matches = []       # (start,end) in current editor, for find-current
            self._match_idx = -1
            self._results = []       # dicts from search_all, aligned with results rows
            self._results_shown = False
            self._build_ui()
            self.protocol("WM_DELETE_WINDOW", self.on_close)
            self._retitle()

        # ---- UI --------------------------------------------------------
        def _build_ui(self):
            bar = ttk.Frame(self)
            bar.pack(side=tk.TOP, fill=tk.X, padx=4, pady=4)
            for label, cmd in [
                ("Open", self.on_open), ("Save", self.on_save),
                ("Save As", self.on_save_as), ("Format XML", self.on_format),
                ("Find", self.show_find), ("Add Part", self.on_add),
                ("Delete Part", self.on_delete), ("Rename Part", self.on_rename),
            ]:
                ttk.Button(bar, text=label, command=cmd).pack(side=tk.LEFT, padx=2)

            # --- find bar (hidden until Ctrl+F / Find) ---
            self.findbar = ttk.Frame(self)
            ttk.Label(self.findbar, text="Find:").pack(side=tk.LEFT, padx=(6, 2))
            self.find_var = tk.StringVar()
            fe = ttk.Entry(self.findbar, textvariable=self.find_var, width=32)
            fe.pack(side=tk.LEFT, padx=2)
            self.find_entry = fe
            fe.bind("<Return>", lambda e: self.find_current(1))
            fe.bind("<Shift-Return>", lambda e: self.find_current(-1))
            fe.bind("<Escape>", lambda e: self.hide_find())
            self.opt_case = tk.BooleanVar(value=False)
            self.opt_regex = tk.BooleanVar(value=False)
            ttk.Checkbutton(self.findbar, text="Aa", variable=self.opt_case).pack(side=tk.LEFT)
            ttk.Checkbutton(self.findbar, text=".*", variable=self.opt_regex).pack(side=tk.LEFT)
            ttk.Button(self.findbar, text="Prev", width=6,
                       command=lambda: self.find_current(-1)).pack(side=tk.LEFT, padx=2)
            ttk.Button(self.findbar, text="Next", width=6,
                       command=lambda: self.find_current(1)).pack(side=tk.LEFT, padx=2)
            ttk.Button(self.findbar, text="All parts",
                       command=self.find_all).pack(side=tk.LEFT, padx=2)
            self.find_count = tk.StringVar(value="")
            ttk.Label(self.findbar, textvariable=self.find_count, width=22).pack(side=tk.LEFT, padx=4)
            ttk.Button(self.findbar, text="x", width=2,
                       command=self.hide_find).pack(side=tk.RIGHT, padx=4)

            paned = ttk.Panedwindow(self, orient=tk.HORIZONTAL)
            paned.pack(fill=tk.BOTH, expand=True, padx=4, pady=(0, 4))

            left = ttk.Frame(paned)
            self.tree = ttk.Treeview(left, show="tree")
            ysb = ttk.Scrollbar(left, orient=tk.VERTICAL, command=self.tree.yview)
            self.tree.configure(yscrollcommand=ysb.set)
            ysb.pack(side=tk.RIGHT, fill=tk.Y)
            self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            self.tree.bind("<<TreeviewSelect>>", self.on_select)
            paned.add(left, weight=1)

            # right side: vertical split (editor over results)
            self.vpaned = ttk.Panedwindow(paned, orient=tk.VERTICAL)
            edit_frame = ttk.Frame(self.vpaned)
            self.editor = tk.Text(edit_frame, wrap="none", undo=True,
                                  font=("Consolas", 10), tabs=("1c",))
            eysb = ttk.Scrollbar(edit_frame, orient=tk.VERTICAL, command=self.editor.yview)
            exsb = ttk.Scrollbar(edit_frame, orient=tk.HORIZONTAL, command=self.editor.xview)
            self.editor.configure(yscrollcommand=eysb.set, xscrollcommand=exsb.set)
            eysb.pack(side=tk.RIGHT, fill=tk.Y)
            exsb.pack(side=tk.BOTTOM, fill=tk.X)
            self.editor.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            self.editor.tag_configure("match", background="#fff3a3")
            self.editor.tag_configure("active", background="#ffb437")
            self.vpaned.add(edit_frame, weight=4)

            # results panel (added to vpaned lazily)
            self.results_frame = ttk.Frame(self.vpaned)
            cols = ("part", "loc", "text")
            self.results = ttk.Treeview(self.results_frame, columns=cols,
                                        show="headings", height=8)
            for c, w, t in (("part", 240, "Part"), ("loc", 80, "Where"),
                            ("text", 640, "Match")):
                self.results.heading(c, text=t)
                self.results.column(c, width=w, anchor="w")
            rsb = ttk.Scrollbar(self.results_frame, orient=tk.VERTICAL,
                                command=self.results.yview)
            self.results.configure(yscrollcommand=rsb.set)
            rsb.pack(side=tk.RIGHT, fill=tk.Y)
            self.results.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            self.results.bind("<Double-1>", self.on_result_activate)
            self.results.bind("<Return>", self.on_result_activate)
            paned.add(self.vpaned, weight=3)

            self.status = tk.StringVar(value="Open a .docx/.dotx to begin.")
            ttk.Label(self, textvariable=self.status, anchor="w",
                      relief=tk.SUNKEN).pack(side=tk.BOTTOM, fill=tk.X)
            self.bind("<Control-s>", lambda e: self.on_save())
            self.bind("<Control-f>", lambda e: self.show_find())

        def _retitle(self):
            name = os.path.basename(self.zip_path) if self.zip_path else "(no file)"
            self.title(f"DOCX Tree Editor - {name}{' *' if self.dirty else ''}")

        # ---- loading ---------------------------------------------------
        def load_file(self, path):
            try:
                self.order, self.parts = read_package(path)
            except Exception as e:  # noqa: BLE001
                messagebox.showerror("Open failed", str(e)); return
            self.zip_path = path
            self.dirty.clear()
            self.current = None
            self.editor.config(state="normal")
            self.editor.delete("1.0", "end")
            self._populate_tree()
            self._retitle()
            self.status.set(f"{len(self.order)} parts.")
            if "word/document.xml" in self.part_to_node:
                self.tree.selection_set(self.part_to_node["word/document.xml"])

        def _populate_tree(self):
            self.tree.delete(*self.tree.get_children())
            self.part_to_node = {}
            folders = {}

            def ensure_folder(path):
                if path == "":
                    return ""
                if path in folders:
                    return folders[path]
                parent, _, nm = path.rpartition("/")
                piid = ensure_folder(parent) if parent else ""
                iid = self.tree.insert(piid, "end", text=nm + "/", open=True)
                folders[path] = iid
                return iid

            for name in self.order:
                parent, _, leaf = name.rpartition("/")
                piid = ensure_folder(parent) if parent else ""
                iid = self.tree.insert(piid, "end",
                                       text=leaf + (" *" if name in self.dirty else ""))
                self.part_to_node[name] = iid

        def _refresh_label(self, name):
            iid = self.part_to_node.get(name)
            if iid:
                leaf = name.rpartition("/")[2]
                self.tree.item(iid, text=leaf + (" *" if name in self.dirty else ""))

        # ---- editor state ----------------------------------------------
        def commit_current(self):
            if self.current is None:
                return
            if is_binary(self.current, self.parts.get(self.current, b"")):
                return
            newb = self.editor.get("1.0", "end-1c").encode("utf-8")
            if newb != self.parts.get(self.current):
                self.parts[self.current] = newb
                self.dirty.add(self.current)
                self._refresh_label(self.current)
                self._retitle()

        def _show_part(self, name):
            if name not in self.parts:
                return
            if name != self.current:
                self.commit_current()
            self.current = name
            data = self.parts[name]
            self.editor.config(state="normal")
            self.editor.delete("1.0", "end")
            self._matches, self._match_idx = [], -1
            if is_binary(name, data):
                self.editor.insert("1.0", hex_preview(data))
                self.editor.config(state="disabled")
                self.status.set(f"{name}  --  {len(data)} bytes (binary, read-only)")
            else:
                self.editor.insert("1.0", data.decode("utf-8"))
                self.editor.edit_reset()
                self.status.set(f"{name}  --  {len(data)} bytes")
            iid = self.part_to_node.get(name)
            if iid and self.tree.selection() != (iid,):
                self.tree.selection_set(iid)
                self.tree.see(iid)

        def on_select(self, _e=None):
            sel = self.tree.selection()
            if not sel:
                return
            name = next((n for n, i in self.part_to_node.items() if i == sel[0]), None)
            if name is None or name == self.current:
                return
            self._show_part(name)

        # ---- find: current part ---------------------------------------
        def show_find(self):
            self.findbar.pack(after=self.children_toolbar(), fill=tk.X)
            self.find_entry.focus_set()
            try:
                sel = self.editor.get("sel.first", "sel.last")
                if sel:
                    self.find_var.set(sel)
            except tk.TclError:
                pass
            self.find_entry.select_range(0, "end")

        def children_toolbar(self):
            # the toolbar frame is the first child packed at top
            return self.pack_slaves()[0]

        def hide_find(self):
            self.findbar.pack_forget()
            self.editor.tag_remove("match", "1.0", "end")
            self.editor.tag_remove("active", "1.0", "end")

        def _compile(self):
            pat = self.find_var.get()
            if not pat:
                return None
            try:
                return compile_pattern(pat, regex=self.opt_regex.get(),
                                       case=self.opt_case.get())
            except re.error as e:
                self.status.set(f"bad regex: {e}")
                return None

        def find_current(self, direction=1):
            rx = self._compile()
            if not rx or self.current is None:
                return
            text = self.editor.get("1.0", "end-1c")
            self.editor.tag_remove("match", "1.0", "end")
            self.editor.tag_remove("active", "1.0", "end")
            self._matches = [(m.start(), m.end()) for m in rx.finditer(text)]
            for s, e in self._matches:
                self.editor.tag_add("match", f"1.0+{s}c", f"1.0+{e}c")
            if not self._matches:
                self.find_count.set("0 matches")
                self._match_idx = -1
                return
            cur = len(self.editor.get("1.0", "insert"))
            if direction > 0:
                idx = next((i for i, (s, _) in enumerate(self._matches) if s >= cur), 0)
            else:
                prev = [i for i, (s, _) in enumerate(self._matches) if s < cur]
                idx = prev[-1] if prev else len(self._matches) - 1
            self._goto_match(idx)

        def _goto_match(self, idx):
            if not self._matches:
                return
            idx %= len(self._matches)
            self._match_idx = idx
            s, e = self._matches[idx]
            self.editor.tag_remove("active", "1.0", "end")
            self.editor.tag_add("active", f"1.0+{s}c", f"1.0+{e}c")
            self.editor.mark_set("insert", f"1.0+{e}c")
            self.editor.see(f"1.0+{s}c")
            self.find_count.set(f"{idx + 1} / {len(self._matches)}")

        # ---- find: all parts ------------------------------------------
        def find_all(self):
            rx = self._compile()
            if not rx or not self.order:
                return
            self.commit_current()
            self.results.delete(*self.results.get_children())
            self._results = search_all(self.order, self.parts, rx)
            for h in self._results:
                loc = f"line {h['line']}" if not h["binary"] else f"off {h['start']}"
                self.results.insert("", "end", values=(h["name"], loc, h["preview"]))
            files = len({h["name"] for h in self._results})
            self.find_count.set(f"{len(self._results)} in {files} part(s)")
            self._show_results(True)
            if not self._results:
                self.status.set("No matches across parts.")

        def _show_results(self, show):
            if show and not self._results_shown:
                self.vpaned.add(self.results_frame, weight=2)
                self._results_shown = True
            elif not show and self._results_shown:
                self.vpaned.forget(self.results_frame)
                self._results_shown = False

        def on_result_activate(self, _e=None):
            sel = self.results.selection()
            if not sel:
                return
            row = sel[0]
            idx = self.results.index(row)
            if idx >= len(self._results):
                return
            h = self._results[idx]
            self._show_part(h["name"])
            if h["binary"]:
                line = HEX_HEADER_LINES + 1 + h["start"] // 16
                self.editor.see(f"{line}.0")
                self.status.set(f"{h['name']} @ byte offset {h['start']} (binary)")
            else:
                s, e = h["start"], h["start"] + h["length"]
                self.editor.tag_remove("active", "1.0", "end")
                self.editor.tag_add("active", f"1.0+{s}c", f"1.0+{e}c")
                self.editor.mark_set("insert", f"1.0+{e}c")
                self.editor.see(f"1.0+{s}c")

        # ---- toolbar: file & parts ------------------------------------
        def on_open(self):
            if not self._confirm_discard():
                return
            path = filedialog.askopenfilename(
                filetypes=[("Word packages", "*.docx *.dotx *.docm *.dotm"),
                           ("Zip", "*.zip"), ("All files", "*.*")])
            if path:
                self.load_file(path)

        def _write_to(self, path):
            self.commit_current()
            try:
                write_package(path, self.order, self.parts)
            except PermissionError:
                messagebox.showerror("Save failed",
                                     "Permission denied. The file is probably open "
                                     "in Word -- close it and try again.")
                return False
            except Exception as e:  # noqa: BLE001
                messagebox.showerror("Save failed", str(e)); return False
            self.zip_path = path
            self.dirty.clear()
            self._populate_tree()
            if self.current in self.part_to_node:
                self.tree.selection_set(self.part_to_node[self.current])
            self._retitle()
            self.status.set(f"Saved {path}")
            return True

        def on_save(self):
            if not self.zip_path:
                return self.on_save_as()
            self._write_to(self.zip_path)

        def on_save_as(self):
            if not self.order:
                return
            path = filedialog.asksaveasfilename(
                defaultextension=os.path.splitext(self.zip_path or ".docx")[1] or ".docx",
                initialfile=os.path.basename(self.zip_path) if self.zip_path else "edited.docx")
            if path:
                self._write_to(path)

        def on_format(self):
            if self.current is None or is_binary(self.current, self.parts[self.current]):
                return
            raw = self.editor.get("1.0", "end-1c")
            try:
                pretty = minidom.parseString(raw).toprettyxml(indent="  ")
                pretty = "\n".join(ln for ln in pretty.splitlines() if ln.strip())
            except Exception as e:  # noqa: BLE001
                messagebox.showerror("Not well-formed XML", str(e)); return
            self.editor.delete("1.0", "end")
            self.editor.insert("1.0", pretty)
            self.status.set("Formatted for reading. NOTE: pretty-printing can add "
                            "whitespace between inline runs -- check w:t before saving.")

        def on_add(self):
            if not self.order:
                messagebox.showinfo("Add Part", "Open a package first."); return
            name = simpledialog.askstring(
                "Add Part", "New part path (e.g. word/glossary/document.xml):")
            if not name:
                return
            name = name.strip().lstrip("/")
            if name in self.parts:
                messagebox.showerror("Add Part", "That part already exists."); return
            self.parts[name] = b""
            self.order.append(name)
            self.dirty.add(name)
            self._populate_tree()
            self.tree.selection_set(self.part_to_node[name])
            self._retitle()

        def on_delete(self):
            if self.current is None:
                return
            if not messagebox.askyesno("Delete Part", f"Delete {self.current}?"):
                return
            name = self.current
            self.order.remove(name)
            self.parts.pop(name, None)
            self.dirty.discard(name)
            self.current = None
            self.editor.config(state="normal")
            self.editor.delete("1.0", "end")
            self._populate_tree()
            self._retitle()
            self.status.set(f"Deleted {name} (Save to persist)")

        def on_rename(self):
            if self.current is None:
                return
            new = simpledialog.askstring("Rename Part", "New path:",
                                         initialvalue=self.current)
            if not new or new == self.current:
                return
            new = new.strip().lstrip("/")
            if new in self.parts:
                messagebox.showerror("Rename Part", "Target already exists."); return
            idx = self.order.index(self.current)
            self.order[idx] = new
            self.parts[new] = self.parts.pop(self.current)
            self.dirty.discard(self.current)
            self.dirty.add(new)
            self.current = new
            self._populate_tree()
            self.tree.selection_set(self.part_to_node[new])
            self._retitle()

        # ---- close -----------------------------------------------------
        def _confirm_discard(self):
            if not self.dirty:
                return True
            return messagebox.askyesno("Unsaved changes", "Discard unsaved changes?")

        def on_close(self):
            self.commit_current()
            if self._confirm_discard():
                self.destroy()

    app = Editor()
    if initial_path:
        app.load_file(initial_path)
    app.mainloop()


def main(argv):
    if "--selftest" in argv:
        return selftest()
    path = next((a for a in argv[1:] if not a.startswith("-")), None)
    if path and not os.path.exists(path):
        print(f"No such file: {path}"); return 1
    launch(path)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
