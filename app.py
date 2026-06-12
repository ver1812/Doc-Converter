import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
from PIL import Image, ImageDraw, ImageFont, ImageTk
import fitz  # PyMuPDF

from converters.pdf_converter import convert_pdf_to_docx
from converters.image_converter import convert_image
from converters.pdf_compressor import compress_pdf

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

SIDEBAR_W = 210
ACCENT    = "#1f6feb"
SIDEBAR_BG = "#1a1a2e"
CARD_BG    = "#16213e"
TOOL_BG    = "#0f0f28"

NAV_ITEMS = [
    ("documents", "  PDF → Word"),
    ("images",    "  Image Converter"),
    ("compress",  "  Compress PDF"),
    ("editor",    "  Edit PDF"),
    ("viewer",    "  PDF Viewer"),
]

_FONT_PATHS = [
    "/System/Library/Fonts/Helvetica.ttc",
    "/System/Library/Fonts/SFNSText.ttf",
    "/Library/Fonts/Arial.ttf",
    "/System/Library/Fonts/Supplemental/Arial.ttf",
]


def _pil_font(size: int) -> ImageFont.FreeTypeFont:
    for fp in _FONT_PATHS:
        try:
            return ImageFont.truetype(fp, max(6, int(size)))
        except Exception:
            pass
    return ImageFont.load_default()


# ══════════════════════════════════════════════════════════════════════════════


class DocConverterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Doc Converter")
        self.geometry("1120x740")
        self.minsize(960, 640)

        # ── Editor state ──────────────────────────────────────
        self._ed_doc: fitz.Document | None = None
        self._ed_src: str = ""
        self._ed_page: int = 0
        self._ed_scale: float = 2.5
        self._ed_annots: dict[int, list] = {}
        self._ed_mode: str = "highlight_yellow"
        self._ed_drag_start: tuple | None = None
        self._ed_drag_rect = None
        self._ed_photo = None          # prevent GC
        self._ed_resize_job = None

        # ── Viewer state ──────────────────────────────────────
        self._vw_doc: fitz.Document | None = None
        self._vw_page: int = 0
        self._vw_book: bool = False
        self._vw_cream: bool = False
        self._vw_photo = None
        self._vw_zoom: float = 1.0
        self._vw_resize_job = None

        self._build_layout()
        self._show_frame("documents")

    # ═══════════════════════════════════════ Layout ══════════

    def _build_layout(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._sidebar = ctk.CTkFrame(self, width=SIDEBAR_W, corner_radius=0, fg_color=SIDEBAR_BG)
        self._sidebar.grid(row=0, column=0, sticky="nsew")
        self._sidebar.grid_propagate(False)
        self._create_sidebar()

        self._content = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self._content.grid(row=0, column=1, sticky="nsew")
        self._content.grid_columnconfigure(0, weight=1)
        self._content.grid_rowconfigure(0, weight=1)

        self._frames: dict[str, ctk.CTkFrame] = {}
        self._frames["documents"] = self._create_document_frame()
        self._frames["images"]    = self._create_image_frame()
        self._frames["compress"]  = self._create_compress_frame()
        self._frames["editor"]    = self._create_editor_frame()
        self._frames["viewer"]    = self._create_viewer_frame()

    def _create_sidebar(self):
        self._sidebar.grid_columnconfigure(0, weight=1)
        self._sidebar.grid_rowconfigure(len(NAV_ITEMS) + 3, weight=1)

        ctk.CTkLabel(
            self._sidebar, text="Doc Converter",
            font=ctk.CTkFont(size=17, weight="bold"), text_color="#e0e0e0",
        ).grid(row=0, column=0, padx=20, pady=(28, 14), sticky="w")

        ctk.CTkFrame(self._sidebar, height=1, fg_color="#333355").grid(
            row=1, column=0, sticky="ew", padx=14, pady=(0, 10)
        )

        self._nav_btns: dict[str, ctk.CTkButton] = {}
        for i, (key, label) in enumerate(NAV_ITEMS, start=2):
            btn = ctk.CTkButton(
                self._sidebar, text=label, anchor="w",
                height=40, corner_radius=8,
                fg_color="transparent", hover_color="#2a2a4a",
                text_color="#c0c0d0", font=ctk.CTkFont(size=13),
                command=lambda k=key: self._show_frame(k),
            )
            btn.grid(row=i, column=0, padx=10, pady=2, sticky="ew")
            self._nav_btns[key] = btn

        ctk.CTkLabel(
            self._sidebar, text="v1.1.0",
            text_color="#555577", font=ctk.CTkFont(size=11),
        ).grid(row=len(NAV_ITEMS) + 4, column=0, padx=20, pady=14, sticky="sw")

    def _show_frame(self, name: str):
        for frame in self._frames.values():
            frame.grid_remove()
        pad = 0 if name == "viewer" else (8 if name == "editor" else 30)
        self._frames[name].grid(row=0, column=0, sticky="nsew", padx=pad, pady=pad)
        for key, btn in self._nav_btns.items():
            btn.configure(
                fg_color=ACCENT if key == name else "transparent",
                text_color="white" if key == name else "#c0c0d0",
            )
        if name == "viewer" and not self._vw_doc:
            self.after(50, self._vw_draw_placeholder)

    # ═══════════════════════════════════════ Shared helpers ══

    def _section_header(self, frame: ctk.CTkFrame, title: str, subtitle: str):
        frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(frame, text=title, font=ctk.CTkFont(size=22, weight="bold")).grid(
            row=0, column=0, pady=(28, 4), sticky="w", padx=28
        )
        ctk.CTkLabel(frame, text=subtitle, text_color="#8888aa",
                      font=ctk.CTkFont(size=13)).grid(
            row=1, column=0, pady=(0, 18), sticky="w", padx=28
        )
        ctk.CTkFrame(frame, height=1, fg_color="#2a2a4a").grid(
            row=2, column=0, sticky="ew", padx=28, pady=(0, 18)
        )

    def _create_file_row(self, parent, row, label, var,
                          btn_text, select_dir=False,
                          filetypes=None, auto_dir_var=None):
        rf = ctk.CTkFrame(parent, fg_color="transparent")
        rf.grid(row=row, column=0, padx=28, pady=(6, 4), sticky="ew")
        rf.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(rf, text=label, font=ctk.CTkFont(size=13),
                      width=120, anchor="w").grid(row=0, column=0, sticky="w")
        ctk.CTkEntry(rf, textvariable=var, placeholder_text="Not selected",
                      height=36).grid(row=0, column=1, sticky="ew", padx=(12, 8))

        def browse():
            path = (filedialog.askdirectory(title="Select Folder") if select_dir
                    else filedialog.askopenfilename(
                        title="Select File",
                        filetypes=filetypes or [("All files", "*.*")]
                    ))
            if path:
                var.set(path)
                if not select_dir and auto_dir_var and not auto_dir_var.get():
                    auto_dir_var.set(os.path.dirname(path))

        ctk.CTkButton(rf, text=btn_text, width=130, height=36,
                       command=browse).grid(row=0, column=2)

    def _make_progress_row(self, parent, btn_row, prog_row, status_row, btn_text, cmd):
        btn = ctk.CTkButton(
            parent, text=btn_text, height=44, corner_radius=8,
            font=ctk.CTkFont(size=14, weight="bold"), command=cmd,
        )
        btn.grid(row=btn_row, column=0, padx=28, pady=(22, 10), sticky="ew")

        prog = ctk.CTkProgressBar(parent, height=8, corner_radius=4)
        prog.set(0)
        prog.grid(row=prog_row, column=0, padx=28, pady=(0, 6), sticky="ew")

        status = ctk.CTkLabel(parent, text="", text_color="#8888aa",
                               font=ctk.CTkFont(size=12))
        status.grid(row=status_row, column=0, padx=28, pady=(0, 24), sticky="w")
        return btn, prog, status

    # ═══════════════════════════════════════ Documents tab ═══

    def _create_document_frame(self) -> ctk.CTkFrame:
        f = ctk.CTkFrame(self._content, corner_radius=12, fg_color=CARD_BG)
        self._section_header(f, "PDF  →  Word",
                              "Convert PDF documents into editable Word (.docx) files.")

        self._pdf_path = ctk.StringVar()
        self._pdf_out_dir = ctk.StringVar()
        self._create_file_row(f, 3, "PDF File", self._pdf_path, "Browse PDF",
                               filetypes=[("PDF files", "*.pdf")],
                               auto_dir_var=self._pdf_out_dir)
        self._create_file_row(f, 5, "Output Folder", self._pdf_out_dir,
                               "Browse Folder", select_dir=True)

        self._pdf_btn, self._pdf_prog, self._pdf_status = \
            self._make_progress_row(f, 7, 8, 9, "Convert to Word",
                                    self._start_pdf_conversion)
        return f

    def _start_pdf_conversion(self):
        src = self._pdf_path.get().strip()
        out_dir = self._pdf_out_dir.get().strip() or (os.path.dirname(src) if src else "")
        if not src or not os.path.isfile(src):
            self._pdf_status.configure(text="Please select a valid PDF.", text_color="#ff6b6b")
            return
        dest = os.path.join(out_dir, os.path.splitext(os.path.basename(src))[0] + ".docx")
        self._pdf_btn.configure(state="disabled")
        self._pdf_prog.set(0); self._pdf_prog.start()
        self._pdf_status.configure(text="Converting…", text_color="#8888aa")

        def task():
            try:
                convert_pdf_to_docx(src, dest)
                self.after(0, lambda: _done(dest, None))
            except Exception as e:
                self.after(0, lambda err=e: _done(None, err))

        def _done(dest, err):
            self._pdf_prog.stop(); self._pdf_prog.set(0 if err else 1)
            self._pdf_btn.configure(state="normal")
            if err:
                self._pdf_status.configure(text=f"Error: {err}", text_color="#ff6b6b")
            else:
                self._pdf_status.configure(text=f"Saved: {dest}", text_color="#4ecca3")

        threading.Thread(target=task, daemon=True).start()

    # ═══════════════════════════════════════ Images tab ══════

    def _create_image_frame(self) -> ctk.CTkFrame:
        f = ctk.CTkFrame(self._content, corner_radius=12, fg_color=CARD_BG)
        self._section_header(f, "Image Converter",
                              "Convert between JPEG, PNG, and HEIC formats.")

        self._img_path = ctk.StringVar()
        self._img_out_dir = ctk.StringVar()
        self._create_file_row(f, 3, "Image File", self._img_path, "Browse Image",
                               filetypes=[("Images", "*.jpg *.jpeg *.png *.heic *.heif")],
                               auto_dir_var=self._img_out_dir)
        self._create_file_row(f, 5, "Output Folder", self._img_out_dir,
                               "Browse Folder", select_dir=True)

        fmt_row = ctk.CTkFrame(f, fg_color="transparent")
        fmt_row.grid(row=6, column=0, padx=28, pady=(10, 0), sticky="ew")
        fmt_row.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(fmt_row, text="Output Format", font=ctk.CTkFont(size=13),
                      width=120, anchor="w").grid(row=0, column=0, sticky="w")
        self._img_fmt = ctk.StringVar(value="PNG")
        ctk.CTkOptionMenu(fmt_row, values=["PNG", "JPEG", "HEIC"],
                           variable=self._img_fmt, width=160).grid(
            row=0, column=1, sticky="w", padx=(12, 0)
        )

        self._img_btn, self._img_prog, self._img_status = \
            self._make_progress_row(f, 8, 9, 10, "Convert Image",
                                    self._start_image_conversion)
        return f

    def _start_image_conversion(self):
        src = self._img_path.get().strip()
        out_dir = self._img_out_dir.get().strip() or (os.path.dirname(src) if src else "")
        fmt = self._img_fmt.get()
        if not src or not os.path.isfile(src):
            self._img_status.configure(text="Please select a valid image.", text_color="#ff6b6b")
            return
        ext = {"JPEG": ".jpg", "PNG": ".png", "HEIC": ".heic"}[fmt]
        dest = os.path.join(out_dir, os.path.splitext(os.path.basename(src))[0] + ext)
        self._img_btn.configure(state="disabled")
        self._img_prog.set(0); self._img_prog.start()
        self._img_status.configure(text="Converting…", text_color="#8888aa")

        def task():
            try:
                convert_image(src, dest, fmt)
                self.after(0, lambda: _done(dest, None))
            except Exception as e:
                self.after(0, lambda err=e: _done(None, err))

        def _done(dest, err):
            self._img_prog.stop(); self._img_prog.set(0 if err else 1)
            self._img_btn.configure(state="normal")
            if err:
                self._img_status.configure(text=f"Error: {err}", text_color="#ff6b6b")
            else:
                self._img_status.configure(text=f"Saved: {dest}", text_color="#4ecca3")

        threading.Thread(target=task, daemon=True).start()

    # ═══════════════════════════════════════ Compress tab ════

    def _create_compress_frame(self) -> ctk.CTkFrame:
        f = ctk.CTkFrame(self._content, corner_radius=12, fg_color=CARD_BG)
        self._section_header(f, "Compress PDF",
                              "Reduce PDF file size by cleaning redundant data and compressing streams.")

        self._cmp_path = ctk.StringVar()
        self._cmp_out_dir = ctk.StringVar()
        self._create_file_row(f, 3, "PDF File", self._cmp_path, "Browse PDF",
                               filetypes=[("PDF files", "*.pdf")],
                               auto_dir_var=self._cmp_out_dir)
        self._create_file_row(f, 5, "Output Folder", self._cmp_out_dir,
                               "Browse Folder", select_dir=True)

        lvl_f = ctk.CTkFrame(f, fg_color="transparent")
        lvl_f.grid(row=6, column=0, padx=28, pady=(14, 0), sticky="ew")
        lvl_f.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(lvl_f, text="Compression", font=ctk.CTkFont(size=13),
                      width=120, anchor="w").grid(row=0, column=0, sticky="w")

        self._cmp_level = ctk.StringVar(value="Medium")
        seg = ctk.CTkSegmentedButton(
            lvl_f, values=["Small", "Medium", "Full Compress"],
            variable=self._cmp_level, width=340,
        )
        seg.set("Medium")
        seg.grid(row=0, column=1, sticky="w", padx=(12, 0))

        ctk.CTkLabel(
            f,
            text="Small: basic cleanup   ·   Medium: + stream & image compression"
                 "   ·   Full Compress: maximum — removes all dead data",
            text_color="#555577", font=ctk.CTkFont(size=11),
        ).grid(row=7, column=0, padx=28, pady=(6, 0), sticky="w")

        self._cmp_btn, self._cmp_prog, self._cmp_status = \
            self._make_progress_row(f, 8, 9, 10, "Compress PDF",
                                    self._start_compress)
        return f

    def _start_compress(self):
        src = self._cmp_path.get().strip()
        out_dir = self._cmp_out_dir.get().strip() or (os.path.dirname(src) if src else "")
        level_map = {"Small": "small", "Medium": "medium", "Full Compress": "full"}
        level = level_map[self._cmp_level.get()]
        if not src or not os.path.isfile(src):
            self._cmp_status.configure(text="Please select a valid PDF.", text_color="#ff6b6b")
            return
        base = os.path.splitext(os.path.basename(src))[0]
        dest = os.path.join(out_dir, f"{base}_compressed_{level}.pdf")
        self._cmp_btn.configure(state="disabled")
        self._cmp_prog.set(0); self._cmp_prog.start()
        self._cmp_status.configure(text="Compressing…", text_color="#8888aa")

        def task():
            try:
                result = compress_pdf(src, dest, level)
                self.after(0, lambda r=result: _done(dest, r, None))
            except Exception as e:
                self.after(0, lambda err=e: _done(None, None, err))

        def _done(dest, result, err):
            self._cmp_prog.stop(); self._cmp_prog.set(0 if err else 1)
            self._cmp_btn.configure(state="normal")
            if err:
                self._cmp_status.configure(text=f"Error: {err}", text_color="#ff6b6b")
            else:
                orig = result["original_kb"]; comp = result["compressed_kb"]
                pct  = result["saved_pct"]
                self._cmp_status.configure(
                    text=f"Saved: {dest}\n{orig} KB  →  {comp} KB  ({pct}% smaller)",
                    text_color="#4ecca3",
                )

        threading.Thread(target=task, daemon=True).start()

    # ═══════════════════════════════════════ Editor tab ══════

    def _create_editor_frame(self) -> ctk.CTkFrame:
        outer = ctk.CTkFrame(self._content, corner_radius=12, fg_color=CARD_BG)
        outer.grid_columnconfigure(1, weight=1)
        outer.grid_rowconfigure(1, weight=1)

        # ── Top file bar ─────────────────────────────────────
        top = ctk.CTkFrame(outer, fg_color="transparent")
        top.grid(row=0, column=0, columnspan=2, padx=12, pady=(12, 0), sticky="ew")
        top.grid_columnconfigure(2, weight=1)

        ctk.CTkLabel(top, text="Edit PDF",
                      font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, padx=(8, 14))

        self._ed_path_var = ctk.StringVar()
        ctk.CTkEntry(top, textvariable=self._ed_path_var,
                      placeholder_text="Select a PDF to edit…", height=34).grid(
            row=0, column=2, sticky="ew", padx=(0, 8))

        def _browse_ed():
            p = filedialog.askopenfilename(title="Open PDF",
                                           filetypes=[("PDF files", "*.pdf")])
            if p:
                self._ed_path_var.set(p)

        ctk.CTkButton(top, text="Browse", width=80, height=34,
                       command=_browse_ed).grid(row=0, column=3, padx=(0, 6))
        ctk.CTkButton(top, text="Load", width=70, height=34,
                       fg_color=ACCENT,
                       command=self._ed_load).grid(row=0, column=4, padx=(0, 6))

        # Page nav
        nav = ctk.CTkFrame(top, fg_color="transparent")
        nav.grid(row=0, column=5, padx=(8, 8))
        ctk.CTkButton(nav, text="◀", width=30, height=34,
                       command=lambda: self._ed_nav(-1)).grid(row=0, column=0)
        self._ed_page_lbl = ctk.CTkLabel(nav, text="–", width=80,
                                          font=ctk.CTkFont(size=11))
        self._ed_page_lbl.grid(row=0, column=1, padx=4)
        ctk.CTkButton(nav, text="▶", width=30, height=34,
                       command=lambda: self._ed_nav(1)).grid(row=0, column=2)

        # ── Tool panel (left) ────────────────────────────────
        tools = ctk.CTkFrame(outer, fg_color=TOOL_BG, corner_radius=8, width=136)
        tools.grid(row=1, column=0, sticky="ns", padx=(12, 0), pady=8)
        tools.grid_propagate(False)
        tools.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(tools, text="TOOLS", font=ctk.CTkFont(size=10, weight="bold"),
                      text_color="#555577").grid(row=0, column=0, pady=(12, 6))

        self._ed_tool_btns: dict[str, ctk.CTkButton] = {}
        _row = 1

        for mode, label in [("edit_text", "✏  Edit Text"), ("add_text", "＋  Add Text")]:
            b = ctk.CTkButton(
                tools, text=label, width=116, height=34, anchor="w",
                fg_color="transparent", hover_color="#2a2a4a",
                command=lambda m=mode: self._ed_set_mode(m))
            b.grid(row=_row, column=0, padx=8, pady=2)
            self._ed_tool_btns[mode] = b
            _row += 1

        ctk.CTkFrame(tools, height=1, fg_color="#2a2a4a").grid(
            row=_row, column=0, sticky="ew", padx=8, pady=6)
        _row += 1
        ctk.CTkLabel(tools, text="HIGHLIGHT", font=ctk.CTkFont(size=10, weight="bold"),
                      text_color="#555577").grid(row=_row, column=0, pady=(0, 4))
        _row += 1

        _hl = [
            ("highlight_yellow", "●  Yellow",  "#d4d400", "#111111"),
            ("highlight_green",  "●  Green",   "#00c060", "#111111"),
            ("highlight_pink",   "●  Pink",    "#e820c0", "#ffffff"),
        ]
        for mode, label, bg, fg in _hl:
            b = ctk.CTkButton(
                tools, text=label, width=116, height=34, anchor="w",
                fg_color=bg, hover_color=bg, text_color=fg,
                command=lambda m=mode: self._ed_set_mode(m))
            b.grid(row=_row, column=0, padx=8, pady=2)
            self._ed_tool_btns[mode] = b
            _row += 1

        ctk.CTkFrame(tools, height=1, fg_color="#2a2a4a").grid(
            row=_row, column=0, sticky="ew", padx=8, pady=6)
        _row += 1

        b = ctk.CTkButton(
            tools, text="◻  Whiteout", width=116, height=34, anchor="w",
            fg_color="#e8e8e8", hover_color="#cccccc", text_color="#111111",
            command=lambda: self._ed_set_mode("whiteout"))
        b.grid(row=_row, column=0, padx=8, pady=2)
        self._ed_tool_btns["whiteout"] = b
        _row += 1

        ctk.CTkFrame(tools, height=1, fg_color="#2a2a4a").grid(
            row=_row, column=0, sticky="ew", padx=8, pady=6)
        _row += 1
        tools.grid_rowconfigure(_row, weight=1)
        _row += 1

        ctk.CTkButton(tools, text="↩  Undo", width=116, height=34,
                       fg_color="#2a2a5a", command=self._ed_undo).grid(
            row=_row, column=0, padx=8, pady=2)
        _row += 1
        ctk.CTkButton(tools, text="💾  Save PDF", width=116, height=34,
                       fg_color=ACCENT, command=self._ed_save).grid(
            row=_row, column=0, padx=8, pady=(2, 14))

        # ── Canvas area ───────────────────────────────────────
        ch = ctk.CTkFrame(outer, fg_color="#252545", corner_radius=8)
        ch.grid(row=1, column=1, sticky="nsew", padx=(8, 12), pady=8)
        ch.grid_rowconfigure(0, weight=1)
        ch.grid_columnconfigure(0, weight=1)

        self._ed_canvas = tk.Canvas(ch, bg="#2e2e50", highlightthickness=0,
                                     cursor="crosshair")
        ed_vsb = tk.Scrollbar(ch, orient="vertical",   command=self._ed_canvas.yview)
        ed_hsb = tk.Scrollbar(ch, orient="horizontal", command=self._ed_canvas.xview)
        self._ed_canvas.configure(yscrollcommand=ed_vsb.set,
                                   xscrollcommand=ed_hsb.set)
        ed_vsb.grid(row=0, column=1, sticky="ns")
        ed_hsb.grid(row=1, column=0, sticky="ew")
        self._ed_canvas.grid(row=0, column=0, sticky="nsew")

        self._ed_canvas.bind("<ButtonPress-1>",   self._ed_press)
        self._ed_canvas.bind("<B1-Motion>",        self._ed_drag)
        self._ed_canvas.bind("<ButtonRelease-1>",  self._ed_release)
        self._ed_canvas.bind("<MouseWheel>", lambda e: self._ed_canvas.yview_scroll(
            -1 if e.delta > 0 else 1, "units"))
        self._ed_canvas.bind("<Configure>", self._ed_on_resize)

        # ── Footer note ───────────────────────────────────────
        ctk.CTkLabel(
            outer,
            text="ⓘ  PDF editing uses overlays — text does not reflow across pages. "
                 "For full paragraph editing, convert to Word first.",
            text_color="#555577", font=ctk.CTkFont(size=11),
        ).grid(row=2, column=0, columnspan=2, padx=16, pady=(0, 8), sticky="w")

        self._ed_set_mode("highlight_yellow")
        return outer

    # ── Editor logic ──────────────────────────────────────────

    def _ed_load(self):
        path = self._ed_path_var.get().strip()
        if not path or not os.path.isfile(path):
            return
        if self._ed_doc:
            self._ed_doc.close()
        self._ed_doc  = fitz.open(path)
        self._ed_src  = path
        self._ed_page = 0
        self._ed_annots = {}
        self._ed_render()

    def _ed_nav(self, delta: int):
        if not self._ed_doc:
            return
        self._ed_page = max(0, min(len(self._ed_doc) - 1, self._ed_page + delta))
        self._ed_render()

    def _ed_render(self):
        if not self._ed_doc:
            return
        page = self._ed_doc[self._ed_page]
        n    = len(self._ed_doc)
        self._ed_page_lbl.configure(text=f"Page {self._ed_page + 1} / {n}")

        # Fit page width to canvas; render at 2× for sharpness then downsample
        cw = max(self._ed_canvas.winfo_width() - 4, 500)
        self._ed_scale = cw / page.rect.width

        render_s = self._ed_scale * 2.0
        pix = page.get_pixmap(matrix=fitz.Matrix(render_s, render_s),
                               alpha=False, colorspace=fitz.csRGB)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        dw = int(page.rect.width  * self._ed_scale)
        dh = int(page.rect.height * self._ed_scale)
        img = img.resize((dw, dh), Image.LANCZOS)

        for annot in self._ed_annots.get(self._ed_page, []):
            img = self._composite(img, annot, self._ed_scale)

        self._ed_photo = ImageTk.PhotoImage(img)
        self._ed_canvas.configure(scrollregion=(0, 0, dw, dh))
        self._ed_canvas.delete("all")
        self._ed_canvas.create_image(0, 0, anchor="nw", image=self._ed_photo)

    def _ed_on_resize(self, _=None):
        if not self._ed_doc:
            return
        if self._ed_resize_job:
            self.after_cancel(self._ed_resize_job)
        self._ed_resize_job = self.after(120, self._ed_render)

    def _composite(self, img: Image.Image, annot: dict, scale: float) -> Image.Image:
        t = annot["type"]
        r = annot.get("rect")

        if t == "highlight" and r:
            x0, y0 = int(r.x0 * scale), int(r.y0 * scale)
            x1, y1 = int(r.x1 * scale), int(r.y1 * scale)
            ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
            c  = annot["color"]
            ImageDraw.Draw(ov).rectangle(
                [x0, y0, x1, y1],
                fill=(int(c[0]*255), int(c[1]*255), int(c[2]*255), 115))
            img = Image.alpha_composite(img.convert("RGBA"), ov).convert("RGB")

        elif t == "whiteout" and r:
            x0, y0 = int(r.x0 * scale), int(r.y0 * scale)
            x1, y1 = int(r.x1 * scale), int(r.y1 * scale)
            ImageDraw.Draw(img).rectangle([x0, y0, x1, y1], fill="white")

        elif t == "replace_text" and r:
            x0, y0 = int(r.x0 * scale), int(r.y0 * scale)
            x1, y1 = int(r.x1 * scale), int(r.y1 * scale)
            d = ImageDraw.Draw(img)
            d.rectangle([x0, y0, x1, y1], fill="white")
            fsize = max(8, int((r.y1 - r.y0) * scale * 0.80))
            d.text((x0 + 1, y0), annot["new_text"],
                   fill="#000000", font=_pil_font(fsize))

        elif t == "add_text":
            x     = int(annot["pdf_x"] * scale)
            y     = int(annot["pdf_y"] * scale)
            fsize = max(8, int(annot.get("fontsize", 12) * scale))
            ImageDraw.Draw(img).text(
                (x, y - fsize), annot["text"],
                fill="#111111", font=_pil_font(fsize))

        return img

    def _ed_set_mode(self, mode: str):
        self._ed_mode = mode
        for m, btn in self._ed_tool_btns.items():
            btn.configure(border_width=2 if m == mode else 0,
                          border_color="white")
        cursor_map = {
            "edit_text": "xterm", "add_text": "plus",
        }
        self._ed_canvas.configure(
            cursor=cursor_map.get(mode, "crosshair"))

    def _ed_press(self, event):
        cx = self._ed_canvas.canvasx(event.x)
        cy = self._ed_canvas.canvasy(event.y)
        mode = self._ed_mode

        if mode in ("highlight_yellow", "highlight_green",
                    "highlight_pink", "whiteout"):
            self._ed_drag_start = (cx, cy)
            outline = {"whiteout": "#ffffff"}.get(mode, "#44aaff")
            self._ed_drag_rect = self._ed_canvas.create_rectangle(
                cx, cy, cx, cy, outline=outline, width=2, dash=(4, 4))
        elif mode == "edit_text":
            self._ed_click_word(cx, cy)
        elif mode == "add_text":
            self._ed_place_text(cx, cy)

    def _ed_drag(self, event):
        if self._ed_drag_rect and self._ed_drag_start:
            cx = self._ed_canvas.canvasx(event.x)
            cy = self._ed_canvas.canvasy(event.y)
            self._ed_canvas.coords(
                self._ed_drag_rect,
                self._ed_drag_start[0], self._ed_drag_start[1], cx, cy)

    def _ed_release(self, event):
        if not self._ed_drag_rect or not self._ed_drag_start:
            return
        cx = self._ed_canvas.canvasx(event.x)
        cy = self._ed_canvas.canvasy(event.y)
        self._ed_canvas.delete(self._ed_drag_rect)
        self._ed_drag_rect = None
        x0, y0 = self._ed_drag_start
        self._ed_drag_start = None

        if abs(cx - x0) < 5 or abs(cy - y0) < 5:
            return

        s = self._ed_scale
        pdf_rect = fitz.Rect(min(x0, cx)/s, min(y0, cy)/s,
                              max(x0, cx)/s, max(y0, cy)/s)
        mode = self._ed_mode
        color_map = {
            "highlight_yellow": (1.00, 0.95, 0.00),
            "highlight_green":  (0.00, 0.80, 0.38),
            "highlight_pink":   (0.95, 0.10, 0.80),
        }
        if mode == "whiteout":
            annot = {"type": "whiteout", "rect": pdf_rect}
        else:
            annot = {"type": "highlight", "rect": pdf_rect,
                     "color": color_map.get(mode, (1, 1, 0))}

        self._ed_annots.setdefault(self._ed_page, []).append(annot)
        self._ed_render()

    def _ed_click_word(self, cx: float, cy: float):
        if not self._ed_doc:
            return
        s = self._ed_scale
        px, py = cx / s, cy / s
        words = self._ed_doc[self._ed_page].get_text("words")
        hit = next((w for w in words if w[0] <= px <= w[2] and w[1] <= py <= w[3]), None)
        if not hit:
            return

        x0, y0, x1, y1, word = hit[0], hit[1], hit[2], hit[3], hit[4]
        width = max(int((x1 - x0) * s) + 50, 90)
        var   = tk.StringVar(value=word)
        entry = ctk.CTkEntry(self._ed_canvas, textvariable=var,
                              width=width, height=26, font=ctk.CTkFont(size=11))
        ew = self._ed_canvas.create_window(
            int(x0 * s), int(y0 * s), anchor="nw", window=entry)

        def confirm(_=None):
            new = var.get().strip()
            self._ed_canvas.delete(ew)
            entry.destroy()
            if new and new != word:
                self._ed_annots.setdefault(self._ed_page, []).append({
                    "type": "replace_text",
                    "rect": fitz.Rect(x0, y0, x1, y1),
                    "old_text": word, "new_text": new,
                })
                self._ed_render()

        entry.bind("<Return>", confirm)
        entry.bind("<Escape>",
                   lambda _: (self._ed_canvas.delete(ew), entry.destroy()))
        entry.focus_set()
        entry.select_range(0, tk.END)

    def _ed_place_text(self, cx: float, cy: float):
        if not self._ed_doc:
            return
        var   = tk.StringVar()
        entry = ctk.CTkEntry(self._ed_canvas, textvariable=var, width=200,
                              height=30, placeholder_text="Type, then Enter…",
                              font=ctk.CTkFont(size=11))
        ew = self._ed_canvas.create_window(cx, cy, anchor="nw", window=entry)

        def confirm(_=None):
            text = var.get().strip()
            self._ed_canvas.delete(ew)
            entry.destroy()
            if text:
                s = self._ed_scale
                self._ed_annots.setdefault(self._ed_page, []).append({
                    "type": "add_text",
                    "pdf_x":   cx / s,
                    "pdf_y":   cy / s + 12,
                    "text":    text,
                    "fontsize": 12,
                })
                self._ed_render()

        entry.bind("<Return>", confirm)
        entry.bind("<Escape>",
                   lambda _: (self._ed_canvas.delete(ew), entry.destroy()))
        entry.focus_set()

    def _ed_undo(self):
        lst = self._ed_annots.get(self._ed_page, [])
        if lst:
            lst.pop()
            self._ed_render()

    def _ed_save(self):
        if not self._ed_doc or not self._ed_src:
            return
        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Save Edited PDF",
            initialfile=(os.path.splitext(os.path.basename(self._ed_src))[0]
                         + "_edited.pdf"),
        )
        if not save_path:
            return

        annots_copy = {k: list(v) for k, v in self._ed_annots.items()}

        def task():
            try:
                doc = fitz.open(self._ed_src)
                for pno, annots in annots_copy.items():
                    page = doc[pno]
                    for a in annots:
                        if a["type"] == "highlight":
                            sh = page.new_shape()
                            sh.draw_rect(a["rect"])
                            c = a["color"]
                            sh.finish(fill=c, fill_opacity=0.40, color=None)
                            sh.commit()
                        elif a["type"] == "whiteout":
                            # Draw a white filled rectangle to cover content
                            sh = page.new_shape()
                            sh.draw_rect(a["rect"])
                            sh.finish(fill=(1, 1, 1), color=None)
                            sh.commit()
                        elif a["type"] == "replace_text":
                            r  = a["rect"]
                            fs = max(7, int((r.y1 - r.y0) * 0.82))
                            # Cover old text with white, then write new text
                            sh = page.new_shape()
                            sh.draw_rect(r)
                            sh.finish(fill=(1, 1, 1), color=None)
                            sh.commit()
                            page.insert_text(
                                fitz.Point(r.x0, r.y1 - 1),
                                a["new_text"], fontname="helv", fontsize=fs)
                        elif a["type"] == "add_text":
                            page.insert_text(
                                fitz.Point(a["pdf_x"], a["pdf_y"]),
                                a["text"], fontname="helv",
                                fontsize=a.get("fontsize", 12))

                doc.save(save_path, garbage=3, deflate=True)
                doc.close()
                self.after(0, lambda p=save_path: messagebox.showinfo(
                    "Saved", f"PDF saved to:\n{p}"))
            except Exception as e:
                self.after(0, lambda err=e: messagebox.showerror(
                    "Save Error", str(err)))

        threading.Thread(target=task, daemon=True).start()

    # ═══════════════════════════════════════ Viewer tab ══════

    def _create_viewer_frame(self) -> ctk.CTkFrame:
        outer = ctk.CTkFrame(self._content, corner_radius=10, fg_color=CARD_BG)
        outer.grid_columnconfigure(1, weight=1)   # canvas column expands
        outer.grid_rowconfigure(0, weight=1)

        # ── Left controls panel ───────────────────────────────
        panel = ctk.CTkFrame(outer, fg_color=TOOL_BG, corner_radius=0, width=150)
        panel.grid(row=0, column=0, sticky="ns")
        panel.grid_propagate(False)
        panel.grid_columnconfigure(0, weight=1)

        _row = 0
        ctk.CTkLabel(panel, text="PDF Viewer",
                      font=ctk.CTkFont(size=13, weight="bold"),
                      text_color="#c0c0d0").grid(
            row=_row, column=0, padx=10, pady=(16, 10)); _row += 1

        ctk.CTkFrame(panel, height=1, fg_color="#2a2a4a").grid(
            row=_row, column=0, sticky="ew", padx=8, pady=(0, 10)); _row += 1

        self._vw_path_var = ctk.StringVar()

        def _browse_vw():
            p = filedialog.askopenfilename(title="Open PDF",
                                           filetypes=[("PDF files", "*.pdf")])
            if p:
                self._vw_path_var.set(p)

        ctk.CTkButton(panel, text="Browse PDF", height=34, anchor="w",
                       command=_browse_vw).grid(
            row=_row, column=0, padx=10, pady=(0, 4), sticky="ew"); _row += 1

        ctk.CTkEntry(panel, textvariable=self._vw_path_var,
                      placeholder_text="No file selected",
                      height=28, font=ctk.CTkFont(size=10)).grid(
            row=_row, column=0, padx=10, pady=(0, 6), sticky="ew"); _row += 1

        ctk.CTkButton(panel, text="Open", height=34, fg_color=ACCENT,
                       command=self._vw_open).grid(
            row=_row, column=0, padx=10, pady=(0, 10), sticky="ew"); _row += 1

        ctk.CTkFrame(panel, height=1, fg_color="#2a2a4a").grid(
            row=_row, column=0, sticky="ew", padx=8, pady=(0, 10)); _row += 1

        ctk.CTkLabel(panel, text="NAVIGATION",
                      font=ctk.CTkFont(size=10, weight="bold"),
                      text_color="#555577").grid(
            row=_row, column=0, padx=10, pady=(0, 6)); _row += 1

        nav_f = ctk.CTkFrame(panel, fg_color="transparent")
        nav_f.grid(row=_row, column=0, padx=10, pady=(0, 4), sticky="ew"); _row += 1
        nav_f.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(nav_f, text="◀", width=32, height=30,
                       command=lambda: self._vw_nav(-1)).grid(row=0, column=0)
        self._vw_page_lbl = ctk.CTkLabel(nav_f, text="–",
                                          font=ctk.CTkFont(size=11))
        self._vw_page_lbl.grid(row=0, column=1)
        ctk.CTkButton(nav_f, text="▶", width=32, height=30,
                       command=lambda: self._vw_nav(1)).grid(row=0, column=2)

        ctk.CTkFrame(panel, height=1, fg_color="#2a2a4a").grid(
            row=_row, column=0, sticky="ew", padx=8, pady=10); _row += 1

        ctk.CTkLabel(panel, text="VIEW",
                      font=ctk.CTkFont(size=10, weight="bold"),
                      text_color="#555577").grid(
            row=_row, column=0, padx=10, pady=(0, 6)); _row += 1

        self._vw_book_btn = ctk.CTkButton(
            panel, text="📖  Book View", height=34, anchor="w",
            fg_color="#252545", command=self._vw_toggle_book)
        self._vw_book_btn.grid(row=_row, column=0, padx=10, pady=(0, 4), sticky="ew"); _row += 1

        self._vw_cream_btn = ctk.CTkButton(
            panel, text="🎨  Cream BG", height=34, anchor="w",
            fg_color="#252545", command=self._vw_toggle_cream)
        self._vw_cream_btn.grid(row=_row, column=0, padx=10, pady=(0, 10), sticky="ew"); _row += 1

        ctk.CTkFrame(panel, height=1, fg_color="#2a2a4a").grid(
            row=_row, column=0, sticky="ew", padx=8, pady=(0, 10)); _row += 1

        ctk.CTkLabel(panel, text="ZOOM",
                      font=ctk.CTkFont(size=10, weight="bold"),
                      text_color="#555577").grid(
            row=_row, column=0, padx=10, pady=(0, 6)); _row += 1

        zoom_f = ctk.CTkFrame(panel, fg_color="transparent")
        zoom_f.grid(row=_row, column=0, padx=10, pady=(0, 16), sticky="ew"); _row += 1
        zoom_f.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(zoom_f, text="−", width=30, height=30,
                       command=self._vw_zoom_out).grid(row=0, column=0)
        self._vw_zoom_lbl = ctk.CTkLabel(zoom_f, text="100%",
                                          font=ctk.CTkFont(size=11))
        self._vw_zoom_lbl.grid(row=0, column=1)
        ctk.CTkButton(zoom_f, text="+", width=30, height=30,
                       command=self._vw_zoom_in).grid(row=0, column=2)

        # ── Canvas (right side) ───────────────────────────────
        ch = ctk.CTkFrame(outer, fg_color="#1e1e38", corner_radius=0)
        ch.grid(row=0, column=1, sticky="nsew")
        ch.grid_rowconfigure(0, weight=1)
        ch.grid_columnconfigure(0, weight=1)

        self._vw_canvas = tk.Canvas(ch, bg="#1e1e38", highlightthickness=0)
        vw_vsb = tk.Scrollbar(ch, orient="vertical",   command=self._vw_canvas.yview)
        vw_hsb = tk.Scrollbar(ch, orient="horizontal", command=self._vw_canvas.xview)
        self._vw_canvas.configure(yscrollcommand=vw_vsb.set, xscrollcommand=vw_hsb.set)
        vw_vsb.grid(row=0, column=1, sticky="ns")
        vw_hsb.grid(row=1, column=0, sticky="ew")
        self._vw_canvas.grid(row=0, column=0, sticky="nsew")

        self._vw_canvas.bind("<MouseWheel>", lambda e: self._vw_canvas.yview_scroll(
            -1 if e.delta > 0 else 1, "units"))
        self._vw_canvas.bind("<Configure>", self._vw_on_resize)

        self._vw_hide_job = None

        return outer

    # ── Viewer logic ──────────────────────────────────────────

    def _vw_draw_placeholder(self):
        self._vw_canvas.delete("all")
        cw = self._vw_canvas.winfo_width()  or 800
        ch = self._vw_canvas.winfo_height() or 600
        self._vw_canvas.create_text(
            cw // 2, ch // 2,
            text="Browse and open a PDF to start reading",
            fill="#444466", font=("Helvetica", 15))

    def _vw_open(self):
        path = self._vw_path_var.get().strip()
        if not path or not os.path.isfile(path):
            return
        if self._vw_doc:
            self._vw_doc.close()
        self._vw_doc  = fitz.open(path)
        self._vw_page = 0
        self._vw_show_controls()
        self._vw_render()

    def _vw_nav(self, delta: int):
        if not self._vw_doc:
            return
        step = 2 if self._vw_book else 1
        self._vw_page = max(
            0, min(len(self._vw_doc) - 1, self._vw_page + delta * step))
        self._vw_render()

    def _vw_toggle_book(self):
        self._vw_book = not self._vw_book
        self._vw_book_btn.configure(
            fg_color=ACCENT if self._vw_book else "#2a2a4a")
        if self._vw_book and self._vw_page % 2 == 1:
            self._vw_page = max(0, self._vw_page - 1)
        self._vw_render()

    def _vw_zoom_in(self):
        self._vw_zoom = min(3.0, round(self._vw_zoom + 0.25, 2))
        self._vw_zoom_lbl.configure(text=f"{int(self._vw_zoom * 100)}%")
        self._vw_render()

    def _vw_zoom_out(self):
        self._vw_zoom = max(0.25, round(self._vw_zoom - 0.25, 2))
        self._vw_zoom_lbl.configure(text=f"{int(self._vw_zoom * 100)}%")
        self._vw_render()

    def _vw_on_resize(self, _event=None):
        if not self._vw_doc:
            return
        if self._vw_resize_job:
            self.after_cancel(self._vw_resize_job)
        self._vw_resize_job = self.after(120, self._vw_render)

    def _vw_toggle_cream(self):
        self._vw_cream = not self._vw_cream
        self._vw_cream_btn.configure(
            fg_color=ACCENT if self._vw_cream else "#2a2a4a")
        self._vw_render()

    def _vw_render(self):
        if not self._vw_doc:
            return
        n = len(self._vw_doc)
        self._vw_page_lbl.configure(text=f"Page {self._vw_page + 1} / {n}")

        # Render at 3× (216 DPI) — supersample then LANCZOS-downsample to display size
        RENDER_SCALE = 3.0
        PAD    = 24
        SHADOW = 8

        def render_page(idx: int) -> Image.Image:
            p   = self._vw_doc[idx]
            pix = p.get_pixmap(matrix=fitz.Matrix(RENDER_SCALE, RENDER_SCALE),
                                alpha=False, colorspace=fitz.csRGB)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            if self._vw_cream:
                cream = Image.new("RGB", img.size, (255, 248, 220))
                img   = Image.blend(img, cream, 0.28)
            return img

        # Canvas width drives the display size; zoom scales relative to fit-width
        cw = max(self._vw_canvas.winfo_width(), 860)

        def fit_to_width(img: Image.Image, target_w: int) -> Image.Image:
            if img.width == target_w:
                return img
            ratio = target_w / img.width
            return img.resize(
                (target_w, int(img.height * ratio)), Image.LANCZOS)

        if self._vw_book and self._vw_page + 1 < n:
            il = render_page(self._vw_page)
            ir = render_page(self._vw_page + 1)
            gap = 16
            # Base fit: both pages side-by-side fill canvas, then apply zoom
            base_w = (cw - PAD * 2 - SHADOW - gap) // 2
            target_w = max(80, int(base_w * self._vw_zoom))
            il = fit_to_width(il, target_w)
            ir = fit_to_width(ir, target_w)
            tw = il.width + gap + ir.width + PAD * 2
            th = max(il.height, ir.height) + PAD * 2 + SHADOW
            bg = Image.new("RGB", (tw, th), "#252545")
            sh_l = Image.new("RGB", (il.width + SHADOW, il.height + SHADOW), "#0a0a20")
            sh_r = Image.new("RGB", (ir.width + SHADOW, ir.height + SHADOW), "#0a0a20")
            bg.paste(sh_l, (PAD + SHADOW,                    PAD + SHADOW))
            bg.paste(sh_r, (PAD + il.width + gap + SHADOW,   PAD + SHADOW))
            bg.paste(il,   (PAD,                             PAD))
            bg.paste(ir,   (PAD + il.width + gap,            PAD))
            img = bg
        else:
            raw = render_page(self._vw_page)
            base_w  = cw - PAD * 2 - SHADOW
            target_w = max(80, int(base_w * self._vw_zoom))
            raw = fit_to_width(raw, target_w)
            tw = raw.width  + PAD * 2 + SHADOW
            th = raw.height + PAD * 2 + SHADOW
            bg = Image.new("RGB", (tw, th), "#252545")
            sh = Image.new("RGB", (raw.width + SHADOW, raw.height + SHADOW), "#0a0a20")
            bg.paste(sh,  (PAD + SHADOW, PAD + SHADOW))
            bg.paste(raw, (PAD, PAD))
            img = bg

        self._vw_photo = ImageTk.PhotoImage(img)
        self._vw_canvas.configure(scrollregion=(0, 0, img.width, img.height))
        self._vw_canvas.delete("all")
        self._vw_canvas.create_image(0, 0, anchor="nw", image=self._vw_photo)
