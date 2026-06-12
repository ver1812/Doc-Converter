import os
import threading
import customtkinter as ctk
from tkinter import filedialog

from converters.pdf_converter import convert_pdf_to_docx
from converters.image_converter import convert_image

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

SIDEBAR_WIDTH = 200
ACCENT = "#1f6feb"
SIDEBAR_BG = "#1a1a2e"
CARD_BG = "#16213e"


class DocConverterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Doc Converter")
        self.geometry("960x620")
        self.minsize(820, 540)
        self._build_layout()
        self._show_frame("documents")

    # ------------------------------------------------------------------ layout

    def _build_layout(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._sidebar = ctk.CTkFrame(self, width=SIDEBAR_WIDTH, corner_radius=0, fg_color=SIDEBAR_BG)
        self._sidebar.grid(row=0, column=0, sticky="nsew")
        self._sidebar.grid_propagate(False)
        self._create_sidebar()

        self._content = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self._content.grid(row=0, column=1, sticky="nsew", padx=0, pady=0)
        self._content.grid_columnconfigure(0, weight=1)
        self._content.grid_rowconfigure(0, weight=1)

        self._frames = {}
        self._frames["documents"] = self._create_document_frame()
        self._frames["images"] = self._create_image_frame()

    def _create_sidebar(self):
        self._sidebar.grid_rowconfigure(10, weight=1)

        logo = ctk.CTkLabel(
            self._sidebar,
            text="Doc Converter",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color="#e0e0e0",
        )
        logo.grid(row=0, column=0, padx=20, pady=(30, 20), sticky="w")

        sep = ctk.CTkFrame(self._sidebar, height=1, fg_color="#333355")
        sep.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 16))

        self._nav_buttons = {}
        nav_items = [("documents", "  Documents"), ("images", "  Images")]
        for i, (key, label) in enumerate(nav_items, start=2):
            btn = ctk.CTkButton(
                self._sidebar,
                text=label,
                anchor="w",
                height=42,
                corner_radius=8,
                fg_color="transparent",
                hover_color="#2a2a4a",
                text_color="#c0c0d0",
                font=ctk.CTkFont(size=14),
                command=lambda k=key: self._show_frame(k),
            )
            btn.grid(row=i, column=0, padx=12, pady=4, sticky="ew")
            self._nav_buttons[key] = btn

        version = ctk.CTkLabel(self._sidebar, text="v1.0.0", text_color="#555577", font=ctk.CTkFont(size=11))
        version.grid(row=11, column=0, padx=20, pady=16, sticky="sw")

    def _show_frame(self, name: str):
        for key, frame in self._frames.items():
            frame.grid_remove()
        self._frames[name].grid(row=0, column=0, sticky="nsew", padx=32, pady=32)

        for key, btn in self._nav_buttons.items():
            if key == name:
                btn.configure(fg_color=ACCENT, text_color="white")
            else:
                btn.configure(fg_color="transparent", text_color="#c0c0d0")

    # --------------------------------------------------------- document frame

    def _create_document_frame(self) -> ctk.CTkFrame:
        frame = ctk.CTkFrame(self._content, corner_radius=12, fg_color=CARD_BG)
        frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(frame, text="PDF  →  Word", font=ctk.CTkFont(size=22, weight="bold")).grid(
            row=0, column=0, pady=(28, 4), sticky="w", padx=28
        )
        ctk.CTkLabel(frame, text="Convert PDF documents into editable Word (.docx) files.",
                     text_color="#8888aa", font=ctk.CTkFont(size=13)).grid(
            row=1, column=0, pady=(0, 20), sticky="w", padx=28
        )

        sep = ctk.CTkFrame(frame, height=1, fg_color="#2a2a4a")
        sep.grid(row=2, column=0, sticky="ew", padx=28, pady=(0, 20))

        # PDF file row
        self._pdf_path = ctk.StringVar(value="")
        self._create_file_row(frame, row=3, label="PDF File", var=self._pdf_path,
                               btn_text="Browse PDF",
                               filetypes=[("PDF files", "*.pdf")],
                               select_dir=False)

        # Output folder row
        self._pdf_out_dir = ctk.StringVar(value="")
        self._create_file_row(frame, row=5, label="Output Folder", var=self._pdf_out_dir,
                               btn_text="Browse Folder", select_dir=True)

        # Convert button
        self._pdf_convert_btn = ctk.CTkButton(
            frame, text="Convert to Word", height=44, corner_radius=8,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._start_pdf_conversion,
        )
        self._pdf_convert_btn.grid(row=7, column=0, padx=28, pady=(24, 12), sticky="ew")

        # Progress + status
        self._pdf_progress = ctk.CTkProgressBar(frame, height=8, corner_radius=4)
        self._pdf_progress.set(0)
        self._pdf_progress.grid(row=8, column=0, padx=28, pady=(0, 8), sticky="ew")

        self._pdf_status = ctk.CTkLabel(frame, text="", text_color="#8888aa", font=ctk.CTkFont(size=12))
        self._pdf_status.grid(row=9, column=0, padx=28, pady=(0, 24), sticky="w")

        return frame

    # ---------------------------------------------------------- image frame

    def _create_image_frame(self) -> ctk.CTkFrame:
        frame = ctk.CTkFrame(self._content, corner_radius=12, fg_color=CARD_BG)
        frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(frame, text="Image Converter", font=ctk.CTkFont(size=22, weight="bold")).grid(
            row=0, column=0, pady=(28, 4), sticky="w", padx=28
        )
        ctk.CTkLabel(frame, text="Convert between JPEG, PNG, and HEIC formats.",
                     text_color="#8888aa", font=ctk.CTkFont(size=13)).grid(
            row=1, column=0, pady=(0, 20), sticky="w", padx=28
        )

        sep = ctk.CTkFrame(frame, height=1, fg_color="#2a2a4a")
        sep.grid(row=2, column=0, sticky="ew", padx=28, pady=(0, 20))

        # Image file row
        self._img_path = ctk.StringVar(value="")
        self._create_file_row(frame, row=3, label="Image File", var=self._img_path,
                               btn_text="Browse Image",
                               filetypes=[("Image files", "*.jpg *.jpeg *.png *.heic *.heif")],
                               select_dir=False)

        # Output folder row
        self._img_out_dir = ctk.StringVar(value="")
        self._create_file_row(frame, row=5, label="Output Folder", var=self._img_out_dir,
                               btn_text="Browse Folder", select_dir=True)

        # Format selector row
        fmt_row = ctk.CTkFrame(frame, fg_color="transparent")
        fmt_row.grid(row=6, column=0, padx=28, pady=(12, 0), sticky="ew")
        fmt_row.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(fmt_row, text="Output Format", font=ctk.CTkFont(size=13), width=120, anchor="w").grid(
            row=0, column=0, sticky="w"
        )
        self._img_format = ctk.StringVar(value="PNG")
        fmt_menu = ctk.CTkOptionMenu(fmt_row, values=["PNG", "JPEG", "HEIC"],
                                      variable=self._img_format, width=160)
        fmt_menu.grid(row=0, column=1, sticky="w", padx=(12, 0))

        # Convert button
        self._img_convert_btn = ctk.CTkButton(
            frame, text="Convert Image", height=44, corner_radius=8,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._start_image_conversion,
        )
        self._img_convert_btn.grid(row=8, column=0, padx=28, pady=(24, 12), sticky="ew")

        # Progress + status
        self._img_progress = ctk.CTkProgressBar(frame, height=8, corner_radius=4)
        self._img_progress.set(0)
        self._img_progress.grid(row=9, column=0, padx=28, pady=(0, 8), sticky="ew")

        self._img_status = ctk.CTkLabel(frame, text="", text_color="#8888aa", font=ctk.CTkFont(size=12))
        self._img_status.grid(row=10, column=0, padx=28, pady=(0, 24), sticky="w")

        return frame

    # ---------------------------------------------------------- shared widget

    def _create_file_row(self, parent, row, label, var, btn_text, select_dir=False, filetypes=None):
        lbl_row = ctk.CTkFrame(parent, fg_color="transparent")
        lbl_row.grid(row=row, column=0, padx=28, pady=(8, 4), sticky="ew")
        lbl_row.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(lbl_row, text=label, font=ctk.CTkFont(size=13), width=120, anchor="w").grid(
            row=0, column=0, sticky="w"
        )

        entry = ctk.CTkEntry(lbl_row, textvariable=var, placeholder_text="Not selected", height=36)
        entry.grid(row=0, column=1, sticky="ew", padx=(12, 8))

        def browse(sd=select_dir, ft=filetypes, v=var):
            if sd:
                path = filedialog.askdirectory(title="Select Output Folder")
            else:
                path = filedialog.askopenfilename(title="Select File", filetypes=ft or [("All files", "*.*")])
            if path:
                v.set(path)
                # Auto-set output dir to same folder as input file
                if not sd:
                    out_dir_var = self._pdf_out_dir if v is self._pdf_path else self._img_out_dir
                    if not out_dir_var.get():
                        out_dir_var.set(os.path.dirname(path))

        ctk.CTkButton(lbl_row, text=btn_text, width=130, height=36, command=browse).grid(row=0, column=2)

    # --------------------------------------------------------- conversions

    def _start_pdf_conversion(self):
        src = self._pdf_path.get().strip()
        out_dir = self._pdf_out_dir.get().strip()
        if not src:
            self._pdf_status.configure(text="Please select a PDF file.", text_color="#ff6b6b")
            return
        if not os.path.isfile(src):
            self._pdf_status.configure(text="File not found.", text_color="#ff6b6b")
            return
        if not out_dir:
            out_dir = os.path.dirname(src)
        basename = os.path.splitext(os.path.basename(src))[0]
        dest = os.path.join(out_dir, basename + ".docx")

        self._pdf_convert_btn.configure(state="disabled")
        self._pdf_progress.set(0)
        self._pdf_status.configure(text="Converting…", text_color="#8888aa")
        self._pdf_progress.start()

        def task():
            try:
                convert_pdf_to_docx(src, dest)
                self.after(0, lambda: self._pdf_done(dest, None))
            except Exception as e:
                self.after(0, lambda err=e: self._pdf_done(None, err))

        threading.Thread(target=task, daemon=True).start()

    def _pdf_done(self, dest, error):
        self._pdf_progress.stop()
        self._pdf_progress.set(1 if not error else 0)
        self._pdf_convert_btn.configure(state="normal")
        if error:
            self._pdf_status.configure(text=f"Error: {error}", text_color="#ff6b6b")
        else:
            self._pdf_status.configure(text=f"Saved: {dest}", text_color="#4ecca3")

    def _start_image_conversion(self):
        src = self._img_path.get().strip()
        out_dir = self._img_out_dir.get().strip()
        fmt = self._img_format.get()
        if not src:
            self._img_status.configure(text="Please select an image file.", text_color="#ff6b6b")
            return
        if not os.path.isfile(src):
            self._img_status.configure(text="File not found.", text_color="#ff6b6b")
            return
        if not out_dir:
            out_dir = os.path.dirname(src)

        ext_map = {"JPEG": ".jpg", "PNG": ".png", "HEIC": ".heic"}
        basename = os.path.splitext(os.path.basename(src))[0]
        dest = os.path.join(out_dir, basename + ext_map[fmt])

        self._img_convert_btn.configure(state="disabled")
        self._img_progress.set(0)
        self._img_status.configure(text="Converting…", text_color="#8888aa")
        self._img_progress.start()

        def task():
            try:
                convert_image(src, dest, fmt)
                self.after(0, lambda: self._img_done(dest, None))
            except Exception as e:
                self.after(0, lambda err=e: self._img_done(None, err))

        threading.Thread(target=task, daemon=True).start()

    def _img_done(self, dest, error):
        self._img_progress.stop()
        self._img_progress.set(1 if not error else 0)
        self._img_convert_btn.configure(state="normal")
        if error:
            self._img_status.configure(text=f"Error: {error}", text_color="#ff6b6b")
        else:
            self._img_status.configure(text=f"Saved: {dest}", text_color="#4ecca3")
