import io
import os
import tkinter as tk
from tkinter import colorchooser, messagebox, ttk
from typing import Callable, Dict, List, Optional, Tuple

import customtkinter as ctk
from PIL import Image, ImageTk

try:
    import docx
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Inches, Pt, RGBColor
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


# ==============================================================================
# DOC / DOCX PREVIEW & IN-APP EDITOR
# ==============================================================================

class DocxEditorComponent(ctk.CTkFrame):
    """
    Vùng xem trước và chỉnh sửa tài liệu DOC / DOCX:
    Hỗ trợ điều chỉnh Font chữ, Cỡ chữ, Căn lề văn bản (Align),
    Định dạng kích thước và vị trí ảnh (Format ảnh),
    Lưu trực tiếp vào file .docx.
    """
    def __init__(self, master, filepath: str, on_change_callback: Optional[Callable] = None):
        super().__init__(master, fg_color="transparent")
        self.filepath = filepath
        self.on_change_callback = on_change_callback

        self.doc: Optional[docx.Document] = None
        self.extracted_images: List[Dict] = []
        self._image_tk_refs: List[ImageTk.PhotoImage] = []

        self._build_toolbar()
        self._build_editor_area()
        self._load_document()

    def _build_toolbar(self):
        toolbar = ctk.CTkFrame(self, height=44, corner_radius=8, fg_color=("#e2e8f0", "#1e293b"))
        toolbar.pack(fill="x", padx=4, pady=(2, 6))

        # 1. Font Family
        ctk.CTkLabel(toolbar, text="Font:", font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")).pack(side="left", padx=(10, 4))
        self.font_family_var = tk.StringVar(value="Arial")
        self.cb_font_family = ctk.CTkComboBox(
            toolbar,
            values=["Arial", "Times New Roman", "Calibri", "Segoe UI", "Tahoma", "Georgia", "Courier New", "Verdana"],
            variable=self.font_family_var,
            width=140,
            height=28,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            command=self._apply_font_family,
        )
        self.cb_font_family.pack(side="left", padx=(0, 8))

        # 2. Font Size
        ctk.CTkLabel(toolbar, text="Cỡ:", font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")).pack(side="left", padx=(0, 4))
        self.font_size_var = tk.StringVar(value="12")
        self.cb_font_size = ctk.CTkComboBox(
            toolbar,
            values=["8", "9", "10", "11", "12", "14", "16", "18", "20", "24", "28", "36", "48"],
            variable=self.font_size_var,
            width=65,
            height=28,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            command=self._apply_font_size,
        )
        self.cb_font_size.pack(side="left", padx=(0, 8))

        # 3. Định dạng Đậm, Nghiêng, Gạch chân
        self.btn_bold = ctk.CTkButton(
            toolbar,
            text="B",
            width=30,
            height=28,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            command=self._toggle_bold,
            fg_color=("#94a3b8", "#334155"),
            hover_color=("#64748b", "#475569"),
        )
        self.btn_bold.pack(side="left", padx=2)

        self.btn_italic = ctk.CTkButton(
            toolbar,
            text="I",
            width=30,
            height=28,
            font=ctk.CTkFont(family="Segoe UI", size=11, slant="italic"),
            command=self._toggle_italic,
            fg_color=("#94a3b8", "#334155"),
            hover_color=("#64748b", "#475569"),
        )
        self.btn_italic.pack(side="left", padx=2)

        self.btn_underline = ctk.CTkButton(
            toolbar,
            text="U",
            width=30,
            height=28,
            font=ctk.CTkFont(family="Segoe UI", size=11, underline=True),
            command=self._toggle_underline,
            fg_color=("#94a3b8", "#334155"),
            hover_color=("#64748b", "#475569"),
        )
        self.btn_underline.pack(side="left", padx=2)

        # Đường phân cách
        sep1 = ctk.CTkFrame(toolbar, width=1, height=22, fg_color=("#cbd5e1", "#475569"))
        sep1.pack(side="left", padx=8)

        # 4. Căn lề Text (Align)
        self.btn_align_left = ctk.CTkButton(
            toolbar,
            text="⬅ Trái",
            width=58,
            height=28,
            command=lambda: self._set_alignment("left"),
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=("#94a3b8", "#334155"),
            hover_color=("#64748b", "#475569"),
        )
        self.btn_align_left.pack(side="left", padx=2)

        self.btn_align_center = ctk.CTkButton(
            toolbar,
            text="↔ Giữa",
            width=58,
            height=28,
            command=lambda: self._set_alignment("center"),
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=("#94a3b8", "#334155"),
            hover_color=("#64748b", "#475569"),
        )
        self.btn_align_center.pack(side="left", padx=2)

        self.btn_align_right = ctk.CTkButton(
            toolbar,
            text="➡ Phải",
            width=58,
            height=28,
            command=lambda: self._set_alignment("right"),
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=("#94a3b8", "#334155"),
            hover_color=("#64748b", "#475569"),
        )
        self.btn_align_right.pack(side="left", padx=2)

        self.btn_align_justify = ctk.CTkButton(
            toolbar,
            text="≡ Đều",
            width=55,
            height=28,
            command=lambda: self._set_alignment("justify"),
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=("#94a3b8", "#334155"),
            hover_color=("#64748b", "#475569"),
        )
        self.btn_align_justify.pack(side="left", padx=2)

        # Đường phân cách 2
        sep2 = ctk.CTkFrame(toolbar, width=1, height=22, fg_color=("#cbd5e1", "#475569"))
        sep2.pack(side="left", padx=8)

        # 5. Format Ảnh (Kích thước & Căn chỉnh)
        self.btn_format_image = ctk.CTkButton(
            toolbar,
            text="🖼️ Format Ảnh",
            width=110,
            height=28,
            command=self._open_image_format_modal,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color=("#8b5cf6", "#7c3aed"),
            hover_color=("#7c3aed", "#6d28d9"),
        )
        self.btn_format_image.pack(side="left", padx=4)

        # 6. Màu chữ
        self.btn_color = ctk.CTkButton(
            toolbar,
            text="🎨 Màu",
            width=65,
            height=28,
            command=self._choose_text_color,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=("#0284c7", "#0369a1"),
            hover_color=("#0369a1", "#0284c7"),
        )
        self.btn_color.pack(side="left", padx=2)

    def _build_editor_area(self):
        # Khung nền tài liệu (giả lập trang giấy Word với thanh cuộn)
        doc_container = ctk.CTkFrame(self, fg_color=("#f1f5f9", "#090d16"), corner_radius=6)
        doc_container.pack(fill="both", expand=True, padx=4, pady=(0, 4))
        doc_container.grid_rowconfigure(0, weight=1)
        doc_container.grid_columnconfigure(0, weight=1)

        self.text_widget = tk.Text(
            doc_container,
            wrap="word",
            undo=True,
            font=("Arial", 12),
            bg="#ffffff",
            fg="#1e293b",
            insertbackground="#2563eb",
            selectbackground="#93c5fd",
            selectforeground="#0f172a",
            padx=28,
            pady=24,
            relief="flat",
            bd=0,
        )
        vsb = ttk.Scrollbar(doc_container, orient="vertical", command=self.text_widget.yview)
        self.text_widget.configure(yscrollcommand=vsb.set)

        self.text_widget.grid(row=0, column=0, sticky="nsew", padx=(10, 0), pady=10)
        vsb.grid(row=0, column=1, sticky="ns", pady=10, padx=(2, 6))

        # Cấu hình các thẻ định dạng
        self.text_widget.tag_configure("align_left", justify="left")
        self.text_widget.tag_configure("align_center", justify="center")
        self.text_widget.tag_configure("align_right", justify="right")
        self.text_widget.tag_configure("align_justify", justify="left")

    def _load_document(self):
        if not DOCX_AVAILABLE:
            self.text_widget.insert("1.0", f"Thư viện python-docx không khả dụng để nạp: {self.filepath}")
            return

        if not os.path.exists(self.filepath):
            self.text_widget.insert("1.0", f"Không tìm thấy file: {self.filepath}")
            return

        try:
            self.doc = docx.Document(self.filepath)
            self.text_widget.delete("1.0", "end")
            self._image_tk_refs.clear()
            self.extracted_images.clear()

            # Trích xuất hình ảnh nếu có trong file docx
            img_counter = 0
            for part in self.doc.part.related_parts.values():
                if "image" in part.content_type:
                    try:
                        img_bytes = part.blob
                        pil_img = Image.open(io.BytesIO(img_bytes))
                        self.extracted_images.append({
                            "id": img_counter,
                            "pil": pil_img,
                            "width": pil_img.width,
                            "height": pil_img.height,
                            "align": "center",
                            "scale": 100,
                        })
                        img_counter += 1
                    except Exception:
                        pass

            # Nạp các đoạn văn (paragraphs)
            for idx, p in enumerate(self.doc.paragraphs):
                text = p.text
                if not text and idx < len(self.extracted_images):
                    # Thêm ảnh nếu đoạn này đại diện cho ảnh
                    pass

                align_tag = "align_left"
                if p.alignment == WD_ALIGN_PARAGRAPH.CENTER:
                    align_tag = "align_center"
                elif p.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
                    align_tag = "align_right"
                elif p.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
                    align_tag = "align_justify"

                start_idx = self.text_widget.index("end-1c")
                self.text_widget.insert("end", text + "\n", (align_tag,))

            # Hiển thị ảnh kèm theo nếu có
            if self.extracted_images:
                self.text_widget.insert("end", "\n--- [HÌNH ẢNH TRONG TÀI LIỆU] ---\n\n", ("align_center",))
                for img_data in self.extracted_images:
                    pil_copy = img_data["pil"].copy()
                    # Thu nhỏ vừa vặn khung xem trước tối đa 350px
                    max_w = 350
                    if pil_copy.width > max_w:
                        ratio = max_w / pil_copy.width
                        pil_copy = pil_copy.resize((max_w, int(pil_copy.height * ratio)), Image.Resampling.LANCZOS)

                    tk_img = ImageTk.PhotoImage(pil_copy)
                    self._image_tk_refs.append(tk_img)

                    lbl_img = tk.Label(
                        self.text_widget,
                        image=tk_img,
                        bg="#ffffff",
                        bd=1,
                        relief="solid",
                        cursor="hand2",
                    )
                    lbl_img.bind("<Button-1>", lambda e, d=img_data: self._open_single_image_format(d))
                    self.text_widget.window_create("end", window=lbl_img)
                    self.text_widget.insert("end", f"\nẢnh #{img_data['id']+1} (Click để đổi kích thước & căn lề)\n\n", ("align_center",))

        except Exception as e:
            # Fallback nếu file .doc hoặc lỗi nạp
            self.text_widget.delete("1.0", "end")
            self.text_widget.insert("1.0", f"Lỗi đọc tài liệu bằng docx: {e}\n\nĐang hiển thị chế độ văn bản thô nếu có thể...")

    def _apply_font_family(self, font_name: str):
        try:
            cur_size = int(self.font_size_var.get())
        except Exception:
            cur_size = 12
        try:
            if self.text_widget.tag_ranges("sel"):
                tag_name = f"font_{font_name}_{cur_size}"
                self.text_widget.tag_configure(tag_name, font=(font_name, cur_size))
                self.text_widget.tag_add(tag_name, "sel.first", "sel.last")
            else:
                self.text_widget.configure(font=(font_name, cur_size))
        except Exception:
            pass

    def _apply_font_size(self, size_str: str):
        try:
            sz = int(size_str)
            cur_font = self.font_family_var.get()
            if self.text_widget.tag_ranges("sel"):
                tag_name = f"font_{cur_font}_{sz}"
                self.text_widget.tag_configure(tag_name, font=(cur_font, sz))
                self.text_widget.tag_add(tag_name, "sel.first", "sel.last")
            else:
                self.text_widget.configure(font=(cur_font, sz))
        except Exception:
            pass

    def _toggle_bold(self):
        cur_font = self.font_family_var.get()
        try:
            sz = int(self.font_size_var.get())
        except Exception:
            sz = 12
        tag_name = f"bold_{cur_font}_{sz}"
        self.text_widget.tag_configure(tag_name, font=(cur_font, sz, "bold"))
        if self.text_widget.tag_ranges("sel"):
            # Kiểm tra xem đã có tag bold chưa để toggle
            if tag_name in self.text_widget.tag_names("sel.first"):
                self.text_widget.tag_remove(tag_name, "sel.first", "sel.last")
            else:
                self.text_widget.tag_add(tag_name, "sel.first", "sel.last")

    def _toggle_italic(self):
        cur_font = self.font_family_var.get()
        try:
            sz = int(self.font_size_var.get())
        except Exception:
            sz = 12
        tag_name = f"italic_{cur_font}_{sz}"
        self.text_widget.tag_configure(tag_name, font=(cur_font, sz, "italic"))
        if self.text_widget.tag_ranges("sel"):
            if tag_name in self.text_widget.tag_names("sel.first"):
                self.text_widget.tag_remove(tag_name, "sel.first", "sel.last")
            else:
                self.text_widget.tag_add(tag_name, "sel.first", "sel.last")

    def _toggle_underline(self):
        tag_name = "underline_text"
        self.text_widget.tag_configure(tag_name, underline=True)
        if self.text_widget.tag_ranges("sel"):
            if tag_name in self.text_widget.tag_names("sel.first"):
                self.text_widget.tag_remove(tag_name, "sel.first", "sel.last")
            else:
                self.text_widget.tag_add(tag_name, "sel.first", "sel.last")

    def _set_alignment(self, align: str):
        align_tag = f"align_{align}"
        # Xóa các tag căn lề cũ
        for a in ["align_left", "align_center", "align_right", "align_justify"]:
            if self.text_widget.tag_ranges("sel"):
                self.text_widget.tag_remove(a, "sel.first", "sel.last")
            else:
                self.text_widget.tag_remove(a, "insert linestart", "insert lineend")

        if self.text_widget.tag_ranges("sel"):
            self.text_widget.tag_add(align_tag, "sel.first", "sel.last")
        else:
            self.text_widget.tag_add(align_tag, "insert linestart", "insert lineend")

    def _choose_text_color(self):
        color = colorchooser.askcolor(title="Chọn màu chữ")[1]
        if color:
            tag_name = f"color_{color}"
            self.text_widget.tag_configure(tag_name, foreground=color)
            if self.text_widget.tag_ranges("sel"):
                self.text_widget.tag_add(tag_name, "sel.first", "sel.last")
            else:
                self.text_widget.configure(fg=color)

    def _open_image_format_modal(self):
        """Mở bảng điều chỉnh định dạng ảnh."""
        if not self.extracted_images:
            messagebox.showinfo("Thông báo", "Tài liệu này không chứa hình ảnh để định dạng.")
            return
        self._open_single_image_format(self.extracted_images[0])

    def _open_single_image_format(self, img_data: dict):
        win = ctk.CTkToplevel(self)
        win.title(f"Format Ảnh #{img_data['id']+1}")
        win.geometry("380x360")
        win.resizable(False, False)
        win.attributes("-topmost", True)

        pad = ctk.CTkFrame(win, fg_color="transparent")
        pad.pack(fill="both", expand=True, padx=20, pady=16)

        ctk.CTkLabel(pad, text=f"🖼️ Định dạng Ảnh #{img_data['id']+1}", font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold")).pack(anchor="w", pady=(0, 10))

        # Xem kích thước gốc
        ctk.CTkLabel(pad, text=f"Kích thước gốc: {img_data['pil'].width} x {img_data['pil'].height} px", text_color="#94a3b8").pack(anchor="w", pady=(0, 10))

        # Tỉ lệ co giãn
        ctk.CTkLabel(pad, text="Tỉ lệ hiển thị / In ấn (%):", font=ctk.CTkFont(family="Segoe UI", size=12)).pack(anchor="w", pady=(4, 2))
        scale_slider = ctk.CTkSlider(pad, from_=20, to=200, number_of_steps=18)
        scale_slider.set(img_data.get("scale", 100))
        scale_slider.pack(fill="x", pady=(0, 4))

        lbl_scale_val = ctk.CTkLabel(pad, text=f"{int(scale_slider.get())}%", font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"))
        lbl_scale_val.pack(anchor="e")
        scale_slider.configure(command=lambda v: lbl_scale_val.configure(text=f"{int(v)}%"))

        # Căn lề ảnh
        ctk.CTkLabel(pad, text="Căn lề vị trí ảnh:", font=ctk.CTkFont(family="Segoe UI", size=12)).pack(anchor="w", pady=(8, 2))
        align_opt = ctk.CTkSegmentedButton(pad, values=["Trái (Left)", "Giữa (Center)", "Phải (Right)"])
        cur_a = img_data.get("align", "center")
        if cur_a == "left":
            align_opt.set("Trái (Left)")
        elif cur_a == "right":
            align_opt.set("Phải (Right)")
        else:
            align_opt.set("Giữa (Center)")
        align_opt.pack(fill="x", pady=(0, 14))

        def _save_img_props():
            new_scale = int(scale_slider.get())
            choice = align_opt.get()
            if "Trái" in choice:
                img_data["align"] = "left"
            elif "Phải" in choice:
                img_data["align"] = "right"
            else:
                img_data["align"] = "center"
            img_data["scale"] = new_scale
            messagebox.showinfo("Thành công", f"Đã cập nhật định dạng ảnh (Tỉ lệ: {new_scale}%, Vị trí: {img_data['align']})!")
            win.destroy()

        ctk.CTkButton(pad, text="Áp dụng thay đổi", command=_save_img_props, fg_color="#2563eb", hover_color="#1d4ed8").pack(fill="x", pady=(10, 0))

    def save_changes(self) -> Tuple[bool, str]:
        """Lưu lại nội dung văn bản, font chữ, cỡ chữ, căn lề và ảnh vào file .docx."""
        if not DOCX_AVAILABLE or self.doc is None:
            return False, "Không thể lưu vì thư viện python-docx không khả dụng hoặc file chưa nạp."

        try:
            # Lấy toàn bộ text từ editor
            content = self.text_widget.get("1.0", "end-1c")
            lines = content.split("\n")

            # Xóa các paragraph cũ và ghi lại nội dung mới đã định dạng
            # (Giữ các element cốt lõi của document)
            p_elements = list(self.doc.paragraphs)
            for p in p_elements:
                p._element.getparent().remove(p._element)

            font_name = self.font_family_var.get()
            try:
                font_size_pt = Pt(float(self.font_size_var.get()))
            except Exception:
                font_size_pt = Pt(12)

            for line in lines:
                if "--- [HÌNH ẢNH TRONG TÀI LIỆU] ---" in line or "Ảnh #" in line:
                    continue
                p = self.doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                run = p.add_run(line)
                run.font.name = font_name
                run.font.size = font_size_pt

            # Cập nhật kích thước ảnh nếu có
            if self.extracted_images and hasattr(self.doc, "inline_shapes"):
                for idx, shape in enumerate(self.doc.inline_shapes):
                    if idx < len(self.extracted_images):
                        scale = self.extracted_images[idx].get("scale", 100) / 100.0
                        if shape.width and shape.height:
                            shape.width = int(shape.width * scale)
                            shape.height = int(shape.height * scale)

            self.doc.save(self.filepath)
            return True, f"Đã lưu thành công các thay đổi vào file: {os.path.basename(self.filepath)}"

        except PermissionError:
            return False, f"Không thể lưu vì file đang được mở trong phần mềm khác (như Microsoft Word). Vui lòng đóng file trong Word rồi nhấn Lưu lại!"
        except Exception as e:
            return False, f"Lỗi khi lưu file docx: {e}"


# ==============================================================================
# XLS / XLSX PREVIEW & IN-APP EDITOR
# ==============================================================================

class XlsxEditorComponent(ctk.CTkFrame):
    """
    Vùng xem trước và chỉnh sửa bảng tính XLS / XLSX:
    - Hỗ trợ chỉnh sửa Font chữ, Cỡ chữ, Căn lề (Align Text).
    - Hỗ trợ Normal View Mode (Lưới bảng tính).
    - Hỗ trợ Page Break Preview View Mode (Xem ngắt trang):
      Cho phép người dùng kéo/điều chỉnh độ rộng từng trang in (Page Width).
    - Lưu trực tiếp vào file .xlsx.
    """
    def __init__(self, master, filepath: str, on_change_callback: Optional[Callable] = None):
        super().__init__(master, fg_color="transparent")
        self.filepath = filepath
        self.on_change_callback = on_change_callback

        self.wb: Optional[openpyxl.Workbook] = None
        self.active_sheet_name: str = ""
        self.sheet_data: Dict[str, List[List]] = {}
        self.col_widths: Dict[str, float] = {}
        self.view_mode: str = "normal"  # "normal" hoặc "page_break"

        # Cài đặt ngắt trang (Cột phân chia trang 1 và trang 2)
        self.page_break_col_index: int = 5  # Mặc định cột E là ngắt trang 1
        self.page_scale_percent: int = 100

        # Theo dõi ô đang chọn hiện tại
        self._current_col_letter: str = "A"
        self._current_row_id: str = "1"
        self._is_bold: bool = False
        self._is_italic: bool = False

        self._build_toolbar()
        self._build_sheet_tabs()
        self._build_editor_area()
        self._load_workbook()

    def _build_toolbar(self):
        toolbar = ctk.CTkFrame(self, height=44, corner_radius=8, fg_color=("#e2e8f0", "#1e293b"))
        toolbar.pack(fill="x", padx=4, pady=(2, 4))

        # 1. Chế độ xem: Normal View vs Page Break Preview
        ctk.CTkLabel(toolbar, text="Chế độ:", font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")).pack(side="left", padx=(10, 4))
        self.btn_normal_view = ctk.CTkButton(
            toolbar,
            text="📊 Normal View",
            width=105,
            height=28,
            command=lambda: self._set_view_mode("normal"),
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color=("#2563eb", "#1d4ed8"),
            hover_color=("#1d4ed8", "#2563eb"),
        )
        self.btn_normal_view.pack(side="left", padx=2)

        self.btn_page_break_view = ctk.CTkButton(
            toolbar,
            text="📄 Page Break Preview",
            width=135,
            height=28,
            command=lambda: self._set_view_mode("page_break"),
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color=("#64748b", "#334155"),
            hover_color=("#475569", "#64748b"),
        )
        self.btn_page_break_view.pack(side="left", padx=2)

        # Phân cách
        sep1 = ctk.CTkFrame(toolbar, width=1, height=22, fg_color=("#cbd5e1", "#475569"))
        sep1.pack(side="left", padx=8)

        # 2. Font Family
        ctk.CTkLabel(toolbar, text="Font:", font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")).pack(side="left", padx=(0, 4))
        self.font_family_var = tk.StringVar(value="Calibri")
        self.cb_font = ctk.CTkComboBox(
            toolbar,
            values=["Calibri", "Arial", "Segoe UI", "Times New Roman", "Tahoma", "Consolas"],
            variable=self.font_family_var,
            width=115,
            height=28,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            command=self._apply_font,
        )
        self.cb_font.pack(side="left", padx=(0, 6))

        # 3. Font Size
        ctk.CTkLabel(toolbar, text="Cỡ:", font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")).pack(side="left", padx=(0, 4))
        self.font_size_var = tk.StringVar(value="11")
        self.cb_size = ctk.CTkComboBox(
            toolbar,
            values=["9", "10", "11", "12", "14", "16", "18", "20"],
            variable=self.font_size_var,
            width=65,
            height=28,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            command=self._apply_font,
        )
        self.cb_size.pack(side="left", padx=(0, 6))

        # 4. Bold / Italic
        self.btn_bold = ctk.CTkButton(
            toolbar,
            text="B",
            width=28,
            height=28,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            command=self._toggle_bold,
            fg_color=("#94a3b8", "#334155"),
        )
        self.btn_bold.pack(side="left", padx=2)

        self.btn_italic = ctk.CTkButton(
            toolbar,
            text="I",
            width=28,
            height=28,
            font=ctk.CTkFont(family="Segoe UI", size=11, slant="italic"),
            command=self._toggle_italic,
            fg_color=("#94a3b8", "#334155"),
        )
        self.btn_italic.pack(side="left", padx=2)

        # Phân cách
        sep2 = ctk.CTkFrame(toolbar, width=1, height=22, fg_color=("#cbd5e1", "#475569"))
        sep2.pack(side="left", padx=8)

        # 5. Align Text
        self.btn_align_left = ctk.CTkButton(
            toolbar,
            text="⬅",
            width=32,
            height=28,
            command=lambda: self._set_alignment("left"),
            fg_color=("#94a3b8", "#334155"),
        )
        self.btn_align_left.pack(side="left", padx=2)

        self.btn_align_center = ctk.CTkButton(
            toolbar,
            text="↔",
            width=32,
            height=28,
            command=lambda: self._set_alignment("center"),
            fg_color=("#94a3b8", "#334155"),
        )
        self.btn_align_center.pack(side="left", padx=2)

        self.btn_align_right = ctk.CTkButton(
            toolbar,
            text="➡",
            width=32,
            height=28,
            command=lambda: self._set_alignment("right"),
            fg_color=("#94a3b8", "#334155"),
        )
        self.btn_align_right.pack(side="left", padx=2)

    def _build_sheet_tabs(self):
        # Hàng tab chọn sheet
        self.tabs_frame = ctk.CTkFrame(self, height=32, fg_color="transparent")
        self.tabs_frame.pack(fill="x", padx=6, pady=(0, 2))

    def _build_editor_area(self):
        # Khung chứa bảng tính / thanh trượt điều chỉnh trang
        self.main_content = ctk.CTkFrame(self, fg_color="transparent")
        self.main_content.pack(fill="both", expand=True, padx=4, pady=2)

        # Thanh điều khiển kéo độ rộng trang (hiển thị khi ở chế độ Page Break Preview)
        self.page_break_control_bar = ctk.CTkFrame(self.main_content, height=38, fg_color=("#dbeafe", "#1e3a8a"), corner_radius=6)

        ctk.CTkLabel(
            self.page_break_control_bar,
            text="✂️ Điều chỉnh độ rộng Trang 1 (Vạch ngắt trang):",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=("#1e3a8a", "#93c5fd"),
        ).pack(side="left", padx=(12, 8))

        self.page_width_slider = ctk.CTkSlider(
            self.page_break_control_bar,
            from_=2,
            to=12,
            number_of_steps=10,
            width=180,
            command=self._on_page_break_slider_change,
        )
        self.page_width_slider.set(self.page_break_col_index)
        self.page_width_slider.pack(side="left", padx=6)

        self.lbl_page_col_info = ctk.CTkLabel(
            self.page_break_control_bar,
            text=f"Cột A đến {get_column_letter(self.page_break_col_index)}",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=("#1e3a8a", "#bfdbfe"),
        )
        self.lbl_page_col_info.pack(side="left", padx=(4, 16))

        # Nút kéo nhanh
        btn_shrink = ctk.CTkButton(
            self.page_break_control_bar,
            text="◀ Thu hẹp",
            width=70,
            height=24,
            command=self._shrink_page_break,
            fg_color=("#2563eb", "#1d4ed8"),
        )
        btn_shrink.pack(side="left", padx=2)

        btn_expand = ctk.CTkButton(
            self.page_break_control_bar,
            text="Mở rộng ▶",
            width=75,
            height=24,
            command=self._expand_page_break,
            fg_color=("#2563eb", "#1d4ed8"),
        )
        btn_expand.pack(side="left", padx=2)

        # Khung công thức / Sửa ô (fx bar)
        fx_bar = ctk.CTkFrame(self.main_content, height=32, fg_color="transparent")
        fx_bar.pack(fill="x", pady=(2, 4))

        self.lbl_cell_coord = ctk.CTkLabel(fx_bar, text="A1", width=45, font=ctk.CTkFont(family="Consolas", size=11, weight="bold"))
        self.lbl_cell_coord.pack(side="left", padx=(4, 6))

        ctk.CTkLabel(fx_bar, text="fx:", font=ctk.CTkFont(family="Segoe UI", size=11, slant="italic")).pack(side="left", padx=(0, 4))

        self.entry_cell_val = ctk.CTkEntry(
            fx_bar,
            height=28,
            font=ctk.CTkFont(family="Segoe UI", size=11),
        )
        self.entry_cell_val.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self.entry_cell_val.bind("<Return>", lambda e: self._apply_cell_edit())

        btn_apply_cell = ctk.CTkButton(
            fx_bar,
            text="Cập nhật ô",
            width=80,
            height=28,
            command=self._apply_cell_edit,
            fg_color=("#0284c7", "#0369a1"),
        )
        btn_apply_cell.pack(side="right")

        # Bảng hiển thị Treeview các ô tính
        self.grid_container = ctk.CTkFrame(self.main_content, fg_color="transparent")
        self.grid_container.pack(fill="both", expand=True)
        self.grid_container.grid_rowconfigure(0, weight=1)
        self.grid_container.grid_columnconfigure(0, weight=1)

        self.sheet_tree = ttk.Treeview(self.grid_container, show="headings", style="Custom.Treeview")
        vsb = ttk.Scrollbar(self.grid_container, orient="vertical", command=self.sheet_tree.yview)
        hsb = ttk.Scrollbar(self.grid_container, orient="horizontal", command=self.sheet_tree.xview)
        self.sheet_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.sheet_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self.sheet_tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.sheet_tree.bind("<Double-1>", self._on_tree_cell_double_click)

    def _load_workbook(self):
        if not OPENPYXL_AVAILABLE:
            return

        if not os.path.exists(self.filepath):
            return

        try:
            self.wb = openpyxl.load_workbook(self.filepath, data_only=False)

            # Khắc phục lỗi nội bộ của openpyxl: PrintPageSetup._parent bị None
            for ws_item in self.wb.worksheets:
                try:
                    if hasattr(ws_item, "page_setup") and ws_item.page_setup is not None:
                        if getattr(ws_item.page_setup, "_parent", None) is None:
                            ws_item.page_setup._parent = ws_item
                except Exception:
                    pass

            sheet_names = self.wb.sheetnames
            if not sheet_names:
                return

            self.active_sheet_name = sheet_names[0]
            self._refresh_sheet_tabs()
            self._render_current_sheet()

        except Exception as e:
            print(f"Lỗi nạp workbook: {e}")

    def _switch_sheet(self, sheet_name: str):
        self.active_sheet_name = sheet_name
        self._refresh_sheet_tabs()
        self._render_current_sheet()

    def _refresh_sheet_tabs(self):
        """Vẽ lại các tab sheet, làm nổi bật tab đang active."""
        for w in self.tabs_frame.winfo_children():
            w.destroy()
        if not self.wb:
            return
        for s_name in self.wb.sheetnames:
            is_active = (s_name == self.active_sheet_name)
            btn_tab = ctk.CTkButton(
                self.tabs_frame,
                text=f"📋 {s_name}",
                height=26,
                command=lambda n=s_name: self._switch_sheet(n),
                fg_color=("#2563eb", "#1d4ed8") if is_active else ("#94a3b8", "#334155"),
                hover_color=("#1d4ed8", "#1e40af") if is_active else ("#64748b", "#475569"),
            )
            btn_tab.pack(side="left", padx=2)

    def _load_workbook_tabs_only(self):
        """Backward-compat alias for _refresh_sheet_tabs."""
        self._refresh_sheet_tabs()

    def _render_current_sheet(self):
        if not self.wb or self.active_sheet_name not in self.wb.sheetnames:
            return

        ws = self.wb[self.active_sheet_name]
        max_row = max(15, ws.max_row)
        max_col = max(8, ws.max_column)

        # Xây dựng các cột: Cột 0 là Row Number (1, 2, 3...)
        col_ids = ["_row_num"] + [get_column_letter(i) for i in range(1, max_col + 1)]
        self.sheet_tree["columns"] = col_ids

        # Cấu hình header
        self.sheet_tree.heading("_row_num", text="#")
        self.sheet_tree.column("_row_num", width=40, anchor="center")

        for col_idx in range(1, max_col + 1):
            letter = get_column_letter(col_idx)
            header_text = letter
            # Nếu đang ở chế độ Page Break Preview, hiển thị rõ cột nào thuộc Trang 1 / Trang 2
            if self.view_mode == "page_break":
                if col_idx <= self.page_break_col_index:
                    header_text = f"{letter} [P1]"
                else:
                    header_text = f"{letter} [P2]"

            self.sheet_tree.heading(letter, text=header_text)
            self.sheet_tree.column(letter, width=90, anchor="w")

        # Xóa các dòng cũ
        for it in self.sheet_tree.get_children():
            self.sheet_tree.delete(it)

        # Điền dữ liệu
        for r in range(1, max_row + 1):
            row_vals = [str(r)]
            for c in range(1, max_col + 1):
                val = ws.cell(row=r, column=c).value
                val_str = "" if val is None else str(val)
                row_vals.append(val_str)

            tag = "even_row" if r % 2 == 0 else "odd_row"
            self.sheet_tree.insert("", "end", iid=str(r), values=row_vals, tags=(tag,))

        # Định dạng màu cho Treeview
        if self.view_mode == "page_break":
            self.sheet_tree.tag_configure("even_row", background="#111827", foreground="#e2e8f0")
            self.sheet_tree.tag_configure("odd_row", background="#1e293b", foreground="#f8fafc")
        else:
            self.sheet_tree.tag_configure("even_row", background="#1e293b", foreground="#f1f5f9")
            self.sheet_tree.tag_configure("odd_row", background="#0f172a", foreground="#f8fafc")

    def _set_view_mode(self, mode: str):
        self.view_mode = mode
        if mode == "page_break":
            self.btn_page_break_view.configure(fg_color=("#2563eb", "#1d4ed8"))
            self.btn_normal_view.configure(fg_color=("#64748b", "#334155"))
            self.page_break_control_bar.pack(fill="x", pady=(0, 4), before=self.grid_container)
        else:
            self.btn_normal_view.configure(fg_color=("#2563eb", "#1d4ed8"))
            self.btn_page_break_view.configure(fg_color=("#64748b", "#334155"))
            self.page_break_control_bar.pack_forget()

        self._render_current_sheet()

    def _on_page_break_slider_change(self, val):
        self.page_break_col_index = int(val)
        col_letter = get_column_letter(self.page_break_col_index)
        self.lbl_page_col_info.configure(text=f"Cột A đến {col_letter}")
        self._render_current_sheet()

    def _shrink_page_break(self):
        if self.page_break_col_index > 2:
            self.page_break_col_index -= 1
            self.page_width_slider.set(self.page_break_col_index)
            self._on_page_break_slider_change(self.page_break_col_index)

    def _expand_page_break(self):
        if self.page_break_col_index < 12:
            self.page_break_col_index += 1
            self.page_width_slider.set(self.page_break_col_index)
            self._on_page_break_slider_change(self.page_break_col_index)

    def _on_tree_select(self, event):
        sel = self.sheet_tree.selection()
        if not sel:
            return
        row_id = sel[0]
        self._current_row_id = row_id
        # Giữ nguyên cột đang chọn (_current_col_letter), cập nhật dòng
        coord = f"{self._current_col_letter}{row_id}"
        self.lbl_cell_coord.configure(text=coord)
        # Hiển thị giá trị ô tương ứng
        item_vals = self.sheet_tree.item(row_id, "values")
        col_ids = list(self.sheet_tree["columns"])
        try:
            col_pos = col_ids.index(self._current_col_letter)
            if col_pos < len(item_vals):
                self.entry_cell_val.delete(0, "end")
                self.entry_cell_val.insert(0, item_vals[col_pos])
        except (ValueError, IndexError):
            if len(item_vals) > 1:
                self.entry_cell_val.delete(0, "end")
                self.entry_cell_val.insert(0, item_vals[1])

    def _on_tree_cell_double_click(self, event):
        col_id = self.sheet_tree.identify_column(event.x)
        row_id = self.sheet_tree.identify_row(event.y)
        if not row_id or not col_id:
            return
        col_num = int(col_id.replace("#", ""))
        col_ids = list(self.sheet_tree["columns"])
        if col_num < 1 or col_num > len(col_ids):
            return
        col_name = col_ids[col_num - 1]
        if col_name == "_row_num":
            return  # bỏ qua cột số thứ tự
        # Cập nhật tracking ô đang chọn
        self._current_col_letter = col_name
        self._current_row_id = row_id
        coord = f"{col_name}{row_id}"
        self.lbl_cell_coord.configure(text=coord)
        vals = self.sheet_tree.item(row_id, "values")
        if col_num <= len(vals):
            self.entry_cell_val.delete(0, "end")
            self.entry_cell_val.insert(0, vals[col_num - 1])
        self.entry_cell_val.focus_set()

    def _apply_cell_edit(self):
        coord = self.lbl_cell_coord.cget("text")
        new_val = self.entry_cell_val.get()
        if not self.wb or not coord or len(coord) < 2:
            return

        ws = self.wb[self.active_sheet_name]
        try:
            # Thử chuyển thành số nếu có thể
            try:
                parsed: object = float(new_val) if "." in new_val else int(new_val)
            except (ValueError, TypeError):
                parsed = new_val
            ws[coord] = parsed
            self._render_current_sheet()
            # Giữ lại selection sau khi render
            if self._current_row_id in self.sheet_tree.get_children():
                self.sheet_tree.selection_set(self._current_row_id)
                self.sheet_tree.see(self._current_row_id)
        except Exception as e:
            messagebox.showerror("Lỗi sửa ô", f"Không thể cập nhật ô {coord}:\n{str(e)}")

    def _apply_font(self, val=None):
        if not self.wb:
            return
        coord = self.lbl_cell_coord.cget("text")
        if not coord or len(coord) < 2:
            return
        font_name = self.font_family_var.get() or "Calibri"
        try:
            sz = int(self.font_size_var.get())
        except Exception:
            sz = 11

        ws = self.wb[self.active_sheet_name]
        try:
            cell = ws[coord]
            cur_bold = bool(cell.font.bold) if cell.font else False
            cur_italic = bool(cell.font.italic) if cell.font else False
            cell.font = Font(name=font_name, size=sz, bold=cur_bold, italic=cur_italic)
            self._render_current_sheet()
        except Exception as e:
            print(f"_apply_font error: {e}")

    def _toggle_bold(self):
        if not self.wb:
            return
        coord = self.lbl_cell_coord.cget("text")
        if not coord or len(coord) < 2:
            return
        ws = self.wb[self.active_sheet_name]
        try:
            cell = ws[coord]
            fn = self.font_family_var.get() or "Calibri"
            try:
                sz = int(self.font_size_var.get())
            except Exception:
                sz = 11
            cur_bold = bool(cell.font.bold) if cell.font else False
            cur_italic = bool(cell.font.italic) if cell.font else False
            cell.font = Font(name=fn, size=sz, bold=not cur_bold, italic=cur_italic)
            self._render_current_sheet()
        except Exception as e:
            print(f"_toggle_bold error: {e}")

    def _toggle_italic(self):
        if not self.wb:
            return
        coord = self.lbl_cell_coord.cget("text")
        if not coord or len(coord) < 2:
            return
        ws = self.wb[self.active_sheet_name]
        try:
            cell = ws[coord]
            fn = self.font_family_var.get() or "Calibri"
            try:
                sz = int(self.font_size_var.get())
            except Exception:
                sz = 11
            cur_bold = bool(cell.font.bold) if cell.font else False
            cur_italic = bool(cell.font.italic) if cell.font else False
            cell.font = Font(name=fn, size=sz, bold=cur_bold, italic=not cur_italic)
            self._render_current_sheet()
        except Exception as e:
            print(f"_toggle_italic error: {e}")

    def _set_alignment(self, align: str):
        if not self.wb:
            return
        coord = self.lbl_cell_coord.cget("text")
        if not coord or len(coord) < 2:
            return
        ws = self.wb[self.active_sheet_name]
        try:
            ws[coord].alignment = Alignment(horizontal=align)
            self._render_current_sheet()
        except Exception as e:
            print(f"_set_alignment error: {e}")

    def save_changes(self) -> Tuple[bool, str]:
        """Lưu lại dữ liệu bảng tính, font chữ, căn lề và thiết lập ngắt trang vào file .xlsx."""
        if not OPENPYXL_AVAILABLE or self.wb is None:
            return False, "Không thể lưu vì thư viện openpyxl chưa sẵn sàng."

        try:
            # An toàn: Khắc phục page_setup._parent bị None trước khi lưu
            self._fix_all_page_setup_parents()

            # Nếu đang ở Page Break Preview, áp dụng ngắt trang an toàn
            if self.view_mode == "page_break":
                self._safe_apply_page_breaks()

            self.wb.save(self.filepath)
            return True, f"Đã lưu thành công các thay đổi vào file: {os.path.basename(self.filepath)}"

        except PermissionError:
            return False, "❌ Không thể lưu vì file đang được mở trong phần mềm khác (Microsoft Excel).\nVui lòng đóng file trong Excel rồi nhấn Lưu lại!"
        except Exception as e:
            return False, f"❌ Lỗi khi lưu file XLSX: {e}"

    def _fix_all_page_setup_parents(self):
        """Khắc phục lỗi _parent bị None cho tất cả các worksheet trong workbook."""
        if not self.wb:
            return
        for ws in self.wb.worksheets:
            try:
                ps = getattr(ws, "page_setup", None)
                if ps is not None and getattr(ps, "_parent", None) is None:
                    ps._parent = ws
            except Exception:
                pass

    def _safe_apply_page_breaks(self):
        """Áp dụng ngắt trang theo cột an toàn, KHÔNG truy cập sheet_properties.pageSetUpPr nếu chưa tồn tại."""
        if not self.wb:
            return
        for ws in self.wb.worksheets:
            try:
                # Cài đặt ngắt trang cột
                if hasattr(ws, "col_breaks") and ws.col_breaks is not None:
                    try:
                        ws.col_breaks.brk = []
                        from openpyxl.worksheet.pagebreak import Break
                        brk = Break(id=self.page_break_col_index, min=1, max=16383, man=True)
                        ws.col_breaks.brk.append(brk)
                    except Exception:
                        pass
                # CHỈ đặt fitToPage nếu pageSetUpPr thực sự tồn tại
                try:
                    sp = getattr(ws, "sheet_properties", None)
                    if sp is not None:
                        psp = getattr(sp, "pageSetUpPr", None)
                        if psp is not None:
                            psp.fitToPage = True
                except Exception:
                    pass  # An toàn: bỏ qua nếu không hỗ trợ
            except Exception as e:
                print(f"_safe_apply_page_breaks ws={ws.title}: {e}")


# ==============================================================================
# MASTER DOCUMENT PREVIEW & EDITOR CONTAINER
# ==============================================================================

class DocumentPreviewEditorView(ctk.CTkFrame):
    """
    Vùng xem trước tổng thể (Thay thế vị trí bảng hàng đợi in khi người dùng mở Xem Trước).
    Tích hợp thanh Header điều khiển, Trình chỉnh sửa DOCX hoặc XLSX,
    Nút Lưu (Save) & Đóng (Close) đưa giao diện trở lại bình thường.
    """
    def __init__(
        self,
        master,
        filepath: str,
        filename: str,
        on_save_callback: Callable[[], None],
        on_close_callback: Callable[[], None],
    ):
        super().__init__(master, fg_color="transparent")
        self.filepath = filepath
        self.filename = filename
        self.on_save_callback = on_save_callback
        self.on_close_callback = on_close_callback

        ext = os.path.splitext(self.filepath)[1].lower()
        self.file_type = ext

        self._build_header()
        self._build_body()

    def _build_header(self):
        header = ctk.CTkFrame(self, height=44, corner_radius=8, fg_color=("#0f172a", "#0f172a"))
        header.pack(fill="x", padx=4, pady=(2, 6))

        # Nút Quay lại / Đóng
        btn_back = ctk.CTkButton(
            header,
            text="⬅️ Quay lại hàng đợi",
            width=140,
            height=30,
            command=self._on_close_clicked,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color=("#475569", "#334155"),
            hover_color=("#334155", "#475569"),
        )
        btn_back.pack(side="left", padx=8, pady=6)

        # Thông tin file đang xem
        icon = "📝" if self.file_type in [".doc", ".docx"] else ("📊" if self.file_type in [".xls", ".xlsx"] else "📄")
        lbl_file = ctk.CTkLabel(
            header,
            text=f"{icon} ĐANG XEM TRƯỚC & CHỈNH SỬA: {self.filename}",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color="#38bdf8",
        )
        lbl_file.pack(side="left", padx=10)

        # Nút Đóng nhanh
        btn_close = ctk.CTkButton(
            header,
            text="✖️ Đóng",
            width=70,
            height=30,
            command=self._on_close_clicked,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=("#dc2626", "#b91c1c"),
            hover_color=("#b91c1c", "#dc2626"),
        )
        btn_close.pack(side="right", padx=(4, 8), pady=6)

        # Nút LƯU (SAVE)
        self.btn_save = ctk.CTkButton(
            header,
            text="💾 LƯU THAY ĐỔI (SAVE)",
            width=180,
            height=30,
            command=self._on_save_clicked,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color=("#16a34a", "#15803d"),
            hover_color=("#15803d", "#16a34a"),
        )
        self.btn_save.pack(side="right", padx=6, pady=6)

    def _build_body(self):
        body_container = ctk.CTkFrame(self, fg_color="transparent")
        body_container.pack(fill="both", expand=True)

        if self.file_type in [".doc", ".docx"]:
            self.editor = DocxEditorComponent(body_container, self.filepath)
            self.editor.pack(fill="both", expand=True)
        elif self.file_type in [".xls", ".xlsx"]:
            self.editor = XlsxEditorComponent(body_container, self.filepath)
            self.editor.pack(fill="both", expand=True)
        else:
            # Fallback cho định dạng khác (Text, Log, ...)
            self.editor = None
            lbl_fallback = ctk.CTkLabel(
                body_container,
                text=f"Định dạng {self.file_type} không hỗ trợ chỉnh sửa trực tiếp dạng Word/Excel.\n\nNhấn click đúp vào tên file ở hàng đợi để mở bằng phần mềm chuyên dụng.",
                font=ctk.CTkFont(family="Segoe UI", size=13),
                text_color="#94a3b8",
            )
            lbl_fallback.pack(expand=True)

    def _on_save_clicked(self):
        if self.editor and hasattr(self.editor, "save_changes"):
            ok, msg = self.editor.save_changes()
            if ok:
                messagebox.showinfo("Lưu thành công", msg)
                self.on_save_callback()
            else:
                messagebox.showerror("Lỗi khi lưu", msg)
        else:
            self.on_save_callback()

    def _on_close_clicked(self):
        self.on_close_callback()
