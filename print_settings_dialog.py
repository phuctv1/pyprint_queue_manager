import os
import tkinter as tk
from tkinter import messagebox, ttk
from typing import Callable, Dict, List, Optional, Tuple

import customtkinter as ctk
from PIL import Image, ImageDraw, ImageTk

from models import PrintItem


def create_setting_icon(icon_name: str, size: Tuple[int, int] = (28, 28)) -> ImageTk.PhotoImage:
    """Tạo icon vector đường nét tinh tế chuẩn Office (Dark theme)."""
    w, h = size
    # Scale x2 để vẽ siêu nét (supersampling antialiasing)
    img = Image.new("RGBA", (w * 2, h * 2), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    stroke = (210, 220, 230, 255)
    accent = (56, 189, 248, 255)  # Cyan/Blue accent
    sub_stroke = (140, 155, 175, 255)
    line_w = 3

    if icon_name == "doc":
        # Trang tài liệu có góc gập
        d.polygon([(8, 4), (38, 4), (48, 14), (48, 52), (8, 52)], outline=stroke, width=line_w)
        d.line([(38, 4), (38, 14), (48, 14)], fill=stroke, width=line_w)
        # Các dòng chữ bên trong
        d.line([(16, 22), (40, 22)], fill=sub_stroke, width=2)
        d.line([(16, 30), (40, 30)], fill=sub_stroke, width=2)
        d.line([(16, 38), (32, 38)], fill=sub_stroke, width=2)

    elif icon_name == "sheets_grid":
        # Bảng tính Excel grid
        d.rectangle([(6, 6), (50, 48)], outline=stroke, width=line_w)
        d.line([(6, 20), (50, 20)], fill=stroke, width=2)
        d.line([(6, 34), (50, 34)], fill=stroke, width=2)
        d.line([(20, 6), (20, 48)], fill=stroke, width=2)
        d.line([(36, 6), (36, 48)], fill=stroke, width=2)

    elif icon_name == "duplex":
        # Hai trang tài liệu lật 2 mặt
        d.rectangle([(6, 8), (36, 50)], outline=stroke, width=line_w)
        d.line([(12, 18), (30, 18)], fill=sub_stroke, width=2)
        d.line([(12, 26), (30, 26)], fill=sub_stroke, width=2)
        # Trang sau hơi chéo / lật
        d.line([(36, 14), (48, 14), (48, 44), (36, 44)], fill=accent, width=2)

    elif icon_name == "simplex":
        # Trang 1 mặt
        d.rectangle([(8, 8), (46, 48)], outline=stroke, width=line_w)
        d.line([(14, 18), (40, 18)], fill=sub_stroke, width=2)
        d.line([(14, 28), (40, 28)], fill=sub_stroke, width=2)
        d.line([(14, 38), (32, 38)], fill=sub_stroke, width=2)

    elif icon_name == "collated":
        # Hai trang đặt cạnh nhau [1] [2]
        d.rectangle([(4, 10), (26, 46)], outline=stroke, width=line_w)
        d.rectangle([(30, 10), (52, 46)], outline=stroke, width=line_w)

    elif icon_name == "portrait":
        # Trang dọc
        d.rectangle([(12, 6), (44, 50)], outline=stroke, width=line_w)
        d.polygon([(36, 6), (44, 14), (36, 14)], fill=sub_stroke)

    elif icon_name == "landscape":
        # Trang ngang
        d.rectangle([(4, 14), (52, 44)], outline=stroke, width=line_w)
        d.polygon([(44, 14), (52, 22), (44, 22)], fill=sub_stroke)

    elif icon_name == "paper_size":
        # Khung khổ giấy
        d.rectangle([(10, 8), (46, 48)], outline=stroke, width=line_w)

    elif icon_name == "margins_normal":
        # Khung lề giấy bình thường
        d.rectangle([(6, 6), (50, 50)], outline=sub_stroke, width=1)
        d.rectangle([(16, 16), (40, 40)], outline=accent, width=2)

    elif icon_name == "margins_narrow":
        # Khung lề giấy hẹp (vùng trong lớn hơn)
        d.rectangle([(6, 6), (50, 50)], outline=sub_stroke, width=1)
        d.rectangle([(10, 10), (46, 46)], outline=accent, width=2)

    elif icon_name == "page_per_sheet":
        # 1 trang trên 1 tờ
        d.rectangle([(6, 6), (50, 50)], outline=stroke, width=line_w)
        d.rectangle([(14, 12), (42, 44)], outline=sub_stroke, width=2)
        d.line([(18, 20), (38, 20)], fill=sub_stroke, width=1)
        d.line([(18, 26), (38, 26)], fill=sub_stroke, width=1)

    elif icon_name == "no_scaling":
        # Không co giãn (100%)
        d.rectangle([(6, 6), (50, 48)], outline=stroke, width=line_w)
        # Vẽ biểu tượng 100
        d.text((12, 20), "100", fill=accent)
    else:
        d.rectangle([(8, 8), (48, 48)], outline=stroke, width=line_w)

    img = img.resize(size, Image.Resampling.LANCZOS)
    return ImageTk.PhotoImage(img)


class OfficeSettingDropdown(ctk.CTkFrame):
    """
    Widget mô phỏng dropdown cài đặt in ấn chuẩn Microsoft Office:
    Gồm Icon bên trái, Tiêu đề (dòng 1), Mô tả phụ (dòng 2), Chevron xổ xuống bên phải.
    """
    def __init__(
        self,
        master,
        icon_name: str,
        options: List[Tuple[str, str]],  # List of (Title, Subtitle)
        initial_index: int = 0,
        on_change: Optional[Callable[[str, str], None]] = None,
        width: int = 280,
    ):
        super().__init__(
            master,
            width=width,
            height=54,
            corner_radius=4,
            fg_color="#18181b",
            border_width=1,
            border_color="#3f3f46",
        )
        self.grid_propagate(False)
        self.pack_propagate(False)

        self.options = options
        self.current_index = initial_index if 0 <= initial_index < len(options) else 0
        self.on_change = on_change
        self._icon_image = create_setting_icon(icon_name, (26, 26))

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Icon cột 0
        self.lbl_icon = tk.Label(
            self,
            image=self._icon_image,
            bg="#18181b",
            bd=0,
            cursor="hand2",
        )
        self.lbl_icon.grid(row=0, column=0, rowspan=2, padx=(8, 10), pady=6, sticky="w")

        # Text container cột 1
        self.text_frame = tk.Frame(self, bg="#18181b", cursor="hand2")
        self.text_frame.grid(row=0, column=1, rowspan=2, sticky="w", pady=4)

        cur_title, cur_sub = self.options[self.current_index]

        self.lbl_title = tk.Label(
            self.text_frame,
            text=cur_title,
            font=("Segoe UI", 10, "normal"),
            fg="#f4f4f5",
            bg="#18181b",
            anchor="w",
            cursor="hand2",
        )
        self.lbl_title.pack(anchor="w")

        self.lbl_sub = tk.Label(
            self.text_frame,
            text=cur_sub if cur_sub else "",
            font=("Segoe UI", 8),
            fg="#a1a1aa",
            bg="#18181b",
            anchor="w",
            cursor="hand2",
        )
        if cur_sub:
            self.lbl_sub.pack(anchor="w")

        # Chevron cột 2
        self.lbl_arrow = tk.Label(
            self,
            text="∨",
            font=("Segoe UI", 9, "bold"),
            fg="#a1a1aa",
            bg="#18181b",
            cursor="hand2",
        )
        self.lbl_arrow.grid(row=0, column=2, rowspan=2, padx=(6, 10), sticky="e")

        # Bind sự kiện click mở menu
        widgets = [self, self.lbl_icon, self.text_frame, self.lbl_title, self.lbl_sub, self.lbl_arrow]
        for w in widgets:
            w.bind("<Button-1>", self._on_click)
            w.bind("<Enter>", self._on_enter)
            w.bind("<Leave>", self._on_leave)

    def _on_enter(self, e):
        self.configure(border_color="#71717a", fg_color="#27272a")
        for w in [self.lbl_icon, self.text_frame, self.lbl_title, self.lbl_sub, self.lbl_arrow]:
            w.configure(bg="#27272a")

    def _on_leave(self, e):
        self.configure(border_color="#3f3f46", fg_color="#18181b")
        for w in [self.lbl_icon, self.text_frame, self.lbl_title, self.lbl_sub, self.lbl_arrow]:
            w.configure(bg="#18181b")

    def _on_click(self, e):
        menu = tk.Menu(self, tearoff=0, bg="#27272a", fg="#f4f4f5", activebackground="#3b82f6", activeforeground="#ffffff", font=("Segoe UI", 9))
        for idx, (title, sub) in enumerate(self.options):
            label = f"{title}"
            if sub:
                label += f" ({sub})"
            menu.add_command(label=label, command=lambda i=idx: self.set_selection(i))

        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height()
        menu.post(x, y)

    def set_selection(self, index: int):
        if 0 <= index < len(self.options):
            self.current_index = index
            title, sub = self.options[index]
            self.lbl_title.configure(text=title)
            if sub:
                self.lbl_sub.configure(text=sub)
                if not self.lbl_sub.winfo_ismapped():
                    self.lbl_sub.pack(anchor="w")
            else:
                self.lbl_sub.pack_forget()

            if self.on_change:
                self.on_change(title, sub)

    def get_selected(self) -> Tuple[str, str]:
        return self.options[self.current_index]


class OfficePrintSettingsDialog(ctk.CTkToplevel):
    """
    Cửa sổ Print Settings:
    - Nếu file là DOC/DOCX: Giao diện chuẩn hình ảnh 2.
    - Nếu file là XLS/XLSX: Giao diện chuẩn hình ảnh 3.
    """
    def __init__(
        self,
        parent,
        item: Optional[PrintItem] = None,
        all_items: Optional[List[PrintItem]] = None,
        on_apply: Optional[Callable[[dict, bool], None]] = None,
    ):
        super().__init__(parent)
        self.parent = parent
        self.item = item
        self.all_items = all_items or []
        self.on_apply = on_apply

        # Xác định chế độ DOC/DOCX hay XLS/XLSX
        self.mode = "docx"  # "docx" hoặc "xlsx"
        if self.item:
            ext = os.path.splitext(self.item.filepath)[1].lower()
            if ext in [".xls", ".xlsx"]:
                self.mode = "xlsx"

        self.title("Print Settings" if not item else f"Print Settings - {item.filename}")
        self.geometry("340x620")
        self.resizable(False, False)
        self.configure(fg_color="#0f0f11")
        self.attributes("-topmost", True)

        # Tập hợp các widget cài đặt để lấy giá trị khi Save
        self.widgets_map = {}

        self._build_ui()
        self.center_window()

    def center_window(self):
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        px = self.parent.winfo_rootx() + (self.parent.winfo_width() - w) // 2
        py = self.parent.winfo_rooty() + (self.parent.winfo_height() - h) // 2
        self.geometry(f"+{max(10, px)}+{max(10, py)}")

    def _build_ui(self):
        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=16, pady=14)

        # Header: Tiêu đề "Settings"
        header_box = ctk.CTkFrame(container, fg_color="transparent")
        header_box.pack(fill="x", pady=(0, 10))

        lbl_settings = tk.Label(
            header_box,
            text="Settings",
            font=("Segoe UI", 13, "bold"),
            fg="#f4f4f5",
            bg="#0f0f11",
        )
        lbl_settings.pack(side="left")

        # Nút chuyển đổi định dạng (nếu người dùng muốn xem dạng Word hoặc Excel)
        mode_btn = tk.Label(
            header_box,
            text=f"[{'Word Mode' if self.mode == 'docx' else 'Excel Mode'}]",
            font=("Segoe UI", 9, "italic"),
            fg="#38bdf8",
            bg="#0f0f11",
            cursor="hand2",
        )
        mode_btn.pack(side="right")
        mode_btn.bind("<Button-1>", lambda e: self._toggle_mode())

        # Vùng chứa các mục settings cuộn
        self.body_frame = ctk.CTkFrame(container, fg_color="transparent")
        self.body_frame.pack(fill="both", expand=True)

        if self.mode == "docx":
            self._build_docx_settings()
        else:
            self._build_xlsx_settings()

        # Link "Page Setup" ở góc phải
        link_box = ctk.CTkFrame(container, fg_color="transparent")
        link_box.pack(fill="x", pady=(8, 12))

        lbl_page_setup = tk.Label(
            link_box,
            text="Page Setup",
            font=("Segoe UI", 10, "underline"),
            fg="#38bdf8",
            bg="#0f0f11",
            cursor="hand2",
        )
        lbl_page_setup.pack(side="right")
        lbl_page_setup.bind("<Button-1>", lambda e: self._open_page_setup_modal())

        # Hàng nút Áp dụng và Đóng
        btn_bar = ctk.CTkFrame(container, fg_color="transparent")
        btn_bar.pack(fill="x")

        btn_apply = ctk.CTkButton(
            btn_bar,
            text="Lưu & Áp Dụng",
            height=32,
            command=self._on_apply_clicked,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color="#2563eb",
            hover_color="#1d4ed8",
        )
        btn_apply.pack(side="left", fill="x", expand=True, padx=(0, 6))

        btn_close = ctk.CTkButton(
            btn_bar,
            text="Đóng",
            height=32,
            width=70,
            command=self.destroy,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color="#3f3f46",
            hover_color="#52525b",
        )
        btn_close.pack(side="right")

    def _toggle_mode(self):
        self.mode = "xlsx" if self.mode == "docx" else "docx"
        for w in self.body_frame.winfo_children():
            w.destroy()
        if self.mode == "docx":
            self._build_docx_settings()
        else:
            self._build_xlsx_settings()

    def _build_docx_settings(self):
        """Xây dựng giao diện cài đặt file DOC/DOCX theo ảnh thứ 2."""
        # 1. Print All Pages
        opt_pages = [
            ("Print All Pages", "The whole thing"),
            ("Print Selection", "Only print current selection"),
            ("Print Current Page", "Only this page"),
            ("Custom Print", "Enter specific pages"),
        ]
        self.dd_print_pages = OfficeSettingDropdown(self.body_frame, "doc", opt_pages, 0)
        self.dd_print_pages.pack(fill="x", pady=3)

        # 2. Pages Input: Pages: [        ] (i)
        pages_input_frame = tk.Frame(self.body_frame, bg="#0f0f11", height=34)
        pages_input_frame.pack(fill="x", pady=(3, 3))
        pages_input_frame.pack_propagate(False)

        lbl_pages = tk.Label(pages_input_frame, text="Pages:", font=("Segoe UI", 10), fg="#f4f4f5", bg="#0f0f11")
        lbl_pages.pack(side="left", padx=(4, 8))

        self.entry_pages = ctk.CTkEntry(
            pages_input_frame,
            height=28,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color="#18181b",
            border_color="#3f3f46",
            border_width=1,
            text_color="#f4f4f5",
            placeholder_text="",
        )
        self.entry_pages.pack(side="left", fill="x", expand=True, padx=(0, 6))

        lbl_info = tk.Label(pages_input_frame, text="ⓘ", font=("Segoe UI", 11, "bold"), fg="#38bdf8", bg="#0f0f11", cursor="hand2")
        lbl_info.pack(side="right", padx=(2, 4))
        lbl_info.bind("<Button-1>", lambda e: messagebox.showinfo(
            "Hướng dẫn nhập trang in",
            "Nhập số trang hoặc khoảng trang cần in cách nhau bằng dấu phẩy.\n\nVí dụ: 1-5, 8, 11-13"
        ))

        # 3. Print on Both Sides
        opt_duplex = [
            ("Print on Both Sides", "Flip pages on long edge"),
            ("Print on Both Sides", "Flip pages on short edge"),
            ("Print One Sided", "Only print on one side of the page"),
        ]
        self.dd_duplex = OfficeSettingDropdown(self.body_frame, "duplex", opt_duplex, 0)
        self.dd_duplex.pack(fill="x", pady=3)

        # 4. Collated
        opt_collated = [
            ("Collated", "1,2,3  1,2,3  1,2,3"),
            ("Uncollated", "1,1,1  2,2,2  3,3,3"),
        ]
        self.dd_collated = OfficeSettingDropdown(self.body_frame, "collated", opt_collated, 0)
        self.dd_collated.pack(fill="x", pady=3)

        # 5. Orientation
        opt_orient = [
            ("Portrait Orientation", ""),
            ("Landscape Orientation", ""),
        ]
        self.dd_orient = OfficeSettingDropdown(self.body_frame, "portrait", opt_orient, 0)
        self.dd_orient.pack(fill="x", pady=3)

        # 6. Paper Size
        opt_size = [
            ("Letter", '8.5" x 11"'),
            ("A4", "210 x 297 mm"),
            ("A3", "297 x 420 mm"),
            ("Legal", '8.5" x 14"'),
        ]
        self.dd_size = OfficeSettingDropdown(self.body_frame, "paper_size", opt_size, 0)
        self.dd_size.pack(fill="x", pady=3)

        # 7. Margins
        opt_margins = [
            ("Normal Margins", 'Top: 1" Bottom: 1" Left: 1" Right: 1"'),
            ("Narrow Margins", 'Top: 0.5" Bottom: 0.5" Left: 0.5" Right: 0.5"'),
            ("Moderate Margins", 'Top: 1" Bottom: 1" Left: 0.75" Right: 0.75"'),
            ("Wide Margins", 'Top: 1" Bottom: 1" Left: 2" Right: 2"'),
        ]
        self.dd_margins = OfficeSettingDropdown(self.body_frame, "margins_normal", opt_margins, 0)
        self.dd_margins.pack(fill="x", pady=3)

        # 8. Page Per Sheet
        opt_pps = [
            ("1 Page Per Sheet", ""),
            ("2 Pages Per Sheet", ""),
            ("4 Pages Per Sheet", ""),
            ("6 Pages Per Sheet", ""),
        ]
        self.dd_pps = OfficeSettingDropdown(self.body_frame, "page_per_sheet", opt_pps, 0)
        self.dd_pps.pack(fill="x", pady=3)

    def _build_xlsx_settings(self):
        """Xây dựng giao diện cài đặt file XLS/XLSX theo ảnh thứ 3."""
        # 1. Print Active Sheets
        opt_sheets = [
            ("Print Active Sheets", "Only print the active sheets"),
            ("Print Entire Workbook", "Print the entire workbook"),
            ("Print Selection", "Only print the current selection"),
        ]
        self.dd_print_pages = OfficeSettingDropdown(self.body_frame, "sheets_grid", opt_sheets, 0)
        self.dd_print_pages.pack(fill="x", pady=3)

        # 2. Pages: [  ] to [  ]
        pages_input_frame = tk.Frame(self.body_frame, bg="#0f0f11", height=34)
        pages_input_frame.pack(fill="x", pady=(3, 3))
        pages_input_frame.pack_propagate(False)

        lbl_pages = tk.Label(pages_input_frame, text="Pages:", font=("Segoe UI", 10), fg="#f4f4f5", bg="#0f0f11")
        lbl_pages.pack(side="left", padx=(4, 6))

        self.spin_from = ttk.Spinbox(pages_input_frame, from_=1, to=999, width=5)
        self.spin_from.set("1")
        self.spin_from.pack(side="left", padx=(0, 4))

        lbl_to = tk.Label(pages_input_frame, text="to", font=("Segoe UI", 10), fg="#f4f4f5", bg="#0f0f11")
        lbl_to.pack(side="left", padx=(2, 4))

        self.spin_to = ttk.Spinbox(pages_input_frame, from_=1, to=999, width=5)
        self.spin_to.set("1")
        self.spin_to.pack(side="left", padx=(0, 4))

        # 3. Print One Sided
        opt_duplex = [
            ("Print One Sided", "Only print on one side of the page"),
            ("Print on Both Sides", "Flip pages on long edge"),
            ("Print on Both Sides", "Flip pages on short edge"),
        ]
        self.dd_duplex = OfficeSettingDropdown(self.body_frame, "simplex", opt_duplex, 0)
        self.dd_duplex.pack(fill="x", pady=3)

        # 4. Collated
        opt_collated = [
            ("Collated", "1,2,3  1,2,3  1,2,3"),
            ("Uncollated", "1,1,1  2,2,2  3,3,3"),
        ]
        self.dd_collated = OfficeSettingDropdown(self.body_frame, "collated", opt_collated, 0)
        self.dd_collated.pack(fill="x", pady=3)

        # 5. Orientation (Landscape)
        opt_orient = [
            ("Landscape Orientation", ""),
            ("Portrait Orientation", ""),
        ]
        self.dd_orient = OfficeSettingDropdown(self.body_frame, "landscape", opt_orient, 0)
        self.dd_orient.pack(fill="x", pady=3)

        # 6. Paper Size
        opt_size = [
            ("Letter", '8.5" x 11"'),
            ("A4", "210 x 297 mm"),
            ("A3", "297 x 420 mm"),
            ("Legal", '8.5" x 14"'),
        ]
        self.dd_size = OfficeSettingDropdown(self.body_frame, "paper_size", opt_size, 0)
        self.dd_size.pack(fill="x", pady=3)

        # 7. Margins (Narrow Margins)
        opt_margins = [
            ("Narrow Margins", 'Top: 0.75" Bottom: 0.75" Left: 0.25" Right: 0.25"'),
            ("Normal Margins", 'Top: 0.75" Bottom: 0.75" Left: 0.7" Right: 0.7"'),
            ("Wide Margins", 'Top: 1" Bottom: 1" Left: 1" Right: 1"'),
        ]
        self.dd_margins = OfficeSettingDropdown(self.body_frame, "margins_narrow", opt_margins, 0)
        self.dd_margins.pack(fill="x", pady=3)

        # 8. Scaling (No Scaling)
        opt_scaling = [
            ("No Scaling", "Print sheets at their actual size"),
            ("Fit Sheet on One Page", "Shrink the printout so that it fits on one page"),
            ("Fit All Columns on One Page", "Shrink the printout so that it is one page wide"),
            ("Fit All Rows on One Page", "Shrink the printout so that it is one page high"),
        ]
        self.dd_scaling = OfficeSettingDropdown(self.body_frame, "no_scaling", opt_scaling, 0)
        self.dd_scaling.pack(fill="x", pady=3)

    def _open_page_setup_modal(self):
        """Mở hộp thoại thiết lập trang in chi tiết (Page Setup)."""
        setup_win = ctk.CTkToplevel(self)
        setup_win.title("Page Setup")
        setup_win.geometry("380x320")
        setup_win.resizable(False, False)
        setup_win.attributes("-topmost", True)

        pad_f = ctk.CTkFrame(setup_win, fg_color="transparent")
        pad_f.pack(fill="both", expand=True, padx=20, pady=16)

        ctk.CTkLabel(pad_f, text="📐 Cấu hình trang in (Page Setup)", font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold")).pack(anchor="w", pady=(0, 10))

        # Khổ giấy
        f_size = ctk.CTkFrame(pad_f, fg_color="transparent")
        f_size.pack(fill="x", pady=4)
        ctk.CTkLabel(f_size, text="Khổ giấy:", width=90, anchor="w").pack(side="left")
        cb_size = ctk.CTkOptionMenu(f_size, values=["Letter (8.5 x 11 in)", "A4 (210 x 297 mm)", "A3", "Legal"])
        cb_size.pack(side="left", fill="x", expand=True)

        # Hướng trang
        f_ori = ctk.CTkFrame(pad_f, fg_color="transparent")
        f_ori.pack(fill="x", pady=4)
        ctk.CTkLabel(f_ori, text="Hướng trang:", width=90, anchor="w").pack(side="left")
        cb_ori = ctk.CTkOptionMenu(f_ori, values=["Portrait (Dọc)", "Landscape (Ngang)"])
        cb_ori.pack(side="left", fill="x", expand=True)

        # Căn lề
        f_mar = ctk.CTkFrame(pad_f, fg_color="transparent")
        f_mar.pack(fill="x", pady=4)
        ctk.CTkLabel(f_mar, text="Căn lề:", width=90, anchor="w").pack(side="left")
        cb_mar = ctk.CTkOptionMenu(f_mar, values=["Tiêu chuẩn (1.0 in)", "Hẹp (0.5 in)", "Rộng (2.0 in)"])
        cb_mar.pack(side="left", fill="x", expand=True)

        # Tỉ lệ
        f_scale = ctk.CTkFrame(pad_f, fg_color="transparent")
        f_scale.pack(fill="x", pady=4)
        ctk.CTkLabel(f_scale, text="Tỉ lệ in (%):", width=90, anchor="w").pack(side="left")
        en_scale = ctk.CTkEntry(f_scale)
        en_scale.insert(0, "100%")
        en_scale.pack(side="left", fill="x", expand=True)

        btn_ok = ctk.CTkButton(
            pad_f,
            text="Đồng ý",
            command=setup_win.destroy,
            fg_color="#2563eb",
            hover_color="#1d4ed8",
        )
        btn_ok.pack(pady=(16, 0))

    def _on_apply_clicked(self):
        """Lưu lại các thiết lập."""
        settings = {
            "mode": self.mode,
            "print_pages": self.dd_print_pages.get_selected(),
            "duplex": self.dd_duplex.get_selected(),
            "collated": self.dd_collated.get_selected(),
            "orientation": self.dd_orient.get_selected(),
            "paper_size": self.dd_size.get_selected(),
            "margins": self.dd_margins.get_selected(),
        }

        if self.mode == "docx":
            settings["pages_range"] = self.entry_pages.get().strip()
            settings["page_per_sheet"] = self.dd_pps.get_selected()
        else:
            settings["page_from"] = self.spin_from.get()
            settings["page_to"] = self.spin_to.get()
            settings["scaling"] = self.dd_scaling.get_selected()

        if self.on_apply:
            self.on_apply(settings, False)

        messagebox.showinfo("Thành công", "Đã cập nhật cài đặt in ấn (Print Settings) cho tài liệu!")
        self.destroy()
