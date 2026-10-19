import datetime
import os
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Optional

import customtkinter as ctk

from models import PrintItem, PrintStatus
from print_engine import EngineState, PrintEngine
from printer_service import PrinterService
from queue_manager import PrintQueueManager
from document_preview_editor import DocumentPreviewEditorView
from print_settings_dialog import OfficePrintSettingsDialog

try:
    import windnd
    DRAG_DROP_AVAILABLE = True
except ImportError:
    DRAG_DROP_AVAILABLE = False


class ModernPrintApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Khởi tạo dữ liệu cốt lõi
        self.queue_mgr = PrintQueueManager()
        self.engine = PrintEngine(self.queue_mgr)
        self.printers_data: List[Dict] = []
        self.selected_printer_name: str = ""
        self.preview_panel: Optional[DocumentPreviewEditorView] = None

        # Cấu hình cửa sổ chính
        self.title("PyPrint Queue Manager - Trình Quản Lý & In Hàng Đợi Tự Động")
        self.geometry("1150x780")
        self.minsize(980, 650)

        # Thiết lập giao diện CustomTkinter
        ctk.set_appearance_mode("Dark")  # "Dark" hoặc "Light"
        ctk.set_default_color_theme("blue")

        # Đăng ký callback từ Engine & QueueManager
        self.queue_mgr.subscribe(self._on_queue_changed)
        self.engine.on_log = self._add_log_safe
        self.engine.on_state_change = self._on_engine_state_changed
        self.engine.on_progress = self._on_progress_update
        self.engine.on_finished = self._on_print_finished

        # Biến trạng thái kéo thả bảng
        self._tree_drag_item = None

        # Khởi tạo giao diện người dùng
        self._setup_styles()
        self._build_ui()
        self._bind_shortcuts()

        # Quét danh sách máy in ban đầu
        self.refresh_printers()

        # Kích hoạt Drag & Drop từ Windows Explorer
        self._setup_drag_and_drop()

        self._add_log("Ứng dụng PyPrint Queue Manager đã khởi động thành công.", "info")

    def _setup_styles(self):
        """Định dạng giao diện Treeview cho khớp với Dark Mode/Light Mode."""
        self.style = ttk.Style()
        self.style.theme_use("clam")

        # Cấu hình Treeview Dark Style
        self.style.configure(
            "Custom.Treeview",
            background="#1e293b",
            foreground="#f1f5f9",
            fieldbackground="#1e293b",
            rowheight=32,
            font=("Segoe UI", 10),
            borderwidth=0,
            relief="flat",
        )
        self.style.configure(
            "Custom.Treeview.Heading",
            background="#0f172a",
            foreground="#38bdf8",
            font=("Segoe UI", 10, "bold"),
            borderwidth=1,
            relief="flat",
            padding=(6, 6),
        )
        self.style.map(
            "Custom.Treeview",
            background=[("selected", "#2563eb")],
            foreground=[("selected", "#ffffff")],
        )

    def _build_ui(self):
        # Frame chính chia làm 3 tầng: Header & Máy In, Hàng Đợi & Thao Tác, Điều Khiển In & Log
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)  # Vùng hàng đợi chiếm phần lớn diện tích

        # 1. KHU VỰC ĐẦU TRANG & MÁY IN
        self._build_header_and_printer_card()

        # 2. KHU VỰC HÀNG ĐỢI FILE
        self._build_queue_card()

        # 3. KHU VỰC ĐIỀU KHIỂN IN & TIẾN ĐỘ
        self._build_controls_card()

        # 4. KHU VỰC NHẬT KÝ HOẠT ĐỘNG (LOG CONSOLE)
        self._build_log_card()

    def _build_header_and_printer_card(self):
        header_card = ctk.CTkFrame(self, corner_radius=12, fg_color=("#f8fafc", "#1e293b"))
        header_card.grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 8))
        header_card.grid_columnconfigure(1, weight=1)

        # Tiêu đề ứng dụng
        title_box = ctk.CTkFrame(header_card, fg_color="transparent")
        title_box.grid(row=0, column=0, columnspan=2, sticky="ew", padx=16, pady=(10, 6))

        title_lbl = ctk.CTkLabel(
            title_box,
            text="🖨️ PyPrint Queue Manager",
            font=ctk.CTkFont(family="Segoe UI", size=20, weight="bold"),
            text_color=("#0f172a", "#38bdf8"),
        )
        title_lbl.pack(side="left")

        subtitle_lbl = ctk.CTkLabel(
            title_box,
            text="Trình Quản Lý & Tự Động Hóa In Ấn Hàng Đợi Không Giới Hạn",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=("#64748b", "#94a3b8"),
        )
        subtitle_lbl.pack(side="left", padx=(12, 0), pady=(4, 0))

        # Chế độ Sáng / Tối Switch
        self.theme_switch = ctk.CTkSwitch(
            title_box,
            text="Giao diện Tối",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            command=self._toggle_theme,
        )
        self.theme_switch.select()
        self.theme_switch.pack(side="right")

        # Đường phân cách mảnh
        sep = ctk.CTkFrame(header_card, height=1, fg_color=("#cbd5e1", "#334155"))
        sep.grid(row=1, column=0, columnspan=2, sticky="ew", padx=16, pady=4)

        # Hàng chọn máy in
        printer_bar = ctk.CTkFrame(header_card, fg_color="transparent")
        printer_bar.grid(row=2, column=0, columnspan=2, sticky="ew", padx=16, pady=(6, 12))
        printer_bar.grid_columnconfigure(1, weight=1)

        printer_lbl = ctk.CTkLabel(
            printer_bar,
            text="Máy in:",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
        )
        printer_lbl.grid(row=0, column=0, sticky="w", padx=(0, 10))

        # Dropdown máy in
        self.printer_combobox = ctk.CTkOptionMenu(
            printer_bar,
            values=["Đang quét máy in..."],
            command=self._on_printer_selected,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            height=34,
            dynamic_resizing=False,
        )
        self.printer_combobox.grid(row=0, column=1, sticky="ew", padx=(0, 10))

        # Nút Làm mới máy in
        self.btn_refresh_printers = ctk.CTkButton(
            printer_bar,
            text="🔄 Quét Lại",
            width=95,
            height=34,
            command=self.refresh_printers,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color=("#0284c7", "#0369a1"),
            hover_color=("#0369a1", "#0284c7"),
        )
        self.btn_refresh_printers.grid(row=0, column=2, padx=(0, 8))

        # Nút Đặt làm mặc định
        self.btn_set_default = ctk.CTkButton(
            printer_bar,
            text="⭐ Đặt Mặc Định",
            width=120,
            height=34,
            command=self._set_current_as_default,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color=("#475569", "#334155"),
            hover_color=("#334155", "#475569"),
        )
        self.btn_set_default.grid(row=0, column=3, padx=(0, 8))

        # Nút Printer Properties (điều chỉnh thuộc tính máy in của Windows)
        self.btn_printer_properties = ctk.CTkButton(
            printer_bar,
            text="⚙️ Printer Properties",
            width=140,
            height=34,
            command=self._open_printer_properties,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color=("#0284c7", "#0369a1"),
            hover_color=("#0369a1", "#0284c7"),
        )
        self.btn_printer_properties.grid(row=0, column=4)

        # Thông tin chi tiết máy in (Badge hiển thị trạng thái, cổng, driver)
        self.printer_details_box = ctk.CTkFrame(header_card, fg_color=("#e2e8f0", "#0f172a"), corner_radius=8)
        self.printer_details_box.grid(row=3, column=0, columnspan=2, sticky="ew", padx=16, pady=(0, 12))

        self.lbl_printer_status = ctk.CTkLabel(
            self.printer_details_box,
            text="Trạng thái: 🟢 Sẵn sàng",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color="#10b981",
        )
        self.lbl_printer_status.pack(side="left", padx=12, pady=6)

        self.lbl_printer_port = ctk.CTkLabel(
            self.printer_details_box,
            text="Cổng: N/A",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=("#475569", "#94a3b8"),
        )
        self.lbl_printer_port.pack(side="left", padx=12, pady=6)

        self.lbl_printer_driver = ctk.CTkLabel(
            self.printer_details_box,
            text="Driver: N/A",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=("#475569", "#94a3b8"),
        )
        self.lbl_printer_driver.pack(side="left", padx=12, pady=6)

        self.lbl_printer_jobs = ctk.CTkLabel(
            self.printer_details_box,
            text="Spooler: 0 lệnh",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=("#475569", "#94a3b8"),
        )
        self.lbl_printer_jobs.pack(side="left", padx=12, pady=6)

    def _build_queue_card(self):
        queue_card = ctk.CTkFrame(self, corner_radius=12, fg_color=("#f8fafc", "#1e293b"))
        queue_card.grid(row=1, column=0, sticky="nsew", padx=16, pady=4)
        queue_card.grid_columnconfigure(0, weight=1)
        queue_card.grid_rowconfigure(1, weight=1)
        self.queue_card = queue_card

        # Thanh công cụ Hàng đợi (Toolbar)
        toolbar = ctk.CTkFrame(queue_card, fg_color="transparent")
        toolbar.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 8))
        toolbar.grid_columnconfigure(1, weight=1)

        # Tiêu đề hàng đợi + thống kê số lượng
        self.lbl_queue_title = ctk.CTkLabel(
            toolbar,
            text="📋 HÀNG ĐỢI IN (0 file)",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color=("#0f172a", "#f8fafc"),
        )
        self.lbl_queue_title.grid(row=0, column=0, sticky="w", padx=(0, 16))

        # Các nút thao tác với file
        btns_frame = ctk.CTkFrame(toolbar, fg_color="transparent")
        btns_frame.grid(row=0, column=1, sticky="e")

        # Nút Print Settings nằm bên trái nút Thêm File (yêu cầu #4)
        self.btn_print_settings = ctk.CTkButton(
            btns_frame,
            text="⚙️ Print Settings",
            width=120,
            height=30,
            command=self._on_print_settings_click,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color=("#6366f1", "#4f46e5"),
            hover_color=("#4f46e5", "#4338ca"),
            state="disabled",
        )
        self.btn_print_settings.pack(side="left", padx=3)

        self.btn_add_files = ctk.CTkButton(
            btns_frame,
            text="➕ Thêm File",
            width=100,
            height=30,
            command=self._on_add_files_click,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color=("#2563eb", "#1d4ed8"),
            hover_color=("#1d4ed8", "#2563eb"),
        )
        self.btn_add_files.pack(side="left", padx=3)

        self.btn_add_folder = ctk.CTkButton(
            btns_frame,
            text="📁 Thư Mục",
            width=90,
            height=30,
            command=self._on_add_folder_click,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color=("#0284c7", "#0369a1"),
            hover_color=("#0369a1", "#0284c7"),
        )
        self.btn_add_folder.pack(side="left", padx=3)

        # Nhóm nút di chuyển thứ tự (Thay đổi vị trí trong queue)
        self.btn_move_up = ctk.CTkButton(
            btns_frame,
            text="⬆️ Lên",
            width=65,
            height=30,
            command=self._on_move_up_click,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color=("#475569", "#334155"),
            hover_color=("#334155", "#475569"),
        )
        self.btn_move_up.pack(side="left", padx=3)

        self.btn_move_down = ctk.CTkButton(
            btns_frame,
            text="⬇️ Xuống",
            width=70,
            height=30,
            command=self._on_move_down_click,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color=("#475569", "#334155"),
            hover_color=("#334155", "#475569"),
        )
        self.btn_move_down.pack(side="left", padx=3)

        self.btn_move_top = ctk.CTkButton(
            btns_frame,
            text="🔝 Đầu",
            width=60,
            height=30,
            command=self._on_move_top_click,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=("#475569", "#334155"),
            hover_color=("#334155", "#475569"),
        )
        self.btn_move_top.pack(side="left", padx=3)

        self.btn_move_bottom = ctk.CTkButton(
            btns_frame,
            text="🔚 Cuối",
            width=60,
            height=30,
            command=self._on_move_bottom_click,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=("#475569", "#334155"),
            hover_color=("#334155", "#475569"),
        )
        self.btn_move_bottom.pack(side="left", padx=3)

        self.btn_remove = ctk.CTkButton(
            btns_frame,
            text="❌ Xóa Chọn",
            width=90,
            height=30,
            command=self._on_remove_selected_click,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=("#b91c1c", "#991b1b"),
            hover_color=("#991b1b", "#b91c1c"),
        )
        self.btn_remove.pack(side="left", padx=3)

        self.btn_clear_all = ctk.CTkButton(
            btns_frame,
            text="🧹 Xóa Hết",
            width=80,
            height=30,
            command=self._on_clear_all_click,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=("#64748b", "#475569"),
            hover_color=("#475569", "#64748b"),
        )
        self.btn_clear_all.pack(side="left", padx=3)

        self.btn_reset_status = ctk.CTkButton(
            btns_frame,
            text="🔄 Đặt Lại",
            width=80,
            height=30,
            command=self._on_reset_status_click,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=("#0891b2", "#0e7490"),
            hover_color=("#0e7490", "#0891b2"),
        )
        self.btn_reset_status.pack(side="left", padx=3)

        # Bảng hiển thị hàng đợi (Treeview) - Được lưu vào self.table_container để có thể thu gọn khi mở Xem Trước
        self.table_container = ctk.CTkFrame(queue_card, fg_color="transparent")
        self.table_container.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 8))
        self.table_container.grid_columnconfigure(0, weight=1)
        self.table_container.grid_rowconfigure(0, weight=1)

        # Thay cột dung lượng bằng cột "preview" (Xem trước) theo yêu cầu #2
        columns = ("stt", "name", "type", "preview", "copies", "status", "detail", "path")
        self.tree = ttk.Treeview(
            self.table_container,
            columns=columns,
            show="headings",
            selectmode="extended",
            style="Custom.Treeview",
        )

        # Định dạng cột
        self.tree.heading("stt", text="STT")
        self.tree.heading("name", text="Tên File")
        self.tree.heading("type", text="Loại")
        self.tree.heading("preview", text="Xem trước")
        self.tree.heading("copies", text="Bản In")
        self.tree.heading("status", text="Trạng Thái")
        self.tree.heading("detail", text="Chi Tiết Xử Lý")
        self.tree.heading("path", text="Đường Dẫn")

        self.tree.column("stt", width=45, anchor="center")
        self.tree.column("name", width=230, anchor="w")
        self.tree.column("type", width=65, anchor="center")
        self.tree.column("preview", width=95, anchor="center")
        self.tree.column("copies", width=70, anchor="center")
        self.tree.column("status", width=120, anchor="w")
        self.tree.column("detail", width=170, anchor="w")
        self.tree.column("path", width=300, anchor="w")

        # Scrollbars
        vsb = ttk.Scrollbar(self.table_container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self.table_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        # Thao tác chuột trên bảng (Kéo thả reorder hàng đợi + Click đúp xem file + Menu ngữ cảnh)
        self.tree.bind("<ButtonPress-1>", self._on_tree_button_press)
        self.tree.bind("<ButtonRelease-1>", self._on_tree_button_release)
        self.tree.bind("<Double-1>", self._on_tree_double_click)
        self.tree.bind("<Button-3>", self._on_tree_right_click)

        # Gợi ý kéo thả file ở chân bảng
        drag_hint = ctk.CTkLabel(
            queue_card,
            text="💡 Gợi ý: Bạn có thể kéo & thả file trực tiếp vào đây | Click vào cột 'Xem trước' để xem & sửa tài liệu trực tiếp",
            font=ctk.CTkFont(family="Segoe UI", size=11, slant="italic"),
            text_color=("#64748b", "#94a3b8"),
        )
        drag_hint.grid(row=2, column=0, sticky="w", padx=16, pady=(0, 8))

        # Menu chuột phải
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="🔍 Xem trước & Chỉnh sửa", command=self._open_preview_selected)
        self.context_menu.add_command(label="🔢 Điều chỉnh số bản in...", command=self._edit_copies_selected)
        self.context_menu.add_command(label="⚙️ Print Settings...", command=self._on_print_settings_click)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="👁️ Mở bằng phần mềm mặc định", command=self._on_tree_double_click)
        self.context_menu.add_command(label="📂 Mở thư mục chứa file", command=self._open_file_directory)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="⬆️ Di chuyển lên trên", command=self._on_move_up_click)
        self.context_menu.add_command(label="⬇️ Di chuyển xuống dưới", command=self._on_move_down_click)
        self.context_menu.add_command(label="🔝 Đưa lên đầu hàng đợi", command=self._on_move_top_click)
        self.context_menu.add_command(label="🔚 Đưa xuống cuối hàng đợi", command=self._on_move_bottom_click)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="❌ Xóa khỏi hàng đợi", command=self._on_remove_selected_click)

    def _build_controls_card(self):
        ctrl_card = ctk.CTkFrame(self, corner_radius=12, fg_color=("#f8fafc", "#1e293b"))
        ctrl_card.grid(row=2, column=0, sticky="ew", padx=16, pady=4)
        ctrl_card.grid_columnconfigure(0, weight=1)

        # Hàng 1: Tiến độ in ấn
        prog_frame = ctk.CTkFrame(ctrl_card, fg_color="transparent")
        prog_frame.grid(row=0, column=0, sticky="ew", padx=16, pady=(10, 4))
        prog_frame.grid_columnconfigure(0, weight=1)

        self.lbl_progress_status = ctk.CTkLabel(
            prog_frame,
            text="Trạng thái: Sẵn sàng bắt đầu",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=("#0f172a", "#f8fafc"),
        )
        self.lbl_progress_status.grid(row=0, column=0, sticky="w")

        self.lbl_progress_percent = ctk.CTkLabel(
            prog_frame,
            text="0%",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=("#2563eb", "#38bdf8"),
        )
        self.lbl_progress_percent.grid(row=0, column=1, sticky="e")

        # Thanh tiến độ
        self.progress_bar = ctk.CTkProgressBar(ctrl_card, height=14, corner_radius=7)
        self.progress_bar.grid(row=1, column=0, sticky="ew", padx=16, pady=(2, 8))
        self.progress_bar.set(0)

        # Hàng 2: Tùy chỉnh & Các nút bấm điều khiển in lớn
        action_bar = ctk.CTkFrame(ctrl_card, fg_color="transparent")
        action_bar.grid(row=2, column=0, sticky="ew", padx=16, pady=(4, 12))
        action_bar.grid_columnconfigure(0, weight=1)

        # Khối bên trái: Cài đặt độ trễ và số bản in
        settings_box = ctk.CTkFrame(action_bar, fg_color="transparent")
        settings_box.grid(row=0, column=0, sticky="w")

        delay_lbl = ctk.CTkLabel(
            settings_box,
            text="Độ trễ giữa các file:",
            font=ctk.CTkFont(family="Segoe UI", size=12),
        )
        delay_lbl.pack(side="left", padx=(0, 6))

        self.delay_slider = ctk.CTkSlider(
            settings_box,
            from_=0.5,
            to=10.0,
            number_of_steps=19,
            width=130,
            command=self._on_delay_slider_change,
        )
        self.delay_slider.set(2.0)
        self.delay_slider.pack(side="left", padx=(0, 6))

        self.lbl_delay_val = ctk.CTkLabel(
            settings_box,
            text="2.0s",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            width=40,
        )
        self.lbl_delay_val.pack(side="left", padx=(0, 16))

        # Checkbox chỉ in file chưa in
        self.chk_pending_only = ctk.CTkCheckBox(
            settings_box,
            text="Bỏ qua file đã hoàn thành",
            font=ctk.CTkFont(family="Segoe UI", size=12),
        )
        self.chk_pending_only.select()
        self.chk_pending_only.pack(side="left", padx=(0, 10))

        # Khối bên phải: Nút BẮT ĐẦU IN, TẠM DỪNG, HỦY
        btn_action_box = ctk.CTkFrame(action_bar, fg_color="transparent")
        btn_action_box.grid(row=0, column=1, sticky="e")

        self.btn_start_print = ctk.CTkButton(
            btn_action_box,
            text="▶ BẮT ĐẦU IN (PRINT)",
            width=190,
            height=40,
            command=self._on_start_print_click,
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            fg_color=("#16a34a", "#15803d"),
            hover_color=("#15803d", "#16a34a"),
        )
        self.btn_start_print.pack(side="left", padx=4)

        self.btn_pause_print = ctk.CTkButton(
            btn_action_box,
            text="⏸ Tạm Dừng",
            width=110,
            height=40,
            command=self._on_pause_print_click,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color=("#ea580c", "#c2410c"),
            hover_color=("#c2410c", "#ea580c"),
            state="disabled",
        )
        self.btn_pause_print.pack(side="left", padx=4)

        self.btn_stop_print = ctk.CTkButton(
            btn_action_box,
            text="⏹ Hủy In",
            width=100,
            height=40,
            command=self._on_stop_print_click,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color=("#dc2626", "#b91c1c"),
            hover_color=("#b91c1c", "#dc2626"),
            state="disabled",
        )
        self.btn_stop_print.pack(side="left", padx=4)

    def _build_log_card(self):
        log_card = ctk.CTkFrame(self, corner_radius=12, fg_color=("#f8fafc", "#1e293b"), height=130)
        log_card.grid(row=3, column=0, sticky="ew", padx=16, pady=(4, 12))
        log_card.grid_columnconfigure(0, weight=1)

        # Thanh tiêu đề log
        log_top = ctk.CTkFrame(log_card, fg_color="transparent")
        log_top.grid(row=0, column=0, sticky="ew", padx=12, pady=(6, 2))
        log_top.grid_columnconfigure(0, weight=1)

        lbl_log_title = ctk.CTkLabel(
            log_top,
            text="📝 NHẬT KÝ HOẠT ĐỘNG THỜI GIAN THỰC",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=("#64748b", "#94a3b8"),
        )
        lbl_log_title.grid(row=0, column=0, sticky="w")

        btn_clear_log = ctk.CTkButton(
            log_top,
            text="Xóa Log",
            width=65,
            height=22,
            font=ctk.CTkFont(family="Segoe UI", size=10),
            command=self._clear_log,
            fg_color=("#64748b", "#334155"),
            hover_color=("#475569", "#64748b"),
        )
        btn_clear_log.grid(row=0, column=1, sticky="e")

        # Khung văn bản Log Console
        self.txt_log = ctk.CTkTextbox(
            log_card,
            height=85,
            font=ctk.CTkFont(family="Consolas", size=11),
            fg_color=("#f1f5f9", "#090d16"),
            text_color=("#0f172a", "#f8fafc"),
            wrap="word",
        )
        self.txt_log.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 8))

    def _setup_drag_and_drop(self):
        """Kích hoạt Drag & Drop file từ Windows Explorer."""
        if DRAG_DROP_AVAILABLE:
            try:
                windnd.hook_dropfiles(self, func=self._on_window_drop_files, force_unicode=True)
                self._add_log("Đã bật tính năng Kéo & Thả file từ Windows Explorer (windnd).", "info")
            except Exception as e:
                self._add_log(f"Không thể kích hoạt drag-and-drop: {e}", "warning")
        else:
            self._add_log("Thư viện windnd chưa sẵn sàng. Bạn có thể dùng nút 'Thêm File'.", "info")

    def _on_window_drop_files(self, files):
        """Xử lý danh sách file được thả từ Windows Explorer."""
        if self.engine.state == EngineState.RUNNING:
            messagebox.showwarning("Cảnh báo", "Đang trong tiến trình in ấn, vui lòng đợi hoàn thành hoặc hủy trước khi thêm file!")
            return

        added_paths = []
        for item in files:
            # Giải mã nếu là bytes
            if isinstance(item, bytes):
                try:
                    path_str = item.decode("utf-8")
                except UnicodeDecodeError:
                    path_str = item.decode("mbcs", errors="replace")
            else:
                path_str = str(item)

            if os.path.isfile(path_str):
                added_paths.append(path_str)
            elif os.path.isdir(path_str):
                # Nếu người dùng thả cả thư mục, duyệt toàn bộ file trong đó
                for root, _, fs in os.walk(path_str):
                    for f in fs:
                        added_paths.append(os.path.join(root, f))

        if added_paths:
            self.queue_mgr.add_files(added_paths)
            self._add_log(f"Đã thêm {len(added_paths)} file từ thao tác Kéo & Thả.", "info")

    def _bind_shortcuts(self):
        """Gán phím tắt tiện lợi."""
        self.bind("<Delete>", lambda e: self._on_remove_selected_click())
        self.bind("<Alt-Up>", lambda e: self._on_move_up_click())
        self.bind("<Alt-Down>", lambda e: self._on_move_down_click())
        self.bind("<Control-o>", lambda e: self._on_add_files_click())
        self.bind("<Control-O>", lambda e: self._on_add_files_click())
        self.bind("<Control-p>", lambda e: self._on_start_print_click())
        self.bind("<Control-P>", lambda e: self._on_start_print_click())

    # ================== MÁY IN LOGIC ==================

    def refresh_printers(self):
        """Quét lại toàn bộ máy in và cập nhật dropdown."""
        self._add_log("Đang quét danh sách máy in kết nối với thiết bị...", "info")
        self.printers_data = PrinterService.get_printers()

        if not self.printers_data:
            self.printer_combobox.configure(values=["Không tìm thấy máy in nào"])
            self.printer_combobox.set("Không tìm thấy máy in nào")
            self.selected_printer_name = ""
            self.lbl_printer_status.configure(text="Trạng thái: 🔴 Không có máy in", text_color="#ef4444")
            self.lbl_printer_port.configure(text="Cổng: N/A")
            self.lbl_printer_driver.configure(text="Driver: N/A")
            self.lbl_printer_jobs.configure(text="Spooler: 0")
            self._add_log("Không tìm thấy máy in nào kết nối với thiết bị!", "warning")
            return

        printer_names = []
        default_index = 0
        for idx, p in enumerate(self.printers_data):
            name_display = p["name"]
            if p["is_default"]:
                name_display += " ⭐ [Mặc định]"
                default_index = idx
            printer_names.append(name_display)

        self.printer_combobox.configure(values=printer_names)

        # Chọn máy in mặc định hoặc máy in đầu tiên
        self.printer_combobox.set(printer_names[default_index])
        self._on_printer_selected(printer_names[default_index])
        self._add_log(f"Đã kết nối thành công với {len(self.printers_data)} máy in.", "success")

    def _on_printer_selected(self, choice_text: str):
        # Trích xuất tên máy in thực tế (bỏ phần ⭐ [Mặc định])
        real_name = choice_text.replace(" ⭐ [Mặc định]", "").strip()
        self.selected_printer_name = real_name

        # Tìm thông tin chi tiết
        info = next((p for p in self.printers_data if p["name"] == real_name), None)
        if info:
            status_txt = info["status_text"]
            if "Sẵn sàng" in status_txt:
                badge = f"Trạng thái: 🟢 {status_txt}"
                color = "#10b981"
            elif "Tạm dừng" in status_txt or "Bận" in status_txt:
                badge = f"Trạng thái: 🟡 {status_txt}"
                color = "#f59e0b"
            else:
                badge = f"Trạng thái: 🔴 {status_txt}"
                color = "#ef4444"

            self.lbl_printer_status.configure(text=badge, text_color=color)
            self.lbl_printer_port.configure(text=f"Cổng: {info['port']}")
            self.lbl_printer_driver.configure(text=f"Driver: {info['driver']}")
            self.lbl_printer_jobs.configure(text=f"Spooler: {info['jobs_count']} lệnh")

    def _set_current_as_default(self):
        if not self.selected_printer_name:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn một máy in trước!")
            return
        ok = PrinterService.set_default_printer(self.selected_printer_name)
        if ok:
            self._add_log(f"Đã đặt '{self.selected_printer_name}' làm máy in mặc định của hệ thống.", "success")
            self.refresh_printers()
        else:
            messagebox.showerror("Lỗi", f"Không thể đặt '{self.selected_printer_name}' làm mặc định.")

    def _open_printer_properties(self):
        """Mở thuộc tính máy in (Printer Properties) của Windows cho máy in đang chọn."""
        if not self.selected_printer_name:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn máy in trước khi mở Printer Properties!")
            return

        self._add_log(f"Đang mở cửa sổ Printer Properties cho '{self.selected_printer_name}'...", "info")
        ok, msg = PrinterService.open_printer_properties(self.selected_printer_name)
        if ok:
            self._add_log(msg, "success")
        else:
            messagebox.showerror("Lỗi", msg)
            self._add_log(msg, "error")

    def _print_test_page(self):
        if not self.selected_printer_name:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn máy in trước khi in thử!")
            return

        answer = messagebox.askyesno(
            "Xác nhận in thử",
            f"Bạn có muốn in 1 trang kiểm tra (Test Page) ra máy in:\n\n'{self.selected_printer_name}' không?"
        )
        if not answer:
            return

        self._add_log(f"Đang gửi trang in kiểm tra tới '{self.selected_printer_name}'...", "info")
        ok, msg = PrinterService.print_test_page(self.selected_printer_name)
        if ok:
            messagebox.showinfo("Thành công", msg)
            self._add_log(msg, "success")
        else:
            messagebox.showerror("Lỗi", msg)
            self._add_log(msg, "error")

    # ================== HÀNG ĐỢI LOGIC & SẮP XẾP ==================

    def _on_add_files_click(self):
        if self.engine.state == EngineState.RUNNING:
            messagebox.showwarning("Cảnh báo", "Không thể thêm file khi đang in!")
            return

        filetypes = [
            ("Tất cả định dạng hỗ trợ", "*.pdf;*.docx;*.doc;*.xlsx;*.xls;*.pptx;*.ppt;*.txt;*.png;*.jpg;*.jpeg;*.bmp"),
            ("Tài liệu PDF (*.pdf)", "*.pdf"),
            ("Tài liệu Word (*.docx;*.doc)", "*.docx;*.doc"),
            ("Bảng tính Excel (*.xlsx;*.xls)", "*.xlsx;*.xls"),
            ("Tệp văn bản (*.txt;*.log;*.csv)", "*.txt;*.log;*.csv"),
            ("Tệp hình ảnh (*.png;*.jpg;*.jpeg;*.bmp)", "*.png;*.jpg;*.jpeg;*.bmp"),
            ("Tất cả các tệp (*.*)", "*.*"),
        ]
        files = filedialog.askopenfilenames(
            title="Chọn các file muốn thêm vào hàng đợi in",
            filetypes=filetypes
        )
        if files:
            self.queue_mgr.add_files(list(files))
            self._add_log(f"Đã thêm {len(files)} file vào hàng đợi.", "info")

    def _on_add_folder_click(self):
        if self.engine.state == EngineState.RUNNING:
            messagebox.showwarning("Cảnh báo", "Không thể thêm file khi đang in!")
            return

        folder = filedialog.askdirectory(title="Chọn thư mục chứa tài liệu in")
        if folder:
            added = self.queue_mgr.add_folder(folder, recursive=False)
            self._add_log(f"Đã thêm {len(added)} file từ thư mục '{os.path.basename(folder)}'.", "info")

    def _get_selected_tree_indices(self) -> List[int]:
        selected = self.tree.selection()
        if not selected:
            return []
        all_children = list(self.tree.get_children())
        indices = [all_children.index(item) for item in selected if item in all_children]
        return sorted(indices)

    def _on_move_up_click(self):
        """Di chuyển file lên trên 1 vị trí."""
        if self.engine.state == EngineState.RUNNING:
            messagebox.showwarning("Cảnh báo", "Không thể thay đổi vị trí khi đang in!")
            return

        indices = self._get_selected_tree_indices()
        if not indices or indices[0] == 0:
            return

        idx = indices[0]
        new_idx = self.queue_mgr.move_up(idx)
        self._select_tree_row_by_index(new_idx)

    def _on_move_down_click(self):
        """Di chuyển file xuống dưới 1 vị trí."""
        if self.engine.state == EngineState.RUNNING:
            messagebox.showwarning("Cảnh báo", "Không thể thay đổi vị trí khi đang in!")
            return

        indices = self._get_selected_tree_indices()
        if not indices:
            return

        idx = indices[-1]
        items_count = len(self.queue_mgr.get_items())
        if idx >= items_count - 1:
            return

        new_idx = self.queue_mgr.move_down(idx)
        self._select_tree_row_by_index(new_idx)

    def _on_move_top_click(self):
        """Đưa file lên đầu hàng đợi."""
        if self.engine.state == EngineState.RUNNING:
            return
        indices = self._get_selected_tree_indices()
        if not indices or indices[0] == 0:
            return
        new_idx = self.queue_mgr.move_to_top(indices[0])
        self._select_tree_row_by_index(new_idx)

    def _on_move_bottom_click(self):
        """Đưa file xuống cuối hàng đợi."""
        if self.engine.state == EngineState.RUNNING:
            return
        indices = self._get_selected_tree_indices()
        if not indices:
            return
        new_idx = self.queue_mgr.move_to_bottom(indices[-1])
        self._select_tree_row_by_index(new_idx)

    def _on_remove_selected_click(self):
        """Xóa các file được chọn khỏi hàng đợi."""
        if self.engine.state == EngineState.RUNNING:
            messagebox.showwarning("Cảnh báo", "Không thể xóa file khi đang in!")
            return

        indices = self._get_selected_tree_indices()
        if not indices:
            return

        count = self.queue_mgr.remove_items_by_indices(indices)
        self._add_log(f"Đã xóa {count} file khỏi hàng đợi.", "info")

    def _on_clear_all_click(self):
        """Xóa toàn bộ hàng đợi."""
        if self.engine.state == EngineState.RUNNING:
            messagebox.showwarning("Cảnh báo", "Không thể xóa toàn bộ khi đang in!")
            return

        if not self.queue_mgr.get_items():
            return

        if messagebox.askyesno("Xác nhận", "Bạn có chắc chắn muốn xóa toàn bộ file khỏi hàng đợi in không?"):
            self.queue_mgr.clear_all()
            self._add_log("Đã xóa sạch hàng đợi in.", "info")

    def _on_reset_status_click(self):
        """Đặt lại trạng thái của tất cả các file về 'Chờ in' để in lại."""
        if self.engine.state == EngineState.RUNNING:
            return
        self.queue_mgr.reset_statuses()
        self.progress_bar.set(0)
        self.lbl_progress_percent.configure(text="0%")
        self.lbl_progress_status.configure(text="Trạng thái: Đã đặt lại toàn bộ về Chờ in")
        self._add_log("Đã đặt lại trạng thái hàng đợi: Sẵn sàng in lại.", "info")

    def _on_queue_changed(self):
        """Được gọi mỗi khi hàng đợi có sự thay đổi (Thêm, Xóa, Đổi vị trí, Đổi trạng thái)."""
        self.after(0, self._render_queue)

    def _render_queue(self):
        """Cập nhật lại toàn bộ Treeview hiển thị hàng đợi."""
        selected_indices = self._get_selected_tree_indices()

        # Xóa các dòng cũ
        for item in self.tree.get_children():
            self.tree.delete(item)

        items = self.queue_mgr.get_items()
        for idx, it in enumerate(items, start=1):
            values = (
                idx,
                it.filename,
                it.file_type.upper(),
                "👁️ Xem & Sửa",
                f"{it.copies} bản",
                it.get_status_badge(),
                it.status_detail,
                it.filepath,
            )
            # Tag định màu trạng thái
            tag = "pending"
            if it.status == PrintStatus.PRINTING:
                tag = "printing"
            elif it.status == PrintStatus.COMPLETED:
                tag = "completed"
            elif it.status == PrintStatus.FAILED:
                tag = "failed"

            self.tree.insert("", "end", iid=str(idx - 1), values=values, tags=(tag,))

        # Màu sắc từng trạng thái
        self.tree.tag_configure("printing", background="#1e3a8a", foreground="#93c5fd")
        self.tree.tag_configure("completed", background="#064e3b", foreground="#6ee7b7")
        self.tree.tag_configure("failed", background="#7f1d1d", foreground="#fca5a5")

        # Cập nhật trạng thái nút Print Settings
        if self.engine.state == EngineState.IDLE:
            self.btn_print_settings.configure(state="normal" if bool(items) else "disabled")

        # Cập nhật thống kê trên thanh tiêu đề
        stats = self.queue_mgr.get_stats()
        self.lbl_queue_title.configure(
            text=f"📋 HÀNG ĐỢI IN ({stats['total']} file - {stats['total_size_str']}) | ⏳ {stats['pending']} | 🖨️ {stats['printing']} | ✅ {stats['completed']} | ❌ {stats['failed']}"
        )

        # Khôi phục vị trí chọn nếu còn hợp lệ
        all_children = self.tree.get_children()
        for s_idx in selected_indices:
            if s_idx < len(all_children):
                self.tree.selection_add(all_children[s_idx])

    def _select_tree_row_by_index(self, index: int):
        all_children = self.tree.get_children()
        if 0 <= index < len(all_children):
            target = all_children[index]
            self.tree.selection_set(target)
            self.tree.see(target)

    # ================== KÉO THẢ TRÊN TREEVIEW ==================

    def _on_tree_button_press(self, event):
        item = self.tree.identify_row(event.y)
        self._tree_drag_item = item

    def _on_tree_button_release(self, event):
        if not self._tree_drag_item:
            return
        if self.engine.state == EngineState.RUNNING:
            self._tree_drag_item = None
            return

        target = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)

        if target and target != self._tree_drag_item:
            all_children = list(self.tree.get_children())
            from_idx = all_children.index(self._tree_drag_item)
            to_idx = all_children.index(target)

            self.queue_mgr.move_item(from_idx, to_idx)
            self._select_tree_row_by_index(to_idx)
            self._add_log(f"Đã thay đổi vị trí file từ #{from_idx + 1} sang #{to_idx + 1}.", "info")
        else:
            # Click trên cùng 1 dòng: Kiểm tra xem người dùng click cột nào
            if target:
                if col == "#4":  # Cột "Xem trước" -> Mở vùng xem trước & chỉnh sửa trực tiếp
                    self._open_preview_for_row(target)
                elif col == "#5":  # Cột "Bản in" -> Mở công cụ điều chỉnh số lượng bản in
                    self._show_copies_editor(target, event.x_root, event.y_root)

        self._tree_drag_item = None

    def _on_tree_double_click(self, event=None):
        """
        Click đúp vào tên file để mở file trực tiếp bằng ứng dụng mặc định của hệ thống.
        Nếu click đúp vào cột Xem trước hoặc Bản in thì gọi tương ứng chức năng đó.
        """
        if event:
            col = self.tree.identify_column(event.x)
            row = self.tree.identify_row(event.y)
            if col == "#4" and row:
                self._open_preview_for_row(row)
                return
            elif col == "#5" and row:
                self._show_copies_editor(row, event.x_root, event.y_root)
                return

        indices = self._get_selected_tree_indices()
        if not indices:
            return
        item = self.queue_mgr.get_item(indices[0])
        if item and os.path.exists(item.filepath):
            try:
                os.startfile(item.filepath)
                self._add_log(f"Đang mở xem file bằng ứng dụng mặc định: {item.filename}", "info")
            except Exception as e:
                messagebox.showerror("Lỗi mở file", str(e))

    # ================== XEM TRƯỚC & CHỈNH SỬA TÀI LIỆU TRỰC TIẾP ==================

    def _open_preview_selected(self):
        indices = self._get_selected_tree_indices()
        if indices:
            self._open_preview_for_row(str(indices[0]))

    def _open_preview_for_row(self, row_id: str):
        try:
            idx = int(row_id)
        except Exception:
            return
        item = self.queue_mgr.get_item(idx)
        if not item or not os.path.exists(item.filepath):
            messagebox.showwarning("Cảnh báo", "File không tồn tại trên ổ đĩa!")
            return

        self._open_document_preview(idx, item)

    def _open_document_preview(self, index: int, item: PrintItem):
        if self.engine.state == EngineState.RUNNING:
            messagebox.showwarning("Cảnh báo", "Không thể chỉnh sửa file khi đang trong tiến trình in!")
            return

        if self.preview_panel:
            try:
                self.preview_panel.destroy()
            except Exception:
                pass
            self.preview_panel = None

        # Thu gọn hàng đợi in vào (ẩn table_container) theo yêu cầu #2
        self.table_container.grid_forget()

        # Mở vùng xem trước và chỉnh sửa
        self.preview_panel = DocumentPreviewEditorView(
            master=self.queue_card,
            filepath=item.filepath,
            filename=item.filename,
            on_save_callback=lambda idx=index: self._on_preview_saved(idx),
            on_close_callback=self._on_preview_closed,
        )
        self.preview_panel.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 8))
        self._add_log(f"Đã mở vùng xem trước & chỉnh sửa cho file: {item.filename}", "info")

    def _on_preview_saved(self, index: int):
        self.queue_mgr.refresh_item(index)
        self._on_preview_closed()
        self._add_log("Đã lưu và đóng vùng xem trước. Hàng đợi in trở lại bình thường.", "success")

    def _on_preview_closed(self):
        if self.preview_panel:
            try:
                self.preview_panel.destroy()
            except Exception:
                pass
            self.preview_panel = None

        # Đưa hàng đợi in trở lại bình thường
        self.table_container.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 8))

    # ================== ĐIỀU CHỈNH SỐ BẢN IN (COPIES) ==================

    def _edit_copies_selected(self):
        indices = self._get_selected_tree_indices()
        if not indices:
            return
        x = self.winfo_rootx() + self.winfo_width() // 2
        y = self.winfo_rooty() + self.winfo_height() // 2
        self._show_copies_editor(str(indices[0]), x, y)

    def _show_copies_editor(self, row_id: str, x_root: int, y_root: int):
        """Mở popup điều chỉnh số lượng bản in (Bản In) trực tiếp."""
        try:
            idx = int(row_id)
        except Exception:
            return
        item = self.queue_mgr.get_item(idx)
        if not item:
            return

        pop = ctk.CTkToplevel(self)
        pop.title("Bản in")
        pop.geometry("240x120")
        pop.resizable(False, False)
        pop.attributes("-topmost", True)
        pop.geometry(f"+{max(10, x_root - 120)}+{max(10, y_root - 60)}")

        f = ctk.CTkFrame(pop, fg_color="transparent")
        f.pack(fill="both", expand=True, padx=14, pady=12)

        short_name = (item.filename[:18] + "..") if len(item.filename) > 20 else item.filename
        ctk.CTkLabel(f, text=f"Số bản in: {short_name}", font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")).pack(anchor="w", pady=(0, 8))

        row_f = ctk.CTkFrame(f, fg_color="transparent")
        row_f.pack(fill="x")

        val_var = tk.IntVar(value=item.copies)

        def _dec():
            if val_var.get() > 1:
                val_var.set(val_var.get() - 1)

        def _inc():
            if val_var.get() < 999:
                val_var.set(val_var.get() + 1)

        btn_minus = ctk.CTkButton(row_f, text="➖", width=32, height=30, command=_dec, fg_color=("#64748b", "#334155"))
        btn_minus.pack(side="left", padx=2)

        en = ctk.CTkEntry(row_f, textvariable=val_var, width=54, height=30, justify="center", font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"))
        en.pack(side="left", padx=4)

        btn_plus = ctk.CTkButton(row_f, text="➕", width=32, height=30, command=_inc, fg_color=("#64748b", "#334155"))
        btn_plus.pack(side="left", padx=2)

        def _confirm():
            try:
                new_val = max(1, int(val_var.get()))
                self.queue_mgr.set_item_copies(idx, new_val)
                self._add_log(f"Đã cập nhật số bản in của '{item.filename}' thành {new_val} bản.", "info")
            except Exception:
                pass
            pop.destroy()

        btn_ok = ctk.CTkButton(row_f, text="OK", width=46, height=30, command=_confirm, fg_color=("#16a34a", "#15803d"), hover_color=("#15803d", "#16a34a"))
        btn_ok.pack(side="left", padx=(6, 0))

        en.bind("<Return>", lambda e: _confirm())
        en.focus_set()

    # ================== PRINT SETTINGS (WORD & EXCEL STYLED) ==================

    def _on_print_settings_click(self):
        """Mở cửa sổ Print Settings theo định dạng tài liệu được chọn."""
        items = self.queue_mgr.get_items()
        if not items:
            messagebox.showinfo("Thông báo", "Vui lòng thêm ít nhất một file vào hàng đợi trước khi mở Print Settings!")
            return

        indices = self._get_selected_tree_indices()
        target_item = items[indices[0]] if indices and indices[0] < len(items) else items[0]

        dialog = OfficePrintSettingsDialog(
            parent=self,
            item=target_item,
            all_items=items,
            on_apply=lambda settings, apply_all: self._on_print_settings_applied(target_item, settings, apply_all),
        )
        self._add_log(f"Đã mở cửa sổ Print Settings cho file: {target_item.filename}", "info")

    def _on_print_settings_applied(self, item: PrintItem, settings: dict, apply_all: bool):
        items = self.queue_mgr.get_items()
        if apply_all:
            for idx, it in enumerate(items):
                ext = os.path.splitext(it.filepath)[1].lower()
                target_ext = os.path.splitext(item.filepath)[1].lower()
                if (ext in [".doc", ".docx"] and target_ext in [".doc", ".docx"]) or \
                   (ext in [".xls", ".xlsx"] and target_ext in [".xls", ".xlsx"]):
                    self.queue_mgr.update_item_settings(idx, settings)
            self._add_log("Đã áp dụng Print Settings cho tất cả file cùng định dạng.", "success")
        else:
            try:
                idx = items.index(item)
                self.queue_mgr.update_item_settings(idx, settings)
                self._add_log(f"Đã lưu Print Settings cho file: {item.filename}", "success")
            except ValueError:
                pass

    def _open_file_directory(self):
        """Mở thư mục chứa file trong Windows Explorer."""
        indices = self._get_selected_tree_indices()
        if not indices:
            return
        item = self.queue_mgr.get_item(indices[0])
        if item and os.path.exists(item.filepath):
            folder = os.path.dirname(os.path.abspath(item.filepath))
            subprocess.run(["explorer", folder])

    def _on_tree_right_click(self, event):
        row = self.tree.identify_row(event.y)
        if row:
            if row not in self.tree.selection():
                self.tree.selection_set(row)
            self.context_menu.post(event.x_root, event.y_root)

    # ================== ĐIỀU KHIỂN TIẾN TRÌNH IN ==================

    def _on_delay_slider_change(self, val):
        self.lbl_delay_val.configure(text=f"{val:.1f}s")

    def _on_start_print_click(self):
        if not self.selected_printer_name:
            messagebox.showwarning("Chưa chọn máy in", "Vui lòng chọn máy in trước khi bắt đầu!")
            return

        items = self.queue_mgr.get_items()
        if not items:
            messagebox.showwarning("Hàng đợi trống", "Vui lòng thêm ít nhất một file vào hàng đợi trước khi in!")
            return

        pending_only = bool(self.chk_pending_only.get())
        delay = float(self.delay_slider.get())

        self.engine.start(
            printer_name=self.selected_printer_name,
            delay_seconds=delay,
            print_all_or_pending_only=pending_only,
        )

    def _on_pause_print_click(self):
        if self.engine.state == EngineState.RUNNING:
            self.engine.pause()
        elif self.engine.state == EngineState.PAUSED:
            self.engine.resume()

    def _on_stop_print_click(self):
        if self.engine.state in [EngineState.RUNNING, EngineState.PAUSED]:
            if messagebox.askyesno("Hủy in", "Bạn có chắc chắn muốn dừng tiến trình in hàng đợi này không?"):
                self.engine.stop()

    def _on_engine_state_changed(self, new_state: EngineState):
        self.after(0, lambda: self._update_controls_for_state(new_state))

    def _update_controls_for_state(self, state: EngineState):
        has_items = bool(self.queue_mgr.get_items())
        if state == EngineState.IDLE:
            self.btn_start_print.configure(state="normal", text="▶ BẮT ĐẦU IN (PRINT)", fg_color=("#16a34a", "#15803d"))
            self.btn_pause_print.configure(state="disabled", text="⏸ Tạm Dừng")
            self.btn_stop_print.configure(state="disabled")
            self.btn_print_settings.configure(state="normal" if has_items else "disabled")
            self.btn_add_files.configure(state="normal")
            self.btn_add_folder.configure(state="normal")
            self.btn_move_up.configure(state="normal")
            self.btn_move_down.configure(state="normal")
            self.btn_move_top.configure(state="normal")
            self.btn_move_bottom.configure(state="normal")
            self.btn_remove.configure(state="normal")
            self.btn_clear_all.configure(state="normal")
            self.btn_reset_status.configure(state="normal")
        elif state == EngineState.RUNNING:
            self.btn_start_print.configure(state="disabled", text="Đang In...")
            self.btn_pause_print.configure(state="normal", text="⏸ Tạm Dừng", fg_color=("#ea580c", "#c2410c"))
            self.btn_stop_print.configure(state="normal")
            self.btn_print_settings.configure(state="disabled")
            self.btn_add_files.configure(state="disabled")
            self.btn_add_folder.configure(state="disabled")
            self.btn_move_up.configure(state="disabled")
            self.btn_move_down.configure(state="disabled")
            self.btn_move_top.configure(state="disabled")
            self.btn_move_bottom.configure(state="disabled")
            self.btn_remove.configure(state="disabled")
            self.btn_clear_all.configure(state="disabled")
            self.btn_reset_status.configure(state="disabled")
        elif state == EngineState.PAUSED:
            self.btn_start_print.configure(state="disabled")
            self.btn_pause_print.configure(state="normal", text="▶ Tiếp Tục In", fg_color=("#0284c7", "#0369a1"))
            self.btn_stop_print.configure(state="normal")
            self.btn_print_settings.configure(state="disabled")

    def _on_progress_update(self, current: int, total: int, percent: float, filename: str):
        self.after(0, lambda: self._apply_progress(current, total, percent, filename))

    def _apply_progress(self, current: int, total: int, percent: float, filename: str):
        self.progress_bar.set(percent)
        p_str = f"{int(percent * 100)}%"
        self.lbl_progress_percent.configure(text=p_str)
        self.lbl_progress_status.configure(
            text=f"Tiến độ: Đang in [{current}/{total}] - {filename}"
        )

    def _on_print_finished(self, summary: dict):
        self.after(0, lambda: self._apply_finished(summary))

    def _apply_finished(self, summary: dict):
        total = summary["total"]
        success = summary["success"]
        failed = summary["failed"]
        stopped = summary["stopped_early"]

        if stopped:
            self.lbl_progress_status.configure(
                text=f"Tiến độ: Đã dừng! (Hoàn thành: {success}/{total}, Lỗi: {failed})"
            )
            messagebox.showwarning(
                "Đã dừng in",
                f"Tiến trình in đã dừng theo yêu cầu.\n\n• Thành công: {success}/{total}\n• Thất bại: {failed}"
            )
        else:
            self.progress_bar.set(1.0)
            self.lbl_progress_percent.configure(text="100%")
            self.lbl_progress_status.configure(
                text=f"Tiến độ: Hoàn thành toàn bộ! ({success}/{total} file thành công)"
            )
            messagebox.showinfo(
                "Hoàn thành in ấn",
                f"🎉 Đã hoàn thành in toàn bộ hàng đợi!\n\n• Số file in thành công: {success}/{total}\n• Số file lỗi: {failed}"
            )

    # ================== LOG & GIAO DIỆN PHỤ ==================

    def _add_log_safe(self, message: str, level: str = "info"):
        self.after(0, lambda: self._add_log(message, level))

    def _add_log(self, message: str, level: str = "info"):
        now_str = datetime.datetime.now().strftime("%H:%M:%S")
        prefix = {
            "info": "ℹ️",
            "success": "✅",
            "warning": "⚠️",
            "error": "❌",
        }.get(level, "•")

        log_line = f"[{now_str}] {prefix} {message}\n"
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", log_line)
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def _clear_log(self):
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.configure(state="disabled")

    def _toggle_theme(self):
        if self.theme_switch.get() == 1:
            ctk.set_appearance_mode("Dark")
            self.style.configure("Custom.Treeview", background="#1e293b", foreground="#f1f5f9", fieldbackground="#1e293b")
            self.style.configure("Custom.Treeview.Heading", background="#0f172a", foreground="#38bdf8")
            self.theme_switch.configure(text="Giao diện Tối")
        else:
            ctk.set_appearance_mode("Light")
            self.style.configure("Custom.Treeview", background="#ffffff", foreground="#0f172a", fieldbackground="#ffffff")
            self.style.configure("Custom.Treeview.Heading", background="#e2e8f0", foreground="#0284c7")
            self.theme_switch.configure(text="Giao diện Sáng")
