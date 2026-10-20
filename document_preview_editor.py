import io
import os
import subprocess
import tempfile
import threading
import tkinter as tk
from tkinter import messagebox, ttk
from typing import Callable, Dict, List, Optional, Tuple
import uuid

import customtkinter as ctk
from PIL import Image, ImageTk

try:
    import pypdfium2
    PYPDFIUM_AVAILABLE = True
except ImportError:
    PYPDFIUM_AVAILABLE = False

try:
    import win32com.client
    import pythoncom
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False

try:
    import docx
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


# ==============================================================================
# 1. BỘ XEM TRƯỚC PDF TƯƠNG TÁC CAO CẤP (PDF DOCUMENT VIEWER)
# ==============================================================================

class PdfDocumentViewerComponent(ctk.CTkFrame):
    """
    Trình xem tài liệu PDF trực quan, mượt mà:
    - Hiển thị từng trang sắc nét chuẩn High-DPI qua pypdfium2.
    - Điều hướng trang: Đầu, Trước, Nhập số trang nhảy trực tiếp, Sau, Cuối.
    - Phóng to / Thu nhỏ: 50% -> 300%, Vừa chiều ngang (Fit Width), Vừa toàn trang (Fit Page).
    - Cuộn trang bằng chuột (Mousewheel) và phím mũi tên.
    """
    def __init__(
        self,
        master,
        filepath: str,
        is_temp_file: bool = False,
        original_filepath: Optional[str] = None,
        source_format_title: str = "PDF"
    ):
        super().__init__(master, fg_color="transparent")
        self.filepath = filepath
        self.is_temp_file = is_temp_file
        self.original_filepath = original_filepath or filepath
        self.source_format_title = source_format_title

        self.pdf: Optional[pypdfium2.PdfDocument] = None
        self.total_pages: int = 0
        self.current_page_idx: int = 0
        self.zoom_factor: float = 1.0  # 1.0 = 100%
        self._current_photo: Optional[ImageTk.PhotoImage] = None

        self._build_toolbar()
        self._build_canvas_area()
        self._load_pdf()

    def _build_toolbar(self):
        toolbar = ctk.CTkFrame(self, height=44, corner_radius=8, fg_color=("#e2e8f0", "#1e293b"))
        toolbar.pack(fill="x", padx=4, pady=(2, 6))

        # Điều hướng trang
        self.btn_first = ctk.CTkButton(
            toolbar, text="⏮ Đầu", width=48, height=28,
            command=self._go_first_page,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=("#64748b", "#334155")
        )
        self.btn_first.pack(side="left", padx=(8, 2))

        self.btn_prev = ctk.CTkButton(
            toolbar, text="◀ Trước", width=62, height=28,
            command=self._go_prev_page,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color=("#2563eb", "#1d4ed8")
        )
        self.btn_prev.pack(side="left", padx=2)

        ctk.CTkLabel(toolbar, text="Trang", font=ctk.CTkFont(family="Segoe UI", size=11)).pack(side="left", padx=(6, 2))

        self.var_page = tk.StringVar(value="1")
        self.ent_page = ctk.CTkEntry(
            toolbar, width=44, height=28, justify="center",
            textvariable=self.var_page,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")
        )
        self.ent_page.pack(side="left", padx=2)
        self.ent_page.bind("<Return>", lambda e: self._on_jump_page())

        self.lbl_total_pages = ctk.CTkLabel(
            toolbar, text="/ 1", font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")
        )
        self.lbl_total_pages.pack(side="left", padx=(2, 6))

        self.btn_next = ctk.CTkButton(
            toolbar, text="Sau ▶", width=62, height=28,
            command=self._go_next_page,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color=("#2563eb", "#1d4ed8")
        )
        self.btn_next.pack(side="left", padx=2)

        self.btn_last = ctk.CTkButton(
            toolbar, text="Cuối ⏭", width=52, height=28,
            command=self._go_last_page,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=("#64748b", "#334155")
        )
        self.btn_last.pack(side="left", padx=2)

        # Phân cách
        sep1 = ctk.CTkFrame(toolbar, width=1, height=22, fg_color=("#cbd5e1", "#475569"))
        sep1.pack(side="left", padx=10)

        # Bộ điều khiển Zoom
        self.btn_zoom_out = ctk.CTkButton(
            toolbar, text="➖", width=30, height=28, command=self._zoom_out,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color=("#64748b", "#334155")
        )
        self.btn_zoom_out.pack(side="left", padx=2)

        self.lbl_zoom = ctk.CTkLabel(
            toolbar, text="100%", width=48,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")
        )
        self.lbl_zoom.pack(side="left", padx=2)

        self.btn_zoom_in = ctk.CTkButton(
            toolbar, text="➕", width=30, height=28, command=self._zoom_in,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color=("#64748b", "#334155")
        )
        self.btn_zoom_in.pack(side="left", padx=2)

        self.btn_fit_width = ctk.CTkButton(
            toolbar, text="↔ Vừa Chiều Ngang", width=125, height=28,
            command=self._fit_width,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=("#475569", "#334155")
        )
        self.btn_fit_width.pack(side="left", padx=4)

        self.btn_fit_page = ctk.CTkButton(
            toolbar, text="📄 Vừa Trang", width=95, height=28,
            command=self._fit_page,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=("#475569", "#334155")
        )
        self.btn_fit_page.pack(side="left", padx=2)

        # Nhãn thông báo định dạng
        lbl_hint = ctk.CTkLabel(
            toolbar,
            text=f"💡 Bản in chuẩn {self.source_format_title} (Bảo toàn 100% định dạng)",
            font=ctk.CTkFont(family="Segoe UI", size=11, slant="italic"),
            text_color="#10b981"
        )
        lbl_hint.pack(side="right", padx=12)

    def _build_canvas_area(self):
        # Khung chứa Canvas và thanh cuộn
        self.container = ctk.CTkFrame(self, fg_color=("#334155", "#0f172a"), corner_radius=6)
        self.container.pack(fill="both", expand=True, padx=4, pady=(0, 4))
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(
            self.container,
            bg="#1e293b",
            highlightthickness=0,
            bd=0
        )
        self.canvas.grid(row=0, column=0, sticky="nsew")

        self.vsb = ttk.Scrollbar(self.container, orient="vertical", command=self.canvas.yview)
        self.vsb.grid(row=0, column=1, sticky="ns")

        self.hsb = ttk.Scrollbar(self.container, orient="horizontal", command=self.canvas.xview)
        self.hsb.grid(row=1, column=0, sticky="ew")

        self.canvas.configure(xscrollcommand=self.hsb.set, yscrollcommand=self.vsb.set)

        # Lắng nghe sự kiện chuột và bàn phím
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Control-MouseWheel>", self._on_ctrl_mousewheel)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

    def _load_pdf(self):
        if not PYPDFIUM_AVAILABLE:
            self._show_error("Thư viện pypdfium2 chưa được cài đặt.")
            return

        if not os.path.exists(self.filepath):
            self._show_error(f"Không tìm thấy tệp: {self.filepath}")
            return

        try:
            self.pdf = pypdfium2.PdfDocument(self.filepath)
            self.total_pages = len(self.pdf)
            if self.total_pages == 0:
                self._show_error("Tài liệu không có trang nào.")
                return

            self.lbl_total_pages.configure(text=f"/ {self.total_pages}")
            self.current_page_idx = 0
            self._render_current_page()

        except Exception as e:
            self._show_error(f"Lỗi khi nạp tài liệu PDF: {e}")

    def _show_error(self, message: str):
        self.canvas.delete("all")
        self.canvas.create_text(
            300, 200,
            text=f"⚠️ {message}",
            fill="#ef4444",
            font=("Segoe UI", 12, "bold")
        )

    def _render_current_page(self):
        if not self.pdf or self.current_page_idx < 0 or self.current_page_idx >= self.total_pages:
            return

        try:
            page = self.pdf[self.current_page_idx]
            # Tính scale render dựa trên zoom_factor (chuẩn 1.5x để nét ở 100%)
            scale = max(0.5, self.zoom_factor * 1.5)
            pil_img = page.render(scale=scale).to_pil()

            self._current_photo = ImageTk.PhotoImage(pil_img)
            w, h = pil_img.size

            self.canvas.delete("all")

            # Căn giữa trang trong Canvas nếu canvas lớn hơn ảnh
            cw = self.canvas.winfo_width()
            ch = self.canvas.winfo_height()
            pos_x = max(20, (cw - w) // 2) if cw > w else 20
            pos_y = max(20, (ch - h) // 2) if ch > h else 20

            # Vẽ bóng đổ nhẹ cho trang giấy
            self.canvas.create_rectangle(
                pos_x - 3, pos_y - 3, pos_x + w + 4, pos_y + h + 4,
                fill="#0b0f19", outline="#090d16"
            )
            # Vẽ trang tài liệu
            self.canvas.create_image(pos_x, pos_y, anchor="nw", image=self._current_photo)

            scroll_w = max(pos_x + w + 40, cw)
            scroll_h = max(pos_y + h + 40, ch)
            self.canvas.configure(scrollregion=(0, 0, scroll_w, scroll_h))

            # Cập nhật số trang hiển thị
            self.var_page.set(str(self.current_page_idx + 1))
            self.lbl_zoom.configure(text=f"{int(self.zoom_factor * 100)}%")

        except Exception as e:
            self._show_error(f"Lỗi hiển thị trang: {e}")

    def _on_canvas_configure(self, event):
        # Tự động căn lại trang khi thay đổi kích thước cửa sổ
        pass

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_ctrl_mousewheel(self, event):
        if event.delta > 0:
            self._zoom_in()
        else:
            self._zoom_out()

    def _zoom_in(self):
        if self.zoom_factor < 3.0:
            self.zoom_factor = min(3.0, round(self.zoom_factor + 0.15, 2))
            self._render_current_page()

    def _zoom_out(self):
        if self.zoom_factor > 0.4:
            self.zoom_factor = max(0.4, round(self.zoom_factor - 0.15, 2))
            self._render_current_page()

    def _fit_width(self):
        if not self.pdf:
            return
        cw = max(300, self.canvas.winfo_width() - 80)
        page = self.pdf[self.current_page_idx]
        base_w = page.get_size()[0]
        if base_w > 0:
            self.zoom_factor = max(0.4, min(3.0, cw / (base_w * 1.5)))
            self._render_current_page()

    def _fit_page(self):
        if not self.pdf:
            return
        ch = max(300, self.canvas.winfo_height() - 80)
        page = self.pdf[self.current_page_idx]
        base_h = page.get_size()[1]
        if base_h > 0:
            self.zoom_factor = max(0.4, min(3.0, ch / (base_h * 1.5)))
            self._render_current_page()

    def _go_prev_page(self):
        if self.current_page_idx > 0:
            self.current_page_idx -= 1
            self._render_current_page()
            self.canvas.yview_moveto(0)

    def _go_next_page(self):
        if self.current_page_idx < self.total_pages - 1:
            self.current_page_idx += 1
            self._render_current_page()
            self.canvas.yview_moveto(0)

    def _go_first_page(self):
        if self.current_page_idx != 0:
            self.current_page_idx = 0
            self._render_current_page()
            self.canvas.yview_moveto(0)

    def _go_last_page(self):
        if self.current_page_idx != self.total_pages - 1:
            self.current_page_idx = self.total_pages - 1
            self._render_current_page()
            self.canvas.yview_moveto(0)

    def _on_jump_page(self):
        try:
            val = int(self.var_page.get())
            if 1 <= val <= self.total_pages:
                self.current_page_idx = val - 1
                self._render_current_page()
                self.canvas.yview_moveto(0)
            else:
                self.var_page.set(str(self.current_page_idx + 1))
        except Exception:
            self.var_page.set(str(self.current_page_idx + 1))

    def destroy(self):
        if self.pdf:
            try:
                self.pdf.close()
            except Exception:
                pass
            self.pdf = None
        if self.is_temp_file and os.path.exists(self.filepath):
            try:
                os.remove(self.filepath)
            except Exception:
                pass
        super().destroy()


# ==============================================================================
# 2. BỘ XEM TRƯỚC WORD & EXCEL CHÂN THỰC 100% (OFFICE FAITHFUL PREVIEW)
# ==============================================================================

class OfficeFaithfulPreviewComponent(ctk.CTkFrame):
    """
    Trình xem tài liệu Word / Excel chân thực 100%:
    - Giữ nguyên hiện trạng của file gốc (mở ReadOnly=True, tuyệt đối không chỉnh sửa tệp gốc).
    - Xuất view PDF tạm thời qua COM Automation để pypdfium2 hiển thị đúng từng font chữ,
      căn lề, bảng biểu, công thức, màu sắc, header, footer của Microsoft Word & Excel.
    - Không bao giờ bị vỡ hay thay đổi format của người dùng.
    """
    def __init__(self, master, filepath: str, on_open_native_callback: Optional[Callable] = None):
        super().__init__(master, fg_color="transparent")
        self.filepath = filepath
        self.on_open_native_callback = on_open_native_callback
        self.temp_pdf_path: Optional[str] = None
        self.pdf_viewer: Optional[PdfDocumentViewerComponent] = None

        ext = os.path.splitext(self.filepath)[1].lower()
        self.doc_type = "Word" if ext in [".docx", ".doc"] else "Excel"

        self._show_loading()
        self.after(50, self._start_export_thread)

    def _show_loading(self):
        self.loading_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.loading_frame.pack(expand=True)

        ctk.CTkLabel(
            self.loading_frame,
            text=f"🔄 Đang kết nối Microsoft {self.doc_type} để dựng bản xem trước chân thực...",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color="#38bdf8"
        ).pack(pady=(20, 8))

        ctk.CTkLabel(
            self.loading_frame,
            text="Hệ thống đảm bảo giữ nguyên 100% định dạng, căn lề, bảng tính và không làm thay đổi tệp gốc.",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color="#94a3b8"
        ).pack()

    def _start_export_thread(self):
        threading.Thread(target=self._export_worker, daemon=True).start()

    def _export_worker(self):
        ext = os.path.splitext(self.filepath)[1].lower()
        temp_pdf = os.path.join(tempfile.gettempdir(), f"pyprint_prev_{uuid.uuid4().hex[:8]}.pdf")
        abs_path = os.path.abspath(self.filepath)

        success = False
        err_msg = ""

        if WIN32COM_AVAILABLE:
            try:
                pythoncom.CoInitialize()
                if ext in [".docx", ".doc"]:
                    word = win32com.client.DispatchEx("Word.Application")
                    word.Visible = False
                    word.DisplayAlerts = 0
                    doc = word.Documents.Open(abs_path, ReadOnly=True, AddToRecentFiles=False, Visible=False)
                    doc.ExportAsFixedFormat(temp_pdf, 17)  # wdExportFormatPDF = 17
                    doc.Close(False)
                    word.Quit()
                    del doc
                    del word
                elif ext in [".xlsx", ".xls"]:
                    excel = win32com.client.DispatchEx("Excel.Application")
                    excel.Visible = False
                    excel.DisplayAlerts = False
                    wb = excel.Workbooks.Open(abs_path, ReadOnly=True)
                    wb.ExportAsFixedFormat(0, temp_pdf)  # xlTypePDF = 0
                    wb.Close(False)
                    excel.Quit()
                    del wb
                    del excel

                if os.path.exists(temp_pdf) and os.path.getsize(temp_pdf) > 0:
                    # Kiểm tra xem PDF có đọc được không
                    try:
                        test_pdf = pypdfium2.PdfDocument(temp_pdf)
                        if len(test_pdf) > 0:
                            test_pdf.close()
                            success = True
                        else:
                            test_pdf.close()
                            err_msg = "Tài liệu trống (Không có trang hoặc dữ liệu in)."
                    except Exception as ex_pdf:
                        err_msg = f"Tài liệu trống hoặc chưa có dữ liệu in ({ex_pdf})"
                else:
                    err_msg = "Không thể xuất bản xem trước PDF từ ứng dụng Office."

            except Exception as e:
                err_msg = str(e)
            finally:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
        else:
            err_msg = "Thư viện COM Windows không khả dụng trên thiết bị này."

        self.after(0, lambda: self._on_export_finished(success, err_msg, temp_pdf))

    def _on_export_finished(self, success: bool, err_msg: str, temp_pdf: str):
        if hasattr(self, "loading_frame") and self.loading_frame.winfo_exists():
            self.loading_frame.destroy()

        if success and os.path.exists(temp_pdf):
            self.temp_pdf_path = temp_pdf
            self.pdf_viewer = PdfDocumentViewerComponent(
                self,
                filepath=temp_pdf,
                is_temp_file=True,
                original_filepath=self.filepath,
                source_format_title=self.doc_type
            )
            self.pdf_viewer.pack(fill="both", expand=True)
        else:
            # Fallback nếu Office COM gặp lỗi hoặc máy tính không có Office
            self._show_fallback(err_msg)

    def _show_fallback(self, error_message: str):
        fallback_frame = ctk.CTkFrame(self, fg_color="transparent")
        fallback_frame.pack(expand=True, fill="both", padx=20, pady=20)

        ctk.CTkLabel(
            fallback_frame,
            text=f"📋 BẢN XEM TRƯỚC {self.doc_type.upper()}",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color="#f8fafc"
        ).pack(pady=(16, 6))

        ctk.CTkLabel(
            fallback_frame,
            text=f"Thông báo: {error_message}\n\nTài liệu vẫn được bảo toàn nguyên vẹn trên ổ đĩa và sẵn sàng in ấn an toàn.",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#94a3b8",
            justify="center"
        ).pack(pady=10)

        btn_box = ctk.CTkFrame(fallback_frame, fg_color="transparent")
        btn_box.pack(pady=12)

        ctk.CTkButton(
            btn_box,
            text=f"✏️ Mở bằng Microsoft {self.doc_type} Gốc",
            command=self._open_in_native_app,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color=("#2563eb", "#1d4ed8"),
            height=34
        ).pack(side="left", padx=6)

    def _open_in_native_app(self):
        try:
            os.startfile(os.path.abspath(self.filepath))
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể mở file: {e}")

    def destroy(self):
        if self.pdf_viewer:
            try:
                self.pdf_viewer.destroy()
            except Exception:
                pass
            self.pdf_viewer = None
        if self.temp_pdf_path and os.path.exists(self.temp_pdf_path):
            try:
                os.remove(self.temp_pdf_path)
            except Exception:
                pass
        super().destroy()


# ==============================================================================
# 3. BỘ XEM TRƯỚC & XOAY/LẬT ẢNH (IMAGE PREVIEW & ROTATE EDITOR)
# ==============================================================================

class ImagePreviewEditorComponent(ctk.CTkFrame):
    """
    Trình xem trước và chỉnh hướng ảnh chuyên dụng:
    - Xoay 90° theo chiều kim đồng hồ, xoay ngược 90°.
    - Lật ngang (Horizontal Flip), Lật dọc (Vertical Flip).
    - Phóng to / Thu nhỏ / Vừa khung hình.
    - Lưu trực tiếp ảnh sau khi xoay/lật vào file gốc.
    """
    def __init__(self, master, filepath: str, on_save_callback: Optional[Callable] = None):
        super().__init__(master, fg_color="transparent")
        self.filepath = filepath
        self.on_save_callback = on_save_callback

        self.original_img: Optional[Image.Image] = None
        self.current_img: Optional[Image.Image] = None
        self.zoom_factor: float = 1.0
        self._photo_ref: Optional[ImageTk.PhotoImage] = None
        self._has_changes: bool = False

        self._build_toolbar()
        self._build_canvas()
        self._load_image()

    def _build_toolbar(self):
        toolbar = ctk.CTkFrame(self, height=44, corner_radius=8, fg_color=("#e2e8f0", "#1e293b"))
        toolbar.pack(fill="x", padx=4, pady=(2, 6))

        # Nhóm xoay ảnh
        ctk.CTkButton(
            toolbar, text="🔄 Xoay Phải 90°", width=110, height=28,
            command=self._rotate_right,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color=("#0284c7", "#0369a1")
        ).pack(side="left", padx=(8, 3))

        ctk.CTkButton(
            toolbar, text="🔄 Xoay Trái 90°", width=110, height=28,
            command=self._rotate_left,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=("#475569", "#334155")
        ).pack(side="left", padx=3)

        # Nhóm lật ảnh
        ctk.CTkButton(
            toolbar, text="↔️ Lật Ngang", width=95, height=28,
            command=self._flip_horizontal,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=("#475569", "#334155")
        ).pack(side="left", padx=3)

        ctk.CTkButton(
            toolbar, text="↕️ Lật Dọc", width=85, height=28,
            command=self._flip_vertical,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=("#475569", "#334155")
        ).pack(side="left", padx=3)

        # Phân cách
        ctk.CTkFrame(toolbar, width=1, height=22, fg_color=("#cbd5e1", "#475569")).pack(side="left", padx=8)

        # Zoom
        ctk.CTkButton(
            toolbar, text="➖", width=28, height=28, command=self._zoom_out,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color=("#64748b", "#334155")
        ).pack(side="left", padx=2)

        self.lbl_zoom = ctk.CTkLabel(toolbar, text="100%", width=46, font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"))
        self.lbl_zoom.pack(side="left", padx=2)

        ctk.CTkButton(
            toolbar, text="➕", width=28, height=28, command=self._zoom_in,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color=("#64748b", "#334155")
        ).pack(side="left", padx=2)

        ctk.CTkButton(
            toolbar, text="Vừa Khung", width=80, height=28, command=self._fit_canvas,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=("#475569", "#334155")
        ).pack(side="left", padx=4)

        # Nút Khôi phục
        ctk.CTkButton(
            toolbar, text="↩️ Đặt Lại Gốc", width=95, height=28, command=self._reset_image,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=("#dc2626", "#b91c1c")
        ).pack(side="left", padx=4)

        # Nhãn kích thước ảnh
        self.lbl_info = ctk.CTkLabel(
            toolbar, text="", font=ctk.CTkFont(family="Segoe UI", size=11), text_color="#94a3b8"
        )
        self.lbl_info.pack(side="right", padx=12)

    def _build_canvas(self):
        container = ctk.CTkFrame(self, fg_color=("#334155", "#0f172a"), corner_radius=6)
        container.pack(fill="both", expand=True, padx=4, pady=(0, 4))
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(container, bg="#1e293b", highlightthickness=0, bd=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")

        vsb = ttk.Scrollbar(container, orient="vertical", command=self.canvas.yview)
        vsb.grid(row=0, column=1, sticky="ns")

        hsb = ttk.Scrollbar(container, orient="horizontal", command=self.canvas.xview)
        hsb.grid(row=1, column=0, sticky="ew")

        self.canvas.configure(xscrollcommand=hsb.set, yscrollcommand=vsb.set)
        self.canvas.bind("<MouseWheel>", lambda e: self.canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

    def _load_image(self):
        try:
            self.original_img = Image.open(self.filepath)
            self.current_img = self.original_img.copy()
            self._fit_canvas()
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể nạp tệp hình ảnh: {e}")

    def _render(self):
        if not self.current_img:
            return

        w, h = self.current_img.size
        disp_w = max(10, int(w * self.zoom_factor))
        disp_h = max(10, int(h * self.zoom_factor))

        resized = self.current_img.resize((disp_w, disp_h), Image.Resampling.LANCZOS)
        self._photo_ref = ImageTk.PhotoImage(resized)

        self.canvas.delete("all")
        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()

        pos_x = max(20, (cw - disp_w) // 2) if cw > disp_w else 20
        pos_y = max(20, (ch - disp_h) // 2) if ch > disp_h else 20

        # Khung viền ảnh
        self.canvas.create_rectangle(pos_x - 1, pos_y - 1, pos_x + disp_w + 1, pos_y + disp_h + 1, outline="#38bdf8", width=1)
        self.canvas.create_image(pos_x, pos_y, anchor="nw", image=self._photo_ref)

        self.canvas.configure(scrollregion=(0, 0, max(pos_x + disp_w + 30, cw), max(pos_y + disp_h + 30, ch)))
        self.lbl_zoom.configure(text=f"{int(self.zoom_factor * 100)}%")
        self.lbl_info.configure(text=f"📐 {w} × {h} px | {os.path.splitext(self.filepath)[1].upper()}")

    def _rotate_right(self):
        if self.current_img:
            self.current_img = self.current_img.rotate(-90, expand=True)
            self._has_changes = True
            self._render()

    def _rotate_left(self):
        if self.current_img:
            self.current_img = self.current_img.rotate(90, expand=True)
            self._has_changes = True
            self._render()

    def _flip_horizontal(self):
        if self.current_img:
            self.current_img = self.current_img.transpose(Image.FLIP_LEFT_RIGHT)
            self._has_changes = True
            self._render()

    def _flip_vertical(self):
        if self.current_img:
            self.current_img = self.current_img.transpose(Image.FLIP_TOP_BOTTOM)
            self._has_changes = True
            self._render()

    def _zoom_in(self):
        if self.zoom_factor < 3.0:
            self.zoom_factor = min(3.0, round(self.zoom_factor + 0.15, 2))
            self._render()

    def _zoom_out(self):
        if self.zoom_factor > 0.2:
            self.zoom_factor = max(0.2, round(self.zoom_factor - 0.15, 2))
            self._render()

    def _fit_canvas(self):
        if not self.current_img:
            return
        cw = max(200, self.canvas.winfo_width() - 60)
        ch = max(200, self.canvas.winfo_height() - 60)
        w, h = self.current_img.size
        ratio = min(cw / w, ch / h)
        self.zoom_factor = min(1.0, max(0.1, ratio))
        self._render()

    def _reset_image(self):
        if self.original_img:
            self.current_img = self.original_img.copy()
            self._has_changes = False
            self._fit_canvas()

    def save_changes(self) -> Tuple[bool, str]:
        if not self._has_changes or not self.current_img:
            return True, "Không có thay đổi cần lưu."
        try:
            self.current_img.save(self.filepath)
            self._has_changes = False
            return True, "Đã lưu thay đổi vào tệp hình ảnh thành công!"
        except Exception as e:
            return False, f"Lỗi khi lưu ảnh: {e}"


# ==============================================================================
# 4. BỘ XEM TRƯỚC & SỬA NHANH VĂN BẢN (TEXT PREVIEW & QUICK EDITOR)
# ==============================================================================

class TextPreviewEditorComponent(ctk.CTkFrame):
    """
    Trình xem và chỉnh sửa tệp văn bản nhanh:
    - Hỗ trợ .txt, .csv, .log, .md, .json, .py...
    - Font Monospace Consolas rõ ràng, hiển thị số dòng và ký tự.
    - Tìm kiếm từ khóa nhanh trong tệp.
    - Lưu trực tiếp với mã hóa UTF-8 an toàn.
    """
    def __init__(self, master, filepath: str, on_save_callback: Optional[Callable] = None):
        super().__init__(master, fg_color="transparent")
        self.filepath = filepath
        self.on_save_callback = on_save_callback

        self._build_toolbar()
        self._build_editor()
        self._load_file()

    def _build_toolbar(self):
        toolbar = ctk.CTkFrame(self, height=44, corner_radius=8, fg_color=("#e2e8f0", "#1e293b"))
        toolbar.pack(fill="x", padx=4, pady=(2, 6))

        # Cỡ chữ
        ctk.CTkLabel(toolbar, text="Cỡ chữ:", font=ctk.CTkFont(family="Segoe UI", size=11)).pack(side="left", padx=(10, 4))
        self.cb_size = ctk.CTkComboBox(
            toolbar, values=["9", "10", "11", "12", "14", "16", "18"],
            width=65, height=28, command=self._apply_font_size
        )
        self.cb_size.set("11")
        self.cb_size.pack(side="left", padx=(0, 10))

        # Tìm kiếm từ khóa
        ctk.CTkLabel(toolbar, text="Tìm kiếm:", font=ctk.CTkFont(family="Segoe UI", size=11)).pack(side="left", padx=(0, 4))
        self.ent_search = ctk.CTkEntry(toolbar, width=160, height=28, placeholder_text="Nhập từ cần tìm...")
        self.ent_search.pack(side="left", padx=(0, 4))
        self.ent_search.bind("<Return>", lambda e: self._find_next())

        ctk.CTkButton(
            toolbar, text="🔍 Tìm", width=55, height=28, command=self._find_next,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=("#0284c7", "#0369a1")
        ).pack(side="left", padx=2)

        # Nhãn thống kê
        self.lbl_stats = ctk.CTkLabel(
            toolbar, text="0 dòng | 0 ký tự | UTF-8", font=ctk.CTkFont(family="Segoe UI", size=11), text_color="#94a3b8"
        )
        self.lbl_stats.pack(side="right", padx=12)

    def _build_editor(self):
        container = ctk.CTkFrame(self, fg_color=("#f1f5f9", "#090d16"), corner_radius=6)
        container.pack(fill="both", expand=True, padx=4, pady=(0, 4))
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.text_box = tk.Text(
            container,
            wrap="word",
            undo=True,
            font=("Consolas", 11),
            bg="#ffffff",
            fg="#0f172a",
            insertbackground="#2563eb",
            selectbackground="#93c5fd",
            padx=16,
            pady=16,
            relief="flat",
            bd=0
        )
        self.text_box.grid(row=0, column=0, sticky="nsew")

        vsb = ttk.Scrollbar(container, orient="vertical", command=self.text_box.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        self.text_box.configure(yscrollcommand=vsb.set)

        self.text_box.tag_configure("highlight", background="#fef08a", foreground="#854d0e")
        self.text_box.bind("<KeyRelease>", lambda e: self._update_stats())

    def _load_file(self):
        try:
            with open(self.filepath, "r", encoding="utf-8", errors="replace") as f:
                content = f.read()
            self.text_box.delete("1.0", "end")
            self.text_box.insert("1.0", content)
            self._update_stats()
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể đọc tệp văn bản: {e}")

    def _apply_font_size(self, size_str):
        try:
            sz = int(size_str)
            self.text_box.configure(font=("Consolas", sz))
        except Exception:
            pass

    def _find_next(self):
        query = self.ent_search.get().strip()
        if not query:
            return

        self.text_box.tag_remove("highlight", "1.0", "end")
        start_pos = self.text_box.index("insert")
        pos = self.text_box.search(query, start_pos, nocase=True, stopindex="end")
        if not pos:
            # Quay lại tìm từ đầu file
            pos = self.text_box.search(query, "1.0", nocase=True, stopindex="end")

        if pos:
            end_pos = f"{pos}+{len(query)}c"
            self.text_box.tag_add("highlight", pos, end_pos)
            self.text_box.see(pos)
            self.text_box.mark_set("insert", end_pos)
        else:
            messagebox.showinfo("Tìm kiếm", f"Không tìm thấy cụm từ '{query}'.")

    def _update_stats(self):
        content = self.text_box.get("1.0", "end-1c")
        lines = content.count("\n") + (1 if content else 0)
        chars = len(content)
        words = len(content.split())
        self.lbl_stats.configure(text=f"📊 {lines} dòng | {words} từ | {chars} ký tự | UTF-8")

    def save_changes(self) -> Tuple[bool, str]:
        try:
            content = self.text_box.get("1.0", "end-1c")
            with open(self.filepath, "w", encoding="utf-8") as f:
                f.write(content)
            return True, "Đã lưu thay đổi vào tệp văn bản thành công!"
        except Exception as e:
            return False, f"Lỗi khi lưu tệp văn bản: {e}"


# ==============================================================================
# 5. KHUNG XEM TRƯỚC TỔNG THỂ (UNIVERSAL DOCUMENT PREVIEW VIEW)
# ==============================================================================

class DocumentPreviewEditorView(ctk.CTkFrame):
    """
    Vùng xem trước tổng thể được tích hợp trực tiếp thay thế vị trí hàng đợi:
    - PDF: Trình xem tương tác đa trang pypdfium2.
    - Word & Excel: Dựng bản xem trước chuẩn mực 100% qua Office COM,
      giữ nguyên hiện trạng tệp gốc (ReadOnly) và không làm vỡ format.
    - Ảnh: Trình xem và xoay 90°, lật ảnh, zoom và lưu thay đổi.
    - Văn bản: Trình đọc & sửa nhanh mã nguồn / văn bản.
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
        self.editor_component = None

        self._build_header()
        self._build_body()

    def _build_header(self):
        header = ctk.CTkFrame(self, height=44, corner_radius=8, fg_color=("#0f172a", "#0f172a"))
        header.pack(fill="x", padx=4, pady=(2, 6))

        # Nút Quay lại hàng đợi
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

        # Biểu tượng định dạng
        icon = "📕" if self.file_type == ".pdf" else (
            "📝" if self.file_type in [".doc", ".docx"] else (
                "📊" if self.file_type in [".xls", ".xlsx"] else (
                    "🖼️" if self.file_type in [".png", ".jpg", ".jpeg", ".bmp", ".webp"] else "📄"
                )
            )
        )

        lbl_file = ctk.CTkLabel(
            header,
            text=f"{icon} BẢN XEM TRƯỚC: {self.filename}",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color="#38bdf8",
        )
        lbl_file.pack(side="left", padx=10)

        # Nút Đóng
        btn_close = ctk.CTkButton(
            header,
            text="✖️ Đóng",
            width=65,
            height=30,
            command=self._on_close_clicked,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=("#dc2626", "#b91c1c"),
            hover_color=("#b91c1c", "#dc2626"),
        )
        btn_close.pack(side="right", padx=(4, 8), pady=6)

        # Nút Mở bằng ứng dụng gốc
        btn_open_native = ctk.CTkButton(
            header,
            text="👁️ Mở App Gốc",
            width=105,
            height=30,
            command=self._open_in_native_app,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=("#0284c7", "#0369a1"),
            hover_color=("#0369a1", "#0284c7"),
        )
        btn_open_native.pack(side="right", padx=4, pady=6)

        # Nút Lưu thay đổi (chỉ hiển thị nếu định dạng hỗ trợ sửa ảnh/text)
        if self.file_type in [".png", ".jpg", ".jpeg", ".bmp", ".webp", ".txt", ".csv", ".log", ".md", ".json", ".xml", ".py"]:
            self.btn_save = ctk.CTkButton(
                header,
                text="💾 LƯU THAY ĐỔI",
                width=135,
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

        # 1. Tệp PDF
        if self.file_type == ".pdf":
            self.editor_component = PdfDocumentViewerComponent(body_container, filepath=self.filepath)
            self.editor_component.pack(fill="both", expand=True)

        # 2. Tệp Word (.docx, .doc) và Excel (.xlsx, .xls)
        elif self.file_type in [".doc", ".docx", ".xls", ".xlsx"]:
            self.editor_component = OfficeFaithfulPreviewComponent(body_container, filepath=self.filepath)
            self.editor_component.pack(fill="both", expand=True)

        # 3. Tệp hình ảnh
        elif self.file_type in [".png", ".jpg", ".jpeg", ".bmp", ".webp", ".tif", ".tiff"]:
            self.editor_component = ImagePreviewEditorComponent(body_container, filepath=self.filepath)
            self.editor_component.pack(fill="both", expand=True)

        # 4. Tệp văn bản thuần
        elif self.file_type in [".txt", ".csv", ".log", ".md", ".json", ".xml", ".ini", ".py", ".sql"]:
            self.editor_component = TextPreviewEditorComponent(body_container, filepath=self.filepath)
            self.editor_component.pack(fill="both", expand=True)

        # 5. Định dạng khác
        else:
            lbl_fallback = ctk.CTkLabel(
                body_container,
                text=f"Tài liệu định dạng {self.file_type.upper()} không có trình xem trước chuyên biệt.\n\nNhấn '👁️ Mở App Gốc' ở góc trên để xem bằng phần mềm mặc định của Windows.",
                font=ctk.CTkFont(family="Segoe UI", size=13),
                text_color="#94a3b8",
            )
            lbl_fallback.pack(expand=True)

    def _open_in_native_app(self):
        try:
            os.startfile(os.path.abspath(self.filepath))
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể mở tệp: {e}")

    def _on_save_clicked(self):
        if self.editor_component and hasattr(self.editor_component, "save_changes"):
            ok, msg = self.editor_component.save_changes()
            if ok:
                messagebox.showinfo("Thành công", msg)
                self.on_save_callback()
            else:
                messagebox.showerror("Lỗi khi lưu", msg)
        else:
            self.on_save_callback()

    def _on_close_clicked(self):
        self.on_close_callback()

    def destroy(self):
        if self.editor_component:
            try:
                self.editor_component.destroy()
            except Exception:
                pass
            self.editor_component = None
        super().destroy()
