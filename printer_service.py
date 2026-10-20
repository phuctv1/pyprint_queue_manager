import os
import subprocess
import time
from typing import Dict, List, Optional, Tuple

try:
    import win32api
    import win32con
    import win32gui
    import win32print
    import win32ui
    import pythoncom
    import win32com.client
    import winreg
    from PIL import Image, ImageWin
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

try:
    import pypdfium2
    PYPDFIUM_AVAILABLE = True
except ImportError:
    PYPDFIUM_AVAILABLE = False


class PrinterService:
    @staticmethod
    def is_available() -> bool:
        return WIN32_AVAILABLE

    @staticmethod
    def is_pdfium_available() -> bool:
        return PYPDFIUM_AVAILABLE

    @staticmethod
    def get_printers() -> List[Dict]:
        """
        Lấy danh sách tất cả máy in đã kết nối với máy tính.
        Trả về danh sách dictionary chứa thông tin chi tiết từng máy in.
        """
        printers_list = []
        if not WIN32_AVAILABLE:
            return printers_list

        try:
            default_printer = win32print.GetDefaultPrinter()
        except Exception:
            default_printer = ""

        try:
            # PRINTER_ENUM_LOCAL: máy in cục bộ (USB, LPT...)
            # PRINTER_ENUM_CONNECTIONS: máy in mạng (Network, Shared, IP...)
            flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            enum_results = win32print.EnumPrinters(flags)

            for item in enum_results:
                printer_name = item[2]
                port_name = "N/A"
                driver_name = "N/A"
                status_code = 0
                status_text = "Sẵn sàng"
                jobs_count = 0

                try:
                    h_printer = win32print.OpenPrinter(printer_name)
                    try:
                        info = win32print.GetPrinter(h_printer, 2)
                        port_name = info.get("pPortName", "N/A")
                        driver_name = info.get("pDriverName", "N/A")
                        status_code = info.get("Status", 0)
                        jobs_count = info.get("cJobs", 0)

                        if status_code & win32print.PRINTER_STATUS_PAUSED:
                            status_text = "Đang tạm dừng"
                        elif status_code & win32print.PRINTER_STATUS_ERROR:
                            status_text = "Bị lỗi"
                        elif status_code & win32print.PRINTER_STATUS_PENDING_DELETION:
                            status_text = "Đang chờ xóa"
                        elif status_code & win32print.PRINTER_STATUS_PAPER_JAM:
                            status_text = "Kẹt giấy"
                        elif status_code & win32print.PRINTER_STATUS_PAPER_OUT:
                            status_text = "Hết giấy"
                        elif status_code & win32print.PRINTER_STATUS_MANUAL_FEED:
                            status_text = "Cần nạp giấy thủ công"
                        elif status_code & win32print.PRINTER_STATUS_PAPER_PROBLEM:
                            status_text = "Sự cố giấy"
                        elif status_code & win32print.PRINTER_STATUS_OFFLINE:
                            status_text = "Ngoại tuyến (Offline)"
                        elif status_code & win32print.PRINTER_STATUS_IO_ACTIVE:
                            status_text = "Đang truyền dữ liệu"
                        elif status_code & win32print.PRINTER_STATUS_BUSY:
                            status_text = "Đang bận"
                        elif status_code & win32print.PRINTER_STATUS_PRINTING:
                            status_text = "Đang in"
                        elif status_code & win32print.PRINTER_STATUS_OUTPUT_BIN_FULL:
                            status_text = "Khay giấy ra đầy"
                        elif status_code & win32print.PRINTER_STATUS_NOT_AVAILABLE:
                            status_text = "Không khả dụng"
                        elif status_code & win32print.PRINTER_STATUS_WAITING:
                            status_text = "Đang chờ"
                        elif status_code & win32print.PRINTER_STATUS_PROCESSING:
                            status_text = "Đang xử lý"
                        elif status_code & win32print.PRINTER_STATUS_INITIALIZING:
                            status_text = "Đang khởi tạo"
                        elif status_code & win32print.PRINTER_STATUS_WARMING_UP:
                            status_text = "Đang làm nóng máy"
                        elif status_code & win32print.PRINTER_STATUS_TONER_LOW:
                            status_text = "Sắp hết mực"
                        elif status_code & win32print.PRINTER_STATUS_NO_TONER:
                            status_text = "Hết mực"
                        elif status_code & win32print.PRINTER_STATUS_PAGE_PUNT:
                            status_text = "Lỗi trang"
                        elif status_code & win32print.PRINTER_STATUS_USER_INTERVENTION:
                            status_text = "Cần người dùng can thiệp"
                        elif status_code & win32print.PRINTER_STATUS_OUT_OF_MEMORY:
                            status_text = "Tràn bộ nhớ"
                        elif status_code & win32print.PRINTER_STATUS_DOOR_OPEN:
                            status_text = "Nắp máy in đang mở"
                        elif status_code & win32print.PRINTER_STATUS_SERVER_UNKNOWN:
                            status_text = "Máy chủ không xác định"
                        elif status_code & win32print.PRINTER_STATUS_POWER_SAVE:
                            status_text = "Chế độ tiết kiệm điện"
                        else:
                            status_text = "Sẵn sàng"
                    finally:
                        win32print.ClosePrinter(h_printer)
                except Exception:
                    pass

                printers_list.append({
                    "name": printer_name,
                    "port": port_name,
                    "driver": driver_name,
                    "is_default": (printer_name == default_printer),
                    "status_code": status_code,
                    "status_text": status_text,
                    "jobs_count": jobs_count,
                })

        except Exception as e:
            print(f"Lỗi khi quét danh sách máy in: {e}")

        return printers_list

    @staticmethod
    def get_default_printer() -> Optional[str]:
        if not WIN32_AVAILABLE:
            return None
        try:
            return win32print.GetDefaultPrinter()
        except Exception:
            return None

    @staticmethod
    def set_default_printer(printer_name: str) -> bool:
        if not WIN32_AVAILABLE:
            return False
        try:
            win32print.SetDefaultPrinter(printer_name)
            return True
        except Exception as e:
            print(f"Không thể đặt máy in mặc định {printer_name}: {e}")
            return False

    @staticmethod
    def get_printer_status(printer_name: str) -> Tuple[str, int]:
        """Trả về (status_text, jobs_count)"""
        if not WIN32_AVAILABLE:
            return "Không khả dụng", 0
        try:
            h = win32print.OpenPrinter(printer_name)
            try:
                info = win32print.GetPrinter(h, 2)
                jobs = info.get("cJobs", 0)
                code = info.get("Status", 0)
                status = "Sẵn sàng"
                if code & win32print.PRINTER_STATUS_PAUSED:
                    status = "Tạm dừng"
                elif code & win32print.PRINTER_STATUS_OFFLINE:
                    status = "Ngoại tuyến"
                elif code & win32print.PRINTER_STATUS_BUSY:
                    status = "Đang bận"
                elif code & win32print.PRINTER_STATUS_PRINTING:
                    status = "Đang in"
                return status, jobs
            finally:
                win32print.ClosePrinter(h)
        except Exception as e:
            return f"Lỗi: {e}", 0

    # =========================================================================
    # DEVMODE & PRINT SETTINGS HELPERS
    # =========================================================================

    @classmethod
    def get_printer_devmode(cls, printer_name: str):
        """Lấy cấu trúc PyDEVMODE hiện tại của máy in."""
        if not WIN32_AVAILABLE or not printer_name:
            return None
        try:
            h = win32print.OpenPrinter(printer_name)
            try:
                info = win32print.GetPrinter(h, 2)
                return info.get("pDevMode")
            finally:
                win32print.ClosePrinter(h)
        except Exception as e:
            print(f"Lỗi lấy DEVMODE máy in '{printer_name}': {e}")
            return None

    @classmethod
    def apply_settings_to_devmode(cls, devmode, settings: Optional[Dict], copies: int = 1):
        """
        Nạp cấu hình từ Print Settings vào cấu trúc PyDEVMODE của Windows:
        - Copies (Số bản in)
        - Orientation (Khổ dọc / Khổ ngang)
        - PaperSize (A4, A3, Letter, Legal...)
        - Duplex (In 1 mặt / In 2 mặt lật cạnh dài / In 2 mặt lật cạnh ngắn)
        - Collate (Soạn bộ 1,2,3... hay 1,1,1...)
        """
        if not devmode:
            return None

        # Thiết lập số bản in
        try:
            if copies and int(copies) > 0:
                devmode.Copies = int(copies)
        except Exception:
            pass

        if not settings:
            return devmode

        # 1. Hướng giấy (Orientation)
        orient = str(settings.get("orientation", "")).lower()
        if "landscape" in orient:
            devmode.Orientation = win32con.DMORIENT_LANDSCAPE
        elif "portrait" in orient:
            devmode.Orientation = win32con.DMORIENT_PORTRAIT

        # 2. Khổ giấy (PaperSize)
        paper = str(settings.get("paper_size", "")).upper()
        if "A4" in paper:
            devmode.PaperSize = win32con.DMPAPER_A4
        elif "A3" in paper:
            devmode.PaperSize = win32con.DMPAPER_A3
        elif "LETTER" in paper:
            devmode.PaperSize = win32con.DMPAPER_LETTER
        elif "LEGAL" in paper:
            devmode.PaperSize = win32con.DMPAPER_LEGAL
        elif "A5" in paper:
            devmode.PaperSize = win32con.DMPAPER_A5
        elif "B5" in paper:
            devmode.PaperSize = win32con.DMPAPER_B5

        # 3. In 2 mặt (Duplex)
        duplex_val = settings.get("duplex")
        if isinstance(duplex_val, tuple):
            duplex_str = f"{duplex_val[0]} {duplex_val[1]}".lower()
        else:
            duplex_str = str(duplex_val).lower()

        if "one sided" in duplex_str or "1 mặt" in duplex_str or "simplex" in duplex_str:
            devmode.Duplex = win32con.DMDUP_SIMPLEX
        elif "short edge" in duplex_str or "cạnh ngắn" in duplex_str:
            devmode.Duplex = win32con.DMDUP_HORIZONTAL
        elif "both sides" in duplex_str or "long edge" in duplex_str or "2 mặt" in duplex_str:
            devmode.Duplex = win32con.DMDUP_VERTICAL

        # 4. Soạn bộ (Collate)
        collate = str(settings.get("collated", "")).lower()
        if "uncollated" in collate or "không" in collate:
            devmode.Collate = win32con.DMCOLLATE_FALSE
        elif "collated" in collate or "soạn" in collate:
            devmode.Collate = win32con.DMCOLLATE_TRUE

        return devmode

    @classmethod
    def get_excel_printer_string(cls, printer_name: str) -> str:
        """Lấy chuỗi định dạng tên máy in cho Excel (ví dụ 'Xerox ... on Ne00:')."""
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows NT\CurrentVersion\Devices"
            )
            val, _ = winreg.QueryValueEx(key, printer_name)
            parts = val.split(",")
            if len(parts) >= 2:
                port = parts[1].strip()
                return f"{printer_name} on {port}"
        except Exception:
            pass
        return printer_name

    @staticmethod
    def parse_page_ranges(range_str: str, max_pages: int) -> List[int]:
        """Phân tích chuỗi khoảng trang như '1-3, 5, 8' thành danh sách chỉ số 0-indexed."""
        if not range_str or not str(range_str).strip():
            return list(range(max_pages))

        s = str(range_str).strip().lower()
        if "all" in s or "whole" in s:
            return list(range(max_pages))

        pages = set()
        parts = s.replace(";", ",").split(",")
        for part in parts:
            part = part.strip()
            if not part:
                continue
            if "-" in part:
                bounds = part.split("-")
                try:
                    start = int(bounds[0].strip())
                    end = int(bounds[1].strip())
                    if start <= end:
                        for p in range(start, end + 1):
                            if 1 <= p <= max_pages:
                                pages.add(p - 1)
                except ValueError:
                    pass
            else:
                try:
                    p = int(part)
                    if 1 <= p <= max_pages:
                        pages.add(p - 1)
                except ValueError:
                    pass

        result = sorted(list(pages))
        return result if result else list(range(max_pages))

    # =========================================================================
    # CORE PRINT ENGINES (PDF GDI, IMAGE GDI, WORD COM, EXCEL COM)
    # =========================================================================

    @classmethod
    def _print_pdf_direct(
        cls,
        pdf_path: str,
        printer_name: str,
        copies: int = 1,
        settings: Optional[Dict] = None
    ) -> Tuple[bool, str]:
        """
        In file PDF trực tiếp qua Windows GDI + pypdfium2:
        - 100% ngầm (Silent), không phụ thuộc Adobe/Edge, không hiện bất kỳ popup nào.
        - Áp dụng DEVMODE: Duplex (in 2 mặt), Orientation (dọc/ngang), Khổ giấy (A4/A3...), Collate, Copies.
        - Tự động lọc khoảng trang theo settings['pages_range'].
        """
        if not PYPDFIUM_AVAILABLE or not WIN32_AVAILABLE:
            return False, "Thư viện pypdfium2 hoặc pywin32 không khả dụng"

        doc = None
        hdc_handle = None
        try:
            doc = pypdfium2.PdfDocument(pdf_path)
            total_pages = len(doc)
            if total_pages == 0:
                return False, "Tài liệu PDF không có trang nào."

            # Phân tích khoảng trang cần in
            range_str = settings.get("pages_range", "") if settings else ""
            target_indices = cls.parse_page_ranges(range_str, total_pages)
            if not target_indices:
                return False, "Không có trang nào phù hợp với khoảng trang đã chọn."

            # Khởi tạo DEVMODE và nạp cài đặt in
            devmode = cls.get_printer_devmode(printer_name)
            if devmode:
                cls.apply_settings_to_devmode(devmode, settings, copies=copies)

            # Tạo DC máy in với DEVMODE
            if devmode:
                hdc_handle = win32gui.CreateDC("WINSPOOL", printer_name, devmode)
            else:
                hdc_handle = win32gui.CreateDC("WINSPOOL", printer_name, None)

            hDC = win32ui.CreateDCFromHandle(hdc_handle)

            pw = hDC.GetDeviceCaps(win32con.HORZRES)
            ph = hDC.GetDeviceCaps(win32con.VERTRES)
            dpi_x = hDC.GetDeviceCaps(win32con.LOGPIXELSX)

            # Scale render sắc nét (tối thiểu 2x, chuẩn theo DPI của máy in)
            scale = max(2.0, min(4.0, dpi_x / 72.0))

            doc_title = f"PyPrint: {os.path.basename(pdf_path)}"
            hDC.StartDoc(doc_title)

            # Kiểm tra xem máy in có tự nhân bản (hardware copies) hay cần lặp software
            num_runs = 1 if (devmode and getattr(devmode, "Copies", 1) == copies and copies > 1) else copies

            for run in range(num_runs):
                for page_idx in target_indices:
                    page = doc[page_idx]
                    hDC.StartPage()

                    pil_img = page.render(scale=scale).to_pil()
                    img_w, img_h = pil_img.size

                    # Căn tỉ lệ vừa khít trang in (Fit to printable area)
                    ratio = min(pw / img_w, ph / img_h)
                    final_w = int(img_w * ratio)
                    final_h = int(img_h * ratio)

                    pos_x = (pw - final_w) // 2
                    pos_y = (ph - final_h) // 2

                    dib = ImageWin.Dib(pil_img)
                    dib.draw(hDC.GetHandleOutput(), (pos_x, pos_y, pos_x + final_w, pos_y + final_h))

                    hDC.EndPage()

            hDC.EndDoc()
            return True, f"In ngầm PDF thành công ({len(target_indices)}/{total_pages} trang, {copies} bản)"

        except Exception as e:
            return False, f"Lỗi in ngầm PDF (GDI): {e}"
        finally:
            if doc:
                try:
                    doc.close()
                except Exception:
                    pass
            if hdc_handle:
                try:
                    win32gui.DeleteDC(hdc_handle)
                except Exception:
                    pass

    @classmethod
    def _print_image_gdi(
        cls,
        image_path: str,
        printer_name: str,
        copies: int = 1,
        settings: Optional[Dict] = None
    ) -> Tuple[bool, str]:
        """In ảnh trực tiếp qua Windows GDI với DEVMODE áp dụng (khổ giấy, hướng giấy, lề, số bản)."""
        hdc_handle = None
        try:
            devmode = cls.get_printer_devmode(printer_name)
            if devmode:
                cls.apply_settings_to_devmode(devmode, settings, copies=copies)
                hdc_handle = win32gui.CreateDC("WINSPOOL", printer_name, devmode)
            else:
                hdc_handle = win32gui.CreateDC("WINSPOOL", printer_name, None)

            hDC = win32ui.CreateDCFromHandle(hdc_handle)

            printable_w = hDC.GetDeviceCaps(win32con.HORZRES)
            printable_h = hDC.GetDeviceCaps(win32con.VERTRES)

            img = Image.open(image_path)
            if img.mode != "RGB":
                img = img.convert("RGB")

            # Tự động xoay ảnh nếu hướng giấy Landscape mà ảnh đang là Portrait
            orient_setting = str(settings.get("orientation", "")).lower() if settings else ""
            if "landscape" in orient_setting and img.size[0] < img.size[1]:
                img = img.rotate(90, expand=True)

            # Tính toán tỉ lệ căn vừa trang in có lề an toàn 4%
            margin_x = int(printable_w * 0.04)
            margin_y = int(printable_h * 0.04)
            target_w = printable_w - (margin_x * 2)
            target_h = printable_h - (margin_y * 2)

            img_w, img_h = img.size
            ratio = min(target_w / img_w, target_h / img_h)

            final_w = int(img_w * ratio)
            final_h = int(img_h * ratio)

            pos_x = margin_x + (target_w - final_w) // 2
            pos_y = margin_y + (target_h - final_h) // 2

            hDC.StartDoc(f"Print Image: {os.path.basename(image_path)}")

            num_runs = 1 if (devmode and getattr(devmode, "Copies", 1) == copies and copies > 1) else copies
            for _ in range(num_runs):
                hDC.StartPage()
                dib = ImageWin.Dib(img)
                dib.draw(hDC.GetHandleOutput(), (pos_x, pos_y, pos_x + final_w, pos_y + final_h))
                hDC.EndPage()

            hDC.EndDoc()
            return True, f"In ảnh GDI thành công ({copies} bản)"
        except Exception as e:
            return False, f"Lỗi GDI: {e}"
        finally:
            if hdc_handle:
                try:
                    win32gui.DeleteDC(hdc_handle)
                except Exception:
                    pass

    @classmethod
    def _print_word_com(
        cls,
        doc_path: str,
        printer_name: str,
        copies: int = 1,
        settings: Optional[Dict] = None
    ) -> Tuple[bool, str]:
        """
        In tài liệu Word (.docx, .doc) ngầm bằng COM Automation:
        - Ẩn hoàn toàn (Visible = False, DisplayAlerts = 0).
        - Thiết lập ActivePrinter, PageSetup (Orientation, PaperSize).
        - Áp dụng khoảng trang (Pages="1-3, 5") và Collate.
        - Đảm bảo giải phóng tài nguyên triệt để (không sót tiến trình WINWORD.EXE).
        """
        word = None
        doc = None
        try:
            pythoncom.CoInitialize()
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0  # wdAlertsNone = 0

            # Thiết lập máy in đích cho Word
            try:
                word.ActivePrinter = printer_name
            except Exception as pe:
                print(f"Cảnh báo đặt ActivePrinter Word: {pe}")

            # Mở tài liệu ở chế độ chỉ đọc, không thêm vào Recent
            abs_path = os.path.abspath(doc_path)
            doc = word.Documents.Open(
                FileName=abs_path,
                ReadOnly=True,
                AddToRecentFiles=False,
                Visible=False
            )

            # Áp dụng cài đặt trang PageSetup nếu có trong settings
            if settings:
                try:
                    # Orientation
                    orient = str(settings.get("orientation", "")).lower()
                    if "landscape" in orient:
                        doc.PageSetup.Orientation = 1  # wdOrientLandscape = 1
                    elif "portrait" in orient:
                        doc.PageSetup.Orientation = 0  # wdOrientPortrait = 0

                    # Paper Size
                    paper = str(settings.get("paper_size", "")).upper()
                    if "A4" in paper:
                        doc.PageSetup.PaperSize = 7  # wdPaperA4 = 7
                    elif "A3" in paper:
                        doc.PageSetup.PaperSize = 6  # wdPaperA3 = 6
                    elif "LETTER" in paper:
                        doc.PageSetup.PaperSize = 0  # wdPaperLetter = 0
                    elif "LEGAL" in paper:
                        doc.PageSetup.PaperSize = 4  # wdPaperLegal = 4
                except Exception as ex_ps:
                    print(f"Bỏ qua PageSetup Word: {ex_ps}")

            # Thiết lập tham số PrintOut
            range_str = settings.get("pages_range", "").strip() if settings else ""
            collate = True
            if settings:
                collate_str = str(settings.get("collated", "")).lower()
                if "uncollated" in collate_str:
                    collate = False

            if range_str and range_str.lower() != "all":
                # wdPrintRangeOfPages = 4
                doc.PrintOut(
                    Background=False,
                    Range=4,
                    Pages=range_str,
                    Copies=copies,
                    Collate=collate
                )
            else:
                # wdPrintAllDocument = 0
                doc.PrintOut(
                    Background=False,
                    Range=0,
                    Copies=copies,
                    Collate=collate
                )

            return True, f"In tài liệu Word thành công qua COM ({copies} bản)"

        except Exception as e:
            return False, f"Lỗi COM Word: {e}"
        finally:
            if doc is not None:
                try:
                    doc.Close(False)  # wdDoNotSaveChanges = 0
                except Exception:
                    pass
                del doc
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass
                del word
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    @classmethod
    def _print_excel_com(
        cls,
        xls_path: str,
        printer_name: str,
        copies: int = 1,
        settings: Optional[Dict] = None
    ) -> Tuple[bool, str]:
        """
        In bảng tính Excel (.xlsx, .xls) ngầm bằng COM Automation:
        - Ẩn hoàn toàn (Visible = False, DisplayAlerts = False).
        - Nhận dạng đúng cổng NeXX của máy in cho Excel ActivePrinter.
        - Áp dụng khoảng trang From/To và PageSetup.
        - Đảm bảo đóng workbook và giải phóng tiến trình EXCEL.EXE.
        """
        excel = None
        wb = None
        try:
            pythoncom.CoInitialize()
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            excel_printer = cls.get_excel_printer_string(printer_name)

            abs_path = os.path.abspath(xls_path)
            wb = excel.Workbooks.Open(Filename=abs_path, ReadOnly=True)

            # Cài đặt hướng giấy và khổ giấy trên các sheet
            if settings:
                try:
                    orient = str(settings.get("orientation", "")).lower()
                    paper = str(settings.get("paper_size", "")).upper()
                    for ws in wb.Worksheets:
                        if "landscape" in orient:
                            ws.PageSetup.Orientation = 2  # xlLandscape = 2
                        elif "portrait" in orient:
                            ws.PageSetup.Orientation = 1  # xlPortrait = 1

                        if "A4" in paper:
                            ws.PageSetup.PaperSize = 9  # xlPaperA4 = 9
                        elif "A3" in paper:
                            ws.PageSetup.PaperSize = 8  # xlPaperA3 = 8
                        elif "LETTER" in paper:
                            ws.PageSetup.PaperSize = 1  # xlPaperLetter = 1
                        elif "LEGAL" in paper:
                            ws.PageSetup.PaperSize = 5  # xlPaperLegal = 5
                except Exception as ex_ps:
                    print(f"Bỏ qua PageSetup Excel: {ex_ps}")

            # Xác định tham số trang From/To
            from_page = None
            to_page = None
            if settings:
                try:
                    pf = settings.get("page_from")
                    pt = settings.get("page_to")
                    if pf:
                        from_page = int(pf)
                    if pt:
                        to_page = int(pt)
                except Exception:
                    pass

            collate = True
            if settings:
                collate_str = str(settings.get("collated", "")).lower()
                if "uncollated" in collate_str:
                    collate = False

            print_args = {
                "Copies": copies,
                "Collate": collate,
                "ActivePrinter": excel_printer
            }
            if from_page is not None:
                print_args["From"] = from_page
            if to_page is not None:
                print_args["To"] = to_page

            # Chế độ in: Active Sheets hay Entire Workbook
            print_pages_setting = str(settings.get("print_pages", "")).lower() if settings else ""
            if "entire workbook" in print_pages_setting:
                wb.PrintOut(**print_args)
            else:
                # Mặc định in ActiveSheet
                wb.ActiveSheet.PrintOut(**print_args)

            return True, f"In bảng tính Excel thành công qua COM ({copies} bản)"

        except Exception as e:
            return False, f"Lỗi COM Excel: {e}"
        finally:
            if wb is not None:
                try:
                    wb.Close(False)
                except Exception:
                    pass
                del wb
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass
                del excel
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    # =========================================================================
    # PUBLIC UNIFIED PRINT DISPATCHER
    # =========================================================================

    @classmethod
    def print_file(
        cls,
        filepath: str,
        printer_name: str,
        copies: int = 1,
        settings: Optional[Dict] = None
    ) -> Tuple[bool, str]:
        """
        In một file cụ thể ra máy in được chỉ định với cơ chế ưu tiên:
        1. PDF: In ngầm GDI siêu tốc bằng pypdfium2 (DEVMODE Duplex, Orientation, PaperSize, Collate, Pages).
        2. Ảnh (PNG, JPG, JPEG, BMP, WEBP, TIFF): In GDI trực tiếp kèm DEVMODE.
        3. Word (DOCX, DOC): In ngầm qua Word COM Automation (Tự động đóng tiến trình, áp dụng PageSetup).
        4. Excel (XLSX, XLS): In ngầm qua Excel COM Automation.
        5. Fallback: ShellExecute / os.startfile kèm nạp DEVMODE và Default Printer.
        """
        if not os.path.exists(filepath):
            return False, f"File không tồn tại: {filepath}"

        abs_path = os.path.abspath(filepath)
        ext = os.path.splitext(abs_path)[1].lower()

        # 1. Định dạng PDF
        if ext == ".pdf":
            success, msg = cls._print_pdf_direct(abs_path, printer_name, copies=copies, settings=settings)
            if success:
                return True, msg
            print(f"Fallback PDF từ GDI sang ShellExecute do: {msg}")

        # 2. Định dạng Ảnh
        elif ext in [".png", ".jpg", ".jpeg", ".bmp", ".webp", ".tif", ".tiff"]:
            success, msg = cls._print_image_gdi(abs_path, printer_name, copies=copies, settings=settings)
            if success:
                return True, msg
            print(f"Fallback Image từ GDI sang ShellExecute do: {msg}")

        # 3. Định dạng Word
        elif ext in [".docx", ".doc"]:
            success, msg = cls._print_word_com(abs_path, printer_name, copies=copies, settings=settings)
            if success:
                return True, msg
            print(f"Fallback Word từ COM sang ShellExecute do: {msg}")

        # 4. Định dạng Excel
        elif ext in [".xlsx", ".xls"]:
            success, msg = cls._print_excel_com(abs_path, printer_name, copies=copies, settings=settings)
            if success:
                return True, msg
            print(f"Fallback Excel từ COM sang ShellExecute do: {msg}")

        # 5. Fallback: In qua ShellExecute / Windows Print API
        original_default = cls.get_default_printer()
        last_err = ""
        success = False

        for c in range(copies):
            # Cách 1: ShellExecute 'printto'
            try:
                ret = win32api.ShellExecute(
                    0,
                    "printto",
                    abs_path,
                    f'"{printer_name}"',
                    ".",
                    0
                )
                if int(ret) > 32:
                    success = True
                else:
                    last_err = f"Mã trả về từ ShellExecute: {ret}"
            except Exception as e:
                last_err = str(e)

            # Cách 2: Nếu printto không được hỗ trợ
            if not success:
                try:
                    if original_default != printer_name:
                        cls.set_default_printer(printer_name)
                        time.sleep(0.3)

                    try:
                        ret = win32api.ShellExecute(0, "print", abs_path, None, ".", 0)
                        if int(ret) > 32:
                            success = True
                    except Exception as e:
                        last_err = str(e)

                    if not success:
                        try:
                            os.startfile(abs_path, "print")
                            success = True
                        except Exception as e:
                            last_err = str(e)
                finally:
                    if original_default and original_default != printer_name:
                        try:
                            time.sleep(1.0)
                            cls.set_default_printer(original_default)
                        except Exception:
                            pass

            if not success:
                return False, f"Không thể gửi lệnh in (Bản {c+1}/{copies}): {last_err}"

            if c < copies - 1:
                time.sleep(1.5)

        return True, "Đã gửi lệnh in thành công đến máy in (Fallback)"

    # =========================================================================
    # TEST PAGE & PRINTER PROPERTIES
    # =========================================================================

    @classmethod
    def print_test_page(cls, printer_name: str) -> Tuple[bool, str]:
        """Tạo và in trang kiểm tra (Test Page) chi tiết."""
        if not WIN32_AVAILABLE:
            return False, "Thư viện Win32 không khả dụng"

        try:
            hDC = win32ui.CreateDC()
            hDC.CreatePrinterDC(printer_name)

            printable_w = hDC.GetDeviceCaps(win32con.HORZRES)
            printable_h = hDC.GetDeviceCaps(win32con.VERTRES)
            dpi_x = hDC.GetDeviceCaps(win32con.LOGPIXELSX)
            dpi_y = hDC.GetDeviceCaps(win32con.LOGPIXELSY)

            hDC.StartDoc("Trang In Kiem Tra - PyPrint Queue Manager")
            hDC.StartPage()

            # Thiết lập font tiêu đề lớn
            font_title = win32ui.CreateFont({
                "name": "Arial",
                "height": int(dpi_y * 0.4),
                "weight": 700,
            })
            hDC.SelectObject(font_title)

            y_cursor = int(dpi_y * 0.8)
            x_left = int(dpi_x * 0.8)

            hDC.TextOut(x_left, y_cursor, "TRANG IN KIỂM TRA MÁY IN")
            y_cursor += int(dpi_y * 0.5)

            # Font phụ đề
            font_sub = win32ui.CreateFont({
                "name": "Arial",
                "height": int(dpi_y * 0.22),
                "weight": 400,
            })
            hDC.SelectObject(font_sub)
            hDC.TextOut(x_left, y_cursor, "PyPrint Queue Manager - Windows Python Print Automation")
            y_cursor += int(dpi_y * 0.3)

            # Vẽ đường kẻ ngang
            hDC.MoveTo(x_left, y_cursor)
            hDC.LineTo(printable_w - x_left, y_cursor)
            y_cursor += int(dpi_y * 0.4)

            # Font nội dung
            font_body = win32ui.CreateFont({
                "name": "Consolas",
                "height": int(dpi_y * 0.2),
                "weight": 400,
            })
            hDC.SelectObject(font_body)

            import datetime
            now_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")

            lines = [
                f"Tên máy in       : {printer_name}",
                f"Thời gian in     : {now_str}",
                f"Độ phân giải DPI : {dpi_x} x {dpi_y} DPI",
                f"Khổ giấy in (px) : {printable_w} x {printable_h} pixels",
                f"Trạng thái       : Kết nối thành công và sẵn sàng in ấn",
                "",
                "Chúc mừng! Máy in của bạn đã kết nối và giao tiếp hoàn hảo",
                "với ứng dụng PyPrint Queue Manager.",
            ]

            for line in lines:
                hDC.TextOut(x_left, y_cursor, line)
                y_cursor += int(dpi_y * 0.28)

            # Khung viền kiểm tra căn chỉnh giấy
            border_margin = int(dpi_x * 0.4)
            hDC.MoveTo(border_margin, border_margin)
            hDC.LineTo(printable_w - border_margin, border_margin)
            hDC.LineTo(printable_w - border_margin, printable_h - border_margin)
            hDC.LineTo(border_margin, printable_h - border_margin)
            hDC.LineTo(border_margin, border_margin)

            hDC.EndPage()
            hDC.EndDoc()
            hDC.DeleteDC()
            return True, f"Đã gửi trang in kiểm tra tới '{printer_name}' thành công!"
        except Exception as e:
            return False, f"Lỗi in trang kiểm tra: {e}"

    @classmethod
    def open_printer_properties(cls, printer_name: str) -> Tuple[bool, str]:
        """
        Mở hộp thoại Printer Properties (Thuộc tính máy in) của Windows cho máy in chỉ định.
        """
        if not printer_name:
            return False, "Chưa chọn máy in."

        try:
            cmd = f'rundll32.exe printui.dll,PrintUIEntry /p /n "{printer_name}"'
            subprocess.Popen(cmd, shell=True)
            return True, f"Đã mở cửa sổ Printer Properties của '{printer_name}'."
        except Exception as e:
            try:
                cmd_pref = f'rundll32.exe printui.dll,PrintUIEntry /e /n "{printer_name}"'
                subprocess.Popen(cmd_pref, shell=True)
                return True, f"Đã mở Preferences của '{printer_name}'."
            except Exception as e2:
                return False, f"Không thể mở Printer Properties: {e2}"
