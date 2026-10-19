import os
import subprocess
import time
from typing import Dict, List, Optional, Tuple

try:
    import win32api
    import win32con
    import win32print
    import win32ui
    from PIL import Image, ImageWin
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False


class PrinterService:
    @staticmethod
    def is_available() -> bool:
        return WIN32_AVAILABLE

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

    @classmethod
    def print_file(cls, filepath: str, printer_name: str, copies: int = 1) -> Tuple[bool, str]:
        """
        In một file cụ thể ra máy in được chỉ định.
        Hỗ trợ số lượng bản in (copies).
        """
        if not os.path.exists(filepath):
            return False, f"File không tồn tại: {filepath}"

        abs_path = os.path.abspath(filepath)
        ext = os.path.splitext(abs_path)[1].lower()

        for c in range(copies):
            success = False
            last_err = ""

            # Cách 1: Đối với file ảnh (PNG, JPG, JPEG, BMP), in trực tiếp qua GDI để chất lượng tốt nhất
            if ext in [".png", ".jpg", ".jpeg", ".bmp", ".webp", ".tif", ".tiff"]:
                success, msg = cls._print_image_gdi(abs_path, printer_name)
                if success:
                    if c < copies - 1:
                        time.sleep(1.0)
                    continue
                else:
                    last_err = msg

            # Cách 2: In bằng ShellExecute 'printto'
            try:
                # Tham số lệnh printto: filepath, "printer_name"
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

            # Cách 3: Nếu printto không được hỗ trợ bởi ứng dụng mặc định
            # Tạm thời chuyển máy in mặc định sang printer_name và gọi 'print'
            if not success:
                original_default = cls.get_default_printer()
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
                        # Fallback tới os.startfile
                        try:
                            os.startfile(abs_path, "print")
                            success = True
                        except Exception as e:
                            last_err = str(e)

                finally:
                    # Trả lại máy in mặc định ban đầu sau khi lệnh được nạp vào spooler
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

        return True, "Đã gửi lệnh in thành công đến máy in"

    @classmethod
    def _print_image_gdi(cls, image_path: str, printer_name: str) -> Tuple[bool, str]:
        """In ảnh trực tiếp không cần mở cửa sổ phần mềm trung gian bằng GDI."""
        try:
            hDC = win32ui.CreateDC()
            hDC.CreatePrinterDC(printer_name)

            printable_w = hDC.GetDeviceCaps(win32con.HORZRES)
            printable_h = hDC.GetDeviceCaps(win32con.VERTRES)

            img = Image.open(image_path)
            if img.mode != "RGB":
                img = img.convert("RGB")

            # Tính toán tỉ lệ căn vừa trang in và có lề trang (5%)
            margin_x = int(printable_w * 0.05)
            margin_y = int(printable_h * 0.05)
            target_w = printable_w - (margin_x * 2)
            target_h = printable_h - (margin_y * 2)

            img_w, img_h = img.size
            ratio = min(target_w / img_w, target_h / img_h)

            final_w = int(img_w * ratio)
            final_h = int(img_h * ratio)

            # Căn giữa trang in
            pos_x = margin_x + (target_w - final_w) // 2
            pos_y = margin_y + (target_h - final_h) // 2

            hDC.StartDoc(f"Print Image: {os.path.basename(image_path)}")
            hDC.StartPage()

            dib = ImageWin.Dib(img)
            dib.draw(hDC.GetHandleOutput(), (pos_x, pos_y, pos_x + final_w, pos_y + final_h))

            hDC.EndPage()
            hDC.EndDoc()
            hDC.DeleteDC()
            return True, "In ảnh GDI thành công"
        except Exception as e:
            return False, f"Lỗi GDI: {e}"

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
            # Lệnh chuẩn Windows mở giao diện Printer Properties
            cmd = f'rundll32.exe printui.dll,PrintUIEntry /p /n "{printer_name}"'
            subprocess.Popen(cmd, shell=True)
            return True, f"Đã mở cửa sổ Printer Properties của '{printer_name}'."
        except Exception as e:
            # Thử phương án thay thế qua Printing Preferences (/e) nếu /p gặp lỗi
            try:
                cmd_pref = f'rundll32.exe printui.dll,PrintUIEntry /e /n "{printer_name}"'
                subprocess.Popen(cmd_pref, shell=True)
                return True, f"Đã mở Preferences của '{printer_name}'."
            except Exception as e2:
                return False, f"Không thể mở Printer Properties: {e2}"
