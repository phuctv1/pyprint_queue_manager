import os
import sys

# Đảm bảo mã hóa chuẩn UTF-8 trên Windows console
try:
    if sys.stdout:
        sys.stdout.reconfigure(encoding="utf-8")
    if sys.stderr:
        sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

# Kích hoạt tính năng High-DPI Awareness trên Windows để giao diện cực nét (không bị mờ font)
try:
    import ctypes
    # Per-monitor DPI aware
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

from app_gui import ModernPrintApp


def main():
    try:
        app = ModernPrintApp()
        app.mainloop()
    except Exception as e:
        import traceback
        traceback.print_exc()
        try:
            import tkinter.messagebox as mb
            mb.showerror("Lỗi Khởi Động", f"Không thể khởi động ứng dụng:\n\n{str(e)}")
        except Exception:
            pass


if __name__ == "__main__":
    main()
