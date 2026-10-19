import threading
import time
from enum import Enum
from typing import Callable, Optional

from models import PrintItem, PrintStatus
from printer_service import PrinterService
from queue_manager import PrintQueueManager


class EngineState(Enum):
    IDLE = "IDLE"
    RUNNING = "RUNNING"
    PAUSED = "PAUSED"


class PrintEngine:
    def __init__(self, queue_manager: PrintQueueManager):
        self.queue_mgr = queue_manager
        self.state = EngineState.IDLE
        self.printer_name: str = ""
        self.delay_between_files: float = 2.0  # Giây

        self._thread: Optional[threading.Thread] = None
        self._stop_event = threading.Event()
        self._pause_event = threading.Event()
        self._pause_event.set()  # set = running (not paused)

        # Callbacks
        self.on_state_change: Optional[Callable[[EngineState], None]] = None
        self.on_log: Optional[Callable[[str, str], None]] = None  # (message, level: info/success/warning/error)
        self.on_progress: Optional[Callable[[int, int, float, str], None]] = None  # (current, total, percent, filename)
        self.on_finished: Optional[Callable[[dict], None]] = None

    def _set_state(self, new_state: EngineState):
        self.state = new_state
        if self.on_state_change:
            try:
                self.on_state_change(new_state)
            except Exception as e:
                print(f"Lỗi callback on_state_change: {e}")

    def log(self, message: str, level: str = "info"):
        if self.on_log:
            try:
                self.on_log(message, level)
            except Exception as e:
                print(f"Lỗi callback log: {e}")

    def start(self, printer_name: str, delay_seconds: float = 2.0, print_all_or_pending_only: bool = True):
        """Bắt đầu tiến trình in ấn tuần tự."""
        if self.state != EngineState.IDLE:
            return

        if not printer_name:
            self.log("LỖI: Chưa chọn máy in!", "error")
            return

        self.printer_name = printer_name
        self.delay_between_files = max(0.5, float(delay_seconds))
        self._stop_event.clear()
        self._pause_event.set()

        self._thread = threading.Thread(
            target=self._worker_loop,
            args=(print_all_or_pending_only,),
            daemon=True
        )
        self._set_state(EngineState.RUNNING)
        self._thread.start()

    def pause(self):
        """Tạm dừng in."""
        if self.state == EngineState.RUNNING:
            self._pause_event.clear()
            self._set_state(EngineState.PAUSED)
            self.log("⏸ Đã tạm dừng tiến trình in hàng đợi.", "warning")

    def resume(self):
        """Tiếp tục in."""
        if self.state == EngineState.PAUSED:
            self._pause_event.set()
            self._set_state(EngineState.RUNNING)
            self.log("▶ Tiếp tục tiến trình in hàng đợi...", "info")

    def stop(self):
        """Hủy toàn bộ tiến trình in."""
        if self.state != EngineState.IDLE:
            self._stop_event.set()
            self._pause_event.set()  # Unblock if paused
            self.log("⏹ Đang gửi lệnh dừng tiến trình in...", "warning")

    def _worker_loop(self, pending_only: bool):
        items = self.queue_mgr.get_items()
        if not items:
            self.log("Hàng đợi in đang trống!", "warning")
            self._set_state(EngineState.IDLE)
            return

        # Xác định danh sách các file cần in
        items_to_print = []
        for idx, item in enumerate(items):
            if pending_only:
                if item.status in [PrintStatus.PENDING, PrintStatus.FAILED]:
                    items_to_print.append((idx, item))
            else:
                items_to_print.append((idx, item))

        total_files = len(items_to_print)
        if total_files == 0:
            self.log("Tất cả các file trong hàng đợi đã được in xong. Nhấn 'Đặt lại trạng thái' nếu muốn in lại.", "info")
            self._set_state(EngineState.IDLE)
            return

        self.log(f"🚀 BẮT ĐẦU IN: {total_files} file đến máy in '{self.printer_name}'", "info")
        self.log(f"⏱ Độ trễ giữa các file: {self.delay_between_files:.1f}s", "info")

        success_count = 0
        failed_count = 0

        for current_idx, (orig_idx, item) in enumerate(items_to_print, start=1):
            # Kiểm tra nếu bị hủy
            if self._stop_event.is_set():
                self.log("Đã dừng tiến trình in theo yêu cầu người dùng.", "warning")
                break

            # Kiểm tra nếu đang tạm dừng
            while not self._pause_event.is_set():
                time.sleep(0.2)
                if self._stop_event.is_set():
                    break

            if self._stop_event.is_set():
                break

            percent = (current_idx - 1) / total_files
            if self.on_progress:
                try:
                    self.on_progress(current_idx, total_files, percent, item.filename)
                except Exception:
                    pass

            # Cập nhật trạng thái item thành Đang in
            self.queue_mgr.update_item_status(item.id, PrintStatus.PRINTING, "Đang gửi đến máy in...")
            self.log(f"[{current_idx}/{total_files}] Đang in file: {item.filename} (Bản in: {item.copies})...", "info")

            # Thực hiện lệnh in thực tế
            start_time = time.time()
            success, msg = PrinterService.print_file(
                filepath=item.filepath,
                printer_name=self.printer_name,
                copies=item.copies
            )
            elapsed = time.time() - start_time

            if success:
                success_count += 1
                item.mark_printed(elapsed)
                self.queue_mgr.update_item_status(item.id, PrintStatus.COMPLETED, f"✅ Hoàn thành ({elapsed:.1f}s)")
                self.log(f"✅ [{current_idx}/{total_files}] Đã in thành công: {item.filename} ({elapsed:.1f}s)", "success")
            else:
                failed_count += 1
                item.error_count += 1
                self.queue_mgr.update_item_status(item.id, PrintStatus.FAILED, f"❌ {msg}")
                self.log(f"❌ [{current_idx}/{total_files}] Lỗi khi in {item.filename}: {msg}", "error")

            # Cập nhật tiến độ sau khi file này xong
            percent_after = current_idx / total_files
            if self.on_progress:
                try:
                    self.on_progress(current_idx, total_files, percent_after, item.filename)
                except Exception:
                    pass

            # Chờ một khoảng trễ an toàn trước khi nạp file kế tiếp
            if current_idx < total_files and not self._stop_event.is_set():
                wait_step = 0.2
                waited = 0.0
                while waited < self.delay_between_files:
                    if self._stop_event.is_set():
                        break
                    time.sleep(wait_step)
                    waited += wait_step

        self._set_state(EngineState.IDLE)
        summary = {
            "total": total_files,
            "success": success_count,
            "failed": failed_count,
            "stopped_early": self._stop_event.is_set()
        }

        if self._stop_event.is_set():
            self.log(f"⏹ ĐÃ DỪNG TIẾN TRÌNH IN! Thành công: {success_count}/{total_files}, Thất bại: {failed_count}", "warning")
        else:
            self.log(f"🎉 HOÀN THÀNH IN ẤN! Thành công: {success_count}/{total_files}, Thất bại: {failed_count}", "success")

        if self.on_finished:
            try:
                self.on_finished(summary)
            except Exception:
                pass
