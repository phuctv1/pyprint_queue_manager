import os
import threading
from typing import Callable, List, Optional

from models import PrintItem, PrintStatus, Priority


class PrintQueueManager:
    def __init__(self):
        self._items: List[PrintItem] = []
        self._lock = threading.RLock()
        self._on_change_callbacks: List[Callable[[], None]] = []

    def subscribe(self, callback: Callable[[], None]):
        """Đăng ký hàm lắng nghe sự thay đổi của hàng đợi để cập nhật UI."""
        if callback not in self._on_change_callbacks:
            self._on_change_callbacks.append(callback)

    def _notify(self):
        for cb in self._on_change_callbacks:
            try:
                cb()
            except Exception as e:
                print(f"Lỗi callback notify: {e}")

    def add_file(self, filepath: str, copies: int = 1) -> Optional[PrintItem]:
        """Thêm 1 file vào hàng đợi. Bỏ qua nếu đã trùng đường dẫn."""
        if not os.path.isfile(filepath):
            return None
        abs_path = os.path.abspath(filepath)
        with self._lock:
            # Kiểm tra trùng lặp
            for existing in self._items:
                if os.path.abspath(existing.filepath) == abs_path:
                    return None  # Đã có trong hàng đợi, bỏ qua
            item = PrintItem(filepath=filepath, copies=copies)
            self._items.append(item)
            self._notify()
            return item

    def add_files(self, filepaths: List[str], copies: int = 1) -> List[PrintItem]:
        """Thêm danh sách file, tự động bỏ qua file trùng đường dẫn."""
        added = []
        skipped_duplicates = 0
        with self._lock:
            existing_paths = {os.path.abspath(it.filepath) for it in self._items}
            for fp in filepaths:
                if not os.path.isfile(fp):
                    continue
                abs_fp = os.path.abspath(fp)
                if abs_fp in existing_paths:
                    skipped_duplicates += 1
                    continue
                item = PrintItem(filepath=fp, copies=copies)
                self._items.append(item)
                existing_paths.add(abs_fp)
                added.append(item)
            if added:
                self._notify()
        return added

    def get_skipped_duplicate_count(self, filepaths: List[str]) -> int:
        """Đếm số file sẽ bị bỏ qua do trùng khi thêm."""
        with self._lock:
            existing_paths = {os.path.abspath(it.filepath) for it in self._items}
        count = 0
        for fp in filepaths:
            if os.path.abspath(fp) in existing_paths:
                count += 1
        return count

    def add_folder(self, folder_path: str, recursive: bool = False, copies: int = 1) -> List[PrintItem]:
        """Thêm tất cả file từ thư mục, bỏ qua trùng."""
        if not os.path.isdir(folder_path):
            return []
        filepaths = []
        if recursive:
            for root, _, files in os.walk(folder_path):
                for f in files:
                    filepaths.append(os.path.join(root, f))
        else:
            for f in os.listdir(folder_path):
                fp = os.path.join(folder_path, f)
                if os.path.isfile(fp):
                    filepaths.append(fp)
        return self.add_files(filepaths, copies=copies)

    def get_items(self) -> List[PrintItem]:
        with self._lock:
            return list(self._items)

    def get_item(self, index: int) -> Optional[PrintItem]:
        with self._lock:
            if 0 <= index < len(self._items):
                return self._items[index]
            return None

    def remove_item(self, index: int) -> bool:
        with self._lock:
            if 0 <= index < len(self._items):
                self._items.pop(index)
                self._notify()
                return True
            return False

    def remove_items_by_indices(self, indices: List[int]) -> int:
        with self._lock:
            indices_sorted = sorted(set(indices), reverse=True)
            removed_count = 0
            for idx in indices_sorted:
                if 0 <= idx < len(self._items):
                    self._items.pop(idx)
                    removed_count += 1
            if removed_count > 0:
                self._notify()
            return removed_count

    def move_up(self, index: int) -> int:
        """Di chuyển file lên 1 vị trí. Trả về vị trí mới."""
        with self._lock:
            if index > 0 and index < len(self._items):
                self._items[index - 1], self._items[index] = self._items[index], self._items[index - 1]
                self._notify()
                return index - 1
            return index

    def move_down(self, index: int) -> int:
        """Di chuyển file xuống 1 vị trí. Trả về vị trí mới."""
        with self._lock:
            if index >= 0 and index < len(self._items) - 1:
                self._items[index], self._items[index + 1] = self._items[index + 1], self._items[index]
                self._notify()
                return index + 1
            return index

    def move_to_top(self, index: int) -> int:
        """Đưa file lên đầu hàng đợi."""
        with self._lock:
            if 0 < index < len(self._items):
                item = self._items.pop(index)
                self._items.insert(0, item)
                self._notify()
                return 0
            return index

    def move_to_bottom(self, index: int) -> int:
        """Đưa file xuống cuối hàng đợi."""
        with self._lock:
            if 0 <= index < len(self._items) - 1:
                item = self._items.pop(index)
                self._items.append(item)
                self._notify()
                return len(self._items) - 1
            return index

    def move_item(self, from_index: int, to_index: int) -> bool:
        """Di chuyển item từ vị trí này sang vị trí khác (kéo thả)."""
        with self._lock:
            if 0 <= from_index < len(self._items) and 0 <= to_index < len(self._items):
                item = self._items.pop(from_index)
                self._items.insert(to_index, item)
                self._notify()
                return True
            return False

    def clear_all(self):
        """Xóa toàn bộ hàng đợi."""
        with self._lock:
            self._items.clear()
            self._notify()

    def clear_completed(self):
        """Xóa các file đã in xong thành công."""
        with self._lock:
            self._items = [item for item in self._items if item.status != PrintStatus.COMPLETED]
            self._notify()

    def clear_failed(self):
        """Xóa các file bị lỗi."""
        with self._lock:
            self._items = [item for item in self._items if item.status != PrintStatus.FAILED]
            self._notify()

    def reset_statuses(self):
        """Đặt lại tất cả trạng thái về Chờ in (để in lại từ đầu)."""
        with self._lock:
            for item in self._items:
                item.status = PrintStatus.PENDING
                item.status_detail = ""
                item.elapsed_seconds = 0.0
            self._notify()

    def reset_failed_statuses(self):
        """Chỉ đặt lại các file lỗi về Chờ in."""
        with self._lock:
            changed = False
            for item in self._items:
                if item.status == PrintStatus.FAILED:
                    item.status = PrintStatus.PENDING
                    item.status_detail = ""
                    item.error_count = 0
                    changed = True
            if changed:
                self._notify()

    def set_item_copies(self, index: int, copies: int) -> bool:
        """Cập nhật số bản in của file tại vị trí index."""
        with self._lock:
            if 0 <= index < len(self._items):
                self._items[index].copies = max(1, int(copies))
                self._notify()
                return True
            return False

    def set_item_priority(self, index: int, priority: Priority) -> bool:
        """Đặt độ ưu tiên cho file tại vị trí index."""
        with self._lock:
            if 0 <= index < len(self._items):
                self._items[index].priority = priority
                self._notify()
                return True
            return False

    def update_item_settings(self, index: int, settings: dict) -> bool:
        """Cập nhật cài đặt in ấn (Print Settings) cho file."""
        with self._lock:
            if 0 <= index < len(self._items):
                self._items[index].print_settings.update(settings)
                self._notify()
                return True
            return False

    def refresh_item(self, index: int):
        """Làm mới thông tin dung lượng file sau khi chỉnh sửa."""
        with self._lock:
            if 0 <= index < len(self._items):
                self._items[index].refresh_file_info()
                self._notify()

    def update_item_status(self, item_id: str, status: PrintStatus, detail: str = ""):
        with self._lock:
            for item in self._items:
                if item.id == item_id:
                    item.status = status
                    item.status_detail = detail
                    self._notify()
                    break

    def get_stats(self) -> dict:
        with self._lock:
            total_count = len(self._items)
            total_size_bytes = sum(item.file_size_bytes for item in self._items)
            pending_count = sum(1 for item in self._items if item.status == PrintStatus.PENDING)
            printing_count = sum(1 for item in self._items if item.status == PrintStatus.PRINTING)
            completed_count = sum(1 for item in self._items if item.status == PrintStatus.COMPLETED)
            failed_count = sum(1 for item in self._items if item.status == PrintStatus.FAILED)
            total_copies = sum(item.copies for item in self._items)
            high_priority_count = sum(
                1 for item in self._items
                if item.priority in (Priority.HIGH, Priority.URGENT)
                and item.status == PrintStatus.PENDING
            )

            return {
                "total": total_count,
                "total_size_bytes": total_size_bytes,
                "total_size_str": PrintItem._format_size(total_size_bytes),
                "pending": pending_count,
                "printing": printing_count,
                "completed": completed_count,
                "failed": failed_count,
                "total_copies": total_copies,
                "high_priority": high_priority_count,
            }

    def has_duplicates_in(self, filepaths: List[str]) -> List[str]:
        """Trả về danh sách path trùng với hàng đợi hiện tại."""
        with self._lock:
            existing = {os.path.abspath(it.filepath) for it in self._items}
        return [fp for fp in filepaths if os.path.abspath(fp) in existing]

    def sort_by_priority(self):
        """Sắp xếp hàng đợi theo độ ưu tiên (URGENT > HIGH > NORMAL), giữ nguyên thứ tự trong cùng nhóm."""
        with self._lock:
            self._items.sort(
                key=lambda it: (-(it.priority.value), 0 if it.status == PrintStatus.PENDING else 1)
            )
            self._notify()

    def sort_by_name(self, ascending: bool = True):
        """Sắp xếp hàng đợi theo tên file."""
        with self._lock:
            self._items.sort(key=lambda it: it.filename.lower(), reverse=not ascending)
            self._notify()

    def sort_by_size(self, ascending: bool = True):
        """Sắp xếp hàng đợi theo dung lượng file."""
        with self._lock:
            self._items.sort(key=lambda it: it.file_size_bytes, reverse=not ascending)
            self._notify()

    def sort_by_type(self):
        """Sắp xếp hàng đợi theo loại file."""
        with self._lock:
            self._items.sort(key=lambda it: it.file_type.lower())
            self._notify()
