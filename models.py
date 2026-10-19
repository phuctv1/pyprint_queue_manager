import datetime
import os
import uuid
from dataclasses import dataclass, field
from enum import Enum
from typing import Optional


class PrintStatus(Enum):
    PENDING = "Chờ in"
    PRINTING = "Đang in..."
    COMPLETED = "Hoàn thành"
    FAILED = "Lỗi"
    SKIPPED = "Đã bỏ qua"


class Priority(Enum):
    NORMAL = 0
    HIGH = 1
    URGENT = 2


@dataclass
class PrintItem:
    filepath: str
    id: str = field(default_factory=lambda: str(uuid.uuid4()))
    filename: str = ""
    file_size_bytes: int = 0
    file_size_str: str = ""
    file_type: str = ""
    status: PrintStatus = PrintStatus.PENDING
    status_detail: str = ""
    copies: int = 1
    print_settings: dict = field(default_factory=dict)
    priority: Priority = Priority.NORMAL
    added_at: datetime.datetime = field(default_factory=datetime.datetime.now)
    printed_at: Optional[datetime.datetime] = None
    elapsed_seconds: float = 0.0
    error_count: int = 0
    notes: str = ""

    def refresh_file_info(self):
        if os.path.exists(self.filepath):
            try:
                self.file_size_bytes = os.path.getsize(self.filepath)
                self.file_size_str = self._format_size(self.file_size_bytes)
            except Exception:
                pass

    def __post_init__(self):
        if not self.filename:
            self.filename = os.path.basename(self.filepath)
        if not self.file_type:
            ext = os.path.splitext(self.filepath)[1].lower()
            self.file_type = ext if ext else ".?"
        if os.path.exists(self.filepath) and self.file_size_bytes == 0:
            try:
                self.file_size_bytes = os.path.getsize(self.filepath)
                self.file_size_str = self._format_size(self.file_size_bytes)
            except Exception:
                self.file_size_str = "0 B"
        elif not self.file_size_str:
            self.file_size_str = self._format_size(self.file_size_bytes)

    @staticmethod
    def _format_size(size_bytes: int) -> str:
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.1f} KB"
        elif size_bytes < 1024 * 1024 * 1024:
            return f"{size_bytes / (1024 * 1024):.1f} MB"
        else:
            return f"{size_bytes / (1024 * 1024 * 1024):.2f} GB"

    def get_status_badge(self) -> str:
        badge_map = {
            PrintStatus.PENDING: "⏳ Chờ in",
            PrintStatus.PRINTING: "🖨️ Đang in...",
            PrintStatus.COMPLETED: "✅ Hoàn thành",
            PrintStatus.FAILED: "❌ Lỗi",
            PrintStatus.SKIPPED: "⏸️ Bỏ qua",
        }
        return badge_map.get(self.status, self.status.value)

    def get_priority_badge(self) -> str:
        return {
            Priority.NORMAL: "",
            Priority.HIGH: "🔥",
            Priority.URGENT: "🚨",
        }.get(self.priority, "")

    def get_type_icon(self) -> str:
        icon_map = {
            ".pdf": "📕",
            ".docx": "📝",
            ".doc": "📝",
            ".xlsx": "📊",
            ".xls": "📊",
            ".pptx": "📋",
            ".ppt": "📋",
            ".txt": "📄",
            ".csv": "📊",
            ".log": "📄",
            ".png": "🖼️",
            ".jpg": "🖼️",
            ".jpeg": "🖼️",
            ".bmp": "🖼️",
            ".webp": "🖼️",
        }
        return icon_map.get(self.file_type.lower(), "📄")

    def get_added_time_str(self) -> str:
        return self.added_at.strftime("%H:%M:%S")

    def is_editable(self) -> bool:
        """Kiểm tra xem file có hỗ trợ chỉnh sửa trực tiếp không."""
        return self.file_type.lower() in [".docx", ".doc", ".xlsx", ".xls"]

    def mark_printed(self, elapsed: float):
        self.printed_at = datetime.datetime.now()
        self.elapsed_seconds = elapsed

    def get_elapsed_str(self) -> str:
        if self.elapsed_seconds > 0:
            return f"{self.elapsed_seconds:.1f}s"
        return ""

    def get_copies_display(self) -> str:
        return f"{self.copies} bản"

    def get_display_name(self, max_len: int = 40) -> str:
        """Trả về tên file rút gọn nếu quá dài."""
        if len(self.filename) <= max_len:
            return self.filename
        name, ext = os.path.splitext(self.filename)
        keep = max_len - len(ext) - 3
        return name[:keep] + "..." + ext if keep > 0 else self.filename[:max_len]
