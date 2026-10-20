# 🖨️ PyPrint Queue Manager

**PyPrint Queue Manager** là ứng dụng desktop viết bằng Python (sử dụng Tkinter / CustomTkinter) với giao diện người dùng hiện đại theo phong cách Windows 11. Chương trình hỗ trợ tự động phát hiện máy in kết nối với thiết bị, quản lý hàng đợi in tài liệu không giới hạn, cho phép linh hoạt sắp xếp vị trí file trong hàng đợi và thực hiện in tuần tự từng file an toàn, mượt mà.

---

## ✨ Tính Năng Nổi Bật

### 1. 🔌 Kết Nối & Quản Lý Máy In
- **Tự động quét**: Tự động nhận diện tất cả các máy in đã kết nối với máy tính (bao gồm máy in USB, máy in mạng LAN, Wi-Fi, máy in ảo PDF, OneNote...).
- **Nhận diện máy in mặc định**: Tự động đánh dấu `⭐ [Mặc định]` cho máy in mặc định của Windows.
- **Hiển thị thông tin chi tiết**: Trạng thái trực quan (🟢 Sẵn sàng, 🟡 Đang bận, 🔴 Ngoại tuyến), cổng kết nối (Port/IP), tên Driver và số lượng lệnh trong Windows Spooler.
- **Nút quét lại (🔄)**: Cập nhật danh sách máy in ngay tức thì khi cắm thêm máy in mới mà không cần khởi động lại app.
- **Đặt làm mặc định & In trang kiểm tra (📄)**: Hỗ trợ in trang test chuyên nghiệp để kiểm tra khả năng kết nối.

### 2. 📋 Hàng Đợi In Không Giới Hạn (Unlimited Queue)
- **Import đa dạng**:
  - Chọn nhiều file cùng lúc qua hộp thoại (Multi-select).
  - Chọn nguyên một thư mục chứa tài liệu (Add Folder).
  - **Kéo & Thả (Drag & Drop)** trực tiếp từ Windows Explorer vào cửa sổ ứng dụng.
- **Hỗ trợ hầu hết các định dạng phổ biến**:
  - Tài liệu: PDF (`.pdf`), Word (`.docx`, `.doc`), Excel (`.xlsx`, `.xls`), PowerPoint (`.pptx`, `.ppt`).
  - Tệp văn bản: `.txt`, `.log`, `.csv`, `.md`, `.json`...
  - Hình ảnh: `.png`, `.jpg`, `.jpeg`, `.bmp`, `.webp`, `.tiff`...
- **Bảng hiển thị trực quan**: Số thứ tự (STT), tên file, định dạng, số trang thực tế, xem trước & sửa nhanh, số bản in, trạng thái hiện tại, chi tiết kết quả và đường dẫn đầy đủ.
- **Đếm trước số trang & Ước tính lượng giấy A4**: Tự động quét số trang của tài liệu (PDF, Word, Excel, Ảnh, Text) ngay khi thêm file; tự động tính toán tổng số tờ giấy A4 cần nạp vào khay máy in theo số bản in và chế độ in 1 mặt/2 mặt.
- **Thao tác nhanh**: Click đúp vào dòng để mở xem trước file, chuột phải để mở thư mục chứa file trong Explorer.

### 3. 👁️ Vùng Xem Trước Đa Năng & Bảo Toàn Định Dạng Tuyệt Đối (Universal Preview)
- **Xem trước Word & Excel chân thực 100% (Không vỡ format)**:
  - Kết nối động cơ Microsoft Office Engine ở chế độ chỉ đọc (`ReadOnly=True`).
  - **Giữ nguyên 100% hiện trạng tệp gốc**: Tuyệt đối không can thiệp hay làm lệch lề, mất bảng biểu, sai font chữ hay thay đổi bố cục ban đầu.
  - Hiển thị chính xác từng trang in như khi mở trực tiếp trong Microsoft Word & Excel.
- **Trình xem PDF tương tác cao cấp**:
  - Render siêu nét qua nhân Chromium PDFium.
  - Chuyển trang linh hoạt: `⏮ Đầu`, `◀ Trước`, `Nhập số trang nhảy trực tiếp`, `▶ Sau`, `⏭ Cuối`.
  - Phóng to/thu nhỏ: `50% -> 300%`, `↔️ Vừa chiều ngang (Fit Width)`, `📄 Vừa toàn trang (Fit Page)`.
  - Hỗ trợ cuộn chuột (Mousewheel) và phím điều hướng mượt mà.
- **Trình xem & Chỉnh hướng Ảnh chuyên dụng (Image Rotate & Flip)**:
  - Hỗ trợ `.png`, `.jpg`, `.jpeg`, `.bmp`, `.webp`, `.tif`...
  - Xoay 90° theo chiều kim đồng hồ, xoay ngược 90°.
  - Lật ngang (Horizontal Flip), lật dọc (Vertical Flip).
  - Nút lưu trực tiếp ảnh đã xoay/lật vào file trên ổ đĩa.
- **Trình đọc & Sửa nhanh Văn bản (Text & Log Editor)**:
  - Xem và sửa nhanh `.txt`, `.csv`, `.log`, `.md`, `.json`, `.xml`, `.py`...
  - Font Monospace Consolas rõ ràng, hiển thị số dòng/từ/ký tự, công cụ tìm kiếm từ khóa và lưu tệp UTF-8 an toàn.

### 4. 🔄 Tùy Biến Thứ Tự File Linh Hoạt
- **Nút điều hướng**:
  - ⬆️ **Lên**: Đẩy file lên 1 bậc.
  - ⬇️ **Xuống**: Đẩy file xuống 1 bậc.
  - 🔝 **Đầu**: Đưa file lên vị trí đầu tiên để ưu tiên in ngay.
  - 🔚 **Cuối**: Chuyển file xuống cuối cùng.
- **Kéo & Thả dòng (Table Drag & Drop)**: Dùng chuột kéo trực tiếp dòng tài liệu đến vị trí mong muốn trong bảng.
- **Phím tắt nhanh**:
  - `Alt + Up`: Di chuyển lên.
  - `Alt + Down`: Di chuyển xuống.
  - `Delete`: Xóa file được chọn.
  - `Ctrl + O`: Mở hộp thoại thêm file.
  - `Ctrl + P`: Bắt đầu in.

### 4. 🚀 Cơ Chế In Tuần Tự & Động Cơ In Lõi Chuyên Nghiệp (Core Printing Engine)
- **Áp dụng Print Settings thực tế vào Windows DEVMODE**:
  - Hộp thoại Print Settings phong cách Office chuẩn mực: Khổ giấy (A4, A3, Letter, Legal...), Hướng giấy (Dọc/Ngang), In 2 mặt Duplex (Lật cạnh dài/ngắn hoặc 1 mặt), Soạn bộ (Collate), Số bản in (Copies).
  - Tất cả thông số được nạp tự động vào cấu trúc `DEVMODE` của Windows Driver trước khi nạp lệnh in.
- **In PDF ngầm 100% (Silent Direct PDF Printing)**:
  - Tích hợp bộ render `pypdfium2` (Chromium PDF Engine) chuyển đổi trực tiếp các trang PDF ra Windows GDI Device Context của máy in.
  - Tốc độ cực nhanh, không phụ thuộc vào Adobe Acrobat hay Microsoft Edge, không bao giờ xuất hiện cửa sổ hay popup làm gián đoạn người dùng.
  - Hỗ trợ chọn khoảng trang in tùy biến (ví dụ: `1-5, 8, 11-13`).
- **In tài liệu Office ngầm qua COM Automation (`win32com.client`)**:
  - Tự động điều khiển Word (`.docx`, `.doc`) và Excel (`.xlsx`, `.xls`) ở chế độ chạy ẩn hoàn toàn (`Visible = False`, `DisplayAlerts = 0`).
  - Tự động nhận diện cổng NeXX của máy in, thiết lập `PageSetup` (hướng giấy, khổ giấy, khoảng trang).
  - Tự động đóng tài liệu và hủy đối tượng COM trong khối `finally`, đảm bảo không để sót tiến trình rác (`WINWORD.EXE`, `EXCEL.EXE`) trong Task Manager.
- **Chạy nền đa luồng (`threading.Thread`)**: Giao diện hoàn toàn mượt mà, không giật lag hoặc bị treo ("Not Responding") khi đang gửi lệnh in.
- **Điều khiển toàn diện**:
  - `▶ Bắt đầu in`: Tiến hành gửi lệnh in tuần tự.
  - `⏸ Tạm dừng`: Dừng tạm thời hàng đợi.
  - `▶ Tiếp tục`: Tiếp tục in từ file đang dở dang.
  - `⏹ Hủy in`: Hủy bỏ tiến trình in còn lại.
- **Độ trễ thông minh (Spooler Delay)**: Cho phép điều chỉnh độ trễ giữa các file (0.5s - 10s) bằng thanh trượt, giúp máy in và Windows Print Spooler kịp nạp buffer an toàn trước khi nạp file tiếp theo.
- **Thanh tiến độ**: Hiển thị phần trăm và số lượng file đã xử lý theo thời gian thực (vd: `3/10 (30%)`).

### 5. 📝 Nhật Ký Hoạt Động Thời Gian Thực (Live Log)
- Ghi lại chi tiết từng thao tác: thời gian, tên file gửi in, cấu hình in áp dụng (khổ giấy, 2 mặt, số trang...), kết quả từ máy in, mã lỗi (nếu có).
- Mã màu trực quan: Xanh lá (thành công), Đỏ (lỗi), Vàng (cảnh báo), Xanh dương (thông tin).

---

## 🛠️ Cài Đặt & Sử Dụng

### 1. Yêu Cầu Hệ Thống
- Hệ điều hành: Windows 10 hoặc Windows 11.
- Python 3.8+ (Đã kiểm tra tương thích trên Python 3.14).

### 2. Cài Đặt Thư Viện Cần Thiết
Mở PowerShell hoặc Command Prompt tại thư mục dự án và chạy:
```bash
pip install -r requirements.txt
```
*(Bao gồm: `customtkinter`, `pywin32`, `pillow`, `windnd`, `python-docx`, `openpyxl`, `pypdfium2`)*

### 3. Khởi Động Ứng Dụng
- Cách 1: Nhấp đúp vào file **`run.bat`** để chạy trực tiếp.
- Cách 2: Chạy qua dòng lệnh:
```bash
python main.py
```

---

## 📁 Cấu Trúc Dự Án

```
pythona/
├── app_gui.py           # Giao diện chính CustomTkinter (Thiết kế Fluent Windows 11)
├── printer_service.py   # Dịch vụ quét máy in, kiểm tra spooler và engine in ấn
├── queue_manager.py     # Quản lý hàng đợi in, logic đổi vị trí, thread-safe
├── print_engine.py      # Bộ điều phối in tuần tự đa luồng (Start, Pause, Resume, Stop)
├── models.py            # Data model cho File Item và trạng thái in ấn
├── main.py              # File chạy chính (Kích hoạt High-DPI và UTF-8)
├── run.bat              # File chạy nhanh cho người dùng Windows
├── requirements.txt     # Danh sách các thư viện phụ thuộc
└── README.md            # Tài liệu hướng dẫn sử dụng tiếng Việt
```
