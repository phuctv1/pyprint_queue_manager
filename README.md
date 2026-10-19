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
- **Bảng hiển thị trực quan**: Số thứ tự (STT), tên file, định dạng, dung lượng, số bản in, trạng thái hiện tại, chi tiết kết quả và đường dẫn đầy đủ.
- **Thao tác nhanh**: Click đúp vào dòng để mở xem trước file, chuột phải để mở thư mục chứa file trong Explorer.

### 3. 🔄 Tùy Biến Thứ Tự File Linh Hoạt
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

### 4. 🚀 Cơ Chế In Tuần Tự (Sequential Printing)
- **Chạy nền đa luồng (`threading.Thread`)**: Giao diện hoàn toàn mượt mà, không giật lag hoặc bị treo ("Not Responding") khi đang gửi lệnh in.
- **In lần lượt từng file**: Lệnh in gửi tuần tự theo đúng thứ tự đã sắp xếp.
- **Điều khiển toàn diện**:
  - `▶ Bắt đầu in`: Tiến hành gửi lệnh in tuần tự.
  - `⏸ Tạm dừng`: Dừng tạm thời hàng đợi.
  - `▶ Tiếp tục`: Tiếp tục in từ file đang dở dang.
  - `⏹ Hủy in`: Hủy bỏ tiến trình in còn lại.
- **Độ trễ thông minh (Spooler Delay)**: Cho phép điều chỉnh độ trễ giữa các file (0.5s - 10s) bằng thanh trượt, giúp máy in và Windows Print Spooler kịp nạp buffer an toàn trước khi nạp file tiếp theo.
- **Thanh tiến độ**: Hiển thị phần trăm và số lượng file đã xử lý theo thời gian thực (vd: `3/10 (30%)`).

### 5. 📝 Nhật Ký Hoạt Động Thời Gian Thực (Live Log)
- Ghi lại chi tiết từng thao tác: thời gian, tên file gửi in, kết quả từ máy in, mã lỗi (nếu có).
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
*(Bao gồm: `customtkinter`, `pywin32`, `pillow`, `windnd`)*

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
