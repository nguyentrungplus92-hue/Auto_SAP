# SAP Auto Tasks - Django Project

Hệ thống tự động quét thư mục và nhập liệu vào SAP.

## Cấu trúc

```
sap_auto/
├── manage.py
├── requirements.txt
├── sap_auto/              # Django project config
│   ├── settings.py
│   ├── urls.py
│   └── wsgi.py
├── tasks/                 # App chính
│   ├── models.py          # TaskConfig + TaskLog
│   ├── views.py           # Dashboard + API
│   ├── urls.py            # Routing
│   ├── handlers.py        # ★ Mỗi task SAP = 1 hàm ở đây
│   ├── admin.py           # Django Admin
│   └── management/
│       └── commands/
│           └── scan_tasks.py   # ★ Background scanner
└── templates/
    ├── base.html
    └── tasks/
        ├── dashboard.html
        └── detail.html
```

## Cài đặt

```bash
# 1. Tạo môi trường ảo
python -m venv venv
venv\Scripts\activate

# 2. Cài thư viện
pip install -r requirements.txt

# 3. Tạo database
python manage.py migrate

# 4. Tạo admin user
python manage.py createsuperuser

# 5. Chạy web server
python manage.py runserver
```

## Cách sử dụng

### Bước 1: Chạy Web Server
```bash
python manage.py runserver 0.0.0.0:8000
```
Truy cập http://localhost:8000 để mở Dashboard.

### Bước 2: Chạy Background Scanner (cửa sổ CMD khác)
```bash
python manage.py scan_tasks
```
Scanner sẽ quét thư mục mỗi 30 giây và tự động chạy task khi phát hiện file mới.

### Bước 3: Thêm Task trên Web
1. Nhấn "Thêm Task"
2. Điền thông tin:
   - **Tên task**: Auto input Exchange Rate into SAP
   - **T-Code**: OB08
   - **Handler**: tasks.handlers.exchange_rate
   - **Đường dẫn**:
     ```
     G:\Accounting\ACS\04.EXCHANGE RATE\FY2026\02.2026\
     # G:\Accounting\ACS\04.EXCHANGE RATE\FYyyyy\mm.yyyy\
     ```
   - **Tên file**:
     ```
     Quotation 04-Feb-2026.xlsx
     # Quotation dd-mmm-yyyy.xlsx
     ```
   - **Regex**: `^Quotation \d{2}-\w{3}-\d{4}\.xlsx$`

### Bước 4: Thêm handler mới
Mở file `tasks/handlers.py` và thêm hàm:

```python
def my_new_task(filepath, session=None):
    """Mô tả task"""
    # Logic xử lý file + nhập SAP
    return {
        "status": "success",
        "message": "Đã xử lý xong",
        "rows": 100
    }
```

Sau đó trên web, tạo task với Handler = `tasks.handlers.my_new_task`

## Chạy cả 2 process cùng lúc (trên Windows)

Tạo file `start.bat`:
```bat
@echo off
echo Starting SAP Auto Tasks...

REM Activate venv
call venv\Scripts\activate

REM Start scanner in background
start "SAP Scanner" cmd /c "python manage.py scan_tasks"

REM Start web server
python manage.py runserver 0.0.0.0:8000
```

## Ghi chú
- Dòng bắt đầu bằng `#` trong đường dẫn/tên file = ghi chú định dạng
- Scanner tự bỏ qua file đã xử lý (kiểm tra trong TaskLog)
- Click trạng thái trên web để bật/tắt auto scan cho từng task
- Xem log chi tiết: click "📋 Log" trên mỗi task card
