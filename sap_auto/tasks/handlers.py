"""
SAP Task Handlers
=================
Mỗi task auto là 1 hàm riêng trong file này.
Hàm nhận (filepath, session) và trả về dict kết quả.

Khi thêm task mới:
1. Viết hàm mới ở đây
2. Tạo TaskConfig trên web, điền handler_module = "tasks.handlers.ten_ham"
"""
import os
import logging
import time
from django.utils import timezone
from .SAP_scripts.sap_base import SapGuiClient
from .SAP_scripts.tcode_ob08 import TCodeOB08

log = logging.getLogger(__name__)


def default_handler(filepath, session=None, task=None):
    """Handler mặc định - chỉ log, không làm gì"""
    log.info(f"[DEFAULT] File: {filepath}")
    return {
        "status": "success",
        "message": "Default handler - no action taken",
        "rows": 0
    }


def exchange_rate(filepath, session=None, task=None):
    """
    OB08 - Auto input Exchange Rate into SAP
    
    Tham số 1 (task.param1): USD/JPY-B20-D20-1, USD/VND-B30-D30-1000
    Format: CURRENCY_PAIR-CELL_NAME-CELL_VALUE-DIVISOR, ...
    """
    log.info(f"[OB08] Bắt đầu xử lý: {os.path.basename(filepath)}")
    start = time.time()

    try:
        # ===== 1. Đọc file Excel =====
        import openpyxl
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb.active

        # ===== 2. Parse tham số 1 =====
        rows_data = []
        today = timezone.localtime(timezone.now()).strftime("%d.%m.%Y")  # Ngày hôm nay theo múi giờ Django

        if task and task.param1:
            # param1 = "USD/JPY-B20-D20-1, USD/VND-B30-D30-1000"
            items = [item.strip() for item in task.param1.split(',')]
            
            for item in items:
                # item = "USD/JPY-B20-D20-1"
                parts = item.split('-')
                if len(parts) != 4:
                    log.warning(f"  Bỏ qua format không hợp lệ: {item}")
                    continue
                
                currency_pair = parts[0].strip()  # "USD/JPY"
                cell_name = parts[1].strip()       # "B20"
                cell_value = parts[2].strip()      # "D20"
                divisor = parts[3].strip()         # "1" hoặc "1000"

                # Lấy giá trị từ Excel
                name_value = ws[cell_name].value
                rate_value = ws[cell_value].value
                
                if rate_value is None:
                    log.warning(f"  Ô {cell_value} không có giá trị")
                    continue

                # Chia cho divisor
                try:
                    divisor_num = float(divisor)
                    if divisor_num != 0:
                        rate_value = f"{float(rate_value) / divisor_num:.3f}"
                except ValueError:
                    log.warning(f"  Divisor không hợp lệ: {divisor}")
                    continue

                # Parse currency pair
                if '/' in currency_pair:
                    currency_from, currency_to = currency_pair.split('/')
                else:
                    currency_from = currency_pair
                    currency_to = 'VND'
                
                rows_data.append({
                    'kurst': 'M',
                    'gdatu': today,
                    'kursm': str(rate_value),
                    'fcurr': currency_to,
                    'tcurr': currency_from,
                })
                
                log.info(f"  {currency_pair}: {cell_name}={name_value}, {cell_value}={rate_value}")

        if not rows_data:
            return {"status": "skipped", "message": "Không có dữ liệu hoặc param1 rỗng", "rows": 0}

        log.info(f"  Đọc được {len(rows_data)} tỷ giá")

        # ===== 3. Kiểm tra SAP User =====
        if not task or not task.sap_user:
            return {
                "status": "success",
                "message": f"Validated (no SAP user): {rows_data}",
                "rows": len(rows_data)
            }

        sap_client = task.sap_user.client
        sap_username = task.sap_user.username
        sap_password = task.sap_user.password
        log.info(f"  SAP User: {sap_client}/{sap_username}")

        # Đợi SAP GUI sẵn sàng
        time.sleep(2)

        # ===== 4. Nhập vào SAP qua GUI Scripting =====
        with SapGuiClient(
            # sap_entry_name=task.param2 or "V2Q",  
            sap_entry_name='V2Q',  
            client_no=sap_client,
            username=sap_username,
            password=sap_password,
        ) as sap:
            
            st = sap.last_login_status
            if sap.is_error(st) or ("name or password is incorrect" in (st.text or "").lower()):
                return {
                    "status": "error",
                    "message": f"SAP Login Error: {st.text}",
                    "rows": 0,
                    "duration": round(time.time() - start, 2)
                }
            
            log.info(f"  Đăng nhập SAP thành công")
            
            # Chạy OB08
            ob08 = TCodeOB08(sap)
            res = ob08.run(rows_data)
            log.info(f"  Kết quả OB08: {res}")

            duration = time.time() - start
            
            # Kiểm tra kết quả
            if res.get('ok'):
                # Thành công → Xóa file
                try:
                    os.remove(filepath)
                    log.info(f"  Đã xóa file: {filepath}")
                except Exception as del_err:
                    log.warning(f"  Không thể xóa file: {del_err}")
                    
                return {
                    "status": "success",
                    "message": res.get('message', f"Đã nhập {len(rows_data)} tỷ giá vào OB08"),
                    "rows": len(rows_data),
                    "duration": round(duration, 2)
                }
            else:
                # Lỗi → Giữ lại file
                return {
                    "status": "error",
                    "message": res.get('message', 'Lỗi không xác định từ OB08'),
                    "rows": 0,
                    "duration": round(duration, 2)
                }

    except Exception as e:
        log.error(f"  LỖI: {e}")
        return {
            "status": "error",
            "message": str(e),
            "rows": 0,
            "duration": round(time.time() - start, 2)
        }


def bom_upload(filepath, session=None, task=None):
    """
    CS01 - Auto upload BOM Master Data
    
    File: BOM_Upload_yyyymmdd.xlsx
    """
    log.info(f"[CS01] Bắt đầu xử lý: {os.path.basename(filepath)}")

    try:
        import openpyxl
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb.active

        row_count = ws.max_row - 1  # trừ header
        
        # TODO: Logic nhập BOM
        # ...

        return {
            "status": "success",
            "message": f"Đã upload {row_count} BOM items",
            "rows": row_count
        }

    except Exception as e:
        return {"status": "error", "message": str(e), "rows": 0}


def goods_receipt(filepath, session=None, task=None):
    """
    MIGO - Auto post Goods Receipt

    File: GR_ddmmmyyyy_batchNN.xlsx
    """
    log.info(f"[MIGO] Bắt đầu xử lý: {os.path.basename(filepath)}")

    try:
        import openpyxl
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb.active

        row_count = ws.max_row - 1

        # TODO: Logic post GR
        # ...

        return {
            "status": "success",
            "message": f"Đã post {row_count} GR items",
            "rows": row_count
        }

    except Exception as e:
        return {"status": "error", "message": str(e), "rows": 0}
