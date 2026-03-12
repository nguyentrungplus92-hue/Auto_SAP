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
from .SAP_scripts.tcode_md12 import TCodeMD12
from .SAP_scripts.tcode_mm02 import TCodeMM02
from .SAP_scripts.tcode_me22 import TCodeME22
from .SAP_scripts.tcode_me12 import TCodeME12
from .SAP_scripts.tcode_mk01 import TCodeMK01


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
    


def md12_unfix_order(filepath, session=None, task=None):
    """
    MD12 - Bỏ tick "Firmly planned" cho Planned Orders
    
    File CSV/Excel: Cột A chứa danh sách Planned Order numbers (không cần header)
    """
    log.info(f"[MD12] Bắt đầu xử lý: {os.path.basename(filepath)}")
    start = time.time()

    try:
        # ===== 1. Đọc file =====
        planned_orders = []
        
        # Kiểm tra định dạng file
        if filepath.lower().endswith('.csv'):
            import csv
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                for row in reader:
                    if row and row[0].strip():
                        value = row[0].strip()
                        # Chỉ lấy nếu là số
                        if value.isdigit():
                            planned_orders.append(value)
        else:
            # Excel file
            import openpyxl
            wb = openpyxl.load_workbook(filepath, data_only=True)
            ws = wb.active
            for row in ws.iter_rows(min_col=1, max_col=1):
                cell = row[0]
                if cell.value:
                    value = str(cell.value).strip()
                    if value.isdigit():
                        planned_orders.append(value)
        
        if not planned_orders:
            return {"status": "skipped", "message": "Không có dữ liệu trong file", "rows": 0}

        log.info(f"  Đọc được {len(planned_orders)} planned orders")

        # ===== 2. Kiểm tra SAP User =====
        if not task or not task.sap_user:
            return {
                "status": "success",
                "message": f"Validated (no SAP user): {len(planned_orders)} orders",
                "rows": len(planned_orders)
            }

        sap_client = task.sap_user.client
        sap_username = task.sap_user.username
        sap_password = task.sap_user.password
        log.info(f"  SAP User: {sap_client}/{sap_username}")

        # Đợi SAP GUI sẵn sàng
        time.sleep(2)

        # ===== 3. Xử lý trong SAP =====
        with SapGuiClient(
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
            
            # Chạy MD12
            md12 = TCodeMD12(sap)
            res = md12.run(planned_orders)
            log.info(f"  Kết quả MD12: {res}")

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
                    "message": res.get('message', f"Đã xử lý {res.get('processed', 0)} planned orders"),
                    "rows": res.get('processed', 0),
                    "duration": round(duration, 2)
                }
            else:
                # Lỗi → Giữ lại file
                return {
                    "status": "error",
                    "message": res.get('message', 'Lỗi không xác định từ MD12'),
                    "rows": res.get('processed', 0),
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
    


def mm02_update_vietnam_name(filepath, session=None, task=None):
    """
    MM02 - Cập nhật mô tả tiếng Việt cho Material
    
    File Excel: 
    - Cột A: Material code
    - Cột B: Mô tả tiếng Việt (Long Text)
    - Không có header, dữ liệu bắt đầu từ dòng 1
    """
    log.info(f"[MM02] Bắt đầu xử lý: {os.path.basename(filepath)}")
    start = time.time()

    try:
        # ===== 1. Đọc file Excel =====
        import openpyxl
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb.active
        
        materials = []
        for row in ws.iter_rows(min_row=1):  # Đọc từ dòng 1 (không có header)
            material = row[0].value  # Cột A
            description = row[1].value if len(row) > 1 else None  # Cột B
            
            if material and description:
                materials.append({
                    'material': str(material).strip(),
                    'description': str(description).strip()
                })
        
        wb.close()
        
        if not materials:
            return {"status": "skipped", "message": "Không có dữ liệu trong file", "rows": 0}

        log.info(f"  Đọc được {len(materials)} materials")

        # ===== 2. Kiểm tra SAP User =====
        if not task or not task.sap_user:
            return {
                "status": "success",
                "message": f"Validated (no SAP user): {len(materials)} materials",
                "rows": len(materials)
            }

        sap_client = task.sap_user.client
        sap_username = task.sap_user.username
        sap_password = task.sap_user.password
        log.info(f"  SAP User: {sap_client}/{sap_username}")

        time.sleep(2)

        # ===== 3. Xử lý trong SAP =====
        with SapGuiClient(
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
            
            mm02 = TCodeMM02(sap)
            res = mm02.run(materials)
            log.info(f"  Kết quả MM02: {res}")

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
                    "message": res.get('message', f"Đã xử lý {res.get('processed', 0)} planned orders"),
                    "rows": res.get('processed', 0),
                    "duration": round(duration, 2)
                }
            else:
                # Lỗi → Giữ lại file
                return {
                    "status": "error",
                    "message": res.get('message', 'Lỗi không xác định từ MM02'),
                    "rows": res.get('processed', 0),
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
    

def mm02_update_dv_tinh(filepath, session=None, task=None):
    """
    MM02 - Cập nhật đơn vị tính (Tab ZU07 - Internal comment)
    
    File Excel: 
    - Cột A: Material code
    - Cột B: Đơn vị tính
    - Không có header, dữ liệu bắt đầu từ dòng 1
    """
    log.info(f"[MM02-ZU07] Bắt đầu xử lý: {os.path.basename(filepath)}")
    start = time.time()

    try:
        # ===== 1. Đọc file Excel =====
        import openpyxl
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb.active
        
        materials = []
        for row in ws.iter_rows(min_row=1):
            material = row[0].value
            description = row[1].value if len(row) > 1 else None
            
            if material and description:
                materials.append({
                    'material': str(material).strip(),
                    'description': str(description).strip()
                })
        
        wb.close()
        
        if not materials:
            return {"status": "skipped", "message": "Không có dữ liệu trong file", "rows": 0}

        log.info(f"  Đọc được {len(materials)} materials")

        # ===== 2. Kiểm tra SAP User =====
        if not task or not task.sap_user:
            return {
                "status": "success",
                "message": f"Validated (no SAP user): {len(materials)} materials",
                "rows": len(materials)
            }

        sap_client = task.sap_user.client
        sap_username = task.sap_user.username
        sap_password = task.sap_user.password
        log.info(f"  SAP User: {sap_client}/{sap_username}")

        time.sleep(2)

        # ===== 3. Xử lý trong SAP =====
        with SapGuiClient(
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
            
            # Chạy MM02 với Tab ZU07
            mm02 = TCodeMM02(sap)
            res = mm02.run_zu07(materials)
            log.info(f"  Kết quả MM02-ZU07: {res}")

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
                    "message": res.get('message', f"Đã xử lý {res.get('processed', 0)} planned orders"),
                    "rows": res.get('processed', 0),
                    "duration": round(duration, 2)
                }
            else:
                # Lỗi → Giữ lại file
                return {
                    "status": "error",
                    "message": res.get('message', 'Lỗi không xác định từ MM02'),
                    "rows": res.get('processed', 0),
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


def me22_change_po_quantity(filepath, session=None, task=None):
    """
    ME22 - Change PO Quantity
    
    File Excel/CSV:
    - Cột A: PO Number (số PO)
    - Cột B: Item Number (số item trong PO)
    - Cột C: Quantity (số lượng mới)
    - Không có header, dữ liệu bắt đầu từ dòng 1
    """
    log.info(f"[ME22] Bắt đầu xử lý: {os.path.basename(filepath)}")
    start = time.time()

    try:
        # ===== 1. Đọc file =====
        po_items = []
        
        if filepath.lower().endswith('.csv'):
            # Đọc CSV
            import csv
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                for row in reader:
                    if len(row) >= 3 and row[0].strip():
                        po_items.append({
                            'po_number': row[0].strip(),
                            'item_number': row[1].strip(),
                            'quantity': row[2].strip()
                        })
        else:
            # Đọc Excel
            import openpyxl
            wb = openpyxl.load_workbook(filepath, data_only=True)
            ws = wb.active
            
            for row in ws.iter_rows(min_row=1):
                po_number = row[0].value if len(row) > 0 else None
                item_number = row[1].value if len(row) > 1 else None
                quantity = row[2].value if len(row) > 2 else None
                
                if po_number and item_number and quantity:
                    po_items.append({
                        'po_number': str(po_number).strip(),
                        'item_number': str(item_number).strip(),
                        'quantity': str(quantity).strip()
                    })
            
            wb.close()
        
        if not po_items:
            return {"status": "skipped", "message": "Không có dữ liệu trong file", "rows": 0}

        log.info(f"  Đọc được {len(po_items)} PO items")

        # ===== 2. Kiểm tra SAP User =====
        if not task or not task.sap_user:
            return {
                "status": "success",
                "message": f"Validated (no SAP user): {len(po_items)} items",
                "rows": len(po_items)
            }

        sap_client = task.sap_user.client
        sap_username = task.sap_user.username
        sap_password = task.sap_user.password
        log.info(f"  SAP User: {sap_client}/{sap_username}")

        time.sleep(2)

        # ===== 3. Xử lý trong SAP =====
        with SapGuiClient(
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
            
            # Chạy ME22
            me22 = TCodeME22(sap)
            res = me22.run(po_items)
            log.info(f"  Kết quả ME22: {res}")

            duration = time.time() - start
            
            if res.get('ok'):
                # Thành công → Xóa file
                try:
                    os.remove(filepath)
                    log.info(f"  Đã xóa file: {filepath}")
                except Exception as del_err:
                    log.warning(f"  Không thể xóa file: {del_err}")
                    
                return {
                    "status": "success",
                    "message": res.get('message', f"Đã cập nhật {res.get('processed', 0)} PO items"),
                    "rows": res.get('processed', 0),
                    "duration": round(duration, 2)
                }
            else:
                return {
                    "status": "error",
                    "message": res.get('message', 'Lỗi không xác định từ ME22'),
                    "rows": res.get('processed', 0),
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


def me12_update_marker(filepath, session=None, task=None):
    """
    ME12 - Update Marker for Info Record
    
    File Excel/CSV:
    - Cột A: Vendor (mã vendor)
    - Cột B: Material (mã material)
    - Cột C: Marker (mã marker) - CÓ THỂ TRỐNG
    - Không có header, dữ liệu bắt đầu từ dòng 1
    """
    log.info(f"[ME12] Bắt đầu xử lý: {os.path.basename(filepath)}")
    start = time.time()

    try:
        # ===== 1. Đọc file =====
        info_records = []
        
        if filepath.lower().endswith('.csv'):
            # Đọc CSV
            import csv
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                for row in reader:
                    # Chỉ cần 2 cột đầu (vendor, material), cột 3 (marker) có thể trống
                    if len(row) >= 2 and row[0].strip() and row[1].strip():
                        info_records.append({
                            'vendor': row[0].strip(),
                            'material': row[1].strip(),
                            'marker': row[2].strip() if len(row) > 2 else ''
                        })
        else:
            # Đọc Excel
            import openpyxl
            wb = openpyxl.load_workbook(filepath, data_only=True)
            ws = wb.active
            
            for row in ws.iter_rows(min_row=1):
                vendor = row[0].value if len(row) > 0 else None
                material = row[1].value if len(row) > 1 else None
                marker = row[2].value if len(row) > 2 else None
                
                # Chỉ cần vendor và material, marker có thể trống/None
                if vendor and material:
                    info_records.append({
                        'vendor': str(vendor).strip(),
                        'material': str(material).strip(),
                        'marker': str(marker).strip() if marker else ''
                    })
            
            wb.close()
        
        if not info_records:
            return {"status": "skipped", "message": "Không có dữ liệu trong file", "rows": 0}

        log.info(f"  Đọc được {len(info_records)} info records")

        # ===== 2. Kiểm tra SAP User =====
        if not task or not task.sap_user:
            return {
                "status": "success",
                "message": f"Validated (no SAP user): {len(info_records)} records",
                "rows": len(info_records)
            }

        sap_client = task.sap_user.client
        sap_username = task.sap_user.username
        sap_password = task.sap_user.password
        log.info(f"  SAP User: {sap_client}/{sap_username}")

        time.sleep(2)

        # ===== 3. Xử lý trong SAP =====
        with SapGuiClient(
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
            
            # Chạy ME12
            me12 = TCodeME12(sap)
            res = me12.run(info_records)
            log.info(f"  Kết quả ME12: {res}")

            duration = time.time() - start
            
            if res.get('ok'):
                # Thành công → Xóa file
                try:
                    os.remove(filepath)
                    log.info(f"  Đã xóa file: {filepath}")
                except Exception as del_err:
                    log.warning(f"  Không thể xóa file: {del_err}")
                    
                return {
                    "status": "success",
                    "message": res.get('message', f"Đã cập nhật {res.get('processed', 0)} info records"),
                    "rows": res.get('processed', 0),
                    "duration": round(duration, 2)
                }
            else:
                return {
                    "status": "error",
                    "message": res.get('message', 'Lỗi không xác định từ ME12'),
                    "rows": res.get('processed', 0),
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


def me12_update_shipping_instruction(filepath, session=None, task=None):
    """
    ME12 - Update Shipping Instruction (EVERS) for Info Record
    
    File Excel/CSV:
    - Cột A: Vendor (mã vendor)
    - Cột B: Material (mã material)
    - Cột C: Purch. Org (mã purchasing organization)
    - Cột D: Plant (mã plant)
    - Cột E: Shipping Instruction (EVERS) - CÓ THỂ TRỐNG
    - Không có header, dữ liệu bắt đầu từ dòng 1
    """
    log.info(f"[ME12-EVERS] Bắt đầu xử lý: {os.path.basename(filepath)}")
    start = time.time()

    try:
        # ===== 1. Đọc file =====
        info_records = []
        
        if filepath.lower().endswith('.csv'):
            # Đọc CSV
            import csv
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                for row in reader:
                    # Cần ít nhất 4 cột (vendor, material, purch_org, plant)
                    if len(row) >= 4 and row[0].strip() and row[1].strip():
                        info_records.append({
                            'vendor': row[0].strip(),
                            'material': row[1].strip(),
                            'purch_org': row[2].strip(),
                            'plant': row[3].strip(),
                            'shipping_instr': row[4].strip() if len(row) > 4 else ''
                        })
        else:
            # Đọc Excel
            import openpyxl
            wb = openpyxl.load_workbook(filepath, data_only=True)
            ws = wb.active
            
            for row in ws.iter_rows(min_row=1):
                vendor = row[0].value if len(row) > 0 else None
                material = row[1].value if len(row) > 1 else None
                purch_org = row[2].value if len(row) > 2 else None
                plant = row[3].value if len(row) > 3 else None
                shipping_instr = row[4].value if len(row) > 4 else None
                
                # Cần vendor, material, purch_org, plant; shipping_instr có thể trống
                if vendor and material and purch_org and plant:
                    info_records.append({
                        'vendor': str(vendor).strip(),
                        'material': str(material).strip(),
                        'purch_org': str(purch_org).strip(),
                        'plant': str(plant).strip(),
                        'shipping_instr': str(shipping_instr).strip() if shipping_instr else ''
                    })
            
            wb.close()
        
        if not info_records:
            return {"status": "skipped", "message": "Không có dữ liệu trong file", "rows": 0}

        log.info(f"  Đọc được {len(info_records)} info records")

        # ===== 2. Kiểm tra SAP User =====
        if not task or not task.sap_user:
            return {
                "status": "success",
                "message": f"Validated (no SAP user): {len(info_records)} records",
                "rows": len(info_records)
            }

        sap_client = task.sap_user.client
        sap_username = task.sap_user.username
        sap_password = task.sap_user.password
        log.info(f"  SAP User: {sap_client}/{sap_username}")

        time.sleep(2)

        # ===== 3. Xử lý trong SAP =====
        with SapGuiClient(
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
            
            # Chạy ME12 - Shipping Instruction
            me12 = TCodeME12(sap)
            res = me12.run_shipping_instruction(info_records)
            log.info(f"  Kết quả ME12-EVERS: {res}")

            duration = time.time() - start
            
            if res.get('ok'):
                # Thành công → Xóa file
                try:
                    os.remove(filepath)
                    log.info(f"  Đã xóa file: {filepath}")
                except Exception as del_err:
                    log.warning(f"  Không thể xóa file: {del_err}")
                    
                return {
                    "status": "success",
                    "message": res.get('message', f"Đã cập nhật {res.get('processed', 0)} info records"),
                    "rows": res.get('processed', 0),
                    "duration": round(duration, 2)
                }
            else:
                return {
                    "status": "error",
                    "message": res.get('message', 'Lỗi không xác định từ ME12'),
                    "rows": res.get('processed', 0),
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



def mk01_create_maker(filepath, session=None, task=None):
    """
    MK01 - Create Vendor (Maker)
    
    File Excel/CSV (không có header, bắt đầu từ dòng 1):
    - Cột A: LIFNR (Vendor code) - BẮT BUỘC
    - Cột B: NAME1 (Name 1) - BẮT BUỘC
    - Cột C: NAME2 (Name 2)
    - Cột D: EMNFR (Manufacturer number)
    - Cột E: STREET (Street)
    - Cột F: CITY2 (District)
    - Cột G: CITY1 (City)
    - Cột H: COUNTRY (Country code, VD: JP) - BẮT BUỘC
    - Cột I: STCD3 (Tax Number 3)
    - Cột J: STCD4 (Tax Number 4)
    - Cột K: INCO1 (Incoterms 1, VD: FOB)
    - Cột L: INCO2 (Incoterms 2, VD: JK)
    
    Fixed values: EKORG=C00A, KTOKK=V090, SORT1=1, WAERS=USD, WEBRE=True
    """
    log.info(f"[MK01] Bắt đầu xử lý: {os.path.basename(filepath)}")
    start = time.time()

    try:
        # ===== 1. Đọc file =====
        vendors = []
        
        if filepath.lower().endswith('.csv'):
            # Đọc CSV
            import csv
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                for row in reader:
                    lifnr = row[0].strip() if len(row) > 0 else ''
                    name1 = row[1].strip() if len(row) > 1 else ''
                    country = row[7].strip() if len(row) > 7 else ''
                    
                    # Bắt buộc: LIFNR, NAME1, COUNTRY
                    if lifnr and name1 and country:
                        vendors.append({
                            'lifnr': lifnr,
                            'name1': name1,
                            'name2': row[2].strip() if len(row) > 2 else '',
                            'emnfr': row[3].strip() if len(row) > 3 else '',
                            'street': row[4].strip() if len(row) > 4 else '',
                            'city2': row[5].strip() if len(row) > 5 else '',
                            'city1': row[6].strip() if len(row) > 6 else '',
                            'country': country,
                            'stcd3': row[8].strip() if len(row) > 8 else '',
                            'stcd4': row[9].strip() if len(row) > 9 else '',
                            'inco1': row[10].strip() if len(row) > 10 else '',
                            'inco2': row[11].strip() if len(row) > 11 else '',
                        })
        else:
            # Đọc Excel
            import openpyxl
            wb = openpyxl.load_workbook(filepath, data_only=True)
            ws = wb.active
            
            for row in ws.iter_rows(min_row=1):
                lifnr = str(row[0].value).strip() if len(row) > 0 and row[0].value else ''
                name1 = str(row[1].value).strip() if len(row) > 1 and row[1].value else ''
                country = str(row[7].value).strip() if len(row) > 7 and row[7].value else ''
                
                # Bắt buộc: LIFNR, NAME1, COUNTRY
                if lifnr and name1 and country:
                    vendors.append({
                        'lifnr': lifnr,
                        'name1': name1,
                        'name2': str(row[2].value).strip() if len(row) > 2 and row[2].value else '',
                        'emnfr': str(row[3].value).strip() if len(row) > 3 and row[3].value else '',
                        'street': str(row[4].value).strip() if len(row) > 4 and row[4].value else '',
                        'city2': str(row[5].value).strip() if len(row) > 5 and row[5].value else '',
                        'city1': str(row[6].value).strip() if len(row) > 6 and row[6].value else '',
                        'country': country,
                        'stcd3': str(row[8].value).strip() if len(row) > 8 and row[8].value else '',
                        'stcd4': str(row[9].value).strip() if len(row) > 9 and row[9].value else '',
                        'inco1': str(row[10].value).strip() if len(row) > 10 and row[10].value else '',
                        'inco2': str(row[11].value).strip() if len(row) > 11 and row[11].value else '',
                    })
            
            wb.close()
        
        if not vendors:
            return {"status": "skipped", "message": "Không có dữ liệu trong file", "rows": 0}

        log.info(f"  Đọc được {len(vendors)} vendors")

        # ===== 2. Kiểm tra SAP User =====
        if not task or not task.sap_user:
            return {
                "status": "success",
                "message": f"Validated (no SAP user): {len(vendors)} vendors",
                "rows": len(vendors)
            }

        sap_client = task.sap_user.client
        sap_username = task.sap_user.username
        sap_password = task.sap_user.password
        log.info(f"  SAP User: {sap_client}/{sap_username}")

        time.sleep(2)

        # ===== 3. Xử lý trong SAP =====
        with SapGuiClient(
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
            
            # Chạy MK01
            mk01 = TCodeMK01(sap)
            res = mk01.run(vendors)
            log.info(f"  Kết quả MK01: {res}")

            duration = time.time() - start
            
            if res.get('ok'):
                # Thành công → Xóa file
                try:
                    os.remove(filepath)
                    log.info(f"  Đã xóa file: {filepath}")
                except Exception as del_err:
                    log.warning(f"  Không thể xóa file: {del_err}")
                    
                return {
                    "status": "success",
                    "message": res.get('message', f"Đã tạo {res.get('processed', 0)} vendors"),
                    "rows": res.get('processed', 0),
                    "duration": round(duration, 2)
                }
            else:
                return {
                    "status": "error",
                    "message": res.get('message', 'Lỗi không xác định từ MK01'),
                    "rows": res.get('processed', 0),
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