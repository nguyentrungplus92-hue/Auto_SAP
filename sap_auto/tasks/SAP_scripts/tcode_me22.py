"""
TCode ME22 - Change Purchase Order
==================================
Thay đổi số lượng / giá trong PO
Chạy tất cả dòng, ghi nhận lỗi từng dòng, không dừng giữa chừng
"""
import time
import logging

log = logging.getLogger(__name__)


class TCodeME22:
    """
    ME22 - Change Purchase Order
    Cập nhật số lượng (Quantity) hoặc giá (Net Price) cho PO Item
    Chạy tất cả dòng, thu thập lỗi, không dừng giữa chừng
    """

    TABLE_ID = "wnd[0]/usr/tblSAPMM06ETC_0120"
    SESSION_DEAD_MSG = 'SAP session bị đóng — user đã login ở nơi khác hoặc mất kết nối'

    def __init__(self, sap_client):
        self.sap = sap_client
        self.session = sap_client.session

    def run(self, po_items: list) -> dict:
        """
        Xử lý danh sách PO items - Cập nhật Quantity
        """
        processed = 0
        total = len(po_items)
        errors = []

        # Mở TCode ME22 lần đầu
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nME22"
        self.session.findById("wnd[0]").sendVKey(0)
        time.sleep(0.5)
        self.sap.wait_ready(self.session, 10)

        for idx, item in enumerate(po_items):
            po_number   = str(item.get('po_number',   '')).strip()
            item_number = str(item.get('item_number', '')).strip()
            quantity    = str(item.get('quantity',    '')).strip()
            row_number  = idx + 2

            # ===== Kiểm tra session trước mỗi dòng =====
            if not self.sap.is_session_alive():
                for remaining_idx in range(idx, len(po_items)):
                    r = po_items[remaining_idx]
                    r_po   = str(r.get('po_number',   '')).strip()
                    r_item = str(r.get('item_number', '')).strip()
                    errors.append({'row': remaining_idx + 2, 'info': f"PO {r_po}/{r_item}" or '-', 'detail': self.SESSION_DEAD_MSG})
                break
            # ============================================

            if not po_number or not item_number or not quantity:
                continue

            result = self._process_single_quantity(po_number, item_number, quantity)

            if result['ok']:
                processed += 1
                log.info(f"  [ME22] PO {po_number}/{item_number}: OK - Qty: {quantity}")
            else:
                error_detail = result.get('message', 'Lỗi không xác định')
                log.warning(f"  [ME22] Lỗi dòng {row_number} (PO {po_number}/{item_number}): {error_detail}")
                errors.append({'row': row_number, 'info': f"PO {po_number}/{item_number}", 'detail': error_detail})

        if errors:
            error_lines = "\n".join([f"  • Dòng {e['row']} ({e['info']}): {e['detail']}" for e in errors])
            message = f"Hoàn tất {processed}/{total} dòng thành công. Lỗi {len(errors)} dòng:\n{error_lines}"
        else:
            message = f"Hoàn tất. Đã cập nhật thành công {processed}/{total} PO items"

        return {'ok': True, 'message': message, 'processed': processed, 'total': total, 'errors': errors}

    def run_price(self, po_items: list) -> dict:
        """
        Xử lý danh sách PO items - Cập nhật Net Price
        """
        processed = 0
        total = len(po_items)
        errors = []

        for idx, item in enumerate(po_items):
            po_number   = str(item.get('po_number',   '')).strip()
            item_number = str(item.get('item_number', '')).strip()
            price       = str(item.get('price',       '')).strip()
            row_number  = idx + 2

            # ===== Kiểm tra session trước mỗi dòng =====
            if not self.sap.is_session_alive():
                for remaining_idx in range(idx, len(po_items)):
                    r = po_items[remaining_idx]
                    r_po   = str(r.get('po_number',   '')).strip()
                    r_item = str(r.get('item_number', '')).strip()
                    errors.append({'row': remaining_idx + 2, 'info': f"PO {r_po}/{r_item}" or '-', 'detail': self.SESSION_DEAD_MSG})
                break
            # ============================================

            if not po_number or not item_number or not price:
                continue

            result = self._process_single_price(po_number, item_number, price)

            if result['ok']:
                processed += 1
                log.info(f"  [ME22] PO {po_number}/{item_number}: OK - Price: {price}")
            else:
                error_detail = result.get('message', 'Lỗi không xác định')
                log.warning(f"  [ME22] Lỗi dòng {row_number} (PO {po_number}/{item_number}): {error_detail}")
                errors.append({'row': row_number, 'info': f"PO {po_number}/{item_number}", 'detail': error_detail})

        if errors:
            error_lines = "\n".join([f"  • Dòng {e['row']} ({e['info']}): {e['detail']}" for e in errors])
            message = f"Hoàn tất {processed}/{total} dòng thành công. Lỗi {len(errors)} dòng:\n{error_lines}"
        else:
            message = f"Hoàn tất. Đã cập nhật giá thành công {processed}/{total} PO items"

        return {'ok': True, 'message': message, 'processed': processed, 'total': total, 'errors': errors}

    def _process_single_quantity(self, po_number: str, item_number: str, quantity: str) -> dict:
        """Xử lý cập nhật Quantity cho 1 PO item"""
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nME22"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)
            self.sap.wait_ready(self.session, 10)

            po_field = self.sap.safe_find(self.session, "wnd[0]/usr/ctxtRM06E-BSTNR")
            if po_field is None:
                return {'ok': False, 'message': 'Không tìm thấy field PO Number'}

            po_field.text = po_number
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.1)
            self.sap.wait_ready(self.session, 10)

            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            item_field = self.sap.safe_find(self.session, "wnd[0]/usr/txtRM06E-EBELP")
            if item_field is None:
                return {'ok': False, 'message': 'Không tìm thấy field Item Number'}

            item_field.text = item_number
            item_field.setFocus()
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.1)
            self.sap.wait_ready(self.session, 10)

            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            qty_field = self.sap.safe_find(self.session, f"{self.TABLE_ID}/txtEKPO-MENGE[5,0]")
            if qty_field is None:
                return {'ok': False, 'message': 'Không tìm thấy field Quantity'}

            qty_field.text = quantity
            qty_field.setFocus()

            for i in range(3):
                self.session.findById("wnd[0]").sendVKey(0)
                time.sleep(0.1)
                self.sap.wait_ready(self.session, 5)

            btn_save = self.sap.safe_find(self.session, "wnd[0]/tbar[0]/btn[11]")
            if btn_save is None:
                return {'ok': False, 'message': 'Không tìm thấy nút Save'}

            btn_save.press()
            time.sleep(0.5)
            self.sap.wait_ready(self.session, 10)

            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            return {'ok': True, 'message': 'OK'}

        except Exception as e:
            return {'ok': False, 'message': str(e)}

    def _process_single_price(self, po_number: str, item_number: str, price: str) -> dict:
        """Xử lý cập nhật Net Price cho 1 PO item"""
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nME22"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)
            self.sap.wait_ready(self.session, 10)

            po_field = self.sap.safe_find(self.session, "wnd[0]/usr/ctxtRM06E-BSTNR")
            if po_field is None:
                return {'ok': False, 'message': 'Không tìm thấy field PO Number'}

            po_field.text = po_number
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.1)
            self.sap.wait_ready(self.session, 10)

            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            item_field = self.sap.safe_find(self.session, "wnd[0]/usr/txtRM06E-EBELP")
            if item_field is None:
                return {'ok': False, 'message': 'Không tìm thấy field Item Number'}

            item_field.text = item_number
            item_field.setFocus()
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.1)
            self.sap.wait_ready(self.session, 10)

            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            price_field = self.sap.safe_find(self.session, f"{self.TABLE_ID}/txtEKPO-NETPR[9,0]")
            if price_field is None:
                return {'ok': False, 'message': 'Không tìm thấy field Net Price'}

            price_field.text = price
            price_field.setFocus()

            for i in range(3):
                self.session.findById("wnd[0]").sendVKey(0)
                time.sleep(0.1)
                self.sap.wait_ready(self.session, 5)

            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            btn_save = self.sap.safe_find(self.session, "wnd[0]/tbar[0]/btn[11]")
            if btn_save is None:
                return {'ok': False, 'message': 'Không tìm thấy nút Save'}

            btn_save.press()
            time.sleep(0.5)
            self.sap.wait_ready(self.session, 10)

            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            return {'ok': True, 'message': 'OK'}

        except Exception as e:
            return {'ok': False, 'message': str(e)}
