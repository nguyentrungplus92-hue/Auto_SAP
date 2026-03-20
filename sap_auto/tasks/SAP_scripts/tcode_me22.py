"""
TCode ME22 - Change Purchase Order
==================================
Thay đổi số lượng trong PO
Nếu gặp lỗi → DỪNG NGAY và báo cáo
"""
import time
import logging

log = logging.getLogger(__name__)


class TCodeME22:
    """
    ME22 - Change Purchase Order
    Cập nhật số lượng (Quantity) cho PO Item
    Nếu gặp lỗi → DỪNG NGAY và báo cáo
    """

    def __init__(self, sap_client):
        self.sap = sap_client
        self.session = sap_client.session

    def run(self, po_items: list) -> dict:
        """
        Xử lý danh sách PO items
        Nếu gặp lỗi → DỪNG NGAY và báo cáo
        
        Args:
            po_items: List dict với keys: po_number, item_number, quantity
            
        Returns:
            dict với keys: ok, status, message, processed, total, error_at_row, error_detail
        """
        processed = 0
        total = len(po_items)
        
        for idx, item in enumerate(po_items):
            po_number = str(item.get('po_number', '')).strip()
            item_number = str(item.get('item_number', '')).strip()
            quantity = str(item.get('quantity', '')).strip()
            row_number = idx + 2  # Dòng trong file (dòng 1 là header)
            
            if not po_number or not item_number or not quantity:
                continue
            
            result = self._process_single_item(po_number, item_number, quantity)
            
            if result['ok']:
                processed += 1
                log.info(f"  [ME22] PO {po_number}/{item_number}: OK - Qty: {quantity}")
            else:
                # GẶP LỖI → DỪNG NGAY
                error_detail = result.get('message', 'Lỗi không xác định')
                log.error(f"  [ME22] DỪNG tại dòng {row_number} (PO {po_number}/{item_number}): {error_detail}")
                
                return {
                    'ok': False,
                    'status': 'error',
                    'message': f"Đã xử lý {processed}/{total}. Lỗi tại dòng {row_number} (PO {po_number}/{item_number}): {error_detail}",
                    'processed': processed,
                    'total': total,
                    'error_at_row': row_number,
                    'error_detail': error_detail
                }
        
        # Tất cả thành công
        return {
            'ok': True,
            'status': 'success',
            'message': f"Hoàn tất. Đã cập nhật {processed}/{total} PO items",
            'processed': processed,
            'total': total,
            'error_at_row': None,
            'error_detail': None
        }

    def _process_single_item(self, po_number: str, item_number: str, quantity: str) -> dict:
        """
        Xử lý 1 PO item trong ME22
        
        Args:
            po_number: Số PO (VD: 4500023437)
            item_number: Số item (VD: 1, 2, 3...)
            quantity: Số lượng mới
            
        Returns:
            dict với keys: ok, message
        """
        try:
            # 1. Mở TCode ME22
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nME22"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)
            
            # 2. Nhập PO Number
            self.session.findById("wnd[0]/usr/ctxtRM06E-BSTNR").text = po_number
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)
            
            # Kiểm tra lỗi (PO không tồn tại)
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}
            
            # 3. Nhập Item Number
            self.session.findById("wnd[0]/usr/txtRM06E-EBELP").text = item_number
            self.session.findById("wnd[0]/usr/txtRM06E-EBELP").setFocus()
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)
            
            # Kiểm tra lỗi (Item không tồn tại)
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}
            
            # 4. Nhập Quantity mới
            try:
                qty_field = self.session.findById("wnd[0]/usr/tblSAPMM06ETC_0120/txtEKPO-MENGE[5,0]")
                qty_field.text = quantity
                qty_field.setFocus()
                self.session.findById("wnd[0]").sendVKey(0)
                time.sleep(1)
                self.sap.wait_ready(self.session, 10)
            except Exception as e:
                return {'ok': False, 'message': f'Không tìm thấy field Quantity: {e}'}
            
            # 5. Save (btn[11])
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)
            
            # Kiểm tra kết quả save
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}
            
            return {'ok': True, 'message': 'OK'}
            
        except Exception as e:
            return {'ok': False, 'message': str(e)}