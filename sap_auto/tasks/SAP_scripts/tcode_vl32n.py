"""
TCode VL32N - Change Inbound Delivery
=====================================
Xóa Inbound Delivery
"""
import time
import logging

log = logging.getLogger(__name__)


class TCodeVL32N:
    """
    VL32N - Change Inbound Delivery
    Xóa Inbound Delivery
    Nếu gặp lỗi → DỪNG NGAY và báo cáo
    """

    def __init__(self, sap_client):
        self.sap = sap_client
        self.session = sap_client.session

    def run(self, deliveries: list) -> dict:
        """
        Xử lý danh sách Inbound Deliveries - Xóa
        Nếu gặp lỗi → DỪNG NGAY và báo cáo
        
        Args:
            deliveries: List dict với keys:
                - delivery: Delivery Number (BẮT BUỘC)
            
        Returns:
            dict với keys: ok, status, message, processed, total, error_at_row, error_detail
        """
        processed = 0
        total = len(deliveries)
        
        for idx, item in enumerate(deliveries):
            delivery = str(item.get('delivery', '')).strip()
            row_number = idx + 2  # Dòng trong file (dòng 1 là header)
            
            if not delivery:
                continue
            
            result = self._process_single_delivery(delivery)
            
            if result['ok']:
                processed += 1
                log.info(f"  [VL32N] {delivery}: OK - Deleted")
            else:
                # GẶP LỖI → DỪNG NGAY
                error_detail = result.get('message', 'Lỗi không xác định')
                log.error(f"  [VL32N] DỪNG tại dòng {row_number} ({delivery}): {error_detail}")
                
                return {
                    'ok': False,
                    'status': 'error',
                    'message': f"Đã xử lý {processed}/{total}. Lỗi tại dòng {row_number} ({delivery}): {error_detail}",
                    'processed': processed,
                    'total': total,
                    'error_at_row': row_number,
                    'error_detail': error_detail
                }
        
        # Tất cả thành công
        return {
            'ok': True,
            'status': 'success',
            'message': f"Hoàn tất. Đã xóa {processed}/{total} Inbound Deliveries",
            'processed': processed,
            'total': total,
            'error_at_row': None,
            'error_detail': None
        }

    def _process_single_delivery(self, delivery: str) -> dict:
        """
        Xử lý xóa 1 Inbound Delivery trong VL32N
        
        Args:
            delivery: Delivery Number (VD: 1800035488)
            
        Returns:
            dict với keys: ok, message
        """
        try:
            # 1. Mở TCode VL32N
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nVL32N"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)
            
            # 2. Nhập Delivery Number
            self.session.findById("wnd[0]/usr/ctxtLIKP-VBELN").text = delivery
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)
            
            # Kiểm tra lỗi (Delivery không tồn tại)
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}
            
            # 3. Nhấn Delete (btn[14])
            try:
                self.session.findById("wnd[0]/tbar[1]/btn[14]").press()
                time.sleep(0.5)
                self.sap.wait_ready(self.session, 10)
            except Exception as e:
                return {'ok': False, 'message': f'Không tìm thấy nút Delete: {e}'}
            
            # 4. Xử lý popup confirm - Nhấn Yes (SPOP-OPTION1)
            try:
                popup = self.sap.safe_find(self.session, "wnd[1]")
                if popup:
                    btn_yes = self.sap.safe_find(self.session, "wnd[1]/usr/btnSPOP-OPTION1")
                    if btn_yes:
                        btn_yes.press()
                        time.sleep(0.5)
                        self.sap.wait_ready(self.session, 10)
                    else:
                        # Thử nhấn Enter nếu không có button
                        self.session.findById("wnd[1]").sendVKey(0)
                        time.sleep(0.5)
                        self.sap.wait_ready(self.session, 10)
            except:
                pass
            
            # 5. Kiểm tra kết quả
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}
            
            # 6. Save (btn[11]) hoặc Exit (btn[15])
            try:
                # Thử Save trước
                btn_save = self.sap.safe_find(self.session, "wnd[0]/tbar[0]/btn[11]")
                if btn_save:
                    btn_save.press()
                    time.sleep(1)
                    self.sap.wait_ready(self.session, 10)
            except:
                pass
            
            # Kiểm tra kết quả cuối cùng
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}
            
            return {'ok': True, 'message': 'OK'}
            
        except Exception as e:
            return {'ok': False, 'message': str(e)}