"""
TCode MD12 - Change Planned Order
=================================
Bỏ tick checkbox "Firmly planned" cho planned order
"""
import time
import logging

log = logging.getLogger(__name__)


class TCodeMD12:
    """
    MD12 - Change Planned Order
    Bỏ tick "Fixierung" (Firmly planned) checkbox
    Nếu gặp lỗi → DỪNG NGAY và báo cáo
    """

    def __init__(self, sap_client):
        self.sap = sap_client
        self.session = sap_client.session

    def run(self, planned_orders: list) -> dict:
        """
        Xử lý danh sách planned orders
        Nếu gặp lỗi → DỪNG NGAY và báo cáo
        
        Args:
            planned_orders: List các planned order number
            
        Returns:
            dict với keys: ok, status, message, processed, total, error_at_row, error_detail
        """
        processed = 0
        total = len(planned_orders)
        
        for idx, order in enumerate(planned_orders):
            order = str(order).strip()
            row_number = idx + 2  # Dòng trong file (dòng 1 là header)
            
            if not order:
                continue
                
            result = self._process_single_order(order)
            
            if result['ok']:
                processed += 1
                log.info(f"  [MD12] {order}: OK")
            else:
                # GẶP LỖI → DỪNG NGAY
                error_detail = result.get('message', 'Lỗi không xác định')
                log.error(f"  [MD12] DỪNG tại dòng {row_number} ({order}): {error_detail}")
                
                return {
                    'ok': False,
                    'status': 'error',
                    'message': f"Đã xử lý {processed}/{total}. Lỗi tại dòng {row_number} ({order}): {error_detail}",
                    'processed': processed,
                    'total': total,
                    'error_at_row': row_number,
                    'error_detail': error_detail
                }
        
        # Tất cả thành công
        return {
            'ok': True,
            'status': 'success',
            'message': f"Hoàn tất. Đã xử lý {processed}/{total} planned orders",
            'processed': processed,
            'total': total,
            'error_at_row': None,
            'error_detail': None
        }

    def _process_single_order(self, planned_order: str) -> dict:
        """
        Xử lý 1 planned order trong MD12
        
        Args:
            planned_order: Số planned order
            
        Returns:
            dict với keys: ok, message
        """
        try:
            # 1. Mở TCode MD12
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nMD12"
            self.session.findById("wnd[0]").sendVKey(0)
            self.sap.wait_ready(self.session, 10)
            
            # 2. Nhập planned order number
            self.session.findById("wnd[0]/usr/txtRM61P-PLNUM").text = planned_order
            
            # 3. Nhấn Enter để mở planned order
            self.session.findById("wnd[0]").sendVKey(0)
            self.sap.wait_ready(self.session, 10)
            
            # Kiểm tra lỗi
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}
            
            # Dismiss popup nếu có
            self.sap.dismiss_popup_if_any()
            
            # 4. Bỏ tick checkbox "Firmly planned" (PLAF-AUFFX)
            checkbox_id = "wnd[0]/usr/tabsTABTC/tabpTAB01/ssubINCLUDE1XX:SAPLM61O:0711/subINCLUDE711_2:SAPLM61O:0810/chkPLAF-AUFFX"
            checkbox = self.sap.safe_find(self.session, checkbox_id)
            
            if checkbox is None:
                return {'ok': False, 'message': 'Không tìm thấy checkbox Firmly planned'}
            
            # Bỏ tick nếu đang được tick
            if checkbox.selected:
                checkbox.selected = False
                log.info(f"    Đã bỏ tick Firmly planned")
            else:
                log.info(f"    Checkbox đã bỏ tick sẵn")
            
            # 5. Nhấn Save (btn[11])
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            self.sap.wait_ready(self.session, 10)
            
            # Kiểm tra kết quả save
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}
            
            # 6. Nhấn Back (btn[3])
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            self.sap.wait_ready(self.session, 5)
            
            return {'ok': True, 'message': 'OK'}
            
        except Exception as e:
            return {'ok': False, 'message': str(e)}