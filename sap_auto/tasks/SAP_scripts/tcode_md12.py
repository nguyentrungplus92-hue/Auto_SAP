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
    """

    def __init__(self, sap_client):
        self.sap = sap_client
        self.session = sap_client.session

    def run(self, planned_orders: list) -> dict:
        """
        Xử lý danh sách planned orders
        
        Args:
            planned_orders: List các planned order number
            
        Returns:
            dict với keys: ok, message, processed, errors
        """
        processed = 0
        errors = []
        
        for order in planned_orders:
            order = str(order).strip()
            if not order:
                continue
                
            try:
                result = self._process_single_order(order)
                if result['ok']:
                    processed += 1
                    log.info(f"  [MD12] {order}: OK")
                else:
                    errors.append(f"{order}: {result['message']}")
                    log.warning(f"  [MD12] {order}: {result['message']}")
            except Exception as e:
                errors.append(f"{order}: {str(e)}")
                log.error(f"  [MD12] {order}: Exception - {e}")
        
        # Tổng kết
        if errors:
            return {
                'ok': processed > 0,
                'message': f"Processed: {processed}, Errors: {len(errors)} - {'; '.join(errors[:3])}",
                'processed': processed,
                'errors': errors
            }
        else:
            return {
                'ok': True,
                'message': f"Hoàn tất. Đã xử lý {processed} planned orders",
                'processed': processed,
                'errors': []
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