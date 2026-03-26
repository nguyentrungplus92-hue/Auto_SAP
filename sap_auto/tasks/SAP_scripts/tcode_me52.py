"""
TCode ME52 - Change Purchase Requisition
=========================================
Xóa Item trong Purchase Requisition
"""
import time
import logging

log = logging.getLogger(__name__)


class TCodeME52:
    """
    ME52 - Change Purchase Requisition
    Xóa Item trong PR
    Nếu gặp lỗi → DỪNG NGAY và báo cáo
    """

    TABLE_ID = "wnd[0]/usr/tblSAPMM06BTC_0106"

    def __init__(self, sap_client):
        self.sap = sap_client
        self.session = sap_client.session

    def run(self, pr_items: list) -> dict:
        """
        Xử lý danh sách PR Items - Xóa
        Nếu gặp lỗi → DỪNG NGAY và báo cáo
        
        Args:
            pr_items: List dict với keys:
                - banfn: Purchase Requisition Number (BẮT BUỘC)
                - bnfpo: Item Number (BẮT BUỘC)
            
        Returns:
            dict với keys: ok, status, message, processed, total, error_at_row, error_detail
        """
        processed = 0
        total = len(pr_items)
        
        # Mở TCode ME52 lần đầu
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nME52"
        self.session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)
        self.sap.wait_ready(self.session, 10)
        
        for idx, item in enumerate(pr_items):
            banfn = str(item.get('banfn', '')).strip()
            bnfpo = str(item.get('bnfpo', '')).strip()
            row_number = idx + 2  # Dòng trong file (dòng 1 là header)
            
            if not banfn or not bnfpo:
                continue
            
            result = self._process_single_item(banfn, bnfpo)
            
            if result['ok']:
                processed += 1
                log.info(f"  [ME52] {banfn}/{bnfpo}: OK - Deleted")
            else:
                # GẶP LỖI → DỪNG NGAY
                error_detail = result.get('message', 'Lỗi không xác định')
                log.error(f"  [ME52] DỪNG tại dòng {row_number} ({banfn}/{bnfpo}): {error_detail}")
                
                return {
                    'ok': False,
                    'status': 'error',
                    'message': f"Đã xử lý {processed}/{total}. Lỗi tại dòng {row_number} ({banfn}/{bnfpo}): {error_detail}",
                    'processed': processed,
                    'total': total,
                    'error_at_row': row_number,
                    'error_detail': error_detail
                }
        
        # Tất cả thành công
        return {
            'ok': True,
            'status': 'success',
            'message': f"Hoàn tất. Đã xóa {processed}/{total} PR Items",
            'processed': processed,
            'total': total,
            'error_at_row': None,
            'error_detail': None
        }

    def _process_single_item(self, banfn: str, bnfpo: str) -> dict:
        """
        Xử lý xóa 1 PR Item trong ME52
        
        Args:
            banfn: Purchase Requisition Number (VD: 1059967849)
            bnfpo: Item Number (VD: 10, 20, 30...)
            
        Returns:
            dict với keys: ok, message
        """
        try:
            # 1. Nhập PR Number
            pr_field = self.sap.safe_find(self.session, "wnd[0]/usr/ctxtEBAN-BANFN")
            if pr_field is None:
                return {'ok': False, 'message': 'Không tìm thấy field PR Number'}
            
            pr_field.text = banfn
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)
            self.sap.wait_ready(self.session, 10)
            
            # Kiểm tra lỗi (PR không tồn tại)
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}
            
            # 2. Nhập Item Number để filter
            item_field = self.sap.safe_find(self.session, "wnd[0]/usr/txtRM06B-BNFPO")
            if item_field is None:
                return {'ok': False, 'message': 'Không tìm thấy field Item Number'}
            
            item_field.text = bnfpo
            item_field.setFocus()
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)
            self.sap.wait_ready(self.session, 10)
            
            # Kiểm tra lỗi (Item không tồn tại)
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}
            
            # 3. Select row đầu tiên trong table (vì đã filter theo BNFPO, chỉ còn 1 row)
            table = self.sap.safe_find(self.session, self.TABLE_ID)
            if table is None:
                return {'ok': False, 'message': 'Không tìm thấy table items'}
            
            try:
                table.getAbsoluteRow(0).selected = True
                time.sleep(0.3)
            except Exception as e:
                return {'ok': False, 'message': f'Không thể select row: {e}'}
            
            # 4. Nhấn Delete (btn[14])
            btn_delete = self.sap.safe_find(self.session, "wnd[0]/tbar[1]/btn[14]")
            if btn_delete is None:
                return {'ok': False, 'message': 'Không tìm thấy nút Delete'}
            
            btn_delete.press()
            time.sleep(0.5)
            self.sap.wait_ready(self.session, 10)
            
            # 5. Xử lý popup confirm - Nhấn Yes (SPOP-OPTION1)
            popup = self.sap.safe_find(self.session, "wnd[1]")
            if popup:
                btn_yes = self.sap.safe_find(self.session, "wnd[1]/usr/btnSPOP-OPTION1")
                if btn_yes:
                    btn_yes.press()
                    time.sleep(0.5)
                    self.sap.wait_ready(self.session, 10)
                else:
                    # Thử nhấn Enter
                    self.session.findById("wnd[1]").sendVKey(0)
                    time.sleep(0.5)
                    self.sap.wait_ready(self.session, 10)
            
            # 6. Kiểm tra lỗi sau khi delete
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}
            
            # 7. Save (btn[11])
            btn_save = self.sap.safe_find(self.session, "wnd[0]/tbar[0]/btn[11]")
            if btn_save is None:
                return {'ok': False, 'message': 'Không tìm thấy nút Save'}
            
            btn_save.press()
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)
            
            # Kiểm tra kết quả Save
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}
            
            return {'ok': True, 'message': 'OK'}
            
        except Exception as e:
            return {'ok': False, 'message': str(e)}