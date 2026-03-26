"""
TCode QA11 - Record Usage Decision
==================================
Nhập Usage Decision cho Inspection Lot
Nếu gặp lỗi → DỪNG NGAY và báo cáo
"""
import time
import logging

log = logging.getLogger(__name__)


class TCodeQA11:
    """
    QA11 - Record Usage Decision
    Nhập Usage Decision Code cho Inspection Lot
    Nếu gặp lỗi → DỪNG NGAY và báo cáo
    """

    # Field IDs
    PRUEFLOS_ID = "wnd[0]/usr/ctxtQALS-PRUEFLOS"
    VCODE_ID = "wnd[0]/usr/tabsUD_DATA/tabpFEHL/ssubSUB_UD_DATA:SAPMQEVA:0103/subUD_DATA:SAPMQEVA:1103/ctxtRQEVA-VCODE"
    VCODEGRP_ID = "wnd[0]/usr/tabsUD_DATA/tabpFEHL/ssubSUB_UD_DATA:SAPMQEVA:0103/subUD_DATA:SAPMQEVA:1103/ctxtRQEVA-VCODEGRP"

    def __init__(self, sap_client):
        self.sap = sap_client
        self.session = sap_client.session

    def run(self, items: list) -> dict:
        """
        Xử lý danh sách Inspection Lots - Nhập Usage Decision
        Nếu gặp lỗi → DỪNG NGAY và báo cáo
        
        Args:
            items: List dict với keys:
                - prueflos: Inspection Lot Number (BẮT BUỘC)
                - vcode: Usage Decision Code (BẮT BUỘC)
                - vcodegrp: Code Group (BẮT BUỘC)
            
        Returns:
            dict với keys: ok, status, message, processed, total, error_at_row, error_detail
        """
        processed = 0
        total = len(items)
        
        for idx, item in enumerate(items):
            prueflos = str(item.get('prueflos', '')).strip()
            vcode = str(item.get('vcode', '')).strip()
            vcodegrp = str(item.get('vcodegrp', '')).strip()
            row_number = idx + 2  # Dòng trong file (dòng 1 là header)
            
            if not prueflos or not vcode or not vcodegrp:
                continue
            
            result = self._process_single_lot(prueflos, vcode, vcodegrp)
            
            if result['ok']:
                processed += 1
                log.info(f"  [QA11] {prueflos}: OK - {vcode}/{vcodegrp}")
            else:
                # GẶP LỖI → DỪNG NGAY
                error_detail = result.get('message', 'Lỗi không xác định')
                log.error(f"  [QA11] DỪNG tại dòng {row_number} ({prueflos}): {error_detail}")
                
                return {
                    'ok': False,
                    'status': 'error',
                    'message': f"Đã xử lý {processed}/{total}. Lỗi tại dòng {row_number} ({prueflos}): {error_detail}",
                    'processed': processed,
                    'total': total,
                    'error_at_row': row_number,
                    'error_detail': error_detail
                }
        
        # Tất cả thành công
        return {
            'ok': True,
            'status': 'success',
            'message': f"Hoàn tất. Đã nhập UD cho {processed}/{total} Inspection Lots",
            'processed': processed,
            'total': total,
            'error_at_row': None,
            'error_detail': None
        }

    def _process_single_lot(self, prueflos: str, vcode: str, vcodegrp: str) -> dict:
        """
        Xử lý nhập Usage Decision cho 1 Inspection Lot
        
        Args:
            prueflos: Inspection Lot Number (VD: 765007)
            vcode: Usage Decision Code (VD: A)
            vcodegrp: Code Group (VD: UD1)
            
        Returns:
            dict với keys: ok, message
        """
        try:
            # 1. Mở TCode QA11 mỗi lần xử lý
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nQA11"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)
            
            # 2. Nhập Inspection Lot Number
            lot_field = self.sap.safe_find(self.session, self.PRUEFLOS_ID)
            if lot_field is None:
                return {'ok': False, 'message': 'Không tìm thấy field Inspection Lot'}
            
            lot_field.text = prueflos
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)
            self.sap.wait_ready(self.session, 10)
            
            # Kiểm tra lỗi (Lot không tồn tại)
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}
            
            # 3. Nhập Usage Decision Code (VCODE)
            vcode_field = self.sap.safe_find(self.session, self.VCODE_ID)
            if vcode_field is None:
                return {'ok': False, 'message': 'Không tìm thấy field UD Code'}
            
            vcode_field.text = vcode
            time.sleep(0.3)
            
            # 4. Nhập Code Group (VCODEGRP)
            vcodegrp_field = self.sap.safe_find(self.session, self.VCODEGRP_ID)
            if vcodegrp_field is None:
                return {'ok': False, 'message': 'Không tìm thấy field Code Group'}
            
            vcodegrp_field.text = vcodegrp
            vcodegrp_field.setFocus()
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)
            self.sap.wait_ready(self.session, 10)
            
            # Kiểm tra lỗi sau khi nhập
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}
            
            # 5. Save (btn[11])
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