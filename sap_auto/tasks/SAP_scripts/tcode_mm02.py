"""
TCode MM02 - Change Material Master
===================================
Cập nhật mô tả tiếng Việt cho Material
Cập nhật External Material Group
Cập nhật Maximum Lot Size
Nếu gặp lỗi → DỪNG NGAY và báo cáo
"""
import time
import logging

log = logging.getLogger(__name__)


class TCodeMM02:
    """
    MM02 - Change Material Master
    Cập nhật Long Text Description (Tab ZU05, ZU07)
    Cập nhật External Material Group (Tab SP01)
    Cập nhật Maximum Lot Size (Tab SP12 - MRP 1)
    Nếu gặp lỗi → DỪNG NGAY và báo cáo
    """

    # Field IDs
    MATNR_ID = "wnd[0]/usr/ctxtRMMG1-MATNR"
    WERKS_POPUP_ID = "wnd[1]/usr/ctxtRMMG1-WERKS"
    BSTMA_ID = "wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2483/txtMARC-BSTMA"

    def __init__(self, sap_client):
        self.sap = sap_client
        self.session = sap_client.session

    def run(self, materials: list) -> dict:
        """
        Xử lý danh sách materials - Tab ZU05 (Basic data text)
        Nếu gặp lỗi → DỪNG NGAY và báo cáo
        """
        return self._process_materials(materials, tab='ZU05')

    def run_zu07(self, materials: list) -> dict:
        """
        Xử lý danh sách materials - Tab ZU07 (Internal comment)
        Nếu gặp lỗi → DỪNG NGAY và báo cáo
        """
        return self._process_materials(materials, tab='ZU07')

    def run_extwg(self, materials: list) -> dict:
        """
        Xử lý danh sách materials - Cập nhật External Material Group
        Nếu gặp lỗi → DỪNG NGAY và báo cáo
        
        Args:
            materials: List dict với keys:
                - matnr: Material Number (BẮT BUỘC)
                - extwg: External Material Group (BẮT BUỘC)
        """
        processed = 0
        total = len(materials)
        
        # Mở TCode MM02 lần đầu
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nMM02"
        self.session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)
        self.sap.wait_ready(self.session, 10)
        
        for idx, item in enumerate(materials):
            matnr = str(item.get('matnr', '')).strip()
            extwg = str(item.get('extwg', '')).strip()
            row_number = idx + 2  # Dòng trong file (dòng 1 là header)
            
            if not matnr or not extwg:
                continue
            
            result = self._process_single_extwg(matnr, extwg)
            
            if result['ok']:
                processed += 1
                log.info(f"  [MM02-EXTWG] {matnr}: OK - {extwg}")
            else:
                # GẶP LỖI → DỪNG NGAY
                error_detail = result.get('message', 'Lỗi không xác định')
                log.error(f"  [MM02-EXTWG] DỪNG tại dòng {row_number} ({matnr}): {error_detail}")
                
                return {
                    'ok': False,
                    'status': 'error',
                    'message': f"Đã xử lý {processed}/{total}. Lỗi tại dòng {row_number} ({matnr}): {error_detail}",
                    'processed': processed,
                    'total': total,
                    'error_at_row': row_number,
                    'error_detail': error_detail
                }
        
        # Tất cả thành công
        return {
            'ok': True,
            'status': 'success',
            'message': f"Hoàn tất. Đã cập nhật {processed}/{total} materials",
            'processed': processed,
            'total': total,
            'error_at_row': None,
            'error_detail': None
        }

    def run_bstma(self, materials: list) -> dict:
        """
        Xử lý danh sách materials - Cập nhật Maximum Lot Size (Tab SP12 - MRP 1)
        Nếu gặp lỗi → DỪNG NGAY và báo cáo
        
        Args:
            materials: List dict với keys:
                - matnr: Material Number (BẮT BUỘC)
                - werks: Plant (BẮT BUỘC)
                - bstma: Maximum Lot Size (BẮT BUỘC)
        """
        processed = 0
        total = len(materials)
        
        for idx, item in enumerate(materials):
            matnr = str(item.get('matnr', '')).strip()
            werks = str(item.get('werks', '')).strip()
            bstma = str(item.get('bstma', '')).strip()
            row_number = idx + 2  # Dòng trong file (dòng 1 là header)
            
            if not matnr or not werks or not bstma:
                continue
            
            result = self._process_single_bstma(matnr, werks, bstma)
            
            if result['ok']:
                processed += 1
                log.info(f"  [MM02-BSTMA] {matnr}/{werks}: OK - {bstma}")
            else:
                # GẶP LỖI → DỪNG NGAY
                error_detail = result.get('message', 'Lỗi không xác định')
                log.error(f"  [MM02-BSTMA] DỪNG tại dòng {row_number} ({matnr}/{werks}): {error_detail}")
                
                return {
                    'ok': False,
                    'status': 'error',
                    'message': f"Đã xử lý {processed}/{total}. Lỗi tại dòng {row_number} ({matnr}/{werks}): {error_detail}",
                    'processed': processed,
                    'total': total,
                    'error_at_row': row_number,
                    'error_detail': error_detail
                }
        
        # Tất cả thành công
        return {
            'ok': True,
            'status': 'success',
            'message': f"Hoàn tất. Đã cập nhật Maximum Lot Size cho {processed}/{total} materials",
            'processed': processed,
            'total': total,
            'error_at_row': None,
            'error_detail': None
        }

    def _process_single_bstma(self, matnr: str, werks: str, bstma: str) -> dict:
        """
        Xử lý cập nhật Maximum Lot Size cho 1 material
        
        Args:
            matnr: Material Number (VD: ARATB2H00132)
            werks: Plant (VD: V501)
            bstma: Maximum Lot Size (VD: 400)
            
        Returns:
            dict với keys: ok, message
        """
        try:
            # 1. Mở TCode MM02 mỗi lần xử lý
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nMM02"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)
            
            # 2. Nhập Material Number
            matnr_field = self.sap.safe_find(self.session, self.MATNR_ID)
            if matnr_field is None:
                return {'ok': False, 'message': 'Không tìm thấy field Material Number'}
            
            matnr_field.text = matnr
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)
            self.sap.wait_ready(self.session, 10)
            
            # Kiểm tra lỗi (Material không tồn tại)
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}
            
            # 3. Xử lý popup Select View - Chọn MRP 1 (row 10)
            popup = self.sap.safe_find(self.session, "wnd[1]")
            if popup:
                try:
                    # Nhấn Page Down 2 lần để scroll đến MRP 1
                    self.session.findById("wnd[1]/tbar[0]/btn[19]").press()
                    time.sleep(0.3)
                    self.session.findById("wnd[1]/tbar[0]/btn[19]").press()
                    time.sleep(0.3)
                    
                    # Select row 10 (MRP 1)
                    table = self.session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW")
                    table.getAbsoluteRow(10).selected = True
                    time.sleep(0.3)
                    
                    # Nhấn Enter để confirm
                    self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                    time.sleep(0.5)
                    self.sap.wait_ready(self.session, 10)
                    log.info(f"    Selected MRP 1 (row 10) from popup")
                except Exception as e:
                    return {'ok': False, 'message': f'Lỗi chọn View MRP 1: {e}'}
            
            # 4. Xử lý popup nhập Plant
            popup2 = self.sap.safe_find(self.session, "wnd[1]")
            if popup2:
                try:
                    werks_field = self.sap.safe_find(self.session, self.WERKS_POPUP_ID)
                    if werks_field:
                        werks_field.text = werks
                        # Nhấn Enter để confirm
                        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        time.sleep(0.5)
                        self.sap.wait_ready(self.session, 10)
                        log.info(f"    Entered Plant: {werks}")
                    else:
                        return {'ok': False, 'message': 'Không tìm thấy field Plant trong popup'}
                except Exception as e:
                    return {'ok': False, 'message': f'Lỗi nhập Plant: {e}'}
            
            # Kiểm tra lỗi sau popup
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}
            
            # 5. Nhập Maximum Lot Size (Tab SP12)
            bstma_field = self.sap.safe_find(self.session, self.BSTMA_ID)
            if bstma_field is None:
                return {'ok': False, 'message': 'Không tìm thấy field Maximum Lot Size'}
            
            bstma_field.text = bstma
            bstma_field.setFocus()
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)
            self.sap.wait_ready(self.session, 10)
            log.info(f"    Đã nhập BSTMA: {bstma}")
            
            # Kiểm tra lỗi sau khi nhập
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}
            
            # 6. Save (btn[11])
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

    def _process_single_extwg(self, matnr: str, extwg: str) -> dict:
        """
        Xử lý cập nhật External Material Group cho 1 material
        
        Args:
            matnr: Material Number (VD: ARARXPB0001)
            extwg: External Material Group (VD: E-COIL/TRANSFORMER)
            
        Returns:
            dict với keys: ok, message
        """
        try:
            # 1. Nhập Material Number
            matnr_field = self.sap.safe_find(self.session, "wnd[0]/usr/ctxtRMMG1-MATNR")
            if matnr_field is None:
                return {'ok': False, 'message': 'Không tìm thấy field Material Number'}
            
            matnr_field.text = matnr
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)
            self.sap.wait_ready(self.session, 10)
            
            # Kiểm tra lỗi (Material không tồn tại)
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}
            
            # 2. Xử lý popup Select View - Chọn Basic Data 1 và Basic Data 2
            popup = self.sap.safe_find(self.session, "wnd[1]")
            if popup:
                try:
                    table = self.session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW")
                    # Select row 0 (Basic Data 1) và row 1 (Basic Data 2)
                    table.getAbsoluteRow(0).selected = True
                    table.getAbsoluteRow(1).selected = True
                    time.sleep(0.3)
                    # Nhấn Enter để confirm
                    self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                    time.sleep(0.5)
                    self.sap.wait_ready(self.session, 10)
                    log.info(f"    Selected Basic Data 1 & 2 from popup")
                except Exception as e:
                    log.warning(f"    Popup error: {e}")
            
            # Kiểm tra lỗi sau popup
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}
            
            # 3. Nhập External Material Group (Tab SP01 - Basic Data 1)
            extwg_field = self.sap.safe_find(
                self.session, 
                "wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2004/subSUB2:SAPLMGD1:2001/ctxtMARA-EXTWG"
            )
            if extwg_field is None:
                return {'ok': False, 'message': 'Không tìm thấy field Ext. Matl Group'}
            
            extwg_field.text = extwg
            extwg_field.setFocus()
            time.sleep(0.3)
            log.info(f"    Đã nhập EXTWG: {extwg}")
            
            # 4. Save (btn[11])
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

    def _process_materials(self, materials: list, tab: str) -> dict:
        """
        Xử lý danh sách materials cho tab được chỉ định
        Nếu gặp lỗi → DỪNG NGAY và báo cáo
        
        Returns:
            dict với keys: ok, status, message, processed, total, error_at_row, error_detail
        """
        processed = 0
        total = len(materials)
        
        for idx, item in enumerate(materials):
            material = str(item.get('material', '')).strip()
            description = str(item.get('description', '')).strip()
            row_number = idx + 2  # Dòng trong file (dòng 1 là header)
            
            if not material or not description:
                continue
            
            result = self._process_single_material(material, description, tab)
            
            if result['ok']:
                processed += 1
                log.info(f"  [MM02-{tab}] {material}: OK")
            else:
                # GẶP LỖI → DỪNG NGAY
                error_detail = result.get('message', 'Lỗi không xác định')
                log.error(f"  [MM02-{tab}] DỪNG tại dòng {row_number} ({material}): {error_detail}")
                
                return {
                    'ok': False,
                    'status': 'error',
                    'message': f"Đã xử lý {processed}/{total}. Lỗi tại dòng {row_number} ({material}): {error_detail}",
                    'processed': processed,
                    'total': total,
                    'error_at_row': row_number,
                    'error_detail': error_detail
                }
        
        # Tất cả thành công
        return {
            'ok': True,
            'status': 'success',
            'message': f"Hoàn tất. Đã cập nhật {processed}/{total} materials",
            'processed': processed,
            'total': total,
            'error_at_row': None,
            'error_detail': None
        }

    def _process_single_material(self, material: str, description: str, tab: str) -> dict:
        """
        Xử lý 1 material trong MM02
        """
        try:
            # 1. Mở TCode MM02
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nMM02"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)
            
            # 2. Nhập Material code
            self.session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = material
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)
            
            # Kiểm tra lỗi
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}
            
            # 3. Xử lý popup Select View (wnd[1]) - Chọn Basic Data 1
            try:
                popup = self.session.findById("wnd[1]")
                if popup:
                    table = self.session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW")
                    table.getAbsoluteRow(0).selected = True
                    time.sleep(1)
                    self.session.findById("wnd[1]").sendVKey(0)
                    time.sleep(1)
                    self.sap.wait_ready(self.session, 5)
                    log.info(f"    Selected Basic Data 1 from popup")
            except Exception as e:
                log.info(f"    No popup or error: {e}")
            
            # 4. Nhấn btn[30] để select view
            try:
                self.session.findById("wnd[0]/tbar[1]/btn[30]").press()
                time.sleep(1)
                self.sap.wait_ready(self.session, 5)
                log.info(f"    Pressed btn[30]")
            except Exception as e:
                log.warning(f"    btn[30] error: {e}")
            
            # 5. Chọn Tab và nhập text dựa vào tab parameter
            if tab == 'ZU05':
                longtext_id = self._select_tab_zu05()
            elif tab == 'ZU07':
                longtext_id = self._select_tab_zu07()
            else:
                return {'ok': False, 'message': f'Tab không hợp lệ: {tab}'}
            
            if not longtext_id:
                return {'ok': False, 'message': f'Không tìm thấy Tab {tab}'}
            
            # 6. Nhập mô tả
            try:
                self.session.findById(longtext_id).text = description
                time.sleep(1)
                log.info(f"    Đã nhập mô tả: {description[:50]}...")
            except Exception as e:
                return {'ok': False, 'message': f'Không tìm thấy field Long Text: {e}'}
            
            # 7. Nhấn Save (btn[11])
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)
            
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}
            
            # 8. Nhấn Back (btn[3])
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            time.sleep(1)
            self.sap.wait_ready(self.session, 5)
            
            return {'ok': True, 'message': 'OK'}
            
        except Exception as e:
            return {'ok': False, 'message': str(e)}

    def _select_tab_zu05(self) -> str:
        """Chọn Tab ZU05 và trả về ID của field Long Text"""
        try:
            self.session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU05").select()
            time.sleep(1)
            self.sap.wait_ready(self.session, 5)
            log.info(f"    Selected Tab ZU05")
            return "wnd[0]/usr/tabsTABSPR1/tabpZU05/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:2031/cntlLONGTEXT_GRUNDD/shellcont/shell"
        except Exception as e:
            log.error(f"    Tab ZU05 error: {e}")
            return None

    def _select_tab_zu07(self) -> str:
        """Chọn Tab ZU07 và trả về ID của field Long Text"""
        try:
            self.session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU07").select()
            time.sleep(1)
            self.sap.wait_ready(self.session, 5)
            log.info(f"    Selected Tab ZU07")
            return "wnd[0]/usr/tabsTABSPR1/tabpZU07/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:2051/cntlLONGTEXT_IVERM/shellcont/shell"
        except Exception as e:
            log.error(f"    Tab ZU07 error: {e}")
            return None