"""
TCode MM02 - Change Material Master
===================================
Cập nhật mô tả tiếng Việt cho Material
"""
import time
import logging

log = logging.getLogger(__name__)


class TCodeMM02:
    """
    MM02 - Change Material Master
    Cập nhật Long Text Description (Tab ZU05, ZU07)
    """

    def __init__(self, sap_client):
        self.sap = sap_client
        self.session = sap_client.session

    def run(self, materials: list) -> dict:
        """
        Xử lý danh sách materials - Tab ZU05 (Basic data text)
        """
        return self._process_materials(materials, tab='ZU05')

    def run_zu07(self, materials: list) -> dict:
        """
        Xử lý danh sách materials - Tab ZU07 (Internal comment)
        """
        return self._process_materials(materials, tab='ZU07')

    def _process_materials(self, materials: list, tab: str) -> dict:
        """
        Xử lý danh sách materials cho tab được chỉ định
        """
        processed = 0
        errors = []
        
        for item in materials:
            material = str(item.get('material', '')).strip()
            description = str(item.get('description', '')).strip()
            
            if not material or not description:
                continue
                
            try:
                result = self._process_single_material(material, description, tab)
                if result['ok']:
                    processed += 1
                    log.info(f"  [MM02-{tab}] {material}: OK")
                else:
                    errors.append(f"{material}: {result['message']}")
                    log.warning(f"  [MM02-{tab}] {material}: {result['message']}")
            except Exception as e:
                errors.append(f"{material}: {str(e)}")
                log.error(f"  [MM02-{tab}] {material}: Exception - {e}")
        
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
                'message': f"Hoàn tất. Đã cập nhật {processed} materials",
                'processed': processed,
                'errors': []
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