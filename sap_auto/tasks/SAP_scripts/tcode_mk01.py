"""
TCode MK01 - Create Vendor (Maker)
==================================
Tạo mới Vendor/Maker trong SAP
"""
import time
import logging

log = logging.getLogger(__name__)


class TCodeMK01:
    """
    MK01 - Create Vendor (Maker)
    Tạo mới Vendor với các thông tin cơ bản
    """

    # Fixed values
    EKORG = "C00A"      # Purchasing Organization
    KTOKK = "V090"      # Account Group
    SORT1 = "1"         # Search Term
    WAERS = "USD"       # Currency
    WEBRE = True        # GR-Based Inv. Verif. (checkbox)

    def __init__(self, sap_client):
        self.sap = sap_client
        self.session = sap_client.session

    def run(self, vendors: list) -> dict:
        """
        Xử lý danh sách vendors
        
        Args:
            vendors: List dict với keys:
                - lifnr: Vendor code (BẮT BUỘC)
                - name1: Name 1 (BẮT BUỘC)
                - name2: Name 2
                - emnfr: Manufacturer number
                - street: Street
                - city2: District
                - city1: City
                - country: Country (VD: JP) (BẮT BUỘC)
                - stcd3: Tax Number 3
                - stcd4: Tax Number 4
                - inco1: Incoterms 1 (VD: FOB)
                - inco2: Incoterms 2 (VD: JK)
            
        Returns:
            dict với keys: ok, message, processed, errors
        """
        processed = 0
        errors = []
        
        for item in vendors:
            lifnr = str(item.get('lifnr', '')).strip()
            name1 = str(item.get('name1', '')).strip()
            name2 = str(item.get('name2', '')).strip() if item.get('name2') else ''
            emnfr = str(item.get('emnfr', '')).strip() if item.get('emnfr') else ''
            street = str(item.get('street', '')).strip() if item.get('street') else ''
            city2 = str(item.get('city2', '')).strip() if item.get('city2') else ''
            city1 = str(item.get('city1', '')).strip() if item.get('city1') else ''
            country = str(item.get('country', '')).strip() if item.get('country') else ''
            stcd3 = str(item.get('stcd3', '')).strip() if item.get('stcd3') else ''
            stcd4 = str(item.get('stcd4', '')).strip() if item.get('stcd4') else ''
            inco1 = str(item.get('inco1', '')).strip() if item.get('inco1') else ''
            inco2 = str(item.get('inco2', '')).strip() if item.get('inco2') else ''
            
            # Bắt buộc phải có lifnr, name1, country
            if not lifnr or not name1 or not country:
                continue
                
            try:
                result = self._process_single_vendor(
                    lifnr, name1, name2, emnfr, street, city2, city1, 
                    country, stcd3, stcd4, inco1, inco2
                )
                if result['ok']:
                    processed += 1
                    log.info(f"  [MK01] {lifnr}: OK - {name1}")
                else:
                    errors.append(f"{lifnr}: {result['message']}")
                    log.warning(f"  [MK01] {lifnr}: {result['message']}")
            except Exception as e:
                errors.append(f"{lifnr}: {str(e)}")
                log.error(f"  [MK01] {lifnr}: Exception - {e}")
        
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
                'message': f"Hoàn tất. Đã tạo {processed} vendors",
                'processed': processed,
                'errors': []
            }

    def _process_single_vendor(self, lifnr, name1, name2, emnfr, street, city2, city1, 
                                country, stcd3, stcd4, inco1, inco2) -> dict:
        """
        Xử lý tạo 1 vendor trong MK01
        """
        try:
            # ===== Screen 1: Initial Screen =====
            # 1. Mở TCode MK01
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nMK01"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)
            
            # 2. Nhập Vendor code
            self.session.findById("wnd[0]/usr/ctxtRF02K-LIFNR").text = lifnr
            
            # 3. Nhập Purchasing Organization (fixed)
            self.session.findById("wnd[0]/usr/ctxtRF02K-EKORG").text = self.EKORG
            
            # 4. Nhập Account Group (fixed)
            self.session.findById("wnd[0]/usr/ctxtRF02K-KTOKK").text = self.KTOKK
            self.session.findById("wnd[0]/usr/ctxtRF02K-KTOKK").setFocus()
            
            # 5. Enter để tiếp tục
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)
            
            # Kiểm tra lỗi (Vendor đã tồn tại)
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}
            
            # ===== Screen 2: Address Screen =====
            # 6. Nhập Name 1
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-NAME1").text = name1
            
            # 7. Nhập Name 2
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-NAME2").text = name2
            
            # 8. Nhập Search Term (fixed)
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-SORT1").text = self.SORT1
            
            # 9. Nhập Street
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-STREET").text = street
            
            # 10. Nhập Post Code (để trống)
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-POST_CODE1").text = ""
            
            # 11. Nhập City
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-CITY1").text = city1
            
            # 12. Nhập Country
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-COUNTRY").text = country
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-COUNTRY").setFocus()
            
            # 13. Nhấn nút Country để apply
            try:
                self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/btnG_D0100_DUMMY_COUNTRY").press()
                time.sleep(0.5)
            except:
                pass
            
            # 14. Nhập District
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-CITY2").text = city2
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-CITY2").setFocus()
            
            # 15. Next screen (btn[8])
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)
            
            # ===== Screen 3: Control Screen =====
            # 16. Nhập Tax Number 3 (STCD3)
            try:
                self.session.findById("wnd[0]/usr/txtLFA1-STCD3").text = stcd3
            except:
                pass
            
            # 17. Nhập Tax Number 4 (STCD4)
            try:
                self.session.findById("wnd[0]/usr/txtLFA1-STCD4").text = stcd4
            except:
                pass
            
            # 18. Nhập Manufacturer Number (EMNFR)
            try:
                self.session.findById("wnd[0]/usr/txtLFA1-EMNFR").text = emnfr
                self.session.findById("wnd[0]/usr/txtLFA1-EMNFR").setFocus()
            except:
                pass
            
            # 19. Enter
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)
            
            # 20. Next screen (btn[8]) - 2 lần
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)
            
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)
            
            # ===== Screen 4: Purchasing Screen =====
            # 21. Tick checkbox GR-Based Inv. Verif. (WEBRE)
            try:
                self.session.findById("wnd[0]/usr/chkLFM1-WEBRE").selected = self.WEBRE
            except:
                pass
            
            # 22. Nhập Currency (fixed)
            try:
                self.session.findById("wnd[0]/usr/ctxtLFM1-WAERS").text = self.WAERS
            except:
                pass
            
            # 23. Nhập Incoterms 1
            try:
                self.session.findById("wnd[0]/usr/ctxtLFM1-INCO1").text = inco1
            except:
                pass
            
            # 24. Nhập Incoterms 2
            try:
                self.session.findById("wnd[0]/usr/txtLFM1-INCO2").text = inco2
            except:
                pass
            
            # 25. Save (btn[11])
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            time.sleep(2)
            self.sap.wait_ready(self.session, 10)
            
            # Kiểm tra kết quả save
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}
            
            return {'ok': True, 'message': 'OK'}
            
        except Exception as e:
            return {'ok': False, 'message': str(e)}