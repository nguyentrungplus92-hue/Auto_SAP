"""
TCode MK01 - Create Vendor (Maker)
==================================
Tạo mới Vendor/Maker trong SAP
Chạy tất cả dòng, ghi nhận lỗi, không dừng giữa chừng
"""
import time
import logging

log = logging.getLogger(__name__)


class TCodeMK01:
    # Fixed values
    EKORG = "C00A"
    KTOKK = "V090"
    SORT1 = "1"
    WAERS = "USD"
    WEBRE = True
    SESSION_DEAD_MSG = 'SAP session bị đóng — user đã login ở nơi khác hoặc mất kết nối'

    def __init__(self, sap_client):
        self.sap = sap_client
        self.session = sap_client.session

    def _safe_reset(self):
        for _ in range(3):
            try:
                self.session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
                time.sleep(0.1)
                continue
            except:
                pass
            try:
                self.session.findById("wnd[1]/tbar[0]/btn[1]").press()
                time.sleep(0.1)
                continue
            except:
                pass
            try:
                self.session.findById("wnd[1]").sendVKey(12)
                time.sleep(0.1)
                continue
            except:
                pass
            break
        try:
            self.session.findById("wnd[0]").sendVKey(12)
            time.sleep(0.1)
        except:
            pass
        try:
            self.session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
            time.sleep(0.1)
        except:
            pass
        try:
            self.session.findById("wnd[1]").sendVKey(12)
            time.sleep(0.1)
        except:
            pass
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.3)
        except:
            pass
        try:
            self.session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
            time.sleep(0.1)
        except:
            pass
        try:
            self.session.findById("wnd[1]").sendVKey(12)
            time.sleep(0.1)
        except:
            pass

    def run(self, vendors: list) -> dict:
        processed = 0
        total = len(vendors)
        errors = []

        for idx, item in enumerate(vendors):
            lifnr   = str(item.get('lifnr',   '')).strip()
            name1   = str(item.get('name1',   '')).strip()
            name2   = str(item.get('name2',   '')).strip() if item.get('name2')   else ''
            emnfr   = str(item.get('emnfr',   '')).strip() if item.get('emnfr')   else ''
            street  = str(item.get('street',  '')).strip() if item.get('street')  else ''
            city2   = str(item.get('city2',   '')).strip() if item.get('city2')   else ''
            city1   = str(item.get('city1',   '')).strip() if item.get('city1')   else ''
            country = str(item.get('country', '')).strip() if item.get('country') else ''
            stcd3   = str(item.get('stcd3',   '')).strip() if item.get('stcd3')   else ''
            stcd4   = str(item.get('stcd4',   '')).strip() if item.get('stcd4')   else ''
            inco1   = str(item.get('inco1',   '')).strip() if item.get('inco1')   else ''
            inco2   = str(item.get('inco2',   '')).strip() if item.get('inco2')   else ''
            row_number = idx + 2

            # ===== Kiểm tra session trước mỗi dòng =====
            if not self.sap.is_session_alive():
                for remaining_idx in range(idx, len(vendors)):
                    r = vendors[remaining_idx]
                    r_lifnr = str(r.get('lifnr', '')).strip()
                    r_name1 = str(r.get('name1', '')).strip()
                    errors.append({'row': remaining_idx + 2, 'info': f"{r_lifnr} - {r_name1}" or '-', 'detail': self.SESSION_DEAD_MSG})
                break
            # ============================================

            if not lifnr or not name1 or not country:
                continue

            try:
                result = self._process_single_vendor(
                    lifnr, name1, name2, emnfr, street, city2, city1,
                    country, stcd3, stcd4, inco1, inco2
                )
            except Exception as e:
                log.error(f"  [MK01] Exception bất ngờ dòng {row_number} ({lifnr}): {e}")
                result = {'ok': False, 'message': f"Exception: {str(e)}"}
                try:
                    self._safe_reset()
                except:
                    pass

            if result['ok']:
                processed += 1
                log.info(f"  [MK01] {lifnr}: OK - {name1}")
            else:
                error_detail = result.get('message', 'Lỗi không xác định')
                log.warning(f"  [MK01] Lỗi dòng {row_number} ({lifnr}): {error_detail}")
                errors.append({'row': row_number, 'info': f"{lifnr} - {name1}", 'detail': error_detail})

        if errors:
            error_lines = "\n".join([f"  • Dòng {e['row']} ({e['info']}): {e['detail']}" for e in errors])
            message = f"Hoàn tất {processed}/{total} dòng thành công. Lỗi {len(errors)} dòng:\n{error_lines}"
        else:
            message = f"Hoàn tất. Đã tạo {processed}/{total} vendors"

        return {'ok': True, 'message': message, 'processed': processed, 'total': total, 'errors': errors}

    def _process_single_vendor(self, lifnr, name1, name2, emnfr, street, city2, city1,
                                country, stcd3, stcd4, inco1, inco2) -> dict:
        try:
            self._safe_reset()

            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nMK01"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)
            self.sap.wait_ready(self.session, 10)

            try:
                self.session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
                time.sleep(0.1)
            except:
                pass
            try:
                self.session.findById("wnd[1]").sendVKey(12)
                time.sleep(0.1)
            except:
                pass

            self.session.findById("wnd[0]/usr/ctxtRF02K-LIFNR").text = lifnr
            self.session.findById("wnd[0]/usr/ctxtRF02K-EKORG").text = self.EKORG
            self.session.findById("wnd[0]/usr/ctxtRF02K-KTOKK").text = self.KTOKK
            self.session.findById("wnd[0]/usr/ctxtRF02K-KTOKK").setFocus()
            self.session.findById("wnd[0]/usr/ctxtRF02K-REF_LIFNR").text = ""
            self.session.findById("wnd[0]/usr/ctxtRF02K-REF_EKORG").text = ""

            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.1)
            self.sap.wait_ready(self.session, 10)

            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-NAME1").text = name1
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-NAME2").text = name2
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-SORT1").text = self.SORT1
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-STREET").text = street
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-POST_CODE1").text = ""
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-CITY1").text = city1
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-COUNTRY").text = country
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-COUNTRY").setFocus()

            try:
                self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/btnG_D0100_DUMMY_COUNTRY").press()
                time.sleep(0.3)
            except:
                pass

            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-CITY2").text = city2
            self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-CITY2").setFocus()

            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            time.sleep(0.3)
            self.sap.wait_ready(self.session, 10)

            try:
                self.session.findById("wnd[0]/usr/txtLFA1-STCD3").text = stcd3
            except:
                pass
            try:
                self.session.findById("wnd[0]/usr/txtLFA1-STCD4").text = stcd4
            except:
                pass
            try:
                self.session.findById("wnd[0]/usr/txtLFA1-EMNFR").text = emnfr
                self.session.findById("wnd[0]/usr/txtLFA1-EMNFR").setFocus()
            except:
                pass

            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.3)

            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            time.sleep(0.3)
            self.sap.wait_ready(self.session, 10)

            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            time.sleep(0.3)
            self.sap.wait_ready(self.session, 10)

            try:
                self.session.findById("wnd[0]/usr/chkLFM1-WEBRE").selected = self.WEBRE
            except:
                pass
            try:
                self.session.findById("wnd[0]/usr/ctxtLFM1-WAERS").text = self.WAERS
            except:
                pass
            try:
                self.session.findById("wnd[0]/usr/ctxtLFM1-INCO1").text = inco1
            except:
                pass
            try:
                self.session.findById("wnd[0]/usr/txtLFM1-INCO2").text = inco2
            except:
                pass

            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            time.sleep(0.5)
            self.sap.wait_ready(self.session, 10)

            try:
                self.session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
                time.sleep(0.1)
            except:
                pass

            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            return {'ok': True, 'message': 'OK'}

        except Exception as e:
            return {'ok': False, 'message': str(e)}
