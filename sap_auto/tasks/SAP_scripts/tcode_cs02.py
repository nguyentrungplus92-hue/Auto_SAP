"""
TCode CS02 - Change BOM
=======================
Upload/Set BOM PMG - Thêm component vào BOM
Xử lý từng dòng độc lập, ghi nhận lỗi, không dừng giữa chừng
"""
import time
import logging

log = logging.getLogger(__name__)


class TCodeCS02:
    """
    CS02 - Change BOM
    Thêm component vào BOM của Material
    Mỗi dòng = 1 lần mở CS02 độc lập → collect lỗi, không dừng giữa chừng
    """

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

    def run(self, bom_items: list) -> dict:
        processed = 0
        total = len(bom_items)
        errors = []

        for idx, item in enumerate(bom_items):
            row_number = idx + 2

            matnr = str(item.get('matnr', '')).strip()
            werks = str(item.get('werks', '')).strip()
            idnrk = str(item.get('idnrk', '')).strip()

            # ===== Kiểm tra session trước mỗi dòng =====
            if not self.sap.is_session_alive():
                for remaining_idx in range(idx, len(bom_items)):
                    r_item = bom_items[remaining_idx]
                    r_matnr = str(r_item.get('matnr', '')).strip()
                    r_idnrk = str(r_item.get('idnrk', '')).strip()
                    errors.append({'row': remaining_idx + 2, 'info': f"{r_matnr}/{r_idnrk}" or '-', 'detail': self.SESSION_DEAD_MSG})
                break
            # ============================================

            if not matnr or not werks or not idnrk:
                continue

            try:
                result = self._process_single_row(item, row_number)
            except Exception as e:
                log.error(f"  [CS02] Exception bất ngờ dòng {row_number}: {e}")
                result = {'ok': False, 'message': f"Exception: {str(e)}"}
                try:
                    self._safe_reset()
                except:
                    pass

            if result['ok']:
                processed += 1
                log.info(f"  [CS02] Dòng {row_number} ({matnr}/{idnrk}): OK")
            else:
                error_detail = result.get('message', 'Lỗi không xác định')
                log.warning(f"  [CS02] Lỗi dòng {row_number} ({matnr}/{idnrk}): {error_detail}")
                errors.append({'row': row_number, 'info': f"{matnr}/{idnrk}", 'detail': error_detail})

        if errors:
            error_lines = "\n".join([f"  • Dòng {e['row']} ({e['info']}): {e['detail']}" for e in errors])
            message = f"Hoàn tất {processed}/{total} dòng thành công. Lỗi {len(errors)} dòng:\n{error_lines}"
        else:
            message = f"Hoàn tất. Đã thêm {processed}/{total} components vào BOM"

        return {'ok': True, 'message': message, 'processed': processed, 'total': total, 'errors': errors}

    def _process_single_row(self, item: dict, row_number: int) -> dict:
        try:
            matnr = str(item.get('matnr', '')).strip()
            werks = str(item.get('werks', '')).strip()
            stlan = str(item.get('stlan', '')).strip() or '1'
            datuv = str(item.get('datuv', '')).strip()
            idnrk = str(item.get('idnrk', '')).strip()
            menge = str(item.get('menge', '')).strip()
            meins = str(item.get('meins', '')).strip()
            postp = str(item.get('postp', '')).strip()
            lgort = str(item.get('lgort', '')).strip()

            self._safe_reset()

            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nCS02"
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

            self.session.findById("wnd[0]/usr/ctxtRC29N-MATNR").text = matnr
            self.session.findById("wnd[0]/usr/ctxtRC29N-WERKS").text = werks
            self.session.findById("wnd[0]/usr/ctxtRC29N-STLAN").text = stlan

            if datuv:
                self.session.findById("wnd[0]/usr/ctxtRC29N-DATUV").text = datuv
                self.session.findById("wnd[0]/usr/ctxtRC29N-DATUV").setFocus()

            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.3)
            self.sap.wait_ready(self.session, 10)

            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.3)
            self.sap.wait_ready(self.session, 10)

            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': f"BOM {matnr}/{werks}: {status.text}"}

            try:
                self.session.findById("wnd[0]/tbar[1]/btn[5]").press()
                time.sleep(0.3)
                self.sap.wait_ready(self.session, 10)
            except:
                pass

            row_idx = 1

            if postp:
                self.session.findById(f"wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/ctxtRC29P-POSTP[1,{row_idx}]").text = postp

            if idnrk:
                self.session.findById(f"wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/ctxtRC29P-IDNRK[2,{row_idx}]").text = idnrk

            if menge:
                self.session.findById(f"wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/txtRC29P-MENGE[4,{row_idx}]").text = menge

            if meins:
                self.session.findById(f"wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/ctxtRC29P-MEINS[5,{row_idx}]").text = meins

            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.3)
            self.sap.wait_ready(self.session, 10)

            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': f"Component {idnrk}: {status.text}"}

            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.3)
            self.sap.wait_ready(self.session, 10)

            if lgort:
                try:
                    self.session.findById(f"wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/ctxtRC29P-POSTP[1,{row_idx}]").setFocus()
                    self.session.findById("wnd[0]").sendVKey(2)
                    time.sleep(0.3)
                    self.sap.wait_ready(self.session, 10)

                    self.session.findById("wnd[0]").sendVKey(0)
                    time.sleep(0.3)

                    try:
                        self.session.findById("wnd[0]/usr/tabsTS_ITEM/tabpPDAT").select()
                        time.sleep(0.1)
                    except:
                        pass

                    self.session.findById("wnd[0]/usr/tabsTS_ITEM/tabpPDAT/ssubSUBPAGE:SAPLCSDI:0840/ctxtRC29P-LGORT").text = lgort
                    self.session.findById("wnd[0]/usr/tabsTS_ITEM/tabpPDAT/ssubSUBPAGE:SAPLCSDI:0840/ctxtRC29P-LGORT").setFocus()

                    self.session.findById("wnd[0]").sendVKey(0)
                    time.sleep(0.1)

                    self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                    time.sleep(0.1)
                    self.sap.wait_ready(self.session, 10)

                except Exception as lgort_err:
                    return {'ok': False, 'message': f"LGORT {lgort}: {str(lgort_err)}"}

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
                return {'ok': False, 'message': f"Save {matnr}/{idnrk}: {status.text}"}

            return {'ok': True, 'message': 'OK'}

        except Exception as e:
            return {'ok': False, 'message': str(e)}
