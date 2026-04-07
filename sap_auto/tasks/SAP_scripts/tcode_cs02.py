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

    def __init__(self, sap_client):
        self.sap = sap_client
        self.session = sap_client.session

    def _safe_reset(self):
        """
        Đưa SAP về trạng thái sạch trước khi xử lý dòng tiếp theo.
        - Thử dismiss tất cả popup (wnd[1], wnd[2])
        - Navigate /n để cancel mọi thứ đang dở
        """
        # Bước 1: Thử dismiss popup wnd[1] nhiều lần
        for _ in range(3):
            try:
                # Thử nút "No / Không lưu"
                self.session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
                time.sleep(0.3)
                continue
            except:
                pass

            try:
                # Thử btn[1] (thường là "No")
                self.session.findById("wnd[1]/tbar[0]/btn[1]").press()
                time.sleep(0.3)
                continue
            except:
                pass

            try:
                # Escape wnd[1]
                self.session.findById("wnd[1]").sendVKey(12)
                time.sleep(0.3)
                continue
            except:
                pass

            break  # Không còn popup

        # Bước 2: Escape wnd[0] để cancel action đang dở
        try:
            self.session.findById("wnd[0]").sendVKey(12)
            time.sleep(0.3)
        except:
            pass

        # Bước 3: Dismiss popup nếu Escape lại trigger popup mới
        try:
            self.session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
            time.sleep(0.3)
        except:
            pass
        try:
            self.session.findById("wnd[1]").sendVKey(12)
            time.sleep(0.3)
        except:
            pass

        # Bước 4: Navigate /n để về màn hình sạch
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)
        except:
            pass

        # Bước 5: Dismiss lần cuối nếu còn popup
        try:
            self.session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
            time.sleep(0.3)
        except:
            pass
        try:
            self.session.findById("wnd[1]").sendVKey(12)
            time.sleep(0.3)
        except:
            pass

    def run(self, bom_items: list) -> dict:
        """
        Xử lý danh sách BOM items — mỗi dòng xử lý độc lập
        Chạy tất cả, thu thập lỗi, không dừng giữa chừng

        Args:
            bom_items: List dict với keys:
                - matnr, werks, stlan, datuv, idnrk, menge, meins, postp, lgort

        Returns:
            dict với keys: ok, message, processed, total, errors
        """
        processed = 0
        total = len(bom_items)
        errors = []

        for idx, item in enumerate(bom_items):
            row_number = idx + 2  # Dòng trong file (dòng 1 là header)

            matnr = str(item.get('matnr', '')).strip()
            werks = str(item.get('werks', '')).strip()
            idnrk = str(item.get('idnrk', '')).strip()

            if not matnr or not werks or not idnrk:
                continue

            # Wrap toàn bộ trong try/except để đảm bảo loop không bao giờ bị vỡ
            try:
                result = self._process_single_row(item, row_number)
            except Exception as e:
                log.error(f"  [CS02] Exception bất ngờ dòng {row_number}: {e}")
                result = {'ok': False, 'message': f"Exception: {str(e)}"}
                # Reset SAP về trạng thái sạch
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
                errors.append({
                    'row': row_number,
                    'info': f"{matnr}/{idnrk}",
                    'detail': error_detail,
                })
                # Tiếp tục dòng tiếp theo — KHÔNG dừng

        # Build message tổng kết
        if errors:
            error_lines = "\n".join([
                f"  • Dòng {e['row']} ({e['info']}): {e['detail']}"
                for e in errors
            ])
            message = f"Hoàn tất {processed}/{total} dòng thành công. Lỗi {len(errors)} dòng:\n{error_lines}"
        else:
            message = f"Hoàn tất. Đã thêm {processed}/{total} components vào BOM"

        return {
            'ok': True,
            'message': message,
            'processed': processed,
            'total': total,
            'errors': errors,
        }

    def _process_single_row(self, item: dict, row_number: int) -> dict:
        """
        Xử lý 1 dòng (1 component) trong CS02
        """
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

            # ===== 1. Reset SAP về trạng thái sạch trước khi bắt đầu =====
            self._safe_reset()

            # ===== 2. Mở TCode CS02 =====
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nCS02"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)

            # Dismiss popup nếu navigate CS02 trigger "unsaved changes"
            try:
                self.session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
                time.sleep(0.3)
            except:
                pass
            try:
                self.session.findById("wnd[1]").sendVKey(12)
                time.sleep(0.3)
            except:
                pass

            # ===== 3. Nhập BOM header =====
            self.session.findById("wnd[0]/usr/ctxtRC29N-MATNR").text = matnr
            self.session.findById("wnd[0]/usr/ctxtRC29N-WERKS").text = werks
            self.session.findById("wnd[0]/usr/ctxtRC29N-STLAN").text = stlan

            if datuv:
                self.session.findById("wnd[0]/usr/ctxtRC29N-DATUV").text = datuv
                self.session.findById("wnd[0]/usr/ctxtRC29N-DATUV").setFocus()

            # Enter 2 lần để vào BOM
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)

            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)

            # Kiểm tra lỗi header
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': f"BOM {matnr}/{werks}: {status.text}"}

            # ===== 4. Vào chế độ edit =====
            try:
                self.session.findById("wnd[0]/tbar[1]/btn[5]").press()
                time.sleep(1)
                self.sap.wait_ready(self.session, 10)
            except:
                pass

            # ===== 5. Thêm component vào dòng 1 =====
            row_idx = 1

            if postp:
                self.session.findById(f"wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/ctxtRC29P-POSTP[1,{row_idx}]").text = postp

            if idnrk:
                self.session.findById(f"wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/ctxtRC29P-IDNRK[2,{row_idx}]").text = idnrk

            if menge:
                self.session.findById(f"wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/txtRC29P-MENGE[4,{row_idx}]").text = menge

            if meins:
                self.session.findById(f"wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/ctxtRC29P-MEINS[5,{row_idx}]").text = meins

            # Enter để xác nhận
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)
            self.sap.wait_ready(self.session, 10)

            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': f"Component {idnrk}: {status.text}"}

            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)
            self.sap.wait_ready(self.session, 10)

            # ===== 6. Nhập Storage Location nếu có =====
            if lgort:
                try:
                    self.session.findById(f"wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/ctxtRC29P-POSTP[1,{row_idx}]").setFocus()
                    self.session.findById("wnd[0]").sendVKey(2)
                    time.sleep(0.5)
                    self.sap.wait_ready(self.session, 10)

                    self.session.findById("wnd[0]").sendVKey(0)
                    time.sleep(0.5)

                    try:
                        self.session.findById("wnd[0]/usr/tabsTS_ITEM/tabpPDAT").select()
                        time.sleep(0.5)
                    except:
                        pass

                    self.session.findById("wnd[0]/usr/tabsTS_ITEM/tabpPDAT/ssubSUBPAGE:SAPLCSDI:0840/ctxtRC29P-LGORT").text = lgort
                    self.session.findById("wnd[0]/usr/tabsTS_ITEM/tabpPDAT/ssubSUBPAGE:SAPLCSDI:0840/ctxtRC29P-LGORT").setFocus()

                    self.session.findById("wnd[0]").sendVKey(0)
                    time.sleep(0.5)

                    self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                    time.sleep(0.5)
                    self.sap.wait_ready(self.session, 10)

                except Exception as lgort_err:
                    return {'ok': False, 'message': f"LGORT {lgort}: {str(lgort_err)}"}

            # ===== 7. Save =====
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            time.sleep(2)
            self.sap.wait_ready(self.session, 10)

            # Dismiss popup sau save nếu có
            try:
                self.session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
                time.sleep(0.3)
            except:
                pass

            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': f"Save {matnr}/{idnrk}: {status.text}"}

            return {'ok': True, 'message': 'OK'}

        except Exception as e:
            return {'ok': False, 'message': str(e)}
