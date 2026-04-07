"""
TCode QA11 - Record Usage Decision
==================================
Nhập Usage Decision cho Inspection Lot
Chạy tất cả dòng, ghi nhận lỗi, không dừng giữa chừng
"""
import time
import logging

log = logging.getLogger(__name__)


class TCodeQA11:
    """
    QA11 - Record Usage Decision
    Nhập Usage Decision Code cho Inspection Lot
    Chạy tất cả dòng, thu thập lỗi, không dừng giữa chừng
    """

    # Field IDs
    PRUEFLOS_ID = "wnd[0]/usr/ctxtQALS-PRUEFLOS"
    VCODE_ID = "wnd[0]/usr/tabsUD_DATA/tabpFEHL/ssubSUB_UD_DATA:SAPMQEVA:0103/subUD_DATA:SAPMQEVA:1103/ctxtRQEVA-VCODE"
    VCODEGRP_ID = "wnd[0]/usr/tabsUD_DATA/tabpFEHL/ssubSUB_UD_DATA:SAPMQEVA:0103/subUD_DATA:SAPMQEVA:1103/ctxtRQEVA-VCODEGRP"

    def __init__(self, sap_client):
        self.sap = sap_client
        self.session = sap_client.session

    def _safe_reset(self):
        """
        Đưa SAP về trạng thái sạch trước khi xử lý dòng tiếp theo
        """
        for _ in range(3):
            try:
                self.session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
                time.sleep(0.3)
                continue
            except:
                pass
            try:
                self.session.findById("wnd[1]/tbar[0]/btn[1]").press()
                time.sleep(0.3)
                continue
            except:
                pass
            try:
                self.session.findById("wnd[1]").sendVKey(12)
                time.sleep(0.3)
                continue
            except:
                pass
            break

        try:
            self.session.findById("wnd[0]").sendVKey(12)
            time.sleep(0.3)
        except:
            pass

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

        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)
        except:
            pass

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

    def run(self, items: list) -> dict:
        """
        Xử lý danh sách Inspection Lots - Nhập Usage Decision
        Chạy tất cả, thu thập lỗi, không dừng giữa chừng

        Args:
            items: List dict với keys:
                - prueflos: Inspection Lot Number (BẮT BUỘC)
                - vcode: Usage Decision Code (BẮT BUỘC)
                - vcodegrp: Code Group (BẮT BUỘC)

        Returns:
            dict với keys: ok, message, processed, total, errors
        """
        processed = 0
        total = len(items)
        errors = []

        for idx, item in enumerate(items):
            prueflos = str(item.get('prueflos', '')).strip()
            vcode    = str(item.get('vcode',    '')).strip()
            vcodegrp = str(item.get('vcodegrp', '')).strip()
            row_number = idx + 2

            if not prueflos or not vcode or not vcodegrp:
                continue

            # Wrap trong try/except để loop không bao giờ bị vỡ
            try:
                result = self._process_single_lot(prueflos, vcode, vcodegrp)
            except Exception as e:
                log.error(f"  [QA11] Exception bất ngờ dòng {row_number} ({prueflos}): {e}")
                result = {'ok': False, 'message': f"Exception: {str(e)}"}
                try:
                    self._safe_reset()
                except:
                    pass

            if result['ok']:
                processed += 1
                log.info(f"  [QA11] {prueflos}: OK - {vcode}/{vcodegrp}")
            else:
                error_detail = result.get('message', 'Lỗi không xác định')
                log.warning(f"  [QA11] Lỗi dòng {row_number} ({prueflos}): {error_detail}")
                errors.append({
                    'row': row_number,
                    'info': prueflos,
                    'detail': error_detail,
                })
                # Tiếp tục dòng tiếp theo

        # Build message tổng kết
        if errors:
            error_lines = "\n".join([
                f"  • Dòng {e['row']} ({e['info']}): {e['detail']}"
                for e in errors
            ])
            message = f"Hoàn tất {processed}/{total} dòng thành công. Lỗi {len(errors)} dòng:\n{error_lines}"
        else:
            message = f"Hoàn tất. Đã nhập UD cho {processed}/{total} Inspection Lots"

        return {
            'ok': True,
            'message': message,
            'processed': processed,
            'total': total,
            'errors': errors,
        }

    def _process_single_lot(self, prueflos: str, vcode: str, vcodegrp: str) -> dict:
        """
        Xử lý nhập Usage Decision cho 1 Inspection Lot
        """
        try:
            # Reset SAP về trạng thái sạch
            self._safe_reset()

            # 1. Mở TCode QA11
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nQA11"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)

            # Dismiss popup nếu navigate trigger "unsaved changes"
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

            # Dismiss popup sau save nếu có
            try:
                self.session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
                time.sleep(0.3)
            except:
                pass

            # Kiểm tra kết quả Save
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            return {'ok': True, 'message': 'OK'}

        except Exception as e:
            return {'ok': False, 'message': str(e)}
