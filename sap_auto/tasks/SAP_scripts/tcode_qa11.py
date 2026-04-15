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

    PRUEFLOS_ID = "wnd[0]/usr/ctxtQALS-PRUEFLOS"
    SESSION_DEAD_MSG = 'SAP session bị đóng — user đã login ở nơi khác hoặc mất kết nối'

    # Không hardcode path — dùng _find_by_suffix() tìm theo tên field

    def __init__(self, sap_client):
        self.sap = sap_client
        self.session = sap_client.session

    def _find_by_suffix(self, container, suffix: str, max_depth: int = 8):
        """
        Tìm field theo tên cuối của ID (VD: 'ctxtRQEVA-VCODE').
        Duyệt đệ quy qua cây component — không cần biết full path,
        dù SAP sinh path mới vẫn tìm được.
        """
        if max_depth <= 0:
            return None
        try:
            count = container.Children.Count
        except:
            return None
        for i in range(count):
            try:
                child = container.Children(i)
                child_id = child.Id or ""
                if child_id.endswith(suffix):
                    log.info(f"    Found by suffix '{suffix}': {child_id}")
                    return child
                found = self._find_by_suffix(child, suffix, max_depth - 1)
                if found is not None:
                    return found
            except:
                continue
        return None

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

    def run(self, items: list) -> dict:
        processed = 0
        total = len(items)
        errors = []

        for idx, item in enumerate(items):
            prueflos = str(item.get('prueflos', '')).strip()
            vcode    = str(item.get('vcode',    '')).strip()
            vcodegrp = str(item.get('vcodegrp', '')).strip()
            row_number = idx + 2

            # ===== Kiểm tra session trước mỗi dòng =====
            if not self.sap.is_session_alive():
                for remaining_idx in range(idx, len(items)):
                    r = items[remaining_idx]
                    r_prueflos = str(r.get('prueflos', '')).strip()
                    errors.append({'row': remaining_idx + 2, 'info': r_prueflos or '-', 'detail': self.SESSION_DEAD_MSG})
                break
            # ============================================

            if not prueflos or not vcode or not vcodegrp:
                continue

            try:
                result = self._process_single_lot(prueflos, vcode, vcodegrp)
            except Exception as e:
                log.error(f"  [QA11] Exception dòng {row_number} ({prueflos}): {e}")
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
                errors.append({'row': row_number, 'info': prueflos, 'detail': error_detail})

        if errors:
            error_lines = "\n".join([f"  • Dòng {e['row']} ({e['info']}): {e['detail']}" for e in errors])
            message = f"Hoàn tất {processed}/{total} dòng thành công. Lỗi {len(errors)} dòng:\n{error_lines}"
        else:
            message = f"Hoàn tất. Đã nhập UD cho {processed}/{total} Inspection Lots"

        return {'ok': True, 'message': message, 'processed': processed, 'total': total, 'errors': errors}

    def _process_single_lot(self, prueflos: str, vcode: str, vcodegrp: str) -> dict:
        try:
            self._safe_reset()

            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nQA11"
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

            lot_field = self.sap.safe_find(self.session, self.PRUEFLOS_ID)
            if lot_field is None:
                return {'ok': False, 'message': 'Không tìm thấy field Inspection Lot'}

            lot_field.text = prueflos
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.3)
            self.sap.wait_ready(self.session, 10)

            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            ud_container = self.sap.safe_find(self.session, "wnd[0]/usr/tabsUD_DATA")
            if ud_container is None:
                return {'ok': False, 'message': 'Không tìm thấy tabsUD_DATA trong màn hình QA11'}

            vcode_field = self._find_by_suffix(ud_container, "ctxtRQEVA-VCODE")
            if vcode_field is None:
                return {'ok': False, 'message': 'Không tìm thấy field VCODE trong tabsUD_DATA'}

            vcode_field.text = vcode
            time.sleep(0.1)

            vcodegrp_field = self._find_by_suffix(ud_container, "ctxtRQEVA-VCODEGRP")
            if vcodegrp_field is None:
                return {'ok': False, 'message': 'Không tìm thấy field VCODEGRP trong tabsUD_DATA'}

            vcodegrp_field.text = vcodegrp
            vcodegrp_field.setFocus()
            vcodegrp_field.caretPosition = len(vcodegrp)
            time.sleep(0.1)

            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            time.sleep(0.5)
            self.sap.wait_ready(self.session, 10)

            try:
                self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                time.sleep(0.1)
                self.sap.wait_ready(self.session, 10)
            except:
                pass

            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            return {'ok': True, 'message': 'OK'}

        except Exception as e:
            return {'ok': False, 'message': str(e)}
