"""
TCode MD12 - Change Planned Order
=================================
Bỏ tick checkbox "Firmly planned" cho planned order
Chạy tất cả dòng, ghi nhận lỗi, không dừng giữa chừng
"""
import time
import logging

log = logging.getLogger(__name__)


class TCodeMD12:
    """
    MD12 - Change Planned Order
    Bỏ tick "Fixierung" (Firmly planned) checkbox
    Chạy tất cả dòng, thu thập lỗi, không dừng giữa chừng
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

    def run(self, planned_orders: list) -> dict:
        processed = 0
        total = len(planned_orders)
        errors = []

        for idx, order in enumerate(planned_orders):
            order = str(order).strip()
            row_number = idx + 2

            # ===== Kiểm tra session trước mỗi dòng =====
            if not self.sap.is_session_alive():
                for remaining_idx in range(idx, len(planned_orders)):
                    r_order = str(planned_orders[remaining_idx]).strip()
                    errors.append({'row': remaining_idx + 2, 'info': r_order or '-', 'detail': self.SESSION_DEAD_MSG})
                break
            # ============================================

            if not order:
                continue

            try:
                result = self._process_single_order(order)
            except Exception as e:
                log.error(f"  [MD12] Exception bất ngờ dòng {row_number} ({order}): {e}")
                result = {'ok': False, 'message': f"Exception: {str(e)}"}
                try:
                    self._safe_reset()
                except:
                    pass

            if result['ok']:
                processed += 1
                log.info(f"  [MD12] {order}: OK")
            else:
                error_detail = result.get('message', 'Lỗi không xác định')
                log.warning(f"  [MD12] Lỗi dòng {row_number} ({order}): {error_detail}")
                errors.append({'row': row_number, 'info': order, 'detail': error_detail})

        if errors:
            error_lines = "\n".join([f"  • Dòng {e['row']} ({e['info']}): {e['detail']}" for e in errors])
            message = f"Hoàn tất {processed}/{total} dòng thành công. Lỗi {len(errors)} dòng:\n{error_lines}"
        else:
            message = f"Hoàn tất. Đã xử lý {processed}/{total} planned orders"

        return {'ok': True, 'message': message, 'processed': processed, 'total': total, 'errors': errors}

    def _process_single_order(self, planned_order: str) -> dict:
        try:
            self._safe_reset()

            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nMD12"
            self.session.findById("wnd[0]").sendVKey(0)
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

            self.session.findById("wnd[0]/usr/txtRM61P-PLNUM").text = planned_order
            self.session.findById("wnd[0]").sendVKey(0)
            self.sap.wait_ready(self.session, 10)

            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            self.sap.dismiss_popup_if_any()

            checkbox_id = "wnd[0]/usr/tabsTABTC/tabpTAB01/ssubINCLUDE1XX:SAPLM61O:0711/subINCLUDE711_2:SAPLM61O:0810/chkPLAF-AUFFX"
            checkbox = self.sap.safe_find(self.session, checkbox_id)

            if checkbox is None:
                return {'ok': False, 'message': 'Không tìm thấy checkbox Firmly planned'}

            if checkbox.selected:
                checkbox.selected = False
                log.info(f"    Đã bỏ tick Firmly planned")
            else:
                log.info(f"    Checkbox đã bỏ tick sẵn")

            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            self.sap.wait_ready(self.session, 10)

            try:
                self.session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
                time.sleep(0.1)
            except:
                pass

            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            self.sap.wait_ready(self.session, 5)

            return {'ok': True, 'message': 'OK'}

        except Exception as e:
            return {'ok': False, 'message': str(e)}
