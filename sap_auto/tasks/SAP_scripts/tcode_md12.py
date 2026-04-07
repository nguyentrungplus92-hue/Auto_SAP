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

    def __init__(self, sap_client):
        self.sap = sap_client
        self.session = sap_client.session

    def _safe_reset(self):
        """
        Đưa SAP về trạng thái sạch trước khi xử lý dòng tiếp theo
        """
        # Thử dismiss popup wnd[1] nhiều lần
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

        # Escape wnd[0] để cancel action đang dở
        try:
            self.session.findById("wnd[0]").sendVKey(12)
            time.sleep(0.3)
        except:
            pass

        # Dismiss popup nếu Escape trigger thêm popup
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

        # Navigate /n về màn hình trắng
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)
        except:
            pass

        # Dismiss lần cuối nếu còn popup
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

    def run(self, planned_orders: list) -> dict:
        """
        Xử lý danh sách planned orders
        Chạy tất cả, thu thập lỗi, không dừng giữa chừng

        Args:
            planned_orders: List các planned order number

        Returns:
            dict với keys: ok, message, processed, total, errors
        """
        processed = 0
        total = len(planned_orders)
        errors = []

        for idx, order in enumerate(planned_orders):
            order = str(order).strip()
            row_number = idx + 2  # Dòng trong file (dòng 1 là header)

            if not order:
                continue

            # Wrap trong try/except để loop không bao giờ bị vỡ
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
                errors.append({
                    'row': row_number,
                    'info': order,
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
            message = f"Hoàn tất. Đã xử lý {processed}/{total} planned orders"

        return {
            'ok': True,
            'message': message,
            'processed': processed,
            'total': total,
            'errors': errors,
        }

    def _process_single_order(self, planned_order: str) -> dict:
        """
        Xử lý 1 planned order trong MD12
        """
        try:
            # Reset SAP về trạng thái sạch trước khi bắt đầu
            self._safe_reset()

            # 1. Mở TCode MD12
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nMD12"
            self.session.findById("wnd[0]").sendVKey(0)
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

            # 2. Nhập planned order number
            self.session.findById("wnd[0]/usr/txtRM61P-PLNUM").text = planned_order

            # 3. Nhấn Enter để mở planned order
            self.session.findById("wnd[0]").sendVKey(0)
            self.sap.wait_ready(self.session, 10)

            # Kiểm tra lỗi
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            # Dismiss popup nếu có
            self.sap.dismiss_popup_if_any()

            # 4. Bỏ tick checkbox "Firmly planned" (PLAF-AUFFX)
            checkbox_id = "wnd[0]/usr/tabsTABTC/tabpTAB01/ssubINCLUDE1XX:SAPLM61O:0711/subINCLUDE711_2:SAPLM61O:0810/chkPLAF-AUFFX"
            checkbox = self.sap.safe_find(self.session, checkbox_id)

            if checkbox is None:
                return {'ok': False, 'message': 'Không tìm thấy checkbox Firmly planned'}

            # Bỏ tick nếu đang được tick
            if checkbox.selected:
                checkbox.selected = False
                log.info(f"    Đã bỏ tick Firmly planned")
            else:
                log.info(f"    Checkbox đã bỏ tick sẵn")

            # 5. Nhấn Save (btn[11])
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            self.sap.wait_ready(self.session, 10)

            # Dismiss popup sau save nếu có
            try:
                self.session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
                time.sleep(0.3)
            except:
                pass

            # Kiểm tra kết quả save
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            # 6. Nhấn Back (btn[3])
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            self.sap.wait_ready(self.session, 5)

            return {'ok': True, 'message': 'OK'}

        except Exception as e:
            return {'ok': False, 'message': str(e)}
