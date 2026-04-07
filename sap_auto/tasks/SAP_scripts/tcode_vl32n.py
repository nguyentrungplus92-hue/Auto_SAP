"""
TCode VL32N - Change Inbound Delivery
=====================================
Xóa Inbound Delivery
Thay đổi Delivery Date
Chạy tất cả dòng, ghi nhận lỗi, không dừng giữa chừng
"""
import time
import logging

log = logging.getLogger(__name__)


class TCodeVL32N:
    """
    VL32N - Change Inbound Delivery
    - Xóa Inbound Delivery
    - Thay đổi Delivery Date
    Chạy tất cả dòng, thu thập lỗi, không dừng giữa chừng
    """

    # Field IDs
    VBELN_ID = "wnd[0]/usr/ctxtLIKP-VBELN"
    LFDAT_ID = "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV50A:1202/ctxtRV50A-LFDAT_LA"

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

    def run(self, deliveries: list) -> dict:
        """
        Xử lý danh sách Inbound Deliveries - Xóa
        Chạy tất cả, thu thập lỗi, không dừng giữa chừng

        Args:
            deliveries: List dict với keys:
                - delivery: Delivery Number (BẮT BUỘC)

        Returns:
            dict với keys: ok, message, processed, total, errors
        """
        processed = 0
        total = len(deliveries)
        errors = []

        for idx, item in enumerate(deliveries):
            delivery = str(item.get('delivery', '')).strip()
            row_number = idx + 2

            if not delivery:
                continue

            try:
                result = self._process_single_delivery(delivery)
            except Exception as e:
                log.error(f"  [VL32N] Exception bất ngờ dòng {row_number} ({delivery}): {e}")
                result = {'ok': False, 'message': f"Exception: {str(e)}"}
                try:
                    self._safe_reset()
                except:
                    pass

            if result['ok']:
                processed += 1
                log.info(f"  [VL32N] {delivery}: OK - Deleted")
            else:
                error_detail = result.get('message', 'Lỗi không xác định')
                log.warning(f"  [VL32N] Lỗi dòng {row_number} ({delivery}): {error_detail}")
                errors.append({
                    'row': row_number,
                    'info': delivery,
                    'detail': error_detail,
                })
                # Tiếp tục dòng tiếp theo

        if errors:
            error_lines = "\n".join([
                f"  • Dòng {e['row']} ({e['info']}): {e['detail']}"
                for e in errors
            ])
            message = f"Hoàn tất {processed}/{total} dòng thành công. Lỗi {len(errors)} dòng:\n{error_lines}"
        else:
            message = f"Hoàn tất. Đã xóa {processed}/{total} Inbound Deliveries"

        return {
            'ok': True,
            'message': message,
            'processed': processed,
            'total': total,
            'errors': errors,
        }

    def run_change_date(self, items: list) -> dict:
        """
        Xử lý danh sách Inbound Deliveries - Thay đổi Delivery Date
        Chạy tất cả, thu thập lỗi, không dừng giữa chừng

        Args:
            items: List dict với keys:
                - vbeln: Delivery Number (BẮT BUỘC)
                - lfdat_la: Delivery Date (dd.mm.yyyy) (BẮT BUỘC)

        Returns:
            dict với keys: ok, message, processed, total, errors
        """
        processed = 0
        total = len(items)
        errors = []

        for idx, item in enumerate(items):
            vbeln    = str(item.get('vbeln',    '')).strip()
            lfdat_la = str(item.get('lfdat_la', '')).strip()
            row_number = idx + 2

            if not vbeln or not lfdat_la:
                continue

            try:
                result = self._process_single_change_date(vbeln, lfdat_la)
            except Exception as e:
                log.error(f"  [VL32N] Exception bất ngờ dòng {row_number} ({vbeln}): {e}")
                result = {'ok': False, 'message': f"Exception: {str(e)}"}
                try:
                    self._safe_reset()
                except:
                    pass

            if result['ok']:
                processed += 1
                log.info(f"  [VL32N] {vbeln}: OK - Date: {lfdat_la}")
            else:
                error_detail = result.get('message', 'Lỗi không xác định')
                log.warning(f"  [VL32N] Lỗi dòng {row_number} ({vbeln}): {error_detail}")
                errors.append({
                    'row': row_number,
                    'info': vbeln,
                    'detail': error_detail,
                })
                # Tiếp tục dòng tiếp theo

        if errors:
            error_lines = "\n".join([
                f"  • Dòng {e['row']} ({e['info']}): {e['detail']}"
                for e in errors
            ])
            message = f"Hoàn tất {processed}/{total} dòng thành công. Lỗi {len(errors)} dòng:\n{error_lines}"
        else:
            message = f"Hoàn tất. Đã cập nhật ngày cho {processed}/{total} Inbound Deliveries"

        return {
            'ok': True,
            'message': message,
            'processed': processed,
            'total': total,
            'errors': errors,
        }

    def _process_single_delivery(self, delivery: str) -> dict:
        """
        Xử lý xóa 1 Inbound Delivery trong VL32N
        """
        try:
            # Reset SAP về trạng thái sạch
            self._safe_reset()

            # 1. Mở TCode VL32N
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nVL32N"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)

            # Dismiss popup nếu có
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

            # 2. Nhập Delivery Number
            delivery_field = self.sap.safe_find(self.session, self.VBELN_ID)
            if delivery_field is None:
                return {'ok': False, 'message': 'Không tìm thấy field Delivery Number'}

            delivery_field.text = delivery
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)

            # Kiểm tra lỗi (Delivery không tồn tại)
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            # 3. Nhấn Delete (btn[14])
            btn_delete = self.sap.safe_find(self.session, "wnd[0]/tbar[1]/btn[14]")
            if btn_delete is None:
                return {'ok': False, 'message': 'Không tìm thấy nút Delete'}

            btn_delete.press()
            time.sleep(0.5)
            self.sap.wait_ready(self.session, 10)

            # 4. Xử lý popup confirm - Nhấn Yes (SPOP-OPTION1)
            popup = self.sap.safe_find(self.session, "wnd[1]")
            if popup:
                btn_yes = self.sap.safe_find(self.session, "wnd[1]/usr/btnSPOP-OPTION1")
                if btn_yes:
                    btn_yes.press()
                    time.sleep(0.5)
                    self.sap.wait_ready(self.session, 10)
                else:
                    self.session.findById("wnd[1]").sendVKey(0)
                    time.sleep(0.5)
                    self.sap.wait_ready(self.session, 10)

            # 5. Kiểm tra kết quả
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            # 6. Save (btn[11])
            btn_save = self.sap.safe_find(self.session, "wnd[0]/tbar[0]/btn[11]")
            if btn_save:
                btn_save.press()
                time.sleep(1)
                self.sap.wait_ready(self.session, 10)

            # Dismiss popup sau save nếu có
            try:
                self.session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
                time.sleep(0.3)
            except:
                pass

            # Kiểm tra kết quả cuối cùng
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            return {'ok': True, 'message': 'OK'}

        except Exception as e:
            return {'ok': False, 'message': str(e)}

    def _process_single_change_date(self, vbeln: str, lfdat_la: str) -> dict:
        """
        Xử lý thay đổi Delivery Date cho 1 Inbound Delivery
        """
        try:
            # Reset SAP về trạng thái sạch
            self._safe_reset()

            # 1. Mở TCode VL32N
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nVL32N"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)

            # Dismiss popup nếu có
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

            # 2. Nhập Delivery Number
            delivery_field = self.sap.safe_find(self.session, self.VBELN_ID)
            if delivery_field is None:
                return {'ok': False, 'message': 'Không tìm thấy field Delivery Number'}

            delivery_field.text = vbeln
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)
            self.sap.wait_ready(self.session, 10)

            # Kiểm tra lỗi (Delivery không tồn tại)
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            # 3. Nhập Delivery Date
            date_field = self.sap.safe_find(self.session, self.LFDAT_ID)
            if date_field is None:
                return {'ok': False, 'message': 'Không tìm thấy field Delivery Date'}

            date_field.text = lfdat_la
            date_field.setFocus()

            # Enter 2 lần (theo script gốc)
            for i in range(2):
                self.session.findById("wnd[0]").sendVKey(0)
                time.sleep(0.3)
                self.sap.wait_ready(self.session, 5)

            # Kiểm tra lỗi sau khi nhập
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            # 4. Save (btn[11])
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
