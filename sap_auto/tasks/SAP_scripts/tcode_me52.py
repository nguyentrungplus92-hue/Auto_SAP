"""
TCode ME52 - Change Purchase Requisition
=========================================
Xóa Item trong Purchase Requisition
Chạy tất cả dòng, ghi nhận lỗi, không dừng giữa chừng
"""
import time
import logging

log = logging.getLogger(__name__)


class TCodeME52:
    """
    ME52 - Change Purchase Requisition
    Xóa Item trong PR
    Chạy tất cả dòng, thu thập lỗi, không dừng giữa chừng
    """

    TABLE_ID = "wnd[0]/usr/tblSAPMM06BTC_0106"

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

    def run(self, pr_items: list) -> dict:
        """
        Xử lý danh sách PR Items - Xóa
        Chạy tất cả, thu thập lỗi, không dừng giữa chừng

        Args:
            pr_items: List dict với keys:
                - banfn: Purchase Requisition Number (BẮT BUỘC)
                - bnfpo: Item Number (BẮT BUỘC)

        Returns:
            dict với keys: ok, message, processed, total, errors
        """
        processed = 0
        total = len(pr_items)
        errors = []

        for idx, item in enumerate(pr_items):
            banfn = str(item.get('banfn', '')).strip()
            bnfpo = str(item.get('bnfpo', '')).strip()
            row_number = idx + 2  # Dòng trong file (dòng 1 là header)

            if not banfn or not bnfpo:
                continue

            # Wrap trong try/except để loop không bao giờ bị vỡ
            try:
                result = self._process_single_item(banfn, bnfpo)
            except Exception as e:
                log.error(f"  [ME52] Exception bất ngờ dòng {row_number} ({banfn}/{bnfpo}): {e}")
                result = {'ok': False, 'message': f"Exception: {str(e)}"}
                try:
                    self._safe_reset()
                except:
                    pass

            if result['ok']:
                processed += 1
                log.info(f"  [ME52] {banfn}/{bnfpo}: OK - Deleted")
            else:
                error_detail = result.get('message', 'Lỗi không xác định')
                log.warning(f"  [ME52] Lỗi dòng {row_number} ({banfn}/{bnfpo}): {error_detail}")
                errors.append({
                    'row': row_number,
                    'info': f"PR {banfn}/{bnfpo}",
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
            message = f"Hoàn tất. Đã xóa {processed}/{total} PR Items"

        return {
            'ok': True,
            'message': message,
            'processed': processed,
            'total': total,
            'errors': errors,
        }

    def _process_single_item(self, banfn: str, bnfpo: str) -> dict:
        """
        Xử lý xóa 1 PR Item trong ME52
        """
        try:
            # Reset SAP về trạng thái sạch
            self._safe_reset()

            # 1. Mở TCode ME52
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nME52"
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

            # 2. Nhập PR Number
            pr_field = self.sap.safe_find(self.session, "wnd[0]/usr/ctxtEBAN-BANFN")
            if pr_field is None:
                return {'ok': False, 'message': 'Không tìm thấy field PR Number'}

            pr_field.text = banfn
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)
            self.sap.wait_ready(self.session, 10)

            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            # 3. Nhập Item Number để filter
            item_field = self.sap.safe_find(self.session, "wnd[0]/usr/txtRM06B-BNFPO")
            if item_field is None:
                return {'ok': False, 'message': 'Không tìm thấy field Item Number'}

            item_field.text = bnfpo
            item_field.setFocus()
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)
            self.sap.wait_ready(self.session, 10)

            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            # 4. Select row đầu tiên trong table
            table = self.sap.safe_find(self.session, self.TABLE_ID)
            if table is None:
                return {'ok': False, 'message': 'Không tìm thấy table items'}

            try:
                table.getAbsoluteRow(0).selected = True
                time.sleep(0.3)
            except Exception as e:
                return {'ok': False, 'message': f'Không thể select row: {e}'}

            # 5. Nhấn Delete (btn[14])
            btn_delete = self.sap.safe_find(self.session, "wnd[0]/tbar[1]/btn[14]")
            if btn_delete is None:
                return {'ok': False, 'message': 'Không tìm thấy nút Delete'}

            btn_delete.press()
            time.sleep(0.5)
            self.sap.wait_ready(self.session, 10)

            # 6. Xử lý popup confirm - Nhấn Yes (SPOP-OPTION1)
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

            # 7. Kiểm tra lỗi sau khi delete
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            # 8. Save (btn[11])
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
