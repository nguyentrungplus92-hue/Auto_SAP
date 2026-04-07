"""
TCode MM02 - Change Material Master
===================================
Cập nhật mô tả tiếng Việt cho Material
Cập nhật External Material Group
Cập nhật Maximum Lot Size
Chạy tất cả dòng, ghi nhận lỗi, không dừng giữa chừng
"""
import time
import logging

log = logging.getLogger(__name__)


class TCodeMM02:
    """
    MM02 - Change Material Master
    Cập nhật Long Text Description (Tab ZU05, ZU07)
    Cập nhật External Material Group (Tab SP01)
    Cập nhật Maximum Lot Size (Tab SP12 - MRP 1)
    Chạy tất cả dòng, thu thập lỗi, không dừng giữa chừng
    """

    # Field IDs
    MATNR_ID = "wnd[0]/usr/ctxtRMMG1-MATNR"
    WERKS_POPUP_ID = "wnd[1]/usr/ctxtRMMG1-WERKS"
    BSTMA_ID = "wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2483/txtMARC-BSTMA"

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

    # ===== Public run methods =====

    def run(self, materials: list) -> dict:
        """Tab ZU05 (Basic data text)"""
        return self._process_materials(materials, tab='ZU05')

    def run_zu07(self, materials: list) -> dict:
        """Tab ZU07 (Internal comment)"""
        return self._process_materials(materials, tab='ZU07')

    def run_extwg(self, materials: list) -> dict:
        """
        Cập nhật External Material Group
        Chạy tất cả, thu thập lỗi, không dừng giữa chừng

        Args:
            materials: List dict với keys:
                - matnr: Material Number (BẮT BUỘC)
                - extwg: External Material Group (BẮT BUỘC)
        """
        processed = 0
        total = len(materials)
        errors = []

        for idx, item in enumerate(materials):
            matnr = str(item.get('matnr', '')).strip()
            extwg = str(item.get('extwg', '')).strip()
            row_number = idx + 2

            if not matnr or not extwg:
                continue

            try:
                result = self._process_single_extwg(matnr, extwg)
            except Exception as e:
                log.error(f"  [MM02-EXTWG] Exception bất ngờ dòng {row_number} ({matnr}): {e}")
                result = {'ok': False, 'message': f"Exception: {str(e)}"}
                try:
                    self._safe_reset()
                except:
                    pass

            if result['ok']:
                processed += 1
                log.info(f"  [MM02-EXTWG] {matnr}: OK - {extwg}")
            else:
                error_detail = result.get('message', 'Lỗi không xác định')
                log.warning(f"  [MM02-EXTWG] Lỗi dòng {row_number} ({matnr}): {error_detail}")
                errors.append({
                    'row': row_number,
                    'info': matnr,
                    'detail': error_detail,
                })

        if errors:
            error_lines = "\n".join([
                f"  • Dòng {e['row']} ({e['info']}): {e['detail']}"
                for e in errors
            ])
            message = f"Hoàn tất {processed}/{total} dòng thành công. Lỗi {len(errors)} dòng:\n{error_lines}"
        else:
            message = f"Hoàn tất. Đã cập nhật {processed}/{total} materials"

        return {
            'ok': True,
            'message': message,
            'processed': processed,
            'total': total,
            'errors': errors,
        }

    def run_bstma(self, materials: list) -> dict:
        """
        Cập nhật Maximum Lot Size (Tab SP12 - MRP 1)
        Chạy tất cả, thu thập lỗi, không dừng giữa chừng

        Args:
            materials: List dict với keys:
                - matnr: Material Number (BẮT BUỘC)
                - werks: Plant (BẮT BUỘC)
                - bstma: Maximum Lot Size (BẮT BUỘC)
        """
        processed = 0
        total = len(materials)
        errors = []

        for idx, item in enumerate(materials):
            matnr = str(item.get('matnr', '')).strip()
            werks = str(item.get('werks', '')).strip()
            bstma = str(item.get('bstma', '')).strip()
            row_number = idx + 2

            if not matnr or not werks or not bstma:
                continue

            try:
                result = self._process_single_bstma(matnr, werks, bstma)
            except Exception as e:
                log.error(f"  [MM02-BSTMA] Exception bất ngờ dòng {row_number} ({matnr}/{werks}): {e}")
                result = {'ok': False, 'message': f"Exception: {str(e)}"}
                try:
                    self._safe_reset()
                except:
                    pass

            if result['ok']:
                processed += 1
                log.info(f"  [MM02-BSTMA] {matnr}/{werks}: OK - {bstma}")
            else:
                error_detail = result.get('message', 'Lỗi không xác định')
                log.warning(f"  [MM02-BSTMA] Lỗi dòng {row_number} ({matnr}/{werks}): {error_detail}")
                errors.append({
                    'row': row_number,
                    'info': f"{matnr}/{werks}",
                    'detail': error_detail,
                })

        if errors:
            error_lines = "\n".join([
                f"  • Dòng {e['row']} ({e['info']}): {e['detail']}"
                for e in errors
            ])
            message = f"Hoàn tất {processed}/{total} dòng thành công. Lỗi {len(errors)} dòng:\n{error_lines}"
        else:
            message = f"Hoàn tất. Đã cập nhật Maximum Lot Size cho {processed}/{total} materials"

        return {
            'ok': True,
            'message': message,
            'processed': processed,
            'total': total,
            'errors': errors,
        }

    # ===== Internal shared loop for ZU05 / ZU07 =====

    def _process_materials(self, materials: list, tab: str) -> dict:
        """
        Xử lý danh sách materials cho tab ZU05 hoặc ZU07
        Chạy tất cả, thu thập lỗi, không dừng giữa chừng
        """
        processed = 0
        total = len(materials)
        errors = []

        for idx, item in enumerate(materials):
            material    = str(item.get('material',    '')).strip()
            description = str(item.get('description', '')).strip()
            row_number  = idx + 2

            if not material or not description:
                continue

            try:
                result = self._process_single_material(material, description, tab)
            except Exception as e:
                log.error(f"  [MM02-{tab}] Exception bất ngờ dòng {row_number} ({material}): {e}")
                result = {'ok': False, 'message': f"Exception: {str(e)}"}
                try:
                    self._safe_reset()
                except:
                    pass

            if result['ok']:
                processed += 1
                log.info(f"  [MM02-{tab}] {material}: OK")
            else:
                error_detail = result.get('message', 'Lỗi không xác định')
                log.warning(f"  [MM02-{tab}] Lỗi dòng {row_number} ({material}): {error_detail}")
                errors.append({
                    'row': row_number,
                    'info': material,
                    'detail': error_detail,
                })

        if errors:
            error_lines = "\n".join([
                f"  • Dòng {e['row']} ({e['info']}): {e['detail']}"
                for e in errors
            ])
            message = f"Hoàn tất {processed}/{total} dòng thành công. Lỗi {len(errors)} dòng:\n{error_lines}"
        else:
            message = f"Hoàn tất. Đã cập nhật {processed}/{total} materials"

        return {
            'ok': True,
            'message': message,
            'processed': processed,
            'total': total,
            'errors': errors,
        }

    # ===== Single-item process methods =====

    def _process_single_material(self, material: str, description: str, tab: str) -> dict:
        """Xử lý 1 material trong MM02 (ZU05 / ZU07)"""
        try:
            # Reset SAP về trạng thái sạch
            self._safe_reset()

            # 1. Mở TCode MM02
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nMM02"
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

            # 2. Nhập Material code + caretPosition — theo script gốc
            matnr_field = self.session.findById("wnd[0]/usr/ctxtRMMG1-MATNR")
            matnr_field.text = material
            matnr_field.caretPosition = len(material)

            # btn[5] = Select Org. Levels → mở thẳng popup Select View
            self.session.findById("wnd[0]/tbar[1]/btn[5]").press()
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)

            # Kiểm tra lỗi (Material không tồn tại)
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            # 3. Xử lý popup Select View
            # btn[5] luôn mở Select View → xử lý trực tiếp, không dùng if popup
            try:
                # btn[19]: Deselect all trước — tránh chọn nhầm view từ lần trước
                self.session.findById("wnd[1]/tbar[0]/btn[19]").press()
                time.sleep(0.3)

                # Select row 0 (Basic Data 1) và row 1 (Basic Data 2)
                table = self.session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW")
                table.getAbsoluteRow(0).selected = True
                table.getAbsoluteRow(1).selected = True

                # setFocus + caretPosition = 0 vào row 1 — theo script gốc
                row1_field = self.session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,1]")
                row1_field.setFocus()
                row1_field.caretPosition = 0
                time.sleep(0.3)

                # Confirm
                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                time.sleep(0.5)
                self.sap.wait_ready(self.session, 10)
                log.info(f"    Selected Basic Data 1 & 2 from popup")
            except Exception as e:
                return {'ok': False, 'message': f'Lỗi xử lý popup Select View: {e}'}

            # 4. Nhấn btn[30] để vào màn hình tab — bắt buộc sau khi confirm popup
            try:
                self.session.findById("wnd[0]/tbar[1]/btn[30]").press()
                time.sleep(1)
                self.sap.wait_ready(self.session, 5)
                log.info(f"    Pressed btn[30]")
            except Exception as e:
                return {'ok': False, 'message': f'Lỗi nhấn btn[30]: {e}'}

            # 5. Chọn Tab và lấy ID field Long Text
            if tab == 'ZU05':
                longtext_id = self._select_tab_zu05()
            elif tab == 'ZU07':
                longtext_id = self._select_tab_zu07()
            else:
                return {'ok': False, 'message': f'Tab không hợp lệ: {tab}'}

            if not longtext_id:
                return {'ok': False, 'message': f'Không tìm thấy Tab {tab}'}

            # 6. Nhập mô tả
            try:
                self.session.findById(longtext_id).text = description
                time.sleep(1)
                log.info(f"    Đã nhập mô tả: {description[:50]}...")
            except Exception as e:
                return {'ok': False, 'message': f'Không tìm thấy field Long Text: {e}'}

            # 7. Save (btn[11])
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)

            # Dismiss popup sau save nếu có
            try:
                self.session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
                time.sleep(0.3)
            except:
                pass

            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            # 8. Back (btn[3])
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            time.sleep(1)
            self.sap.wait_ready(self.session, 5)

            return {'ok': True, 'message': 'OK'}

        except Exception as e:
            return {'ok': False, 'message': str(e)}

    def _process_single_extwg(self, matnr: str, extwg: str) -> dict:
        """Xử lý cập nhật External Material Group cho 1 material"""
        try:
            # Reset SAP về trạng thái sạch
            self._safe_reset()

            # 1. Mở TCode MM02
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nMM02"
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

            # 2. Nhập Material Number
            matnr_field = self.sap.safe_find(self.session, "wnd[0]/usr/ctxtRMMG1-MATNR")
            if matnr_field is None:
                return {'ok': False, 'message': 'Không tìm thấy field Material Number'}

            matnr_field.text = matnr
            # Set caretPosition để SAP nhận diện đủ text — theo script gốc
            matnr_field.caretPosition = len(matnr)

            # Dùng btn[5] để navigate — theo script gốc
            self.session.findById("wnd[0]/tbar[1]/btn[5]").press()
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)

            # Kiểm tra lỗi (Material không tồn tại)
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            # 3. Xử lý popup Select View
            # btn[5] luôn mở Select View → không cần check if popup, xử lý trực tiếp
            try:
                # btn[19]: Deselect all trước — tránh chọn nhầm view từ lần trước
                self.session.findById("wnd[1]/tbar[0]/btn[19]").press()
                time.sleep(0.3)

                # Select row 0 (Basic Data 1) và row 1 (Basic Data 2)
                table = self.session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW")
                table.getAbsoluteRow(0).selected = True
                table.getAbsoluteRow(1).selected = True

                # setFocus + caretPosition = 0 vào row 1 — theo script gốc
                row1_field = self.session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,1]")
                row1_field.setFocus()
                row1_field.caretPosition = 0
                time.sleep(0.3)

                # Confirm
                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                time.sleep(0.5)
                self.sap.wait_ready(self.session, 10)
                log.info(f"    Selected Basic Data 1 & 2 from popup")
            except Exception as e:
                return {'ok': False, 'message': f'Lỗi xử lý popup Select View: {e}'}

            # Kiểm tra lỗi sau popup
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            # 4. Nhập External Material Group (Tab SP01 - Basic Data 1)
            extwg_field = self.sap.safe_find(
                self.session,
                "wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2004/subSUB2:SAPLMGD1:2001/ctxtMARA-EXTWG"
            )
            if extwg_field is None:
                return {'ok': False, 'message': 'Không tìm thấy field Ext. Matl Group'}

            extwg_field.text = extwg
            extwg_field.setFocus()
            time.sleep(0.3)
            log.info(f"    Đã nhập EXTWG: {extwg}")

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

    def _process_single_bstma(self, matnr: str, werks: str, bstma: str) -> dict:
        """Xử lý cập nhật Maximum Lot Size cho 1 material"""
        try:
            # Reset SAP về trạng thái sạch
            self._safe_reset()

            # 1. Mở TCode MM02
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nMM02"
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

            # 2. Nhập Material Number + caretPosition — theo pattern chuẩn
            matnr_field = self.sap.safe_find(self.session, self.MATNR_ID)
            if matnr_field is None:
                return {'ok': False, 'message': 'Không tìm thấy field Material Number'}

            matnr_field.text = matnr
            matnr_field.caretPosition = len(matnr)

            # btn[5] = Select Org. Levels → mở thẳng popup Select View
            self.session.findById("wnd[0]/tbar[1]/btn[5]").press()
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)

            # Kiểm tra lỗi (Material không tồn tại)
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            # 3. Xử lý popup Select View - Tìm và chọn MRP 1
            # btn[5] luôn mở Select View → xử lý trực tiếp, không dùng if popup
            try:
                # btn[19]: Deselect all trước
                self.session.findById("wnd[1]/tbar[0]/btn[19]").press()
                time.sleep(0.3)

                # Tìm và select row MRP 1
                mrp1_row = self._find_view_row("MRP 1")
                if mrp1_row is None:
                    return {'ok': False, 'message': 'Không tìm thấy view MRP 1 trong danh sách'}

                table = self.session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW")
                table.getAbsoluteRow(mrp1_row).selected = True

                # setFocus + caretPosition vào row đã chọn
                row_field = self.session.findById(f"wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,{mrp1_row}]")
                row_field.setFocus()
                row_field.caretPosition = 0
                time.sleep(0.3)

                # Confirm
                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                time.sleep(0.5)
                self.sap.wait_ready(self.session, 10)
                log.info(f"    Selected MRP 1 (row {mrp1_row}) from popup")
            except Exception as e:
                return {'ok': False, 'message': f'Lỗi xử lý popup Select View: {e}'}

            # 4. Xử lý popup nhập Plant — luôn xuất hiện sau khi chọn MRP 1
            try:
                werks_field = self.sap.safe_find(self.session, self.WERKS_POPUP_ID)
                if werks_field:
                    werks_field.text = werks
                    self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                    time.sleep(0.5)
                    self.sap.wait_ready(self.session, 10)
                    log.info(f"    Entered Plant: {werks}")
                else:
                    return {'ok': False, 'message': 'Không tìm thấy field Plant trong popup'}
            except Exception as e:
                return {'ok': False, 'message': f'Lỗi nhập Plant: {e}'}

            # Kiểm tra lỗi sau popup
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            # 5. Nhập Maximum Lot Size (Tab SP12)
            bstma_field = self.sap.safe_find(self.session, self.BSTMA_ID)
            if bstma_field is None:
                return {'ok': False, 'message': 'Không tìm thấy field Maximum Lot Size'}

            bstma_field.text = bstma
            bstma_field.setFocus()
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)
            self.sap.wait_ready(self.session, 10)
            log.info(f"    Đã nhập BSTMA: {bstma}")

            # Kiểm tra lỗi sau khi nhập
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            # 6. Save (btn[11])
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

    # ===== Helper methods =====

    def _find_view_row(self, view_name: str) -> int:
        """
        Tìm row index của view trong popup Select View
        """
        try:
            table = self.session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW")
            visible_rows = table.VisibleRowCount
            row_count = table.RowCount

            log.info(f"    Searching for '{view_name}' in {row_count} rows...")

            for i in range(row_count):
                if i >= visible_rows:
                    scroll_pos = i - visible_rows + 1
                    table.verticalScrollbar.position = scroll_pos
                    time.sleep(0.2)

                try:
                    cell_id = f"wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,{i}]"
                    cell = self.sap.safe_find(self.session, cell_id)
                    if cell:
                        cell_text = cell.text.strip()
                        if view_name.lower() in cell_text.lower():
                            log.info(f"    Found '{view_name}' at row {i}: '{cell_text}'")
                            table.verticalScrollbar.position = 0
                            time.sleep(0.2)
                            return i
                except Exception as e:
                    log.debug(f"    Row {i} error: {e}")
                    continue

            try:
                table.verticalScrollbar.position = 0
            except:
                pass

            log.warning(f"    View '{view_name}' not found in table")
            return None

        except Exception as e:
            log.error(f"    Error finding view row: {e}")
            return None

    def _select_tab_zu05(self) -> str:
        """Chọn Tab ZU05 và trả về ID của field Long Text"""
        try:
            self.session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU05").select()
            time.sleep(1)
            self.sap.wait_ready(self.session, 5)
            log.info(f"    Selected Tab ZU05")
            return "wnd[0]/usr/tabsTABSPR1/tabpZU05/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:2031/cntlLONGTEXT_GRUNDD/shellcont/shell"
        except Exception as e:
            log.error(f"    Tab ZU05 error: {e}")
            return None

    def _select_tab_zu07(self) -> str:
        """Chọn Tab ZU07 và trả về ID của field Long Text"""
        try:
            self.session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU07").select()
            time.sleep(1)
            self.sap.wait_ready(self.session, 5)
            log.info(f"    Selected Tab ZU07")
            return "wnd[0]/usr/tabsTABSPR1/tabpZU07/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:2051/cntlLONGTEXT_IVERM/shellcont/shell"
        except Exception as e:
            log.error(f"    Tab ZU07 error: {e}")
            return None
