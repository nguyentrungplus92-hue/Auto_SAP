"""
TCode ME12 - Change Info Record
===============================
Cập nhật Marker (KOLIF) hoặc Shipping Instruction (EVERS) cho Info Record
Chạy tất cả dòng, ghi nhận lỗi, không dừng giữa chừng
"""
import time
import logging

log = logging.getLogger(__name__)


class TCodeME12:
    """
    ME12 - Change Info Record
    - Cập nhật Marker (KOLIF) cho Vendor/Material
    - Cập nhật Shipping Instruction (EVERS) cho Vendor/Material/Purch.Org/Plant
    Chạy tất cả dòng, thu thập lỗi, không dừng giữa chừng
    """

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

    # ===== Method 1: Update Marker =====

    def run(self, info_records: list) -> dict:
        """
        Xử lý danh sách info records - Update Marker
        Chạy tất cả, thu thập lỗi, không dừng giữa chừng

        Args:
            info_records: List dict với keys: vendor, material, marker (marker có thể trống)

        Returns:
            dict với keys: ok, message, processed, total, errors
        """
        processed = 0
        total = len(info_records)
        errors = []

        for idx, item in enumerate(info_records):
            vendor = str(item.get('vendor', '')).strip()
            material = str(item.get('material', '')).strip()
            marker = str(item.get('marker', '')).strip() if item.get('marker') else ''
            row_number = idx + 2

            if not vendor or not material:
                continue

            try:
                result = self._process_marker(vendor, material, marker)
            except Exception as e:
                log.error(f"  [ME12] Exception bất ngờ dòng {row_number} ({vendor}/{material}): {e}")
                result = {'ok': False, 'message': f"Exception: {str(e)}"}
                try:
                    self._safe_reset()
                except:
                    pass

            if result['ok']:
                processed += 1
                log.info(f"  [ME12] {vendor}/{material}: OK - Marker: '{marker}'")
            else:
                error_detail = result.get('message', 'Lỗi không xác định')
                log.warning(f"  [ME12] Lỗi dòng {row_number} ({vendor}/{material}): {error_detail}")
                errors.append({
                    'row': row_number,
                    'info': f"{vendor}/{material}",
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
            message = f"Hoàn tất. Đã cập nhật {processed}/{total} info records"

        return {
            'ok': True,
            'message': message,
            'processed': processed,
            'total': total,
            'errors': errors,
        }

    def _process_marker(self, vendor: str, material: str, marker: str) -> dict:
        """Xử lý 1 info record - Update Marker"""
        try:
            # Reset SAP về trạng thái sạch
            self._safe_reset()

            # 1. Mở TCode ME12
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nME12"
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

            # 2. Nhập Vendor
            self.session.findById("wnd[0]/usr/ctxtEINA-LIFNR").text = vendor

            # 3. Nhập Material
            self.session.findById("wnd[0]/usr/ctxtEINA-MATNR").text = material
            self.session.findById("wnd[0]/usr/ctxtEINA-MATNR").setFocus()

            # Các trường PRG,Plat, Info để blank
            self.session.findById("wnd[0]/usr/ctxtEINE-EKORG").text = ""
            self.session.findById("wnd[0]/usr/ctxtEINE-WERKS").text = ""
            self.session.findById("wnd[0]/usr/ctxtEINA-INFNR").text = ""

            # Chọn Standard
            self.session.findById("wnd[0]/usr/radRM06I-NORMB").setFocus

            # 4. Nhấn Enter để load info record
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)

            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            # 5. Nhập Marker (KOLIF)
            try:
                self.session.findById("wnd[0]/usr/ctxtEINA-KOLIF").text = marker
                self.session.findById("wnd[0]/usr/ctxtEINA-KOLIF").setFocus()
            except Exception as e:
                return {'ok': False, 'message': f'Không tìm thấy field Marker: {e}'}

            # 6. Save
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)

            # Dismiss popup sau save
            try:
                self.session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
                time.sleep(0.3)
            except:
                pass

            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            return {'ok': True, 'message': 'OK'}

        except Exception as e:
            return {'ok': False, 'message': str(e)}

    # ===== Method 2: Update Shipping Instruction =====

    def run_shipping_instruction(self, info_records: list) -> dict:
        """
        Xử lý danh sách info records - Update Shipping Instruction
        Chạy tất cả, thu thập lỗi, không dừng giữa chừng

        Args:
            info_records: List dict với keys: vendor, material, purch_org, plant, shipping_instr
                         (shipping_instr có thể trống)

        Returns:
            dict với keys: ok, message, processed, total, errors
        """
        processed = 0
        total = len(info_records)
        errors = []

        for idx, item in enumerate(info_records):
            vendor = str(item.get('vendor', '')).strip()
            material = str(item.get('material', '')).strip()
            purch_org = str(item.get('purch_org', '')).strip()
            plant = str(item.get('plant', '')).strip()
            shipping_instr = str(item.get('shipping_instr', '')).strip() if item.get('shipping_instr') else ''
            row_number = idx + 2

            if not vendor or not material or not purch_org or not plant:
                continue

            try:
                result = self._process_shipping_instruction(vendor, material, purch_org, plant, shipping_instr)
            except Exception as e:
                log.error(f"  [ME12] Exception bất ngờ dòng {row_number} ({vendor}/{material}): {e}")
                result = {'ok': False, 'message': f"Exception: {str(e)}"}
                try:
                    self._safe_reset()
                except:
                    pass

            if result['ok']:
                processed += 1
                log.info(f"  [ME12] {vendor}/{material}/{purch_org}/{plant}: OK - EVERS: '{shipping_instr}'")
            else:
                error_detail = result.get('message', 'Lỗi không xác định')
                log.warning(f"  [ME12] Lỗi dòng {row_number} ({vendor}/{material}): {error_detail}")
                errors.append({
                    'row': row_number,
                    'info': f"{vendor}/{material}/{purch_org}/{plant}",
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
            message = f"Hoàn tất. Đã cập nhật {processed}/{total} info records"

        return {
            'ok': True,
            'message': message,
            'processed': processed,
            'total': total,
            'errors': errors,
        }

    def _process_shipping_instruction(self, vendor: str, material: str, purch_org: str, plant: str, shipping_instr: str) -> dict:
        """
        Xử lý 1 info record - Update Shipping Instruction (EVERS)
        """
        try:
            # Reset SAP về trạng thái sạch
            self._safe_reset()

            # 1. Mở TCode ME12
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nME12"
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

            # 2. Nhập Vendor
            self.session.findById("wnd[0]/usr/ctxtEINA-LIFNR").text = vendor

            # 3. Nhập Material
            self.session.findById("wnd[0]/usr/ctxtEINA-MATNR").text = material

            # 4. Nhập Purchasing Organization
            self.session.findById("wnd[0]/usr/ctxtEINE-EKORG").text = purch_org

            # 5. Nhập Plant
            self.session.findById("wnd[0]/usr/ctxtEINE-WERKS").text = plant
            self.session.findById("wnd[0]/usr/ctxtEINE-WERKS").setFocus()

            # Info record để Blank
            self.session.findById("wnd[0]/usr/ctxtEINA-INFNR").text = ""

            # Chọn Standard
            self.session.findById("wnd[0]/usr/radRM06I-NORMB").setFocus

            # 6. Nhấn Enter để load info record
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)

            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            # 7. Nhấn btn[7] để vào màn hình Conditions/Shipping
            try:
                self.session.findById("wnd[0]/tbar[1]/btn[7]").press()
                time.sleep(1)
                self.sap.wait_ready(self.session, 10)
            except Exception as e:
                return {'ok': False, 'message': f'Không nhấn được btn[7]: {e}'}

            # 8. Nhập Shipping Instruction (EVERS)
            try:
                self.session.findById("wnd[0]/usr/ctxtEINE-EVERS").text = shipping_instr
                self.session.findById("wnd[0]/usr/ctxtEINE-EVERS").setFocus()
            except Exception as e:
                return {'ok': False, 'message': f'Không tìm thấy field EVERS: {e}'}

            # 9. Save
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)

            # Dismiss popup sau save
            try:
                self.session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
                time.sleep(0.3)
            except:
                pass

            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}

            return {'ok': True, 'message': 'OK'}

        except Exception as e:
            return {'ok': False, 'message': str(e)}
