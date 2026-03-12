"""
TCode ME12 - Change Info Record
===============================
Cập nhật Marker (KOLIF) hoặc Shipping Instruction (EVERS) cho Info Record
"""
import time
import logging

log = logging.getLogger(__name__)


class TCodeME12:
    """
    ME12 - Change Info Record
    - Cập nhật Marker (KOLIF) cho Vendor/Material
    - Cập nhật Shipping Instruction (EVERS) cho Vendor/Material/Purch.Org/Plant
    """

    def __init__(self, sap_client):
        self.sap = sap_client
        self.session = sap_client.session

    # ===== Method 1: Update Marker =====
    def run(self, info_records: list) -> dict:
        """
        Xử lý danh sách info records - Update Marker
        
        Args:
            info_records: List dict với keys: vendor, material, marker (marker có thể trống)
            
        Returns:
            dict với keys: ok, message, processed, errors
        """
        processed = 0
        errors = []
        
        # Mở TCode ME12 lần đầu
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nME12"
        self.session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)
        self.sap.wait_ready(self.session, 10)
        
        for item in info_records:
            vendor = str(item.get('vendor', '')).strip()
            material = str(item.get('material', '')).strip()
            marker = str(item.get('marker', '')).strip() if item.get('marker') else ''
            
            if not vendor or not material:
                continue
                
            try:
                result = self._process_marker(vendor, material, marker)
                if result['ok']:
                    processed += 1
                    log.info(f"  [ME12] {vendor}/{material}: OK - Marker: '{marker}'")
                else:
                    errors.append(f"{vendor}/{material}: {result['message']}")
                    log.warning(f"  [ME12] {vendor}/{material}: {result['message']}")
            except Exception as e:
                errors.append(f"{vendor}/{material}: {str(e)}")
                log.error(f"  [ME12] {vendor}/{material}: Exception - {e}")
        
        if errors:
            return {
                'ok': processed > 0,
                'message': f"Processed: {processed}, Errors: {len(errors)} - {'; '.join(errors[:3])}",
                'processed': processed,
                'errors': errors
            }
        else:
            return {
                'ok': True,
                'message': f"Hoàn tất. Đã cập nhật {processed} info records",
                'processed': processed,
                'errors': []
            }

    def _process_marker(self, vendor: str, material: str, marker: str) -> dict:
        """Xử lý 1 info record - Update Marker"""
        try:
            # 1. Nhập Vendor
            self.session.findById("wnd[0]/usr/ctxtEINA-LIFNR").text = vendor
            
            # 2. Nhập Material
            self.session.findById("wnd[0]/usr/ctxtEINA-MATNR").text = material
            self.session.findById("wnd[0]/usr/ctxtEINA-MATNR").setFocus()
            
            # 3. Nhấn Enter để load info record
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)
            
            # Kiểm tra lỗi
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}
            
            # 4. Nhập Marker (KOLIF)
            try:
                self.session.findById("wnd[0]/usr/ctxtEINA-KOLIF").text = marker
                self.session.findById("wnd[0]/usr/ctxtEINA-KOLIF").setFocus()
            except Exception as e:
                return {'ok': False, 'message': f'Không tìm thấy field Marker: {e}'}
            
            # 5. Save
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)
            
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
        
        Args:
            info_records: List dict với keys: vendor, material, purch_org, plant, shipping_instr
                         (shipping_instr có thể trống)
            
        Returns:
            dict với keys: ok, message, processed, errors
        """
        processed = 0
        errors = []
        
        # Mở TCode ME12 lần đầu
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nME12"
        self.session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)
        self.sap.wait_ready(self.session, 10)
        
        for item in info_records:
            vendor = str(item.get('vendor', '')).strip()
            material = str(item.get('material', '')).strip()
            purch_org = str(item.get('purch_org', '')).strip()
            plant = str(item.get('plant', '')).strip()
            shipping_instr = str(item.get('shipping_instr', '')).strip() if item.get('shipping_instr') else ''
            
            if not vendor or not material or not purch_org or not plant:
                continue
                
            try:
                result = self._process_shipping_instruction(vendor, material, purch_org, plant, shipping_instr)
                if result['ok']:
                    processed += 1
                    log.info(f"  [ME12] {vendor}/{material}/{purch_org}/{plant}: OK - EVERS: '{shipping_instr}'")
                else:
                    errors.append(f"{vendor}/{material}: {result['message']}")
                    log.warning(f"  [ME12] {vendor}/{material}: {result['message']}")
            except Exception as e:
                errors.append(f"{vendor}/{material}: {str(e)}")
                log.error(f"  [ME12] {vendor}/{material}: Exception - {e}")
        
        if errors:
            return {
                'ok': processed > 0,
                'message': f"Processed: {processed}, Errors: {len(errors)} - {'; '.join(errors[:3])}",
                'processed': processed,
                'errors': errors
            }
        else:
            return {
                'ok': True,
                'message': f"Hoàn tất. Đã cập nhật {processed} info records",
                'processed': processed,
                'errors': []
            }

    def _process_shipping_instruction(self, vendor: str, material: str, purch_org: str, plant: str, shipping_instr: str) -> dict:
        """
        Xử lý 1 info record - Update Shipping Instruction (EVERS)
        
        Args:
            vendor: Mã vendor (VD: VA001)
            material: Mã material (VD: 2SB1219ARL)
            purch_org: Purchasing Organization (VD: C001)
            plant: Plant (VD: VC01)
            shipping_instr: Shipping Instruction (VD: V3) - có thể trống
        """
        try:
            # 1. Nhập Vendor
            self.session.findById("wnd[0]/usr/ctxtEINA-LIFNR").text = vendor
            
            # 2. Nhập Material
            self.session.findById("wnd[0]/usr/ctxtEINA-MATNR").text = material
            
            # 3. Nhập Purchasing Organization
            self.session.findById("wnd[0]/usr/ctxtEINE-EKORG").text = purch_org
            
            # 4. Nhập Plant
            self.session.findById("wnd[0]/usr/ctxtEINE-WERKS").text = plant
            self.session.findById("wnd[0]/usr/ctxtEINE-WERKS").setFocus()
            
            # 5. Nhấn Enter để load info record
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)
            
            # Kiểm tra lỗi
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}
            
            # 6. Nhấn btn[7] để vào màn hình Conditions/Shipping
            try:
                self.session.findById("wnd[0]/tbar[1]/btn[7]").press()
                time.sleep(1)
                self.sap.wait_ready(self.session, 10)
            except Exception as e:
                return {'ok': False, 'message': f'Không nhấn được btn[7]: {e}'}
            
            # 7. Nhập Shipping Instruction (EVERS)
            try:
                self.session.findById("wnd[0]/usr/ctxtEINE-EVERS").text = shipping_instr
                self.session.findById("wnd[0]/usr/ctxtEINE-EVERS").setFocus()
            except Exception as e:
                return {'ok': False, 'message': f'Không tìm thấy field EVERS: {e}'}
            
            # 8. Save
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)
            
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {'ok': False, 'message': status.text}
            
            return {'ok': True, 'message': 'OK'}
            
        except Exception as e:
            return {'ok': False, 'message': str(e)}