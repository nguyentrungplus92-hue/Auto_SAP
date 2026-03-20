"""
TCode CS02 - Change BOM
=======================
Upload/Set BOM PMG - Thêm component vào BOM
"""
import time
import logging

log = logging.getLogger(__name__)


class TCodeCS02:
    """
    CS02 - Change BOM
    Thêm component vào BOM của Material
    """

    def __init__(self, sap_client):
        self.sap = sap_client
        self.session = sap_client.session

    def run(self, bom_items: list) -> dict:
        """
        Xử lý danh sách BOM items
        Nếu gặp lỗi → DỪNG NGAY và báo cáo
        
        Args:
            bom_items: List dict với keys:
                - matnr, werks, stlan, datuv, idnrk, menge, meins, postp, lgort
            
        Returns:
            dict với keys: ok, status, message, processed, total, error_at_row, error_detail
        """
        processed = 0
        total = len(bom_items)
        
        # Group by BOM header (matnr + werks + stlan + datuv)
        bom_groups = {}
        for idx, item in enumerate(bom_items):
            matnr = str(item.get('matnr', '')).strip()
            werks = str(item.get('werks', '')).strip()
            stlan = str(item.get('stlan', '')).strip()
            datuv = str(item.get('datuv', '')).strip()
            
            if not matnr or not werks:
                continue
            
            key = f"{matnr}|{werks}|{stlan}|{datuv}"
            if key not in bom_groups:
                bom_groups[key] = {
                    'matnr': matnr,
                    'werks': werks,
                    'stlan': stlan,
                    'datuv': datuv,
                    'components': [],
                    'start_row': idx + 2  # +2 vì dòng 1 là header, index bắt đầu từ 0
                }
            
            bom_groups[key]['components'].append({
                'idnrk': str(item.get('idnrk', '')).strip(),
                'menge': str(item.get('menge', '')).strip(),
                'meins': str(item.get('meins', '')).strip(),
                'postp': str(item.get('postp', '')).strip(),
                'lgort': str(item.get('lgort', '')).strip(),
                'row_number': idx + 2  # Số dòng trong file (bắt đầu từ 2)
            })
        
        # Process each BOM - DỪNG NGAY khi gặp lỗi
        for key, bom_data in bom_groups.items():
            result = self._process_single_bom(bom_data)
            
            if result['ok']:
                processed += len(bom_data['components'])
                log.info(f"  [CS02] {bom_data['matnr']}/{bom_data['werks']}: OK - {len(bom_data['components'])} components")
            else:
                # GẶP LỖI → DỪNG NGAY
                error_row = result.get('error_at_row', bom_data['start_row'])
                error_detail = result.get('message', 'Lỗi không xác định')
                
                log.error(f"  [CS02] DỪNG tại dòng {error_row}: {error_detail}")
                
                return {
                    'ok': False,
                    'status': 'error',
                    'message': f"Đã xử lý {processed}/{total} items. Lỗi tại dòng {error_row}: {error_detail}",
                    'processed': processed,
                    'total': total,
                    'error_at_row': error_row,
                    'error_detail': error_detail
                }
        
        # Tất cả thành công
        return {
            'ok': True,
            'status': 'success',
            'message': f"Hoàn tất. Đã thêm {processed}/{total} components vào BOM",
            'processed': processed,
            'total': total,
            'error_at_row': None,
            'error_detail': None
        }

    def _process_single_bom(self, bom_data: dict) -> dict:
        """
        Xử lý 1 BOM trong CS02
        """
        try:
            matnr = bom_data['matnr']
            werks = bom_data['werks']
            stlan = bom_data['stlan'] or '1'
            datuv = bom_data['datuv']
            components = bom_data['components']
            
            # ===== 1. Mở TCode CS02 =====
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nCS02"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
            self.sap.wait_ready(self.session, 10)
            
            # ===== 2. Nhập header =====
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
            
            # Kiểm tra lỗi
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {
                    'ok': False,
                    'message': f"BOM {matnr}/{werks}: {status.text}",
                    'error_at_row': bom_data['start_row']
                }
            
            # ===== 3. Nhấn btn[5] để vào chế độ edit items =====
            try:
                self.session.findById("wnd[0]/tbar[1]/btn[5]").press()
                time.sleep(1)
                self.sap.wait_ready(self.session, 10)
            except:
                pass
            
            # ===== 4. Thêm từng component =====
            for idx, comp in enumerate(components):
                row_idx = idx + 1
                row_number = comp['row_number']
                
                try:
                    # Item category (POSTP)
                    if comp['postp']:
                        self.session.findById(f"wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/ctxtRC29P-POSTP[1,{row_idx}]").text = comp['postp']
                    
                    # Component material (IDNRK)
                    if comp['idnrk']:
                        self.session.findById(f"wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/ctxtRC29P-IDNRK[2,{row_idx}]").text = comp['idnrk']
                    
                    # Quantity (MENGE)
                    if comp['menge']:
                        self.session.findById(f"wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/txtRC29P-MENGE[4,{row_idx}]").text = comp['menge']
                    
                    # Unit (MEINS)
                    if comp['meins']:
                        self.session.findById(f"wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/ctxtRC29P-MEINS[5,{row_idx}]").text = comp['meins']
                    
                    # Enter để xác nhận dòng
                    self.session.findById("wnd[0]").sendVKey(0)
                    time.sleep(0.5)
                    self.sap.wait_ready(self.session, 10)
                    
                    # Kiểm tra lỗi sau khi nhập
                    status = self.sap.status(self.session)
                    if self.sap.is_error(status):
                        return {
                            'ok': False,
                            'message': f"Component {comp['idnrk']}: {status.text}",
                            'error_at_row': row_number
                        }
                    
                    self.session.findById("wnd[0]").sendVKey(0)
                    time.sleep(0.5)
                    self.sap.wait_ready(self.session, 10)
                    
                    # ===== 5. Mở item detail để nhập Storage Location =====
                    if comp['lgort']:
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
                            
                            self.session.findById("wnd[0]/usr/tabsTS_ITEM/tabpPDAT/ssubSUBPAGE:SAPLCSDI:0840/ctxtRC29P-LGORT").text = comp['lgort']
                            self.session.findById("wnd[0]/usr/tabsTS_ITEM/tabpPDAT/ssubSUBPAGE:SAPLCSDI:0840/ctxtRC29P-LGORT").setFocus()
                            
                            self.session.findById("wnd[0]").sendVKey(0)
                            time.sleep(0.5)
                            
                            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                            time.sleep(0.5)
                            self.sap.wait_ready(self.session, 10)
                            
                        except Exception as lgort_err:
                            return {
                                'ok': False,
                                'message': f"LGORT {comp['lgort']}: {str(lgort_err)}",
                                'error_at_row': row_number
                            }
                    
                except Exception as comp_err:
                    return {
                        'ok': False,
                        'message': f"Component {comp['idnrk']}: {str(comp_err)}",
                        'error_at_row': row_number
                    }
            
            # ===== 6. Save (btn[11]) =====
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            time.sleep(2)
            self.sap.wait_ready(self.session, 10)
            
            # Kiểm tra kết quả save
            status = self.sap.status(self.session)
            if self.sap.is_error(status):
                return {
                    'ok': False,
                    'message': f"Save BOM {matnr}: {status.text}",
                    'error_at_row': bom_data['start_row']
                }
            
            return {'ok': True, 'message': 'OK'}
            
        except Exception as e:
            return {
                'ok': False,
                'message': str(e),
                'error_at_row': bom_data.get('start_row', 0)
            }