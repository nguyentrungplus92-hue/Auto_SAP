"""
TCode MASS - Mass Update External Material Group (EXTWG)
Bam sat script SAP goc.
"""
import time
import logging

log = logging.getLogger(__name__)


class TCodeMASS:
    VERSION = "v5-variant"  # Dung Z_EXTWG variant de bo qua man hinh chon table

    SESSION_DEAD_MSG = "SAP session bi dong — user da login o noi khac hoac mat ket noi"

    HEADER_FIELD = "wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/ctxtHEADER_STRUC-FIELD2-VALUE-LEFT[2,0]"
    DATA_TABLE   = "wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/subSUB_DATA:SAPLMASSINTERFACE:0212/tblSAPLMASSINTERFACETCTRL_TABLE"
    FIELD_TABLE  = "wnd[1]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER"

    def __init__(self, sap_client):
        self.sap = sap_client
        self.session = sap_client.session

    def _safe_reset(self):
        self.sap.dismiss_system_popup()
        for _ in range(3):
            try:
                self.session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
                time.sleep(0.1); continue
            except: pass
            try:
                self.session.findById("wnd[1]/tbar[0]/btn[1]").press()
                time.sleep(0.1); continue
            except: pass
            try:
                self.session.findById("wnd[1]").sendVKey(12)
                time.sleep(0.1); continue
            except: pass
            break
        try:
            self.session.findById("wnd[0]").sendVKey(12)
            time.sleep(0.1)
        except: pass
        try:
            self.session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
            time.sleep(0.1)
        except: pass
        try:
            self.session.findById("wnd[1]").sendVKey(12)
            time.sleep(0.1)
        except: pass
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.3)
        except: pass
        try:
            self.session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
            time.sleep(0.1)
        except: pass
        try:
            self.session.findById("wnd[1]").sendVKey(12)
            time.sleep(0.1)
        except: pass

    def run(self, materials: list) -> dict:
        if not materials:
            return {'ok': True, 'message': 'Khong co du lieu', 'processed': 0, 'total': 0, 'errors': []}

        from collections import defaultdict
        groups = defaultdict(list)
        for item in materials:
            extwg = str(item.get('extwg', '')).strip()
            matnr = str(item.get('matnr', '')).strip()
            if extwg and matnr:
                groups[extwg].append(matnr)

        log.warning(f"  [MASS] {self.VERSION} - Tong {len(materials)} materials, {len(groups)} nhom EXTWG")

        total = len(materials)
        processed = 0
        errors = []
        group_list = list(groups.items())
        BATCH_SIZE = 8

        for g_idx, (extwg, matnr_list) in enumerate(group_list):
            if not self.sap.is_session_alive():
                for future_extwg, future_list in group_list[g_idx:]:
                    for m in future_list:
                        errors.append({'row': '', 'info': m, 'detail': self.SESSION_DEAD_MSG})
                break

            batches = [matnr_list[i:i+BATCH_SIZE] for i in range(0, len(matnr_list), BATCH_SIZE)]
            log.warning(f"  [MASS] EXTWG='{extwg}': {len(matnr_list)} materials, {len(batches)} batch")

            for b_idx, batch in enumerate(batches):
                if not self.sap.is_session_alive():
                    for m in batch:
                        errors.append({'row': '', 'info': m, 'detail': self.SESSION_DEAD_MSG})
                    break

                log.warning(f"  [MASS] Batch {b_idx+1}/{len(batches)}: {batch}")

                try:
                    result = self._process_group(extwg, batch)
                except Exception as e:
                    log.warning(f"  [MASS] Exception batch {b_idx+1}: {e}")
                    result = {'ok': False, 'message': str(e), 'count': 0, 'errors': 0, 'error_details': []}
                    try: self._safe_reset()
                    except: pass

                if result.get('ok'):
                    count = result.get('count', len(batch))
                    processed += count
                    for ed in result.get('error_details', []):
                        errors.append({'row': '', 'info': f"{ed['matnr']} / {extwg}", 'detail': ed['detail']})
                else:
                    msg = result.get('message', 'Loi khong xac dinh')
                    log.warning(f"  [MASS] Batch {b_idx+1} LOI: {msg}")
                    for matnr in batch:
                        errors.append({'row': '', 'info': f"{matnr} / {extwg}", 'detail': msg})

        if errors:
            error_lines = "\n".join([f"  • {e['info']}: {e['detail']}" for e in errors])
            message = f"Hoàn tất {processed}/{total} dòng thành công. Lỗi {len(errors)} dòng:\n{error_lines}"
        else:
            message = f"Hoàn tất. Đã cập nhật thành công {processed}/{total} materials"

        return {'ok': True, 'message': message, 'processed': processed, 'total': total, 'errors': errors}

    def _process_group(self, extwg: str, matnr_list: list) -> dict:
        try:
            self._safe_reset()

            # ===== Mo MASS voi Variant Z_EXTWG =====
            # Dung variant de bo qua man hinh chon table (nhanh hon)
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "MASS"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)
            self.sap.wait_ready(self.session, 15)

            obj_field = self.sap.safe_find(self.session, "wnd[0]/usr/ctxtMASSSCREEN-OBJECT")
            if obj_field is None:
                return {'ok': False, 'message': 'Khong tim thay field Object Type'}
            obj_field.text = "BUS1001"

            var_field = self.sap.safe_find(self.session, "wnd[0]/usr/ctxtMASSSCREEN-VARNAME")
            if var_field is not None:
                var_field.text = "Z_EXTWG"
                try:
                    var_field.setFocus()
                    var_field.caretPosition = 7
                except: pass
            else:
                log.warning("    Khong tim thay field Variant — bo qua")

            # Execute — voi variant se nhay thang vao man hinh selection
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            time.sleep(0.3)
            self.sap.wait_ready(self.session, 15)

            # ===== Neu khong co variant: chon table row 0 (fallback) =====
            tab_table = "wnd[0]/usr/tabsTAB/tabpTAB1/ssubSUBTAB:SAPMMSDL:1000/tblSAPMMSDLTC_TAB"
            table = self.sap.safe_find(self.session, tab_table)
            if table is not None:
                log.warning("    Man hinh chon table (variant chua ap dung) — chon row 0")
                table.getAbsoluteRow(0).selected = True
                try:
                    self.session.findById(f"{tab_table}/txtMASSTABS-TABTXT[0,0]").setFocus()
                except: pass
                self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
                time.sleep(0.3)
                self.sap.wait_ready(self.session, 15)

            # Mo Free Selection
            free_sel = "wnd[0]/usr/tabsTAB/tabpCHAN/ssubSUB_ALL:SAPLMASS_SEL_DIALOG:0200/ssubSUB_SEL:SAPLMASSFREESELECTIONS:1000/sub:SAPLMASSFREESELECTIONS:1000/btnMASSFREESEL-MORE[0,69]"
            btn = self.sap.safe_find(self.session, free_sel)
            if btn is None:
                return {'ok': False, 'message': 'Khong tim thay nut Free Selection'}
            btn.press()
            time.sleep(0.3)
            self.sap.wait_ready(self.session, 15)

            # Nhap MATNR vao wnd[1]
            result = self._enter_matnr_wnd1(matnr_list)
            if not result['ok']:
                return result

            # wnd[1]: btn[8] = Execute selection
            self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
            time.sleep(0.3)
            self.sap.wait_ready(self.session, 15)

            # Execute MASS chinh
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            time.sleep(0.5)
            self.sap.wait_ready(self.session, 30)

            status = self.sap.status(self.session)
            log.warning(f"    [MASS] After Execute: '{status.type}:{status.text}'")

            # Variant Z_EXTWG da co san field Ext. Material Group
            # Khong can vao FDAW de tim va chon field nua

            # Dien EXTWG
            header_fld = self.sap.safe_find(self.session, self.HEADER_FIELD)
            if header_fld is None:
                return {'ok': False, 'message': 'Khong tim thay header field EXTWG'}
            header_fld.text = extwg

            # selectAllColumns
            try:
                data_table = self.sap.safe_find(self.session, self.DATA_TABLE)
                if data_table:
                    try:
                        data_table.selectAllColumns()
                        log.warning("    selectAllColumns() OK")
                    except:
                        try:
                            col_count = data_table.columnCount
                            for c in range(col_count):
                                try: data_table.columns.elementAt(c).selected = True
                                except: pass
                            log.warning(f"    selectAllColumns fallback: {col_count} cols")
                        except: pass
            except: pass

            # Apply (FDAE)
            fdae = self.sap.safe_find(
                self.session,
                "wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/btnFDAE"
            )
            if fdae is None:
                return {'ok': False, 'message': 'Khong tim thay FDAE'}
            fdae.press()
            time.sleep(0.1)
            self.sap.wait_ready(self.session, 15)

            # Save
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            time.sleep(0.3)
            self.sap.wait_ready(self.session, 30)

            status = self.sap.status(self.session)
            log.warning(f"    [MASS] After Save: '{status.type}:{status.text}'")

            # Loi toan batch
            if self.sap.is_error(status):
                return {
                    'ok': True,
                    'count': 0,
                    'errors': len(matnr_list),
                    'error_details': [{'matnr': m, 'detail': status.text} for m in matnr_list],
                    'message': f'Loi Save: {status.text}',
                }

            # Doc ket qua
            msg_result = self._parse_mass_result(len(matnr_list))

            # Back
            try:
                self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                time.sleep(0.3)
                self.sap.wait_ready(self.session, 10)
            except: pass

            return {
                'ok': True,
                'count': msg_result['success'],
                'errors': msg_result['errors'],
                'error_details': msg_result['error_details'],
                'message': 'OK',
            }

        except Exception as e:
            try: self._safe_reset()
            except: pass
            import traceback
            log.warning(f"    [MASS] Exception: {traceback.format_exc()}")
            return {'ok': False, 'message': str(e)}

    def _enter_matnr_wnd1(self, matnr_list: list) -> dict:
        """
        Nhap MATNR truc tiep vao cac cell trong wnd[1] free selection.
        Khong dung clipboard, khong phu thuoc focus chuot.
        Toi da 8 cell.
        """
        CELL_BASE = "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL-SLOW_I[1,{i}]"

        # Clear selection cu
        try:
            self.session.findById("wnd[1]/tbar[0]/btn[16]").press()
            time.sleep(0.1)
        except Exception as e:
            log.warning(f"    btn[16] clear loi: {e}")

        # Nhap tung MATNR
        for i, matnr in enumerate(matnr_list[:8]):
            cell_id = CELL_BASE.format(i=i)
            try:
                self.session.findById(cell_id).text = str(matnr).strip()
                log.warning(f"    cell[{i}] = '{matnr}'")
            except Exception as e:
                log.warning(f"    cell[{i}] loi: {e}")

        return {'ok': True}

    def _parse_mass_result(self, expected_count: int) -> dict:
        result = {'success': expected_count, 'errors': 0, 'error_details': [], 'error_materials': []}
        MSG_TABLE = "wnd[0]/usr/tblSAPLMASSMSGLISTTC_MSG"

        try:
            msg_table = self.sap.safe_find(self.session, MSG_TABLE)
            if msg_table is None:
                return result

            try: total_rows = msg_table.rowCount
            except: total_rows = 200

            for vis_row in range(min(total_rows, 50)):
                icon_src = ''
                icon_cell = self.sap.safe_find(self.session, f"{MSG_TABLE}/lblLIGHT[0,{vis_row}]")
                if icon_cell:
                    try:
                        for prop in ['ImageSource', 'text', 'tooltip']:
                            val = getattr(icon_cell, prop, None)
                            if val:
                                icon_src = str(val).strip()
                                break
                    except: pass

                msg_text = ''
                cell = self.sap.safe_find(self.session, f"{MSG_TABLE}/txtREOML-MSGTXT[1,{vis_row}]")
                if cell:
                    try: msg_text = (cell.text or '').strip()
                    except: pass

                # Dung tai separator hoac het data
                if not msg_text or msg_text.startswith('___'):
                    log.warning(f"      [STOP at row {vis_row}]")
                    break

                log.warning(f"      [MSG {vis_row}] icon='{icon_src[:30]}' txt='{msg_text[:80]}'")

                if ' : ' in msg_text:
                    matnr, detail = msg_text.split(' : ', 1)
                    matnr = matnr.strip()
                    detail = detail.strip()
                else:
                    matnr, detail = '', msg_text

                icon_upper = icon_src.upper()
                is_error = ('RED' in icon_upper or 'STOP' in icon_upper or 'ERROR' in icon_upper)

                if is_error and matnr:
                    result['error_details'].append({'matnr': matnr, 'detail': detail})
                    result['error_materials'].append(matnr)

            result['errors'] = len(result['error_details'])
            result['success'] = max(0, expected_count - result['errors'])

        except Exception as e:
            log.warning(f"    _parse_mass_result error: {e}")

        return result
