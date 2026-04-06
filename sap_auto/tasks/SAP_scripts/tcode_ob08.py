from typing import List, Dict, Any
from .sap_base import SapGuiClient


class TCodeOB08:
    TABLE_ID = "wnd[0]/usr/tblSAPL0SAPTCTRL_V_TCURR"

    def __init__(self, sap: SapGuiClient):
        self.sap = sap

    def run(self, rows: List[Dict[str, str]], export_path: str = None) -> Dict[str, Any]:
        """
        rows: [{"kurst","gdatu","kursm","fcurr","tcurr"}, ...]
        export_path: Đường dẫn folder để export RTF (VD: "C:\\Users\\70N7678\\Desktop\\test\\")
        """
        st = self.sap.run_tcode("OB08")
        if self.sap.is_error(st):
            return {"ok": False, "message": f"Open OB08 error: {st.text}", "status": st.__dict__}

        # Kiểm tra popup sau khi mở TCode
        popup_text = self._get_popup_text()
        if popup_text:
            self._close_popup()
            return {"ok": False, "message": popup_text, "status": {"popup": True}}

        # New entries (optional)
        btn_new = self.sap.safe_find(self.sap.session, "wnd[0]/tbar[1]/btn[5]")
        if btn_new is not None:
            self.sap.press("wnd[0]/tbar[1]/btn[5]")
            st2 = self.sap.status(self.sap.session)
            if self.sap.is_error(st2):
                return {"ok": False, "message": f"New Entries error: {st2.text}", "status": st2.__dict__}

        # Kiểm tra popup sau khi nhấn New Entries
        popup_text = self._get_popup_text()
        if popup_text:
            self._close_popup()
            return {"ok": False, "message": popup_text, "status": {"popup": True}}

        # Wait for table (avoid "control not found")
        tbl = self.sap.table_wait(self.TABLE_ID, timeout_s=20)
        if tbl is None:
            # Kiểm tra popup trước khi báo lỗi
            popup_text = self._get_popup_text()
            if popup_text:
                self._close_popup()
                return {"ok": False, "message": popup_text, "status": {"popup": True}}
            
            st_now = self.sap.status(self.sap.session)
            return {
                "ok": False,
                "message": f"Không tìm thấy table OB08: {self.TABLE_ID} (layout khác / chưa vào đúng màn / popup).",
                "status": st_now.__dict__,
            }

        visible_count = int(getattr(tbl, "VisibleRowCount", 1) or 1)

        for i, r in enumerate(rows):
            top = (i // visible_count) * visible_count
            self.sap.table_scroll_to(self.TABLE_ID, top)
            visible_row = i - top

            self.sap.table_set_cell(self.TABLE_ID, "V_TCURR-KURST", 0, visible_row, r["kurst"])
            self.sap.table_set_cell(self.TABLE_ID, "V_TCURR-GDATU", 1, visible_row, r["gdatu"])
            self.sap.table_set_cell(self.TABLE_ID, "RFCU9-KURSM",  2, visible_row, r["kursm"])
            self.sap.table_set_cell(self.TABLE_ID, "V_TCURR-FCURR", 5, visible_row, r["fcurr"])
            self.sap.table_set_cell(self.TABLE_ID, "V_TCURR-TCURR", 10, visible_row, r["tcurr"])

        # Validate
        st3 = self.sap.send_enter()
        if self.sap.is_error(st3):
            return {"ok": False, "message": f"Validate error: {st3.text}", "status": st3.__dict__}

        # Save
        btn_save = self.sap.safe_find(self.sap.session, "wnd[0]/tbar[0]/btn[11]")
        if btn_save is None:
            st_now = self.sap.status(self.sap.session)
            return {"ok": False, "message": "Không tìm thấy nút Save", "status": st_now.__dict__}

        self.sap.press("wnd[0]/tbar[0]/btn[11]")
        st4 = self.sap.status(self.sap.session)

        ok = (st4.type.upper() == "S") and ("data was saved" in (st4.text or "").lower())
        if ok:
            # === Export RTF nếu có export_path ===
            if export_path:
                # Lấy ngày từ row đầu tiên (gdatu đã có format dd.mm.yyyy)
                date_str = rows[0].get("gdatu", "") if rows else ""
                self._export_rtf(export_path, date_str)
            
            filled_rates = []
            for r in rows:
                pair = f"{r.get('tcurr','')}/{r.get('fcurr','')}".strip("/")
                rate = r.get("kursm", "")
                filled_rates.append(f"{pair}-{rate}")

            return {
                "ok": True,
                "message": f"Hoàn tất. ({st4.text}) | Rates: {', '.join(filled_rates)}",
                "status": st4.__dict__,
                "filled_rates": filled_rates,
                "rows": len(rows),
            }

        return {"ok": False, "message": f"Error/Warning: [{st4.type}] {st4.text}", "status": st4.__dict__}

    def _export_rtf(self, export_path: str, date_str: str):
        """
        Export dữ liệu exchange rate ra file RTF.
        Chạy ngầm, không báo lỗi nếu thất bại.
        
        Args:
            export_path: Đường dẫn folder (VD: "C:\\Users\\70N7678\\Desktop\\test\\")
            date_str: Ngày theo format dd.mm.yyyy (lấy từ gdatu)
        """
        try:
            # Dùng date_str từ rows_data (đã có format dd.mm.yyyy)
            today = date_str
            filename = f"{today}.RTF"
            
            # Đảm bảo path kết thúc bằng \
            if not export_path.endswith("\\"):
                export_path += "\\"
            
            # 1. Mở OB08
            self.sap.session.findById("wnd[0]/tbar[0]/okcd").text = "/nOB08"
            self.sap.session.findById("wnd[0]").sendVKey(0)
            self.sap.wait_ready(self.sap.session, 10)
            
            # 2. Nhấn F8 (VKey 86) - Execute/Display
            self.sap.session.findById("wnd[0]").sendVKey(86)
            self.sap.wait_ready(self.sap.session, 5)
            
            # 3. Click vào cell để focus (row 7, col 15)
            try:
                lbl = self.sap.safe_find(self.sap.session, "wnd[0]/usr/lbl[15,7]")
                if lbl:
                    lbl.setFocus()
            except:
                pass
            
            # 4. Nhấn F2 (VKey 2) - Display detail
            self.sap.session.findById("wnd[0]").sendVKey(2)
            self.sap.wait_ready(self.sap.session, 5)
            
            # 5. Nhấn nút Selection (btn[29])
            btn_sel = self.sap.safe_find(self.sap.session, "wnd[0]/tbar[1]/btn[29]")
            if btn_sel:
                btn_sel.press()
                self.sap.wait_ready(self.sap.session, 5)
            
            # 6. Nhập ngày filter
            date_field = self.sap.safe_find(self.sap.session, "wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW")
            if date_field:
                date_field.text = today
                # Nhấn Enter để confirm
                self.sap.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                self.sap.wait_ready(self.sap.session, 5)
            
            # 7. Menu: List > Export > Local file (menu[0]/menu[1]/menu[0])
            try:
                self.sap.session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[0]").select()
                self.sap.wait_ready(self.sap.session, 5)
            except:
                pass
            
            # 8. Chọn Rich Text Format và With column headings
            try:
                # Select radio "Rich text format"
                rad_rtf = self.sap.safe_find(self.sap.session, "wnd[1]/usr/radRSAQTEXT-SAVETEXT")
                if rad_rtf:
                    rad_rtf.setFocus()
                    rad_rtf.select()
                
                # Check "with column headings"
                chk_col = self.sap.safe_find(self.sap.session, "wnd[1]/usr/chkRSAQTEXT-SAVCOL")
                if chk_col:
                    chk_col.selected = True
                
                # Nhấn Enter
                self.sap.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                self.sap.wait_ready(self.sap.session, 5)
            except:
                pass
            
            # 9. Nhập path và filename
            try:
                path_field = self.sap.safe_find(self.sap.session, "wnd[1]/usr/ctxtDY_PATH")
                if path_field:
                    path_field.text = export_path
                
                name_field = self.sap.safe_find(self.sap.session, "wnd[1]/usr/ctxtDY_FILENAME")
                if name_field:
                    name_field.text = filename
                
                # Nhấn Generate
                self.sap.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                self.sap.wait_ready(self.sap.session, 5)
            except:
                pass
            
            # 10. Thoát về Easy Access
            self.sap.run_tcode("/n")
                
        except Exception:
            # Không báo lỗi, chỉ bỏ qua
            pass

    def _get_popup_text(self) -> str:
        """
        Lấy nội dung popup. Nhấn Info để lấy chi tiết nếu có.
        """
        try:
            popup = self.sap.safe_find(self.sap.session, "wnd[1]")
            if popup is None:
                return None
            
            # Lấy title
            title = ""
            try:
                title = popup.text or ""
            except:
                pass
            
            # Thử nhấn nút Info (BUTTON_3) để lấy chi tiết
            detail_text = self._get_info_detail()
            
            if detail_text:
                return f"[{title}] {detail_text}"
            
            # Fallback: trả về message theo title
            if "locked" in title.lower():
                return "Dữ liệu đang bị khóa bởi user khác. Vui lòng thử lại sau."
            
            if title:
                return f"[{title}]"
            
            return "Có popup không xác định"
            
        except:
            return None

    def _get_info_detail(self) -> str:
        """
        Nhấn nút Info để lấy chi tiết, rồi đóng popup Info.
        """
        try:
            # Nhấn nút Info (BUTTON_3)
            btn_info = self.sap.safe_find(self.sap.session, "wnd[1]/usr/btnBUTTON_3")
            if btn_info is None:
                return None
            
            btn_info.press()
            self.sap.wait_ready(self.sap.session, 5)
            
            # Đọc text từ wnd[2] (popup Info)
            texts = []
            for row in range(10):
                for col in range(10):
                    try:
                        lbl_id = f"wnd[2]/usr/lbl[{col},{row}]"
                        elem = self.sap.safe_find(self.sap.session, lbl_id)
                        if elem and elem.text:
                            txt = elem.text.strip()
                            if txt and txt not in texts:
                                texts.append(txt)
                    except:
                        pass
            
            # Đóng popup Info (nhấn nút đóng hoặc Enter)
            try:
                btn_close = self.sap.safe_find(self.sap.session, "wnd[2]/tbar[0]/btn[0]")
                if btn_close:
                    btn_close.press()
                else:
                    self.sap.session.findById("wnd[2]").sendVKey(0)
            except:
                pass
            
            self.sap.wait_ready(self.sap.session, 3)
            
            if texts:
                return " ".join(texts)
            
            return None
            
        except:
            return None

    def _close_popup(self):
        """
        Đóng popup bằng cách nhấn No hoặc ESC.
        """
        try:
            # Thử nhấn nút No (BUTTON_2)
            btn_no = self.sap.safe_find(self.sap.session, "wnd[1]/usr/btnBUTTON_2")
            if btn_no:
                btn_no.press()
                return
            
            # Thử nhấn nút Cancel
            btn_cancel = self.sap.safe_find(self.sap.session, "wnd[1]/tbar[0]/btn[12]")
            if btn_cancel:
                btn_cancel.press()
                return
            
            # Nhấn ESC
            self.sap.session.findById("wnd[1]").sendVKey(12)
        except:
            pass
