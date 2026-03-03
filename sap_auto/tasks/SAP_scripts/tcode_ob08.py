from typing import List, Dict, Any
from .sap_base import SapGuiClient


class TCodeOB08:
    TABLE_ID = "wnd[0]/usr/tblSAPL0SAPTCTRL_V_TCURR"

    def __init__(self, sap: SapGuiClient):
        self.sap = sap

    def run(self, rows: List[Dict[str, str]]) -> Dict[str, Any]:
        """
        rows: [{"kurst","gdatu","kursm","fcurr","tcurr"}, ...]
        """
        st = self.sap.run_tcode("OB08")
        if self.sap.is_error(st):
            return {"ok": False, "message": f"Open OB08 error: {st.text}", "status": st.__dict__}

        # New entries (optional)
        btn_new = self.sap.safe_find(self.sap.session, "wnd[0]/tbar[1]/btn[5]")
        if btn_new is not None:
            self.sap.press("wnd[0]/tbar[1]/btn[5]")
            st2 = self.sap.status(self.sap.session)
            if self.sap.is_error(st2):
                return {"ok": False, "message": f"New Entries error: {st2.text}", "status": st2.__dict__}

        # Wait for table (avoid "control not found")
        tbl = self.sap.table_wait(self.TABLE_ID, timeout_s=20)
        if tbl is None:
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
            # return {"ok": True, "message": f"Hoàn tất. ({st4.text})", "status": st4.__dict__}
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