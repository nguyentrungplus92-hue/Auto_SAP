from sap_base import SapGuiClient
from tcode_ob08 import TCodeOB08

rows = [
    {"kurst": "m", "gdatu": "07.01.2026", "kursm": "123", "fcurr": "vnd", "tcurr": "usd"},
    {"kurst": "m", "gdatu": "07.01.2026", "kursm": "122", "fcurr": "eur", "tcurr": "usd"},
]

with SapGuiClient(
    sap_entry_name="V2Q",
    client_no="250",
    username="ITS-2015030",
    password="Trung999",
) as sap:

    st = sap.last_login_status
    if sap.is_error(st) or ("name or password is incorrect" in (st.text or "").lower()):
        print({"ok": False, "message": f"Error (Login): {st.text}", "status": st.__dict__})
    else:
        ob08 = TCodeOB08(sap)
        res = ob08.run(rows)
        print(res)