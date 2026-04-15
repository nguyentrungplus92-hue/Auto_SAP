"""
SAP Runner - Standalone script, KHÔNG có Django dependency.
Nhận JSON từ stdin, chạy SAP job, trả JSON về stdout.

Cách dùng:
    echo '{"sap_entry": "V2Q", ...}' | python sap_runner.py
"""
import sys
import os
import json

def main():
    try:
        raw = sys.stdin.read()
        data = json.loads(raw)

        sap_scripts_dir = data['sap_scripts_dir']
        sap_entry       = data['sap_entry']
        creds           = data['creds']
        tcode_module    = data['tcode_module']
        tcode_class     = data['tcode_class']
        tcode_method    = data['tcode_method']
        items           = data['items']

        # Thêm SAP_scripts vào sys.path
        if sap_scripts_dir not in sys.path:
            sys.path.insert(0, sap_scripts_dir)

        import pythoncom
        pythoncom.CoInitialize()

        from sap_base import SapGuiClient

        # Import tcode class động
        mod = __import__(tcode_module, fromlist=[tcode_class])
        TCodeClass = getattr(mod, tcode_class)

        with SapGuiClient(
            sap_entry_name=sap_entry,
            client_no=creds['client'],
            username=creds['username'],
            password=creds['password'],
        ) as sap:
            st = sap.last_login_status
            if sap.is_error(st) or ('name or password is incorrect' in (st.text or '').lower()):
                print(json.dumps({'ok': False, 'login_error': True, 'message': f"SAP Login Error: {st.text}"}))
                return

            tcode = TCodeClass(sap)
            method = getattr(tcode, tcode_method)
            res = method(items)
            print(json.dumps({'ok': True, 'res': res}))

    except Exception as e:
        import traceback
        print(json.dumps({
            'ok': False,
            'login_error': False,
            'message': str(e),
            'traceback': traceback.format_exc()
        }))


if __name__ == '__main__':
    main()
