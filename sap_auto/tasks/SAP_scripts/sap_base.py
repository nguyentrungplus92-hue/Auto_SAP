import time
import subprocess
from dataclasses import dataclass

import psutil
import pythoncom
import win32com.client


@dataclass
class SapStatus:
    type: str = ""   # S/W/E/A...
    text: str = ""


class SapGuiClient:
    """
    Behavior:
    - If there is an existing logged-in session AND it is logged in with EXACT username == self.username:
        * DO NOT OpenConnection again (avoid extra login window)
        * Create a NEW session (window) for the script from that connection
        * /n to reset to Easy Access
        * Run tcodes in that new session
        * Close ONLY that new session at the end
    - Otherwise (no logged-in session, OR logged-in with different user, OR can't read current user):
        * OpenConnection(entry)
        * Login with self.username/self.password
        * Run tcodes
        * /nex at the end
    """

    def __init__(
        self,
        sap_entry_name: str,
        client_no: str,
        username: str,
        password: str,
        saplogon_path: str = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",
        wait_open_s: int = 8,
    ):
        self.sap_entry_name = sap_entry_name
        self.client_no = client_no
        self.username = username
        self.password = password
        self.saplogon_path = saplogon_path
        self.wait_open_s = wait_open_s

        self.app = None
        self.connection = None
        self.session = None
        self.last_login_status = SapStatus()

        self.sap_was_running_at_start = False
        self.logged_in_at_start = False
        self.created_session_for_script = False

    # -------- process helpers --------
    @staticmethod
    def _is_running(exe: str) -> bool:
        for p in psutil.process_iter(["name"]):
            try:
                name = (p.info.get("name") or "").lower()
                if name == exe.lower():
                    return True
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
        return False

    def ensure_sap_running(self) -> None:
        self.sap_was_running_at_start = (
            self._is_running("saplogon.exe") or self._is_running("sapgui.exe") or self._is_running("SAPgui.exe")
        )
        if not self.sap_was_running_at_start:
            subprocess.Popen([self.saplogon_path])
            time.sleep(self.wait_open_s)

    # -------- SAP helpers --------
    @staticmethod
    def wait_ready(session, timeout_s: int = 20) -> None:
        t0 = time.time()
        while time.time() - t0 < timeout_s:
            try:
                if session is not None and session.Busy is False:
                    return
            except Exception:
                pass
            time.sleep(0.3)

    @staticmethod
    def safe_find(session, obj_id: str):
        try:
            return session.findById(obj_id, False)
        except Exception:
            return None

    @staticmethod
    def status(session) -> SapStatus:
        try:
            sbar = session.findById("wnd[0]/sbar")
            return SapStatus(
                type=(getattr(sbar, "MessageType", "") or ""),
                text=(getattr(sbar, "Text", "") or ""),
            )
        except Exception:
            return SapStatus()

    @staticmethod
    def is_error(st: SapStatus) -> bool:
        return st.type.upper() in ("E", "A")

    def dismiss_popup_if_any(self) -> None:
        if self.session is None:
            return
        wnd1 = self.safe_find(self.session, "wnd[1]")
        if wnd1 is not None:
            try:
                self.session.findById("wnd[1]").sendVKey(0)
                self.wait_ready(self.session, 10)
            except Exception:
                pass

    def reset_to_easy_access(self) -> None:
        if self.session is None:
            return
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            self.session.findById("wnd[0]").sendVKey(0)
            self.wait_ready(self.session, 10)
        except Exception:
            pass
        self.dismiss_popup_if_any()

    # -------- user/session detection --------
    def _is_logged_in_session(self, sess) -> bool:
        """
        If login screen is shown, txtRSYST-BNAME exists.
        If not exists => past logon screen (considered logged-in).
        """
        try:
            return self.safe_find(sess, "wnd[0]/usr/txtRSYST-BNAME") is None
        except Exception:
            return False

    def _get_logged_user(self, sess) -> str:
        """
        Return current logged-in SAP username if available.
        Most reliable is sess.Info.User.
        If cannot read, return "" (we treat as NOT MATCH to be strict).
        """
        try:
            info = getattr(sess, "Info", None)
            if info is None:
                return ""
            u = getattr(info, "User", "") or ""
            return str(u).strip()
        except Exception:
            return ""

    def _find_any_logged_in_matching_user(self):
        """
        Find an already-logged-in session whose user == self.username.
        If session is logged in but with different user -> ignore
        If cannot read user -> ignore (strict)
        """
        target = (self.username or "").strip().lower()

        try:
            for conn in self.app.Children:
                try:
                    for sess in conn.Children:
                        if sess is None:
                            continue
                        self.wait_ready(sess, 3)

                        if not self._is_logged_in_session(sess):
                            continue

                        logged_user = self._get_logged_user(sess).lower()
                        if logged_user and logged_user == target:
                            return conn, sess
                except Exception:
                    continue
        except Exception:
            pass

        return None, None

    def _wait_any_session(self, timeout_s: int = 25):
        """
        Wait until ANY session exists (after OpenConnection).
        Returns (conn, sess) or (None, None).
        """
        t0 = time.time()
        while time.time() - t0 < timeout_s:
            # enumerate
            try:
                for conn in self.app.Children:
                    try:
                        if conn.Children.Count > 0:
                            return conn, conn.Children(0)
                    except Exception:
                        continue
            except Exception:
                pass

            # active fallback
            try:
                conn = self.app.ActiveConnection
                sess = conn.ActiveSession
                if sess is not None:
                    return conn, sess
            except Exception:
                pass

            time.sleep(0.5)

        return None, None

    def _create_new_session_from_connection(self, conn):
        """Create a new session window from an existing logged-in connection."""
        base = None
        try:
            if conn.Children.Count > 0:
                base = conn.Children(0)
        except Exception:
            base = None

        if base is None:
            raise RuntimeError("Không tìm thấy base session để CreateSession().")

        before = int(conn.Children.Count)
        base.CreateSession()
        time.sleep(1)

        t0 = time.time()
        while time.time() - t0 < 10:
            after = int(conn.Children.Count)
            if after > before:
                new_sess = conn.Children(conn.Children.Count - 1)
                _ = new_sess.findById("wnd[0]/tbar[0]/okcd", False)  # probe
                return new_sess
            time.sleep(0.5)

        raise RuntimeError("CreateSession() không tạo được session mới usable.")

    # -------- connect/login --------
    def connect(self) -> None:
        pythoncom.CoInitialize()
        self.ensure_sap_running()

        sap_gui = win32com.client.GetObject("SAPGUI")
        self.app = sap_gui.GetScriptingEngine

        # 1) If ANY logged-in session exists AND matches exact username -> reuse (DO NOT OpenConnection)
        conn, sess = self._find_any_logged_in_matching_user()
        if conn is not None and sess is not None:
            self.connection = conn
            self.session = sess
            self.logged_in_at_start = True

            # Create a NEW session for script
            self.session = self._create_new_session_from_connection(self.connection)
            self.created_session_for_script = True
            self.wait_ready(self.session, 20)
            self.dismiss_popup_if_any()
            self.reset_to_easy_access()
            return

        # 2) Otherwise -> OpenConnection + login as self.username
        self.logged_in_at_start = False
        try:
            self.app.OpenConnection(self.sap_entry_name, True)
        except Exception:
            pass

        self.connection, self.session = self._wait_any_session(timeout_s=25)
        if self.session is None:
            raise RuntimeError("Không lấy được SAP session sau OpenConnection().")

        self.wait_ready(self.session, 20)
        self.dismiss_popup_if_any()

    def login_if_needed(self) -> SapStatus:
        if self.session is None:
            return SapStatus("E", "No session")

        if self.safe_find(self.session, "wnd[0]/usr/txtRSYST-BNAME") is None:
            return SapStatus("S", "Already logged in")

        self.session.findById("wnd[0]/usr/txtRSYST-MANDT").text = self.client_no
        self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = self.username
        self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = self.password
        self.session.findById("wnd[0]").sendVKey(0)

        self.wait_ready(self.session, 20)
        self.dismiss_popup_if_any()
        return self.status(self.session)

    # -------- Generic actions --------
    def run_tcode(self, tcode: str) -> SapStatus:
        self.session.findById("wnd[0]/tbar[0]/okcd").text = tcode
        self.session.findById("wnd[0]").sendVKey(0)
        self.wait_ready(self.session, 20)
        self.dismiss_popup_if_any()
        return self.status(self.session)

    def press(self, obj_id: str) -> None:
        self.session.findById(obj_id).press()
        self.wait_ready(self.session, 20)
        self.dismiss_popup_if_any()

    def set_text(self, obj_id: str, value: str) -> None:
        self.session.findById(obj_id).text = value

    def send_enter(self) -> SapStatus:
        self.session.findById("wnd[0]").sendVKey(0)
        self.wait_ready(self.session, 20)
        self.dismiss_popup_if_any()
        return self.status(self.session)

    # -------- Table helpers --------
    def table_wait(self, table_id: str, timeout_s: int = 15):
        t0 = time.time()
        while time.time() - t0 < timeout_s:
            obj = self.safe_find(self.session, table_id)
            if obj is not None:
                return obj
            time.sleep(0.3)
        return None

    def table_scroll_to(self, table_id: str, top_row: int) -> None:
        tbl = self.safe_find(self.session, table_id)
        if tbl is None:
            return
        try:
            tbl.VerticalScrollbar.Position = int(top_row)
            time.sleep(0.2)
        except Exception:
            pass

    def table_set_cell(self, table_id: str, field: str, col: int, row: int, value: str) -> None:
        cell_id = f"{table_id}/ctxt{field}[{col},{row}]"
        try:
            self.session.findById(cell_id).text = value
            return
        except Exception:
            pass
        cell_id = f"{table_id}/txt{field}[{col},{row}]"
        self.session.findById(cell_id).text = value

    # -------- finalize --------
    def finalize_close(self) -> None:
        if self.session is None:
            return

        if self.logged_in_at_start:
            # exit tcode in script session
            try:
                self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                self.session.findById("wnd[0]").sendVKey(0)
                self.wait_ready(self.session, 5)
            except Exception:
                pass

            # close only script window
            if self.created_session_for_script:
                try:
                    self.session.findById("wnd[0]").Close()
                except Exception:
                    pass

                # choose No if asked to save
                try:
                    wnd1 = self.session.findById("wnd[1]", False)
                    if wnd1 is not None:
                        btn_no = self.session.findById("wnd[1]/usr/btnSPOP-OPTION2", False)
                        if btn_no is not None:
                            btn_no.press()
                except Exception:
                    pass
            return

        # script opened login -> exit SAP
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
            self.session.findById("wnd[0]").sendVKey(0)
            self.wait_ready(self.session, 5)
        except Exception:
            try:
                self.session.findById("wnd[0]").Close()
            except Exception:
                pass

        try:
            wnd1 = self.session.findById("wnd[1]", False)
            if wnd1 is not None:
                btn_yes = self.session.findById("wnd[1]/usr/btnSPOP-OPTION1", False)
                if btn_yes is not None:
                    btn_yes.press()
        except Exception:
            pass

    def __enter__(self):
        self.connect()
        self.last_login_status = self.login_if_needed()
        return self

    def __exit__(self, exc_type, exc, tb):
        self.finalize_close()
        return False