import time
import threading
import subprocess
from dataclasses import dataclass

import psutil
import pythoncom
import win32com.client

log_import = None
try:
    import logging
    log_import = logging.getLogger(__name__)
except:
    pass

def _log(msg):
    if log_import:
        log_import.info(msg)

# Lock chỉ bảo vệ đoạn mở connection (~3-5 giây)
# Tránh race condition khi nhiều user bấm Chạy gần cùng lúc
# Sau khi có session rồi → unlock, tcode chạy song song hoàn toàn
_connect_lock = threading.Lock()


@dataclass
class SapStatus:
    type: str = ""
    text: str = ""


class SapGuiClient:

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
        """Chờ SAP hết busy — giữ nguyên như doc4, không check popup."""
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

    def _check_and_dismiss_system_popup(self) -> bool:
        """
        Tìm và đóng popup "User session closed by system".

        Chiến lược dismiss-then-verify:
        1. Tìm popup theo keyword
        2. Dismiss (click OK)
        3. Verify: nếu session mình vẫn sống → popup của user khác → False
                   nếu session mình chết → popup của mình → True

        Returns:
            True  — popup của session này → session đã chết
            False — không có popup hoặc popup của session khác
        """
        try:
            import win32gui
            import win32con

            SESSION_KEYWORDS = [
                'session closed', 'closed by system', 'user session',
                'session wurde', 'abgemeldet', 'beendet',
            ]

            dismissed = [False]

            def _get_all_text(hwnd):
                texts = []
                own_text = (win32gui.GetWindowText(hwnd) or '').strip()
                if own_text:
                    texts.append(own_text.lower())
                def _collect_child(child_hwnd, _):
                    t = (win32gui.GetWindowText(child_hwnd) or '').strip()
                    if t:
                        texts.append(t.lower())
                try:
                    win32gui.EnumChildWindows(hwnd, _collect_child, None)
                except Exception:
                    pass
                return ' '.join(texts)

            def _find_ok_button(hwnd):
                ok_buttons = []
                def _find(child_hwnd, _):
                    try:
                        cls = win32gui.GetClassName(child_hwnd) or ''
                        if cls.lower() != 'button':
                            return
                        txt = (win32gui.GetWindowText(child_hwnd) or '').strip().upper()
                        if txt in ('OK', '&OK'):
                            ok_buttons.append(child_hwnd)
                    except Exception:
                        pass
                try:
                    win32gui.EnumChildWindows(hwnd, _find, None)
                except Exception:
                    pass
                return ok_buttons[0] if ok_buttons else None

            def _click_popup(hwnd):
                """Thử nhiều cách click OK trên popup."""
                import win32api

                # Bước 1: Đưa popup lên foreground trước
                try:
                    win32gui.ShowWindow(hwnd, 9)   # SW_RESTORE
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(0.2)
                except Exception:
                    pass

                # Bước 2: Thử click OK button
                ok_hwnd = _find_ok_button(hwnd)
                if ok_hwnd:
                    # Cách 2a: BM_CLICK
                    try:
                        win32gui.SendMessage(ok_hwnd, win32con.BM_CLICK, 0, 0)
                        time.sleep(0.3)
                        if not win32gui.IsWindow(hwnd):  # Popup đã đóng
                            _log("[SAP] Click OK thành công (BM_CLICK)")
                            return True
                    except Exception:
                        pass

                    # Cách 2b: WM_LBUTTONDOWN + WM_LBUTTONUP
                    try:
                        WM_LBUTTONDOWN = 0x0201
                        WM_LBUTTONUP   = 0x0202
                        win32gui.PostMessage(ok_hwnd, WM_LBUTTONDOWN, 0x0001, 0)
                        win32gui.PostMessage(ok_hwnd, WM_LBUTTONUP, 0, 0)
                        time.sleep(0.3)
                        if not win32gui.IsWindow(hwnd):
                            _log("[SAP] Click OK thành công (LBUTTONDOWN/UP)")
                            return True
                    except Exception:
                        pass

                # Bước 3: WM_COMMAND IDOK gửi thẳng vào dialog
                try:
                    WM_COMMAND = 0x0111
                    IDOK = 1
                    win32gui.SendMessage(hwnd, WM_COMMAND, IDOK, 0)
                    time.sleep(0.3)
                    if not win32gui.IsWindow(hwnd):
                        _log("[SAP] Click OK thành công (WM_COMMAND IDOK)")
                        return True
                except Exception:
                    pass

                # Bước 4: Enter key
                try:
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(0.2)
                    win32api.keybd_event(0x0D, 0, 0, 0)
                    win32api.keybd_event(0x0D, 0, 0x0002, 0)
                    time.sleep(0.3)
                    _log("[SAP] Gửi Enter key (fallback cuối)")
                    return True
                except Exception as e:
                    _log(f"[SAP] Tất cả phương pháp click đều thất bại: {e}")
                    return False

            def _enum_windows(hwnd, _):
                if dismissed[0]:
                    return
                try:
                    if not win32gui.IsWindowVisible(hwnd):
                        return
                    all_text = _get_all_text(hwnd)
                    if not any(kw in all_text for kw in SESSION_KEYWORDS):
                        return
                    if 'sap' not in all_text:
                        return

                    _log(f"[SAP] Phát hiện popup system: '{all_text[:80]}'")

                    if _click_popup(hwnd):
                        dismissed[0] = True

                except Exception:
                    pass

            win32gui.EnumWindows(_enum_windows, None)

            if not dismissed[0]:
                return False  # Không có popup

            # Verify: đợi rồi kiểm tra session còn sống không
            # Retry nhiều lần vì SAP cần thời gian để thực sự disconnect
            for i in range(6):           # Retry 6 lần × 0.5s = tổng 3s
                time.sleep(0.5)
                try:
                    _ = self.session.findById("wnd[0]").text
                    if i < 5:
                        continue         # Session vẫn sống, thử lại
                    # Sau 3s vẫn sống → popup của user khác
                    _log("[SAP] Popup dismissed, session vẫn sống sau 3s → popup của session khác")
                    return False
                except Exception:
                    # Session COM fail → xác nhận popup là của mình
                    _log("[SAP] Popup dismissed, session COM fail → session này đã chết")
                    return True

            return False

        except Exception as e:
            _log(f"[SAP] _check_and_dismiss_system_popup error: {e}")
            return False

    def is_session_alive(self) -> bool:
        """
        Kiểm tra SAP session còn sống không.
        Giống doc4 — chỉ gọi giữa các row, không gọi trong wait_ready.
        """
        if self.session is None:
            return False

        if self._check_and_dismiss_system_popup():
            return False

        try:
            _ = self.session.findById("wnd[0]").text
            return True
        except Exception:
            return False

    def dismiss_system_popup(self) -> bool:
        """Public method cho tcode gọi trong _safe_reset()"""
        return self._check_and_dismiss_system_popup()

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
        try:
            return self.safe_find(sess, "wnd[0]/usr/txtRSYST-BNAME") is None
        except Exception:
            return False

    def _get_logged_user(self, sess) -> str:
        try:
            info = getattr(sess, "Info", None)
            if info is None:
                return ""
            u = getattr(info, "User", "") or ""
            return str(u).strip()
        except Exception:
            return ""

    def _find_any_logged_in_matching_user(self):
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

    def _wait_new_connection(self, before_count: int, timeout_s: int = 25):
        """
        Chờ connection MỚI sau OpenConnection().
        Dùng count để tránh lấy nhầm session của user khác đang chạy song song.
        """
        t0 = time.time()
        while time.time() - t0 < timeout_s:
            try:
                current_count = self.app.Children.Count
                if current_count > before_count:
                    new_conn = self.app.Children(current_count - 1)
                    if new_conn.Children.Count > 0:
                        return new_conn, new_conn.Children(0)
            except Exception:
                pass
            time.sleep(0.5)
        return self._wait_any_session(timeout_s=5)

    def _wait_any_session(self, timeout_s: int = 25):
        t0 = time.time()
        while time.time() - t0 < timeout_s:
            try:
                for conn in self.app.Children:
                    try:
                        if conn.Children.Count > 0:
                            return conn, conn.Children(0)
                    except Exception:
                        continue
            except Exception:
                pass
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
                _ = new_sess.findById("wnd[0]/tbar[0]/okcd", False)
                return new_sess
            time.sleep(0.5)

        raise RuntimeError("CreateSession() không tạo được session mới usable.")

    # -------- connect/login --------
    def connect(self) -> None:
        pythoncom.CoInitialize()
        self.ensure_sap_running()

        time.sleep(1)

        sap_gui = win32com.client.GetObject("SAPGUI")
        self.app = sap_gui.GetScriptingEngine

        conn, sess = self._find_any_logged_in_matching_user()
        if conn is not None and sess is not None:
            self.connection = conn
            self.session = sess
            self.logged_in_at_start = True

            self.session = self._create_new_session_from_connection(self.connection)
            self.created_session_for_script = True
            self.wait_ready(self.session, 20)
            self.dismiss_popup_if_any()
            self.reset_to_easy_access()
            return

        self.logged_in_at_start = False

        # Lock chỉ bảo vệ đoạn before_count → OpenConnection → lấy session mới
        # Đảm bảo mỗi user lấy đúng connection của mình khi nhiều user connect cùng lúc
        with _connect_lock:
            before_count = 0
            try:
                before_count = self.app.Children.Count
            except Exception:
                pass

            try:
                self.app.OpenConnection(self.sap_entry_name, True)
            except Exception:
                pass

            self.connection, self.session = self._wait_new_connection(
                before_count=before_count, timeout_s=25
            )

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

        try:
            opt1 = self.safe_find(self.session, "wnd[1]/usr/radMULTI_LOGON_OPT1")
            if opt1 is not None:
                opt1.select()
                opt1.setFocus()
                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                self.wait_ready(self.session, 20)
        except Exception:
            pass

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
            try:
                self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                self.session.findById("wnd[0]").sendVKey(0)
                self.wait_ready(self.session, 5)
            except Exception:
                pass

            if self.created_session_for_script:
                try:
                    self.session.findById("wnd[0]").Close()
                except Exception:
                    pass
                try:
                    wnd1 = self.session.findById("wnd[1]", False)
                    if wnd1 is not None:
                        btn_no = self.session.findById("wnd[1]/usr/btnSPOP-OPTION2", False)
                        if btn_no is not None:
                            btn_no.press()
                except Exception:
                    pass
            return

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
