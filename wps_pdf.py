"""An opt-in loopback receiver for presentation state from the WPS PDF add-in."""

from dataclasses import replace
import hmac
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
import os
from pathlib import Path
import secrets
import sys
import tempfile
import threading
import time
import xml.etree.ElementTree as ET

from slideshow import SlideSnapshot, _process_name, _screen_info


ADDIN_NAME = "ppt-time-pdf"
ADDIN_VERSION = "1.0"
ASSETS = ("index.html", "index_TaskPane.html", "ribbon.xml", "reader.js")
ACTIVE_SECONDS = 3.5


def _atomic_write(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(dir=path.parent, prefix=".ppt-time-", delete=False) as temp:
        temp.write(data)
        tmp_path = Path(temp.name)
    try:
        os.replace(tmp_path, path)
    finally:
        tmp_path.unlink(missing_ok=True)


def _register(path, enable):
    """Change only our entry in WPS's per-user add-in manifest."""
    original = path.read_bytes() if path.exists() else None
    if original is None and not enable:
        return
    try:
        parser = ET.XMLParser(target=ET.TreeBuilder(insert_comments=True, insert_pis=True))
        root = ET.fromstring(original, parser=parser) if original else ET.Element("jsplugins")
    except ET.ParseError as exc:
        raise ValueError("WPS 加载项清单无法读取，请先修复 publish.xml。") from exc
    if root.tag != "jsplugins":
        raise ValueError("WPS 加载项清单格式未知。")
    old = [item for item in root if item.tag == "jsplugin" and item.get("name") == ADDIN_NAME
           and item.get("type", "").lower() == "pdf"]
    if not old and not enable:
        return
    for item in old:
        root.remove(item)
    if enable:
        ET.SubElement(root, "jsplugin", name=ADDIN_NAME, type="pdf", version=ADDIN_VERSION,
                      enable="enable_dev", url="file://")
    current = path.read_bytes() if path.exists() else None
    if original != current:
        raise OSError("WPS 加载项清单刚刚被其他程序修改，请重试。")
    _atomic_write(path, ET.tostring(root, encoding="utf-8", xml_declaration=True))


def _wps_pdf_window(document, candidate):
    if sys.platform != "win32":
        return 0
    import win32api
    import win32con
    import win32gui
    name = Path(document).name.casefold()

    def matches(hwnd):
        if not hwnd or not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd):
            return False
        try:
            monitor = win32api.MonitorFromWindow(hwnd, win32con.MONITOR_DEFAULTTONULL)
            if not monitor:
                return False
            bounds = win32api.GetMonitorInfo(monitor)["Monitor"]
            if any(abs(a - b) > 4 for a, b in zip(win32gui.GetWindowRect(hwnd), bounds)):
                return False
            process = _process_name(hwnd)
            if process == "wpspdf.exe":
                return True
            # Some versions consolidate several editors into wps.exe. Only
            # direct commands there if the named PDF really owns this window.
            title = win32gui.GetWindowText(hwnd).casefold()
            return process == "wps.exe" and bool(name and name in title)
        except OSError:
            return False

    for hwnd in (candidate, win32gui.GetForegroundWindow()):
        if matches(hwnd):
            return hwnd
    found = []
    def visit(hwnd, _):
        if matches(hwnd):
            found.append(hwnd)
        return True
    win32gui.EnumWindows(visit, None)
    return found[0] if len(found) == 1 else 0


class _ReportHandler(BaseHTTPRequestHandler):
    def setup(self):
        super().setup()
        self.connection.settimeout(2.0)

    def log_message(self, *_):
        pass

    def _reply(self, status):
        self.send_response(status)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", "0")
        self.end_headers()

    def do_OPTIONS(self):
        if self.path != "/report":
            self._reply(404)
            return
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "content-type")
        self.send_header("Access-Control-Allow-Private-Network", "true")
        self.send_header("Content-Length", "0")
        self.end_headers()

    def do_POST(self):
        if self.path != "/report":
            self._reply(404)
            return
        try:
            if self.headers.get("Transfer-Encoding"):
                raise ValueError()
            length = int(self.headers["Content-Length"])
            if not 0 < length <= 8192:
                raise ValueError()
            message = json.loads(self.rfile.read(length))
            status = self.server.bridge._receive(message)
        except (ValueError, TypeError, UnicodeError, OSError):
            status = 400
        self._reply(status)


class WpsPdfBridge:
    def __init__(self, config_dir, addin_root=None):
        self.state_file = Path(config_dir) / "wps-pdf-bridge.json"
        roaming = Path(os.environ.get("APPDATA", Path.home()))
        self.addin_root = Path(addin_root) if addin_root is not None else roaming / "kingsoft" / "wps" / "jsaddons"
        self.server = None
        self.lock = threading.Lock()
        self.samples = {}
        self.key = ""
        self.issue = ""

    def _state(self):
        try:
            state = json.loads(self.state_file.read_text(encoding="utf-8"))
            key, port = state["key"], state["port"]
            if isinstance(key, str) and len(key) >= 32 and type(port) is int and 0 < port < 65536:
                return key, port
        except (OSError, ValueError, KeyError, TypeError):
            pass
        return "", 0

    def _listen(self, key, port):
        server = ThreadingHTTPServer(("127.0.0.1", port), _ReportHandler)
        server.daemon_threads = True
        server.bridge = self
        self.key = key
        self.server = server
        self.issue = ""
        threading.Thread(target=server.serve_forever, kwargs={"poll_interval": 0.1},
                         name="ppt time · WPS PDF", daemon=True).start()

    def start(self):
        key, port = self._state()
        if sys.platform == "win32" and key and self.server is None:
            try:
                self._listen(key, port)
            except OSError:
                self.issue = "本机连接端口不可用，可重新安装加载项以更换端口。"

    def install(self):
        if sys.platform != "win32":
            raise ValueError("WPS PDF 加载项适用于 Windows。")
        key, port = self._state()
        newly_started = self.server is None
        if newly_started:
            key = key or secrets.token_urlsafe(32)
            try:
                self._listen(key, port)
            except OSError:
                if not port:
                    raise
                self._listen(key, 0)
        try:
            port = self.server.server_address[1]
            local = self.addin_root / f"{ADDIN_NAME}_{ADDIN_VERSION}"
            assets = Path(__file__).parent / "wps_pdf_addin"
            for name in ASSETS:
                _atomic_write(local / name, (assets / name).read_bytes())
            config = {"url":f"http://127.0.0.1:{port}/report", "key":self.key}
            script = "window.pptTimePdfBridge = " + json.dumps(config) + ";\n"
            _atomic_write(local / "connection.js", script.encode())
            _register(self.addin_root / "publish.xml", True)
            _atomic_write(self.state_file, json.dumps({"key":self.key, "port":port}).encode())
        except Exception:
            if newly_started:
                self.close()
            raise
        return "已安装 WPS PDF 加载项。请完全退出并重新打开 WPS，然后放映 PDF。"

    def uninstall(self):
        local = self.addin_root / f"{ADDIN_NAME}_{ADDIN_VERSION}"
        _register(self.addin_root / "publish.xml", False)
        self.close()
        self.state_file.unlink(missing_ok=True)
        for name in (*ASSETS, "connection.js"):
            (local / name).unlink(missing_ok=True)
        if local.exists() and not any(local.iterdir()):
            local.rmdir()
        self.issue = ""
        return "已移除 ppt time 的 WPS PDF 加载项。重启 WPS 后生效。"

    def _receive(self, data):
        if not isinstance(data, dict):
            return 400
        token = data.get("key")
        if not self.key or not isinstance(token, str) or not token.isascii() or not hmac.compare_digest(token, self.key):
            return 403
        session, seq, active = (data.get(k) for k in ("session", "seq", "active"))
        if not isinstance(session, str) or not 1 <= len(session) <= 64 or type(seq) is not int or seq < 1 or type(active) is not bool:
            return 400
        doc = data.get("document") or ""
        if not isinstance(doc, str) or len(doc) > 2048:
            return 400
        current, total = data.get("current"), data.get("total")
        if type(current) is not int or type(total) is not int or not 1 <= current <= total <= 1000000:
            current = total = None
        snapshot = SlideSnapshot(active, current, total, document=doc)
        now = time.monotonic()
        with self.lock:
            self.samples = {k:v for k,v in self.samples.items() if now - v[0] < 30}
            old = self.samples.get(session)
            if old and old[1] >= seq:
                return 204
            if old is None and len(self.samples) >= 32:
                return 429
            self.samples[session] = (now, seq, snapshot)
        return 204

    def _fresh(self):
        now = time.monotonic()
        with self.lock:
            return [s for t, _, s in self.samples.values() if now - t < ACTIVE_SECONDS]

    def merge(self, external):
        if self.server is None:
            return external
        fresh = self._fresh()
        playing = [s for s in fresh if s.active]
        if len(playing) != 1:
            return external
        snapshot = playing[0]
        if external and external.active and external.hwnd:
            if _process_name(external.hwnd) not in {"wpspdf.exe", "wps.exe"} or external.page_text:
                return external
        hwnd = _wps_pdf_window(snapshot.document, external.hwnd if external else 0)
        if hwnd:
            screen, point = _screen_info(hwnd)
            return replace(snapshot, hwnd=hwnd, screen_name=screen, screen_point=point)
        return snapshot

    def status(self):
        if self.issue:
            return self.issue
        if self.server is None:
            return "WPS PDF 加载项尚未启用。"
        fresh = self._fresh()
        playing = [s for s in fresh if s.active]
        if len(playing) > 1:
            return "多个 WPS PDF 正在放映，等待确定当前文稿。"
        if playing:
            page = playing[0].page_text or "未读取到页码"
            return f"WPS PDF 放映中 · {page}"
        return "已连接 WPS PDF，等待放映。" if fresh else "加载项已启用，等待 WPS PDF 连接。"

    def close(self):
        if self.server is not None:
            self.server.shutdown()
            self.server.server_close()
            self.server = None
        self.key = ""
        with self.lock:
            self.samples.clear()
