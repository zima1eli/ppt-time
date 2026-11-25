"""Per-user Microsoft PowerPoint event receiver and COM add-in registration."""

from dataclasses import dataclass, replace
import hashlib
import hmac
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
import ntpath
import os
from pathlib import Path
import re
import secrets
import sys
import threading
import time
from slideshow import (SlideSnapshot, _DECK_FILE, _deck_named_in, _native_root, _page_quality,
                       _pid, _process_name, _screen_info, _window_caption)
from wps_pdf import _atomic_write


PROGID = "PptTime.PowerPointBridge"
CLSID = "{F7A81B92-8FBF-4535-B76E-62DD79855F90}"
DLL = PROGID + ".dll"
ASSEMBLY = PROGID + ", Version=1.0.2.0, Culture=neutral, PublicKeyToken=null"
COM_CLASS = PROGID + ".Connect"
OWNER = "ppt time / PowerPoint event bridge"
MARKER = "PptTimeRegistration"
PROG_KEY = "Software\\Classes\\" + PROGID
CLS_KEY = "Software\\Classes\\CLSID\\" + CLSID
OFFICE_KEY = "Software\\Microsoft\\Office\\PowerPoint\\Addins\\" + PROGID
POWERPOINT_CLASS = r"Software\Classes\PowerPoint.Application\CLSID"
POWERPOINT_CLSID = "{91493441-5A91-11CF-8700-00AA0060263B}"
PREF_KEY = r"Software\ppt time\PowerPointBridge"
PREF_NAME = "RegistrationOwner"
INPROC = CLS_KEY + "\\InprocServer32"
ROOTS = (PROG_KEY, CLS_KEY, OFFICE_KEY)
EXPIRES = 3.5
_MISSING = object()


class _Registry:
    """Explicit registry views. Per-user paths shared by WOW64 remain idempotent."""
    def __init__(self):
        import winreg as reg
        self.reg = reg
        native64 = sys.maxsize > 2**32 or bool(os.environ.get("PROCESSOR_ARCHITEW6432"))
        self.views = (reg.KEY_WOW64_64KEY, reg.KEY_WOW64_32KEY) if native64 else (0,)

    def exists(self, view, key, hive="user"):
        reg = self.reg
        root = reg.HKEY_CURRENT_USER if hive == "user" else reg.HKEY_LOCAL_MACHINE
        try:
            with reg.OpenKey(root, key, 0, reg.KEY_READ | view):
                return True
        except FileNotFoundError:
            return False

    def get(self, view, key, name, hive="user"):
        reg = self.reg
        root = reg.HKEY_CURRENT_USER if hive == "user" else reg.HKEY_LOCAL_MACHINE
        try:
            with reg.OpenKey(root, key, 0, reg.KEY_READ | view) as handle:
                return reg.QueryValueEx(handle, name)[0]
        except FileNotFoundError:
            return _MISSING

    def set(self, view, key, name, value):
        reg = self.reg
        with reg.CreateKeyEx(reg.HKEY_CURRENT_USER, key, 0, reg.KEY_SET_VALUE | view) as handle:
            reg.SetValueEx(handle, name, 0, reg.REG_DWORD if type(value) is int else reg.REG_SZ, value)

    def remove(self, view, key, name):
        reg = self.reg
        try:
            with reg.OpenKey(reg.HKEY_CURRENT_USER, key, 0, reg.KEY_SET_VALUE | view) as handle:
                reg.DeleteValue(handle, name)
        except FileNotFoundError:
            pass

    def prune(self, view, key):
        reg = self.reg
        try:
            with reg.OpenKey(reg.HKEY_CURRENT_USER, key, 0, reg.KEY_READ | view) as handle:
                subkeys, values, _ = reg.QueryInfoKey(handle)
            if subkeys == 0 and values == 0:
                reg.DeleteKeyEx(reg.HKEY_CURRENT_USER, key, view)
        except FileNotFoundError:
            pass


def _policy_issue(registry):
    # Read the version of PowerPoint registered on this machine, not obsolete policies for other versions.
    for view in registry.views:
        version = registry.get(view, r"Software\Classes\PowerPoint.Application\CurVer", "", "machine")
        major = str(version).rsplit(".", 1)[-1]
        if major not in {"14", "15", "16"}:
            major = "16"
        key = rf"Software\Policies\Microsoft\Office\{major}.0\PowerPoint\Resiliency"
        for hive in ("user", "machine"):
            for prefix in (r"Software\Policies\Microsoft", r"Software\Microsoft"):
                trust = rf"{prefix}\Office\{major}.0\PowerPoint\Security"
                if str(registry.get(view, trust, "DisableAllAddins", hive)) == "1":
                    return "PowerPoint 信任中心已禁用应用程序加载项。"
                if str(registry.get(view, trust, "RequireAddinSig", hive)) == "1":
                    return "PowerPoint 信任中心要求加载项具有受信任发布者签名。"
            required = registry.get(view, key, "RestrictToList", hive)
            entry = registry.get(view, key + r"\AddinList", PROGID, hive)
            if str(entry) == "0" or (str(required) == "1" and str(entry) not in {"1", "2"}):
                return "PowerPoint 的组织策略限制了此加载项。请联系管理员。"
    return ""


def _registration_plan(dll):
    return {
        PROG_KEY: {MARKER: OWNER, "": COM_CLASS},
        PROG_KEY + r"\CLSID": {"": CLSID},
        CLS_KEY: {MARKER: OWNER, "": COM_CLASS},
        CLS_KEY + r"\ProgId": {"": PROGID},
        INPROC: {"": "mscoree.dll", "ThreadingModel": "Both", "Class": COM_CLASS,
                 "Assembly": ASSEMBLY, "RuntimeVersion": "v4.0.30319", "CodeBase": dll.resolve().as_uri()},
        OFFICE_KEY: {MARKER: OWNER, "FriendlyName": "ppt time · PowerPoint 放映页码",
                     "Description": "向当前用户的 ppt time 报告本地放映窗口和页码", "LoadBehavior": 3},
    }


def _office_enabled(reg, views):
    for view in views:
        behavior = reg.get(view, OFFICE_KEY, "LoadBehavior")
        if type(behavior) is int and behavior & 1:
            return True
    return False


def _check_ownership(reg):
    for view in reg.views:
        for key in ROOTS:
            if reg.exists(view, key) and reg.get(view, key, MARKER) != OWNER:
                raise ValueError("同名注册项已存在且由其他程序管理，无法覆盖。")
            if reg.exists(view, key, "machine"):
                raise ValueError("计算机级已有此加载项，请通过其原安装方式管理。")


def _write_registration(reg, plan):
    _check_ownership(reg)
    edits, created = [], []
    try:
        for view in reg.views:
            for key, values in plan.items():
                if not reg.exists(view, key):
                    created.append((view, key))
                for name, new in values.items():
                    old = reg.get(view, key, name)
                    if old != new:
                        reg.set(view, key, name, new)
                        edits.append((view, key, name, old))
    except OSError:
        for view, key, name, old in reversed(edits):
            if old is _MISSING:
                reg.remove(view, key, name)
            else:
                reg.set(view, key, name, old)
        for view, key in reversed(created):
            reg.prune(view, key)
        raise


def _remove_registration(reg, plan):
    # Take the ownership snapshot first (the Office key may alias across views).
    owned = {(view, key) for view in reg.views for key in ROOTS if reg.get(view, key, MARKER) == OWNER}
    for view in reg.views:
        for key, values in reversed(tuple(plan.items())):
            if not any((view, root) in owned and (key == root or key.startswith(root + "\\")) for root in ROOTS):
                continue
            for name in values:
                reg.remove(view, key, name)
            reg.prune(view, key)


@dataclass(frozen=True)
class _Report:
    stamp: float
    seq: int
    pid: int
    snapshot: SlideSnapshot


class _Handler(BaseHTTPRequestHandler):
    def setup(self):
        super().setup()
        self.connection.settimeout(2.0)

    def log_message(self, *_):
        pass

    def do_POST(self):
        port = self.server.server_address[1]
        status = 400
        if self.headers.get("Host") != f"127.0.0.1:{port}" or self.headers.get("Origin"):
            status = 403
        elif self.path != "/report":
            status = 404
        elif not self.headers.get("Transfer-Encoding"):
            try:
                length = int(self.headers.get("Content-Length", "0"))
                if 0 < length <= 16384:
                    status = self.server.bridge._receive(json.loads(self.rfile.read(length)))
            except (ValueError, UnicodeError, OSError):
                pass
        self.send_response(status)
        self.send_header("Content-Length", "0")
        self.send_header("Cache-Control", "no-store")
        self.end_headers()


def _valid_show(hwnd, pid):
    import win32gui
    try:
        if not hwnd or _pid(hwnd) != pid or not win32gui.IsWindowVisible(hwnd) or _process_name(hwnd) != "powerpnt.exe":
            return False
        if win32gui.GetAncestor(hwnd, 2) != hwnd:
            return False
        cls = win32gui.GetClassName(hwnd).lower()
        if cls == "screenclass":
            return True
        # An event HWND has extra evidence, but never attach to the normal editor/presenter tools.
        caption = _window_caption(hwnd).lower()
        return (cls not in {"pptframeclass", "paneclassdc"} and
                any(word in caption for word in ("slide show", "slideshow", "幻灯片放映")) and
                not any(word in caption for word in ("presenter view", "演讲者视图")))
    except (OSError, win32gui.error):
        return False


def _show_windows(pid):
    import win32gui
    shows = []
    def visit(hwnd, _):
        if _valid_show(hwnd, pid):
            shows.append(hwnd)
        return True
    win32gui.EnumWindows(visit, None)
    return shows


def _matches_doc(a, b):
    if not a or not b:
        return False
    left, right = (ntpath.normcase(ntpath.normpath(x)) for x in (a,b))
    if ntpath.isabs(left) and ntpath.isabs(right):
        return left == right
    return ntpath.basename(left) == ntpath.basename(right)


def _report_audience(report, target, target_document):
    """Resolve Office's handle to the actual audience frame for this report."""
    pid = report.pid
    if target and (_pid(target) != pid or not _valid_show(target, pid)):
        return 0
    raw = report.snapshot.hwnd
    if raw:
        raw_pid = _pid(raw)
        if raw_pid and raw_pid != pid:
            return 0
        if _valid_show(raw, pid):
            return raw if not target or raw == target else 0
        root = _native_root(raw)
        if root and raw_pid == pid and _valid_show(root, pid):
            return root if not target or root == target else 0
    # Some Office views return the editor HWND, a stale handle, or zero.
    # Require both a unique audience in the reporting PID and its document.
    windows = _show_windows(pid)
    if len(windows) != 1 or (target and windows[0] != target):
        return 0
    audience = windows[0]
    document = report.snapshot.document
    if _matches_doc(target_document, document) or (document and _deck_named_in(_window_caption(audience), document)):
        return audience
    return 0


class PowerPointBridge:
    def __init__(self, config_dir, bundle_dir=None, *, registry=None, runtime_dir=None, install_dir=None):
        self.config_dir = Path(config_dir)
        appdata = Path(os.environ.get("APPDATA", Path.home()))
        local = Path(os.environ.get("LOCALAPPDATA", Path.home()))
        self.endpoint = (Path(runtime_dir) if runtime_dir else appdata / "ppt time") / "powerpoint-bridge.json"
        self.install_dir = Path(install_dir) if install_dir else local / "ppt time" / "PowerPointAddin"
        self.bundle_dir = Path(bundle_dir) if bundle_dir else Path(__file__).parent / "powerpoint_addin"
        self.registry = registry
        self.server = None
        self.key = ""
        self.issue = ""
        self.samples = {}
        self.lock = threading.Lock()

    def _registry(self):
        if self.registry is None:
            self.registry = _Registry()
        return self.registry

    def _start_receiver(self):
        if self.server is not None:
            return
        server = ThreadingHTTPServer(("127.0.0.1", 0), _Handler)
        server.daemon_threads = True
        server.bridge = self
        token = secrets.token_urlsafe(32)
        config = {"url": f"http://127.0.0.1:{server.server_address[1]}/report", "key": token}
        try:
            _atomic_write(self.endpoint, json.dumps(config).encode())
        except OSError:
            server.server_close()
            raise
        self.key, self.server = token, server
        self.issue = ""
        self.samples.clear()
        threading.Thread(target=server.serve_forever, kwargs={"poll_interval": 0.1},
                         daemon=True, name="ppt time PowerPoint events").start()

    def start(self):
        if sys.platform != "win32":
            return
        try:
            reg = self._registry()
            owned = [v for v in reg.views if reg.get(v, OFFICE_KEY, MARKER) == OWNER]
            if not owned:
                if self._preference(reg) == "disabled":
                    self.issue = "已在 ppt time 中停用 PowerPoint 页码加载项。"
                return
            if not _office_enabled(reg, owned):
                self.issue = "PowerPoint 已停用此加载项。可在其「COM 加载项」中查看加载状态。"
                return
            self.issue = _policy_issue(reg)
            if not self.issue:
                self._start_receiver()
        except OSError as exc:
            self.issue = "无法开启 PowerPoint 本机联动：" + str(exc)

    def _preference(self, reg):
        return reg.get(reg.views[0], PREF_KEY, PREF_NAME)

    def _register(self, reg, owner):
        src = self.bundle_dir / DLL
        if not src.is_file():
            raise ValueError("此软件包缺少 PowerPoint 加载项 DLL。")
        data = src.read_bytes()
        # Immutable locations let a currently open Office process finish with its old DLL.
        dll = self.install_dir / hashlib.sha256(data).hexdigest() / DLL
        copied = not dll.exists()
        if copied:
            _atomic_write(dll, data)
        elif dll.read_bytes() != data:
            raise ValueError("已安装的加载项内容不匹配。")
        plan = _registration_plan(dll)
        plan[PREF_KEY] = {PREF_NAME: owner}
        try:
            _write_registration(reg, plan)
        except (OSError, ValueError):
            if copied and not any(reg.get(v, INPROC, "CodeBase") == plan[INPROC]["CodeBase"]
                                  for v in reg.views):
                try:
                    dll.unlink(missing_ok=True)
                    dll.parent.rmdir()
                except OSError:
                    pass
            raise

    def install(self):
        if sys.platform != "win32":
            raise ValueError("Microsoft PowerPoint COM 加载项适用于 Windows。")
        reg = self._registry()
        issue = _policy_issue(reg)
        if issue:
            raise ValueError(issue)
        _check_ownership(reg)
        fresh = self.server is None
        self._start_receiver()
        try:
            self._register(reg, "manual")
        except (OSError, ValueError):
            if fresh:
                self.close()
            raise
        return "已为当前用户安装 PowerPoint 放映页码加载项。请完全退出再打开 PowerPoint，并开始放映。"

    def install_for_setup(self):
        """Register for future PowerPoint launches without taking over a running timer."""
        reg = self._registry()
        preference = self._preference(reg)
        if preference == "disabled":
            return "已保留此前在 ppt time 中停用页码联动的选择。"
        if preference == "manual":
            return "已保留手动管理的 PowerPoint 页码联动。"
        issue = _policy_issue(reg)
        if issue:
            return issue
        try:
            _check_ownership(reg)
        except ValueError as exc:
            return str(exc)
        owned = [v for v in reg.views if reg.get(v, OFFICE_KEY, MARKER) == OWNER]
        if owned:
            if not _office_enabled(reg, owned):
                return "已保留 PowerPoint 中停用加载项的状态。"
            return "已保留现有 PowerPoint 页码联动。"
        if preference == "setup":
            return "已保留此前移除 PowerPoint 加载项注册的选择。"
        installed = any(str(reg.get(v, POWERPOINT_CLASS, "", hive)).upper() == POWERPOINT_CLSID
                        for v in reg.views for hive in ("user", "machine"))
        if not installed:
            return "未检测到 Microsoft PowerPoint 桌面版，已跳过页码加载项。"
        self._register(reg, "setup")
        return "已为当前用户启用页码联动，请重新打开 PowerPoint。"

    def uninstall(self):
        if sys.platform != "win32":
            raise ValueError("Microsoft PowerPoint COM 加载项适用于 Windows。")
        reg = self._registry()
        for view in reg.views:
            reg.set(view, PREF_KEY, PREF_NAME, "disabled")
        _remove_registration(reg, _registration_plan(self.bundle_dir / DLL))
        self.close()
        locked = self._remove_dlls()
        self.issue = "已在 ppt time 中停用 PowerPoint 页码加载项。"
        return ("已取消 PowerPoint 加载项注册。完全退出 PowerPoint 后可再次移除仍被占用的文件。"
                if locked else "已移除本软件的 PowerPoint 加载项。重启 PowerPoint 后生效。")

    def remove_for_setup(self):
        reg = self._registry()
        if self._preference(reg) != "setup":
            return "没有需要本安装程序移除的 PowerPoint 页码联动。"
        _remove_registration(reg, _registration_plan(self.bundle_dir / DLL))
        locked = self._remove_dlls()
        for view in reg.views:
            reg.remove(view, PREF_KEY, PREF_NAME)
            reg.prune(view, PREF_KEY)
        return ("已取消 PowerPoint 加载项注册，当前运行的 PowerPoint 仍在使用旧文件。" if locked
                else "已移除安装程序配置的 PowerPoint 页码联动。")

    def _remove_dlls(self):
        # Delete only copies whose bytes identify our exact install layout.
        locked = False
        if self.install_dir.is_dir():
            for dll in self.install_dir.glob("*/" + DLL):
                if not re.fullmatch("[0-9a-f]{64}", dll.parent.name):
                    continue
                try:
                    if hashlib.sha256(dll.read_bytes()).hexdigest() == dll.parent.name:
                        dll.unlink()
                        if not any(dll.parent.iterdir()):
                            dll.parent.rmdir()
                except OSError:
                    locked = True
        return locked

    def _receive(self, data):
        if not isinstance(data, dict):
            return 400
        key = data.get("key")
        if not isinstance(key, str) or not key.isascii() or not self.key or not hmac.compare_digest(key, self.key):
            return 403
        session, seq, pid, hwnd = (data.get(k) for k in ("session", "seq", "pid", "hwnd"))
        active, ended = data.get("active"), data.get("ended", False)
        if (not isinstance(session, str) or not 1 <= len(session) <= 96 or type(seq) is not int or seq < 1
                or type(pid) is not int or not 1 <= pid <= 0xffffffff or type(active) is not bool or type(ended) is not bool):
            return 400
        if hwnd is None:
            hwnd = 0
        if type(hwnd) is not int or not -(2**31) <= hwnd < 2**32:
            return 400
        doc = data.get("document") or ""
        if not isinstance(doc, str) or len(doc) > 4096 or "\0" in doc:
            return 400
        current = total = position = None
        if active and not ended:
            def page_number(name):
                value = data.get(name)
                return value if type(value) is int and 1 <= value <= 1000000 else None
            current, total, position = map(page_number, ("current", "total", "show_position"))
            if current is not None and total is not None and current > total:
                current = None
        snapshot = SlideSnapshot(active, current, total, document=doc, ended=ended, hwnd=hwnd & 0xffffffff,
                                 show_position=position)
        now = time.monotonic()
        with self.lock:
            if self.key != key:
                return 403
            self.samples = {k:v for k,v in self.samples.items() if now - v.stamp < 60}
            old = self.samples.get((pid, session))
            if old and old.seq >= seq:
                return 204
            if not old and len(self.samples) >= 64:
                return 429
            self.samples[(pid, session)] = _Report(now, seq, pid, snapshot)
        return 204

    def _fresh(self):
        now = time.monotonic()
        with self.lock:
            return [r for r in self.samples.values() if now - r.stamp < EXPIRES]

    def merge(self, external):
        if self.server is None or sys.platform != "win32" or (external and external.ended):
            return external
        active = [r for r in self._fresh() if r.snapshot.active]
        if not active:
            return external
        target = external.hwnd if external and external.active else 0
        target_document = external.document if external and external.active else ""
        if target and _process_name(target) != "powerpnt.exe":
            return external
        if (external and external.active and not target and target_document and
                ntpath.splitext(target_document)[1].lower() not in {".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".ppsm"}):
            return external
        candidates = []
        for report in active:
            snapshot = report.snapshot
            if target_document and snapshot.document and not _matches_doc(target_document, snapshot.document):
                continue
            hwnd = _report_audience(report, target, target_document)
            if not hwnd:
                continue
            caption = _window_caption(hwnd)
            if snapshot.document and _DECK_FILE.search(caption) and not _deck_named_in(caption, snapshot.document):
                continue
            if external and external.active and not target and not _matches_doc(target_document, snapshot.document):
                continue
            candidates.append((report, hwnd))
        if len(candidates) != 1:
            return external
        report, hwnd = candidates[0]
        snapshot = report.snapshot
        if not snapshot.ended and _page_quality(external) > _page_quality(snapshot):
            return external
        screen, point = _screen_info(hwnd)
        return replace(snapshot, hwnd=hwnd, screen_name=screen or (external.screen_name if external else ""),
                       screen_point=point or (external.screen_point if external else None),
                       document=snapshot.document or target_document)

    def status(self):
        if self.issue:
            return self.issue
        if self.server is None:
            return "Microsoft PowerPoint 页码加载项尚未启用。"
        fresh = self._fresh()
        playing = [r for r in fresh if r.snapshot.active]
        if len(playing) == 1:
            return "PowerPoint 放映中 · " + (playing[0].snapshot.page_text or "未读取到页码")
        if playing:
            return "已收到多个 PowerPoint 放映窗口的报告，按实际放映窗口匹配。"
        return ("已连接 Microsoft PowerPoint，等待放映。" if fresh else
                "已启用，等待 Microsoft PowerPoint 连接。可在 PowerPoint「选项 → 加载项 → COM 加载项」查看加载状态。")

    def close(self):
        old = self.key
        with self.lock:
            self.key = ""
            self.samples.clear()
        if self.server:
            self.server.shutdown()
            self.server.server_close()
            self.server = None
        if old:
            try:
                if json.loads(self.endpoint.read_text(encoding="utf-8")).get("key") == old:
                    self.endpoint.unlink(missing_ok=True)
            except (OSError, ValueError, AttributeError):
                pass
