"""Per-user WPS Presentation add-in and local slide reports."""

from dataclasses import dataclass, replace
import hmac
import json
import ntpath
import os
from pathlib import Path
import re
import secrets
import socket
import sys
import time
import xml.etree.ElementTree as ET

from powerpoint_bridge import _Registry
from slideshow import (SlideSnapshot, _DECK_FILE, _deck_named_in, _is_wps_show_surface, _native_root,
                       _page_quality, _pid, _process_name, _screen_info, _unique_wps_show, _window_caption)
from wps_pdf import WpsPdfBridge, _atomic_write


ADDIN_NAME = "ppt-time-wpp"
ADDIN_VERSION = "1.0.0"
ASSETS = ("index.html", "ribbon.xml", "main.js")
PREF_KEY = r"Software\ppt time\WpsPresentationBridge"
PREF_NAME = "RegistrationOwner"
_UNCHANGED = object()


def _wpp_at(root):
    if not root:
        return False
    directory = Path(os.path.expandvars(root.strip().strip('"')))
    if not directory.is_absolute() or directory == directory.parent:
        return False
    return ((directory / "wpp.exe").is_file() or (directory / "office6" / "wpp.exe").is_file() or
            any(p.is_file() for p in directory.glob("[0-9]*/office6/wpp.exe")))


def _reg_text(reg, root, path, name, view):
    try:
        with reg.OpenKey(root, path, 0, reg.KEY_READ | view) as handle:
            value = reg.QueryValueEx(handle, name)[0]
            return value if isinstance(value, str) else ""
    except OSError:
        return ""


def _wps_installed():
    import winreg
    import win32ts
    try:
        session = win32ts.ProcessIdToSessionId(os.getpid())
        if any(s == session and name and name.casefold() == "wpp.exe"
               for s, _, name, _ in win32ts.WTSEnumerateProcesses()):
            return True
    except (OSError, win32ts.error):
        pass
    app_path = r"Software\Microsoft\Windows\CurrentVersion\App Paths\wpp.exe"
    uninstall = r"Software\Microsoft\Windows\CurrentVersion\Uninstall"
    for root in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
        for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
            exe = Path(os.path.expandvars(_reg_text(winreg, root, app_path, None, view).strip('"')))
            if exe.name.casefold() == "wpp.exe" and exe.is_file():
                return True
            if _wpp_at(_reg_text(winreg, root, r"Software\Kingsoft\Office\6.0\Common", "InstallRoot", view)):
                return True
            try:
                with winreg.OpenKey(root, uninstall, 0, winreg.KEY_READ | view) as handle:
                    entries = [winreg.EnumKey(handle, i) for i in range(winreg.QueryInfoKey(handle)[0])]
            except OSError:
                entries = []
            for entry in entries:
                path = uninstall + "\\" + entry
                name = _reg_text(winreg, root, path, "DisplayName", view).casefold()
                if name.startswith(("wps office", "kingsoft office", "金山办公", "wps 办公")):
                    if _wpp_at(_reg_text(winreg, root, path, "InstallLocation", view)):
                        return True
    for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
        for progid in ("KWPP.Application", "WPP.Application", "PowerPoint.Application"):
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid + r"\CLSID", 0, winreg.KEY_READ | view) as handle:
                    clsid = winreg.QueryValueEx(handle, None)[0]
                key = rf"CLSID\{clsid}\LocalServer32"
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key, 0, winreg.KEY_READ | view) as handle:
                    command = winreg.QueryValueEx(handle, None)[0]
                match = re.match(r'^\s*"([^"]+)"|^\s*(.+?\.exe)(?=\s|$)', command, re.I)
                if match:
                    exe = Path(os.path.expandvars(match.group(1) or match.group(2)))
                    if exe.name.casefold() in {"wpp.exe", "wps.exe"} and exe.is_file():
                        return True
            except OSError:
                pass
    for var in ("LOCALAPPDATA", "ProgramFiles", "ProgramFiles(x86)"):
        if os.environ.get(var) and _wpp_at(str(Path(os.environ[var]) / "Kingsoft" / "WPS Office")):
            return True
    return False


@dataclass(frozen=True)
class _WppReport(SlideSnapshot):
    owner_hwnd: int = 0


def _read(path):
    try:
        return path.read_bytes()
    except FileNotFoundError:
        return None


def _manifest(original, folder, enable):
    if original is None and not enable:
        return None
    try:
        parser = ET.XMLParser(target=ET.TreeBuilder(insert_comments=True, insert_pis=True))
        root = ET.fromstring(original, parser=parser) if original is not None else ET.Element("jsplugins")
    except ET.ParseError as exc:
        raise ValueError("无法读取 WPS 加载项清单 publish.xml。") from exc
    if root.tag != "jsplugins":
        raise ValueError("WPS 加载项清单格式未知。")
    uri = folder.resolve().as_uri()
    own = []
    for item in root:
        if item.get("name") != ADDIN_NAME or item.get("type", "").lower() != "wpp":
            continue
        if item.tag != "jsplugin" or item.get("url", "").rstrip("/") != uri:
            raise ValueError("WPS 已有同名的演示加载项，位置与 ppt time 安装目录不同。")
        own.append(item)
    if (not own and not enable) or (enable and len(own) == 1 and own[0].get("version") == ADDIN_VERSION):
        return original
    for item in own:
        root.remove(item)
    if enable:
        ET.SubElement(root, "jsplugin", name=ADDIN_NAME, type="wpp", version=ADDIN_VERSION, url=uri)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _registration_disabled(xml):
    return any(item.get("name") == ADDIN_NAME and item.get("type", "").lower() == "wpp"
               and item.get("enable", "").lower() in {"false", "0", "disable", "disabled"}
               for item in ET.fromstring(xml))


def _change(path, previous, data):
    if _read(path) != previous:
        raise OSError("WPS 加载项配置正在被其他程序修改，请重试。")
    if data is None:
        path.unlink(missing_ok=True)
    elif previous != data:
        _atomic_write(path, data)


def _same_document(left, right):
    if not left or not right:
        return False
    # Keep both parents for full file paths, UNC paths and WPS cloud URLs.
    left, right = (ntpath.normcase(ntpath.normpath(name)) for name in (left, right))
    if ntpath.dirname(left) and ntpath.dirname(right):
        return left == right
    return ntpath.basename(left) == ntpath.basename(right)


def _audience(report, target, document):
    """Return a real WPP audience HWND, including when the JS API returned zero."""
    import win32gui
    try:
        raw = report.hwnd
        exact = _native_root(raw) if raw else 0
        if raw and not exact:
            return 0
        owner_pid = 0
        owner_hwnd = getattr(report, "owner_hwnd", 0)
        if owner_hwnd:
            if _process_name(owner_hwnd) not in {"wpp.exe", "wps.exe"} or not _pid(owner_hwnd):
                return 0
            owner_pid = _pid(owner_hwnd)
        if exact:
            if _process_name(exact) not in {"wpp.exe", "wps.exe"} or not _pid(exact):
                return 0
            if owner_pid and owner_pid != _pid(exact):
                return 0
            owner_pid = _pid(exact)
            if not _is_wps_show_surface(exact):
                # WPS sometimes returns the editor handle for the show's HWND.
                exact = 0
        if target and exact and exact != target:
            return 0
        hwnd = target or exact
        if not hwnd:
            possible = []
            def visit(top, _):
                if ((not owner_pid or _pid(top) == owner_pid) and _process_name(top) in {"wpp.exe", "wps.exe"}
                        and _is_wps_show_surface(top)):
                    possible.append(top)
                return True
            win32gui.EnumWindows(visit, None)
            if len(possible) > 1 and report.document:
                possible = [h for h in possible if _deck_named_in(_window_caption(h), report.document)]
            if len(possible) != 1:
                return 0
            hwnd = possible[0]
        owner = _process_name(hwnd)
        if (owner not in {"wpp.exe", "wps.exe"} or (owner_pid and _pid(hwnd) != owner_pid)
                or not _is_wps_show_surface(hwnd)):
            return 0
        title = _window_caption(hwnd)
        title_match = _deck_named_in(title, report.document)
        if report.document and _DECK_FILE.search(title) and not title_match:
            return 0
        if not exact:
            if owner == "wps.exe" and not (owner_pid or title_match or _same_document(report.document, document)):
                return 0
            if not _unique_wps_show(hwnd, _pid(hwnd), check_processes=not owner_pid, document=report.document):
                return 0
        return hwnd
    except (OSError, win32gui.error):
        return 0


class WpsPresentationBridge(WpsPdfBridge):
    """Shares the loopback transport with PDF, with its own token, port and samples."""

    def __init__(self, config_dir, addin_root=None, *, registry=None):
        super().__init__(config_dir, addin_root)
        self.state_file = Path(config_dir) / "wps-presentation-bridge.json"
        self.folder = self.addin_root / f"{ADDIN_NAME}_{ADDIN_VERSION}"
        self.manifest = self.addin_root / "publish.xml"
        self.registry = registry
        self.asset_issue = ""

    def _owner_registry(self):
        if self.registry is None:
            self.registry = _Registry()
        return self.registry

    def _owner(self):
        return self._owner_registry().get(0, PREF_KEY, PREF_NAME)

    def _set_owner(self, owner):
        reg = self._owner_registry()
        if owner is None:
            reg.remove(0, PREF_KEY, PREF_NAME)
            reg.prune(0, PREF_KEY)
        else:
            reg.set(0, PREF_KEY, PREF_NAME, owner)

    def start(self):
        key, _ = self._state()
        if sys.platform != "win32" or not key or self.server:
            return
        try:
            original = _read(self.manifest)
            if not key.isascii() or _manifest(original, self.folder, False) == original:
                self.issue = "WPS 演示加载项配置信息不完整，请在页码联动菜单中重新启用。"
                return
            if _registration_disabled(original):
                self.issue = "WPS 已停用演示页码加载项。"
                return
            self._refresh_assets()
            super().start()
        except (OSError, ValueError) as exc:
            self.issue = str(exc)

    def _assets(self):
        source = Path(__file__).parent / "wps_presentation_addin"
        return {self.folder / name: (source / name).read_bytes() for name in ASSETS}

    def _refresh_assets(self):
        # The current user owns this local manifest entry and the paired endpoint.
        # Preserve the key/port while upgrading also for portable installs.
        try:
            self._apply([(path, _read(path), data) for path, data in self._assets().items()])
            self.asset_issue = ""
        except OSError:
            self.asset_issue = "未能更新本地 WPS 加载项文件，已保留现有连接。"

    def _apply(self, changes, owner=_UNCHANGED):
        applied = []
        try:
            for path, before, data in changes:
                _change(path, before, data)
                if before != data:
                    applied.append((path, data, before))
            if owner is not _UNCHANGED:
                self._set_owner(owner)
        except Exception:
            for path, data, before in reversed(applied):
                if _read(path) == data:
                    _change(path, data, before)
            raise

    def _install(self, owner):
        if sys.platform != "win32":
            raise ValueError("WPS 演示加载项适用于 Windows。")
        previous = _read(self.manifest)
        manifest = _manifest(previous, self.folder, True)
        files = self._assets()
        key, port = self._state()
        if not key and _manifest(previous, self.folder, False) == previous:
            if any((self.folder / name).exists() for name in (*ASSETS, "connection.js")):
                raise ValueError("同名 WPS 演示加载项目录已有文件，无法确认属于 ppt time。")
        key = key if key and key.isascii() else secrets.token_urlsafe(32)
        newly_started = owner == "manual" and self.server is None
        if newly_started:
            try:
                self._listen(key, port)
            except OSError:
                if not port:
                    raise
                self._listen(key, 0)
        try:
            if self.server:
                key, port = self.key, self.server.server_address[1]
            elif not port:
                # Setup exits before the timer starts. Never replace a live timer's endpoint.
                with socket.socket() as reserved:
                    reserved.bind(("127.0.0.1", 0))
                    port = reserved.getsockname()[1]
            config = dict(url=f"http://127.0.0.1:{port}/report", key=key)
            files[self.folder / "connection.js"] = ("window.pptTimeWppBridge = " + json.dumps(config) + ";\n").encode()
            files[self.state_file] = json.dumps(dict(key=key, port=port)).encode()
            changes = [(path, _read(path), data) for path, data in files.items()]
            changes.append((self.manifest, previous, manifest))
            self._apply(changes, owner)
            self.asset_issue = ""
        except Exception:
            if newly_started:
                self.close()
            raise
        return "已安装 WPS 演示页码加载项。请保存文稿，完全退出并重新打开 WPS，然后按 F5 放映。"

    def install(self):
        return self._install("manual")

    def install_for_setup(self):
        if sys.platform != "win32" or not _wps_installed():
            return "未检测到 WPS 演示桌面版，已跳过页码加载项。"
        owner = self._owner()
        if owner == "disabled":
            return "已保留手动管理的 WPS 演示页码联动设置。"
        previous = _read(self.manifest)
        registered = _manifest(previous, self.folder, False) != previous
        if registered and _registration_disabled(previous):
            return "已保留 WPS 中停用加载项的设置。"
        if owner == "setup" and not registered:
            return "已保留此前移除 WPS 演示加载项的选择。"
        if owner == "manual" or (owner != "setup" and (registered or self.state_file.exists())):
            # Before 1.0.8 every WPP registration was created manually and had no owner marker.
            if registered and self._state()[0]:
                self._refresh_assets()
            return self.asset_issue or "已保留现有 WPS 演示页码联动。"
        return self._install("setup")

    def _remove(self, for_setup):
        previous = _read(self.manifest)
        remaining = _manifest(previous, self.folder, False)
        registered = remaining != previous
        if not registered and not self.state_file.exists():
            self._set_owner("disabled")
            self.close()
            self.issue = ""
            return "WPS 演示页码加载项尚未启用。"
        # Rollback must not re-enable a plugin the user already removed in WPS.
        owner = None if for_setup and registered else "disabled"
        if for_setup and registered and _registration_disabled(previous):
            owner = "disabled"
        changes = [(self.manifest, previous, remaining)]
        paths = [self.folder / name for name in (*ASSETS, "connection.js")]
        changes.extend((p, _read(p), None) for p in [*paths, self.state_file])
        self._apply(changes, owner)
        self.close()
        if self.folder.is_dir() and not any(self.folder.iterdir()):
            self.folder.rmdir()
        self.issue = ""
        self.asset_issue = ""
        return "已移除 ppt time 的 WPS 演示加载项。重新打开 WPS 后生效。"

    def uninstall(self):
        return self._remove(False)

    def remove_for_setup(self):
        if sys.platform != "win32" or self._owner() != "setup":
            return "没有需要本安装程序移除的 WPS 演示页码联动。"
        return self._remove(True)

    def _receive(self, data):
        if not isinstance(data, dict):
            return 400
        token = data.get("key")
        if not isinstance(token, str) or not token.isascii() or not self.key or not hmac.compare_digest(token, self.key):
            return 403
        session, seq, active = (data.get(k) for k in ("session", "seq", "active"))
        ended, hwnd, doc = data.get("ended", False), data.get("hwnd", 0), data.get("document", "")
        owner_hwnd = data.get("owner_hwnd", 0)
        if (not isinstance(session, str) or not 1 <= len(session) <= 96 or type(seq) is not int or seq < 1
                or type(active) is not bool or type(ended) is not bool or type(hwnd) is not int
                or not -(2**31) <= hwnd < 2**32 or type(owner_hwnd) is not int or not -(2**31) <= owner_hwnd < 2**32
                or not isinstance(doc, str) or len(doc) > 4096 or "\0" in doc):
            return 400
        current = total = position = None
        if active and not ended:
            def number(name):
                value = data.get(name)
                return value if type(value) is int and 1 <= value <= 1000000 else None
            current, total, position = map(number, ("current", "total", "show_position"))
            if current and total and current > total:
                current = None
        sample = _WppReport(active, current, total, document=doc, ended=ended, hwnd=hwnd & 0xffffffff,
                            show_position=position, owner_hwnd=owner_hwnd & 0xffffffff)
        now = time.monotonic()
        with self.lock:
            if token != self.key:
                return 403
            self.samples = {k: v for k, v in self.samples.items() if now - v[0] < 60}
            old = self.samples.get(session)
            if old and seq <= old[1]:
                return 204
            if old is None and len(self.samples) >= 32:
                return 429
            self.samples[session] = (now, seq, sample)
        return 204

    def merge(self, external):
        if sys.platform != "win32" or self.server is None or (external and external.ended):
            return external
        target = external.hwnd if external and external.active else 0
        document = external.document if external and external.active else ""
        if external and external.active and not target and not document:
            return external
        extension = ntpath.splitext(document)[1]
        if extension and not _DECK_FILE.fullmatch(extension):
            return external
        candidates = []
        for sample in self._fresh():
            if not sample.active or (document and sample.document and not _same_document(document, sample.document)):
                continue
            hwnd = _audience(sample, target, document)
            if hwnd:
                candidates.append((sample, hwnd))
        if len(candidates) != 1:
            return external
        sample, hwnd = candidates[0]
        if not sample.ended and _page_quality(external) > _page_quality(sample):
            return external
        screen, point = _screen_info(hwnd)
        return replace(sample, hwnd=hwnd, document=sample.document or document,
                       screen_name=screen or (external.screen_name if external else ""),
                       screen_point=point or (external.screen_point if external else None))

    def status(self):
        if self.issue:
            return self.issue
        if self.server is None:
            return "WPS 演示页码加载项尚未启用。"
        fresh = self._fresh()
        active = [s for s in fresh if s.active]
        if len(active) == 1:
            status = "已收到 WPS 演示放映报告 · " + (active[0].page_text or "未读取到页码")
        elif active:
            status = "已收到多个 WPS 演示放映报告，按实际放映窗口匹配。"
        else:
            status = "已连接 WPS 演示，等待放映。" if fresh else "已启用，等待 WPS 演示加载项连接。"
        return status + (" " + self.asset_issue if self.asset_issue else "")
