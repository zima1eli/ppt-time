"""Read a running presentation without activating its window or opening Office."""

import ntpath
import re
import subprocess
import sys
import threading
import time
from itertools import islice
from urllib.parse import unquote, urlsplit
from dataclasses import dataclass, replace

from PySide6.QtCore import QObject, Qt, Signal, Slot
from page_diagnostics import (capture_page_read, discovery_notes, note_detected_window, note_page_read,
                              note_reader_failure, note_window_candidate, note_wps_audience, reuse_discovery_notes)


@dataclass(frozen=True)
class SlideSnapshot:
    active: bool = False
    current: int | None = None
    total: int | None = None
    screen_name: str = ""
    screen_point: tuple[int, int] | None = None
    document: str = ""
    ended: bool = False
    hwnd: int = 0
    show_position: int | None = None

    @property
    def page_text(self):
        if self.active and not self.ended:
            if self.current and self.current > 0:
                if self.total is None:
                    return f"第 {self.current} 页"
                if self.current <= self.total:
                    return f"{self.current}/{self.total}"
            if self.show_position and self.show_position > 0:
                return f"放映第 {self.show_position} 页"
        return ""


def _page_quality(snapshot):
    if snapshot is None or not snapshot.active:
        return 0
    if snapshot.current and snapshot.total and 1 <= snapshot.current <= snapshot.total:
        return 3
    if snapshot.current:
        return 2
    return 1 if getattr(snapshot, "show_position", None) else 0


def _screen_info(hwnd):
    if not hwnd:
        return "", None
    try:
        import win32api
        monitor = win32api.MonitorFromWindow(hwnd, 0)
        if monitor:
            info = win32api.GetMonitorInfo(monitor)
            left, top, _, _ = info["Monitor"]
            # Qt can expose a model name instead of \\.\DISPLAYn. On Windows its
            # logical screen origin matches the native origin even at mixed scales.
            return info["Device"], (left + 1, top + 1)
    except Exception:
        pass
    return "", None


def _pid(hwnd):
    import win32process
    if not hwnd:
        return 0
    try:
        thread_id, process_id = win32process.GetWindowThreadProcessId(hwnd)
        # On failure the thread ID is zero and Windows leaves the PID output undefined.
        return process_id if thread_id else 0
    except Exception:
        return 0


def _office_hwnd(obj):
    try:
        # The Office automation interface exposes HWND as a signed 32-bit Long.
        return int(obj.HWND) & 0xffffffff
    except Exception:
        return 0


@dataclass(frozen=True)
class _WpsApplication:
    application: object
    shared_class: bool = False  # The Microsoft-compatible CLSID can be published by either suite.


def _registered_wps_server(clsid, view):
    """Check the executable registered in the same registry view as the ProgID."""
    import winreg
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, fr"CLSID\{clsid}\LocalServer32",
                            0, winreg.KEY_READ | view) as key:
            try:
                executable = str(winreg.QueryValueEx(key, "ServerExecutable")[0]).strip().strip('"')
            except FileNotFoundError:
                command = str(winreg.QueryValueEx(key, None)[0]).strip()
                first = re.match(r'^"([^"]+)"|^(.+?\.exe)(?=\s|$)', command, re.I)
                if not first:
                    return False
                executable = first.group(1) or first.group(2)
            executable = ntpath.expandvars(executable)
            return ntpath.basename(executable).casefold() in {"wpp.exe", "wps.exe"}
    except OSError:
        return False


def _running_office_apps(*, identify_wps=False):
    import pythoncom
    import winreg
    from win32com.client import dynamic

    note_page_read("active_object", "checking")
    classes = {}
    progids = ("KWPP.Application", "WPP.Application", "PowerPoint.Application") if identify_wps else (
        "PowerPoint.Application", "KWPP.Application", "WPP.Application")
    for progid in progids:
        # A 64-bit timer may be running alongside a 32-bit Office/WPS build.
        for view in (0, winreg.KEY_WOW64_32KEY, winreg.KEY_WOW64_64KEY):
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid + r"\CLSID", 0,
                                    winreg.KEY_READ | view) as key:
                    clsid = str(winreg.QueryValueEx(key, None)[0]).casefold()
                shared = clsid == _POWERPOINT_CLSID.casefold()
                classes[clsid] = classes.get(clsid, False) or (identify_wps and (
                    (progid != "PowerPoint.Application" and not shared) or _registered_wps_server(clsid, view)))
            except FileNotFoundError:
                continue
            except (OSError, pythoncom.com_error) as error:
                note_page_read("active_object", "error", error)
    for clsid, wps_registration in classes.items():
        try:
            obj = pythoncom.GetActiveObject(clsid).QueryInterface(pythoncom.IID_IDispatch)
            app = dynamic.Dispatch(obj)
            note_page_read("active_object", "ok")
            if wps_registration:
                yield _WpsApplication(app, clsid == _POWERPOINT_CLSID.casefold())
            else:
                yield app
        except (OSError, pythoncom.com_error) as error:
            note_page_read("active_object", "error", error)
    note_page_read("active_object", "empty")


@dataclass(frozen=True)
class _DocumentShow:
    window: object
    HWND: int  # Application HWND if the document-level view doesn't expose one.
    document: str = ""  # A running file moniker can identify its own window.
    application: object = None


def _office_collection_items(collection, count):
    """Yield Office collection members, returning whether all requested members were read."""
    limit = min(count, 128)
    missing = False
    for index in range(1, limit + 1):
        try:
            member = collection.Item(index)
        except Exception as error:
            note_page_read("collection_item", "error", error)
            # A late-bound collection can expose its default Item member
            # through DISPID_VALUE even when named access fails.
            try:
                member = collection(index)
            except Exception:
                member = None
        if member is None:
            missing = True
        else:
            note_page_read("collection_item", "ok")
            yield member
    if not missing:
        return True

    # Request the COM enumerator explicitly. Ordinary Python iteration may
    # fall back to __getitem__(0), which isn't an Office collection index.
    try:
        enumerated = 0
        for member in islice(collection.__iter__(), limit):
            if member is not None:
                enumerated += 1
                note_page_read("collection_enum", "ok")
                yield member
        if not enumerated:
            note_page_read("collection_enum", "empty")
        return enumerated == limit
    except Exception as error:
        note_page_read("collection_enum", "error", error)
        return False


def _app_show_windows(app):
    """A protected or closing member must not hide other running shows."""
    if isinstance(app, _DocumentShow):
        yield app.window
        return
    note_page_read("collection", "checking")
    try:
        collection = app.SlideShowWindows
        count = int(collection.Count)
        note_page_read("collection", "ok" if count else "empty", value=count)
        complete = yield from _office_collection_items(collection, count)
        if complete and count:
            return
    except Exception as error:
        note_page_read("collection", "error", error)
    yield from _document_show_windows(app)


def _document_show_windows(app):
    # This is the document's actual running show, never its editing window.
    try:
        window = app.ActivePresentation.SlideShowWindow
        if window is not None:
            note_page_read("document_show", "ok")
            yield window
    except Exception as error:
        note_page_read("document_show", "error", error)
    # Explorer may have opened several decks in one PowerPoint process. The
    # active editor is not necessarily the presentation running on the projector.
    try:
        presentations = app.Presentations
        count = int(presentations.Count)
    except Exception:
        return
    for document in _office_collection_items(presentations, count):
        try:
            window = document.SlideShowWindow
            if window is not None:
                note_page_read("document_show", "ok")
                yield window
        except Exception:
            continue  # A document that is only open for editing has no show.


_DECK_FILE = re.compile(r"\.(?:ppt[xm]?|pps[xm]?|dps|dpt)\b", re.I)


def _window_caption(hwnd):
    if hwnd:
        try:
            import win32gui
            return win32gui.GetWindowText(hwnd)
        except Exception:
            pass
    return ""


def _deck_named_in(caption, name):
    if not caption or not name:
        return False
    if name.lower().startswith(("https://", "http://", "file://")):
        name = unquote(urlsplit(name).path)
    caption = re.sub(r"(?:https?://|file:///)[^\s\]】）]+", lambda m: unquote(m[0]), caption, flags=re.I)
    base = ntpath.basename(name)
    if not base:
        return False
    # Office can show a localized suffix like （只读） immediately after the name.
    # When extensions are hidden, only a complete bracketed document title counts.
    edges = r"[\s\[\]()【】（）《》「」“”\"'，,;；：:\\/—–-]"
    if re.search(r"(?:^|" + edges + ")" + re.escape(base) + r"(?=$|" + edges + ")", caption, re.I):
        return True
    stem = ntpath.splitext(base)[0].casefold()
    return bool(stem and not _DECK_FILE.search(caption) and
                any(part.strip().casefold() == stem for part in re.findall(r"[\[【]([^\]】]+)[\]】]", caption)))


def _window_document_name(window):
    for attr in ("FullName", "Name"):
        try:
            name = getattr(window.Presentation, attr)
            if isinstance(name, str) and name:
                return name
        except Exception:
            continue
    return ""


def _document_window_apps(hwnd, pid):
    """PowerPoint exposes DocumentWindow on paneClassDC via OBJID_NATIVEOM."""
    import ctypes
    from ctypes import wintypes
    from comtypes import GUID
    from comtypes.automation import IDispatch
    import pythoncom
    import win32con
    import win32gui
    from win32com.client import dynamic

    note_page_read("native_object", "checking")
    panes = []
    def inspect_child(child, _):
        if (len(panes) < 32 and child not in panes and _pid(child) == pid
                and win32gui.GetClassName(child).casefold() == "paneclassdc"):
            panes.append(child)
        return True

    def inspect_top(top, _):
        try:
            if _pid(top) == pid and len(panes) < 32:
                inspect_child(top, None)
                win32gui.EnumChildWindows(top, inspect_child, None)
        except Exception:
            pass
        return True

    # Prefer descendants of the identified audience window. The editor/backend
    # may expose NativeOM on a separate top-level frame in the same process.
    inspect_top(hwnd, None)
    try:
        win32gui.EnumWindows(inspect_top, None)
    except Exception:
        pass
    accessible = ctypes.oledll.oleacc.AccessibleObjectFromWindow
    accessible.argtypes = (wintypes.HWND, wintypes.DWORD, ctypes.POINTER(GUID),
                           ctypes.POINTER(ctypes.POINTER(IDispatch)))
    accessible.restype = ctypes.HRESULT
    for child in panes:
        try:
            if _pid(child) != pid:
                continue
            # Screen out a hung editor before asking oleacc to perform the native
            # accessibility handshake and unmarshal the document window.
            win32gui.SendMessageTimeout(child, win32con.WM_NULL, 0, 0, win32con.SMTO_ABORTIFHUNG, 200)
            dispatch = ctypes.POINTER(IDispatch)()
            accessible(child, 0xfffffff0, ctypes.byref(IDispatch._iid_), ctypes.byref(dispatch))
            if not dispatch:
                continue
            # NativeOM can expose valid dispatch members absent from ITypeInfo.
            # ObjectFromAddress queries IDispatch for its own COM reference.
            address = ctypes.cast(dispatch, ctypes.c_void_p).value
            pane = dynamic.Dispatch(pythoncom.ObjectFromAddress(address, pythoncom.IID_IDispatch))
            try:
                show = pane.Presentation.SlideShowWindow
            except Exception:
                show = None
            if show is not None:
                note_page_read("native_object", "ok")
                yield _DocumentShow(show, child)
            app = pane.Application
            if app is not None:
                note_page_read("native_object", "ok")
                yield app
        except Exception as error:
            note_page_read("native_object", "error", error)
    note_page_read("native_object", "empty")


def _running_file_apps(pid, caption):
    import pythoncom
    from win32com.client import dynamic

    note_page_read("rot", "checking")
    rot = pythoncom.GetRunningObjectTable()
    entries = rot.EnumRunning()
    context = pythoncom.CreateBindCtx(0)
    deadline = time.monotonic() + 0.25
    for _ in range(96):
        if time.monotonic() >= deadline:
            break
        batch = entries.Next(1)
        if not batch:
            break
        moniker = batch[0]
        try:
            if moniker.IsSystemMoniker() != pythoncom.MKSYS_FILEMONIKER:
                continue
            name = moniker.GetDisplayName(context, None)
            if not _DECK_FILE.fullmatch(ntpath.splitext(name)[1]):
                continue
            if _DECK_FILE.search(caption) and not _deck_named_in(caption, name):
                continue
            # Unlike BindToMoniker or Dispatch, ROT.GetObject cannot open a
            # file. A stale moniker simply fails. No deck contents are read.
            raw = rot.GetObject(moniker).QueryInterface(pythoncom.IID_IDispatch)
            document = dynamic.Dispatch(raw)
            try:
                app = document.Application
            except Exception:
                app = None
            app_hwnd = _office_hwnd(app)
            app_pid = _pid(app_hwnd)
            if not app_pid or app_pid == pid:
                try:
                    show = document.SlideShowWindow
                except Exception:
                    show = None
                if show is not None:
                    note_page_read("rot", "ok")
                    yield _DocumentShow(show, app_hwnd, name, app)
                if app is not None:
                    note_page_read("rot", "ok")
                    yield app
        except Exception as error:
            note_page_read("rot", "error", error)
    note_page_read("rot", "empty")


_APP_DISCOVERY = threading.local()
_POWERPOINT_CLSID = "{91493441-5A91-11CF-8700-00AA0060263B}"


def _powerpoint_target(hwnd, pid):
    import win32gui
    try:
        return bool(hwnd and pid and win32gui.IsWindowVisible(hwnd)
                    and win32gui.GetClassName(hwnd).casefold() == "screenclass"
                    and _pid(hwnd) == pid and _process_name(hwnd) == "powerpnt.exe")
    except Exception:
        return False


def _same_com_logon(pid):
    import win32api
    import win32con
    import win32security as security

    def identity(proc):
        token = security.OpenProcessToken(proc, win32con.TOKEN_QUERY)
        try:
            return (security.GetTokenInformation(token, security.TokenStatistics)["AuthenticationId"],
                    security.GetTokenInformation(token, security.TokenIntegrityLevel)[0])
        finally:
            token.Close()

    process = win32api.OpenProcess(0x1000, False, pid)  # PROCESS_QUERY_LIMITED_INFORMATION
    try:
        return identity(process) == identity(win32api.GetCurrentProcess())
    finally:
        process.Close()


def _powerpoint_factory_window(hwnd, pid, caption):
    """Use PowerPoint's MultiUse class factory when its ROT objects aren't available."""
    import pythoncom
    from win32com.client import dynamic

    identity = (hwnd, pid)
    if not _powerpoint_target(hwnd, pid):
        return None, ("", None)
    cache = getattr(_APP_DISCOVERY, "factory", None)
    now = time.monotonic()
    if cache and cache[0] == identity and (cache[1] is not None or now < cache[2]):
        reuse_discovery_notes(cache[3])
        app = cache[1]
        if app is None:
            return None, ("", None)
    else:
        app = None
        try:
            same_logon = _same_com_logon(pid)
            note_page_read("com_identity", "ok" if same_logon else "different")
        except Exception as error:
            same_logon = False
            note_page_read("com_identity", "error", error)
        if same_logon and _powerpoint_target(hwnd, pid):
            note_page_read("powerpoint_factory", "checking")
            try:
                raw = pythoncom.CoCreateInstance(_POWERPOINT_CLSID, None, pythoncom.CLSCTX_LOCAL_SERVER,
                                                 pythoncom.IID_IDispatch)
                app = dynamic.Dispatch(raw)
            except Exception as error:
                note_page_read("powerpoint_factory", "error", error)
        else:
            note_page_read("powerpoint_factory", "skipped")
            _APP_DISCOVERY.factory = (identity, None, now + 3,
                                      discovery_notes(("com_identity", "powerpoint_factory")))
            return None, ("", None)
    if app is not None and _powerpoint_target(hwnd, pid):
        window, screen = _match_show((app,), hwnd, pid, caption)
        actual = _office_hwnd(window)
        root = _native_root(actual)
        same_show = (not actual or not _pid(actual) or actual == hwnd or
                     (root and root == _native_root(hwnd)) or
                     (not _powerpoint_target(actual, pid) and _deck_named_in(caption, _window_document_name(window))))
        if window is not None and same_show and _powerpoint_target(hwnd, pid):
            note_page_read("powerpoint_factory", "ok")
            _APP_DISCOVERY.factory = identity, app, 0, discovery_notes(("com_identity",))
            return window, screen if screen[0] else _screen_info(hwnd)
    # Release a stale proxy and back off while Office is busy or closing.
    note_page_read("powerpoint_factory", "empty")
    _APP_DISCOVERY.factory = (identity, None, time.monotonic() + 3,
                              discovery_notes(("com_identity", "powerpoint_factory")))
    return None, ("", None)


def _release_other_show(hwnd, pid):
    # COM proxies belong to the reader's apartment. Release references on the
    # reader thread as soon as the slideshow changes or the audience exits.
    for key in ("cache", "factory"):
        cached = getattr(_APP_DISCOVERY, key, None)
        if cached and cached[0][:2] != (hwnd, pid):
            delattr(_APP_DISCOVERY, key)


def _secondary_apps(hwnd, pid, caption):
    # COM proxies stay on the calling thread. Refresh them if the window
    # switches, or periodically if an application has republished its ROT entry.
    identity = (hwnd, pid, caption)
    cache = getattr(_APP_DISCOVERY, "cache", None)
    if cache and cache[0] == identity and time.monotonic() < cache[1]:
        reuse_discovery_notes(cache[3])
        return cache[2]
    apps = []
    if _process_name(hwnd) == "powerpnt.exe":
        apps.extend(_document_window_apps(hwnd, pid))
    try:
        apps.extend(_running_file_apps(pid, caption))
    except Exception as error:
        note_page_read("rot", "error", error)
    _APP_DISCOVERY.cache = (identity, time.monotonic() + (3.0 if apps else 1.3), apps, discovery_notes())
    return apps


def _windows_show_candidates(hwnd):
    """Try the live views in order, doing later COM work only if requested."""
    pid = _pid(hwnd)
    _release_other_show(hwnd, pid)
    if hwnd and not pid:
        note_page_read("process", "unavailable")
        return
    caption = _window_caption(hwnd)
    owner = _process_name(hwnd) if hwnd else ""
    if owner in {"wpp.exe", "wps.exe"}:
        matched, screen = _match_show(_running_office_apps(identify_wps=True), hwnd, pid, caption, wps_target=True)
        yield matched, screen
        # The reader resumes here only if the matched collection view did not
        # supply a complete page. WPS can publish a separate, usable document view.
        if matched is not None:
            try:
                direct = matched.Presentation.SlideShowWindow
            except Exception:
                direct = None
            if direct is not None:
                # This view belongs to the matched document. An explicitly different
                # live audience window would be a second show of that presentation.
                actual = _office_hwnd(direct)
                if not actual or actual == hwnd or not _pid(actual) or _native_root(actual) == hwnd:
                    note_page_read("document_show", "ok")
                    yield direct, screen
            doc_name = _window_document_name(matched)
            if doc_name:
                yield _match_show(_running_office_apps(identify_wps=True), hwnd, pid, caption,
                                  wps_target=True, document_name=doc_name)
    else:
        yield _match_show(_running_office_apps(), hwnd, pid, caption)
    if owner in {"powerpnt.exe", "wpp.exe", "wps.exe"}:
        secondary = _secondary_apps(hwnd, pid, caption)
        # Try the running document's own view even when the application's
        # collection returned the correct HWND but an unreadable View.Slide.
        for documents in (True, False):
            sources = (s for s in secondary if isinstance(s, _DocumentShow) == documents)
            yield _match_show(sources, hwnd, pid, caption, wps_target=owner in {"wpp.exe", "wps.exe"})
    if owner == "powerpnt.exe":
        yield _powerpoint_factory_window(hwnd, pid, caption)


def windows_show_window(hwnd=0):
    """Get the matching live COM slideshow window in the caller's apartment."""
    try:
        for window, screen in _windows_show_candidates(hwnd):
            if window is not None:
                note_page_read("window", "ok")
                return window, screen
        note_page_read("window", "empty")
        return None, ("", None)
    finally:
        _release_other_show(0, 0)


def _native_root(hwnd):
    if sys.platform == "win32" and hwnd:
        import win32con
        import win32gui
        try:
            return win32gui.GetAncestor(hwnd, win32con.GA_ROOT)
        except (OSError, win32gui.error):
            pass
    return 0


def _sole_powerpoint_audience(hwnd, pid):
    """Resolve a handle-less Office view only when Windows can prove its owner."""
    if sys.platform != "win32" or not _powerpoint_target(hwnd, pid):
        return False
    import win32gui
    import win32ts
    try:
        session = win32ts.ProcessIdToSessionId(pid)
        processes = {p for s, p, name, _ in win32ts.WTSEnumerateProcesses()
                     if s == session and name and name.casefold() == "powerpnt.exe"}
        if processes != {pid}:
            return False
        audiences = []
        def check(top, _):
            if win32gui.IsWindowVisible(top) and win32gui.GetClassName(top).casefold() == "screenclass" and _pid(top) == pid:
                audiences.append(top)
            return True
        win32gui.EnumWindows(check, None)
        return audiences == [hwnd]
    except (OSError, win32gui.error):
        return False


def is_presenter_window(class_name, caption):
    markers = ("presenter view", "presenter display", "演讲者视图", "演示者视图", "演講者檢視",
               "簡報者檢視", "演講者視圖", "演示者模式", "演讲者模式")
    return "presenterview" in class_name or any(marker in caption for marker in markers)


def _wps_fullscreen_surface(hwnd):
    if sys.platform != "win32":
        return False
    import win32api
    import win32con
    import win32gui
    try:
        # WPS's editor can maximize a frameless Qt top-level window. Its opaque
        # monitor-sized client does not make it an additional audience display.
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
        if style & win32con.WS_MAXIMIZE and style & (win32con.WS_CAPTION | win32con.WS_THICKFRAME):
            return False
        if style & (win32con.WS_CAPTION | win32con.WS_THICKFRAME) and re.search(
                r"\.(?:ppt[xm]?|pps[xm]?|dps|dpt)\s+[-–—]\s+(?:wps\s*(?:office|presentation|演示)|金山演示)\s*$",
                _window_caption(hwnd), re.I):
            return False
        monitor = win32api.MonitorFromWindow(hwnd, win32con.MONITOR_DEFAULTTONULL)
        bounds = win32api.GetMonitorInfo(monitor)["Monitor"]
        return all(abs(a-b) <= 4 for a,b in zip(win32gui.GetWindowRect(hwnd), bounds))
    except (OSError, win32gui.error, TypeError):
        return False


def _wps_owned_overlay(hwnd):
    # Transparent, non-activating tool windows owned by another window in the
    # same process carry ink/controls. They cannot receive slideshow input.
    if sys.platform != "win32":
        return False
    import win32con
    import win32gui
    import pywintypes
    try:
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        flags = win32con.WS_EX_TOOLWINDOW | win32con.WS_EX_TRANSPARENT | win32con.WS_EX_NOACTIVATE
        if style & flags != flags:
            return False
        owner = win32gui.GetWindow(hwnd, win32con.GW_OWNER)
        return bool(owner and _pid(owner) and _pid(owner) == _pid(hwnd))
    except (OSError, pywintypes.error, AttributeError):
        return False


def wps_show_rank(hwnd, class_name, title):
    """Both the detector and COM matcher use this audience-window classification."""
    caption = title.casefold().split(" - [", 1)[0]
    caption = " - ".join(part for part in caption.split(" - ") if not _DECK_FILE.search(part))
    if (_wps_owned_overlay(hwnd) or class_name == "kxwpppresenterwindow" or is_presenter_window(class_name, caption) or
            any(marker in caption for marker in ("presenter console", "slideshow controls", "放映控制台", "演示控制台", "放映总控台"))):
        return 0
    titled = any(label in caption for label in ("slide show", "slideshow", "放映"))
    if class_name in {"kslideshow", "wps_slideshow", "wpp_slideshow_class", "screenclass",
                      "kxwppslideview", "kxwppslideshowwidget"}:
        return 4 if titled else 3
    if class_name in {"kxwppmainwindow", "pp11frameclass", "pp12frameclass", "pptframeclass"}:
        return 0
    if not titled and re.search(
            r"\.(?:doc[xm]?|dot[xm]?|xlsx?|xlsm|et|pdf|wps|wpt|odt)\s+[-–—]\s+(?:wps(?:\s+office|\s*文字|\s*表格|\s*pdf)?|金山\S*)\s*$",
            title, re.I):
        return 0
    return 2 if titled else int(_wps_fullscreen_surface(hwnd))


def _is_wps_show_surface(hwnd):
    import win32gui
    return (win32gui.IsWindowVisible(hwnd) and not win32gui.IsIconic(hwnd) and
            wps_show_rank(hwnd, win32gui.GetClassName(hwnd).casefold(), _window_caption(hwnd)) > 0)


def _unique_wps_show(hwnd, pid, *, check_processes, document=""):
    """Resolve a handle-less show to a physical audience on the user's desktop."""
    process_count = audience_count = same_screen = None

    def result(reason):
        note_wps_audience(reason, processes=process_count, audiences=audience_count, on_target_screen=same_screen)
        return reason in {"unique", "document"}

    if sys.platform != "win32" or not pid or _pid(hwnd) != pid:
        return result("target_gone")
    if _process_name(hwnd) not in {"wpp.exe", "wps.exe"}:
        return result("unavailable")
    import win32gui
    import win32ts
    try:
        possible_owners = {pid}
        if check_processes:
            session = win32ts.ProcessIdToSessionId(pid)
            possible_owners = {p for s, p, name, _ in win32ts.WTSEnumerateProcesses()
                               if s == session and name and name.casefold() in {"wps.exe", "wpp.exe"}}
            process_count = len(possible_owners)
            if pid not in possible_owners:
                return result("unavailable")
        shows = set()
        def collect(top, _):
            if _pid(top) in possible_owners and _is_wps_show_surface(top):
                shows.add(top)
            return True
        win32gui.EnumWindows(collect, None)
        audience_count = len(shows)
        target_screen = _screen_info(hwnd)[0]
        if target_screen:
            same_screen = sum(_screen_info(top)[0] == target_screen for top in shows)
        if _pid(hwnd) != pid:
            return result("target_gone")
        if hwnd not in shows:
            return result("target_not_show")
        if len(shows) == 1:
            return result("unique")
        # Both the target and the other audiences must identify their decks.
        # A second unnamed Qt projection can be showing the very same document.
        if document and _deck_named_in(_window_caption(hwnd), document):
            for other in shows - {hwnd}:
                title = _window_caption(other)
                if not _DECK_FILE.search(title) or _deck_named_in(title, document):
                    return result("multiple")
            return result("document")
        return result("multiple")
    except (OSError, win32gui.error):
        return result("unavailable")


def _presentation_vendor(app):
    if app is None:
        return ""
    try:
        executable = ntpath.basename(str(app.FullName)).casefold()
        if executable in {"wpp.exe", "wps.exe"}:
            return "wps"
        if executable in {"powerpnt.exe", "pptview.exe"}:
            return "powerpoint"
        if executable.endswith((".exe", ".bin")):
            return "other"
    except Exception:
        pass
    try:
        name = re.sub(r"[\s_-]+", "", str(app.Name).casefold())
        if name.startswith(("wps", "kingsoft", "金山")):
            return "wps"
        if name.startswith(("libreoffice", "openoffice", "apacheopenoffice", "keynote",
                            "microsoftword", "microsoftexcel", "adobeacrobat")):
            return "other"
    except Exception:
        pass
    try:
        # WPS can expose the Office-compatible Application.Name while its
        # installation directory or the type library still identifies Kingsoft.
        path = ntpath.normpath(str(app.Path)).casefold()
        if {"kingsoft", "wps office", "wpsoffice"}.intersection(path.split("\\")):
            return "wps"
    except Exception:
        pass
    try:
        typelib, _ = app._oleobj_.GetTypeInfo().GetContainingTypeLib()
        library_name, description, _, _ = typelib.GetDocumentation(-1)
        if re.search(r"kingsoft|金山|\bwps\b", f"{library_name} {description}", re.I):
            return "wps"
    except Exception:
        pass
    return ""


def _wps_compatibility_source(hwnd, pid, document):
    """A handle-less, unbranded live show must have an independently observable owner."""
    if sys.platform != "win32" or not pid or _pid(hwnd) != pid:
        return False
    import win32gui
    import win32ts
    try:
        session = win32ts.ProcessIdToSessionId(pid)
        processes = {p: name.casefold() for s, p, name, _ in win32ts.WTSEnumerateProcesses()
                     if s == session and name}
        if processes.get(pid) not in {"wpp.exe", "wps.exe"}:
            return False
        if any(name in {"soffice.bin", "pptview.exe"} for name in processes.values()):
            return False
        powerpoint = {p for p, name in processes.items() if name == "powerpnt.exe"}
        if not powerpoint:
            return True  # The caller also requires exactly one WPS process/audience and COM show.
        if not document:
            return False

        # A second suite can be open for editing. Bind by deck only if Windows
        # confirms it in WPS and every Microsoft process is an unrelated editor.
        named_wps = False
        conflicting = False
        editors = set()
        def inspect(top, _):
            nonlocal named_wps, conflicting
            owner = _pid(top)
            if owner != pid and owner not in powerpoint:
                return True
            class_name = win32gui.GetClassName(top).casefold()
            title = _window_caption(top)
            if owner in powerpoint:
                if class_name == "screenclass" or is_presenter_window(class_name, title.casefold()) or _deck_named_in(title, document):
                    conflicting = True
                if class_name == "pptframeclass" and win32gui.IsWindowVisible(top):
                    editors.add(owner)
            elif win32gui.IsWindowVisible(top) and _deck_named_in(title, document):
                named_wps = True
            return True
        win32gui.EnumWindows(inspect, None)
        return named_wps and not conflicting and editors == powerpoint
    except (OSError, win32gui.error):
        return False


def _ole_window_handle(obj):
    try:
        import pythoncom
        handle = int(obj._oleobj_.QueryInterface(pythoncom.IID_IOleWindow).GetWindow())
        return handle if _pid(handle) else 0
    except Exception:
        return 0


def _wps_document_hwnd(window, app, native_show_hwnd):
    # Some WPS releases omit Application.HWND and SlideShowWindow.HWND but
    # still expose the editor's DocumentWindow.HWND for the same running file.
    def live_handle(obj):
        handle = _office_hwnd(obj)
        if _pid(handle):
            return handle
        return _ole_window_handle(obj)

    if native_show_hwnd:
        return native_show_hwnd
    try:
        windows = window.Presentation.Windows
        for index in range(1, min(int(windows.Count), 8) + 1):
            handle = live_handle(windows.Item(index))
            if handle:
                return handle
    except Exception:
        pass
    try:
        return live_handle(app.ActiveWindow)
    except Exception:
        return 0


def _match_show(apps, hwnd, target_pid, title, *, wps_target=False, document_name=""):
    target_screen, _ = _screen_info(hwnd)
    named_target = bool(_DECK_FILE.search(title))
    bracketed_title = bool(re.search(r"[\[【][^\]】]+[\]】]", title))
    target_root = _native_root(hwnd)
    sole_audience = None
    seen_hwnds = set()
    seen_dispatch = []
    best_window = None
    best_screen = ("", None)
    best_rank = 0
    wps_audiences = {}
    compatibility = {}
    for source in apps:
        registered_source = isinstance(source, _WpsApplication)
        app = source.application if registered_source else source
        suite_app = app.application if isinstance(app, _DocumentShow) else app
        windows = _document_show_windows(app) if document_name else _app_show_windows(app)
        for window in windows:
            try:
                if document_name:
                    actual_name = _window_document_name(window)
                    if (not actual_name or ntpath.normcase(ntpath.normpath(actual_name)) !=
                            ntpath.normcase(ntpath.normpath(document_name))):
                        continue
                show_hwnd = _office_hwnd(window)
                native_show_hwnd = 0
                if wps_target and not _pid(show_hwnd):
                    # IOleWindow on the show itself can expose its exact audience
                    # even when WPS's automation HWND property is zero.
                    native_show_hwnd = _ole_window_handle(window)
                    if native_show_hwnd and (native_show_hwnd == hwnd or
                                             (target_root and _native_root(native_show_hwnd) == target_root)):
                        note_window_candidate(native_show_hwnd, 0, "exact" if native_show_hwnd == hwnd else "related")
                        return window, _screen_info(hwnd)
                    if native_show_hwnd and _pid(native_show_hwnd) != target_pid:
                        note_window_candidate(native_show_hwnd, 0, "wrong_process")
                        continue
                    other_root = _native_root(native_show_hwnd) if native_show_hwnd else 0
                    if other_root and _is_wps_show_surface(other_root):
                        note_window_candidate(native_show_hwnd, 0, "other_show")
                        continue
                if hwnd and show_hwnd == hwnd:
                    note_window_candidate(show_hwnd, 0, "exact")
                    return window, _screen_info(hwnd)
                app_hwnd = _office_hwnd(app)
                doc = _window_document_name(window) or (app.document if isinstance(app, _DocumentShow) else "")
                matches_name = _deck_named_in(title, doc)
                if doc and named_target and not matches_name:
                    note_window_candidate(show_hwnd, app_hwnd, "different_deck")
                    continue
                if hwnd and target_root and target_root == _native_root(show_hwnd):
                    note_window_candidate(show_hwnd, app_hwnd, "related")
                    return window, _screen_info(hwnd)
                if doc and bracketed_title and not matches_name:
                    note_window_candidate(show_hwnd, app_hwnd, "different_deck")
                    continue
                show_pid = _pid(show_hwnd) if target_pid and show_hwnd else 0
                if show_hwnd and not show_pid and target_pid:
                    show_hwnd = 0  # Office can retain a presentation HWND after Windows destroys it.
                screen = _screen_info(show_hwnd)
                evidence = "same_process"
                if target_pid:
                    owner_pid = show_pid or _pid(app_hwnd)
                    if wps_target and not show_pid:
                        if not owner_pid:
                            app_hwnd = _wps_document_hwnd(window, suite_app, native_show_hwnd)
                            owner_pid = _pid(app_hwnd)
                        if owner_pid and owner_pid != target_pid:
                            note_window_candidate(show_hwnd, app_hwnd, "wrong_process")
                            continue
                        compat = False
                        if not owner_pid:
                            vendor = _presentation_vendor(suite_app)
                            if vendor in {"powerpoint", "other"}:
                                note_window_candidate(show_hwnd, app_hwnd, "wrong_suite")
                                continue
                            if vendor != "wps" and not (registered_source and not source.shared_class):
                                if doc not in compatibility:
                                    compatibility[doc] = _wps_compatibility_source(hwnd, target_pid, doc)
                                if not compatibility[doc]:
                                    note_window_candidate(show_hwnd, app_hwnd, "unknown_owner")
                                    continue
                                compat = True
                        check_processes = not owner_pid
                        audience_key = (check_processes, doc)
                        if audience_key not in wps_audiences:
                            wps_audiences[audience_key] = _unique_wps_show(
                                hwnd, target_pid, check_processes=check_processes, document=doc)
                        if not wps_audiences[audience_key]:
                            note_window_candidate(show_hwnd, app_hwnd, "wps_ambiguous")
                            continue
                        note_page_read("wps_source", "ok")
                        evidence = "wps_compat" if compat else ("wps_unique" if check_processes else "wps_document")
                        owner_pid = target_pid
                    if owner_pid and owner_pid != target_pid:
                        note_window_candidate(show_hwnd, app_hwnd, "wrong_process")
                        continue
                    if not owner_pid:
                        if sole_audience is None:
                            sole_audience = _sole_powerpoint_audience(hwnd, target_pid)
                        if isinstance(app, _DocumentShow):
                            ppt_source = bool(app.document and matches_name)
                        else:
                            try:
                                ppt_source = str(app.Name).strip().casefold() in {"powerpoint", "microsoft powerpoint", "microsoft office powerpoint"}
                            except Exception:
                                ppt_source = False
                        if not (sole_audience and ppt_source):
                            note_window_candidate(show_hwnd, app_hwnd, "unknown_owner")
                            continue
                        evidence = "unique_instance"
                    if sys.platform == "win32" and show_hwnd and (
                            _powerpoint_target(show_hwnd, target_pid) or
                            (wps_target and _native_root(show_hwnd) == show_hwnd and _is_wps_show_surface(show_hwnd))):
                        note_window_candidate(show_hwnd, app_hwnd, "other_show")
                        continue
                if named_target and not matches_name:
                    note_window_candidate(show_hwnd, app_hwnd, "different_deck")
                    continue
                if show_hwnd:
                    if show_hwnd in seen_hwnds:
                        continue
                    seen_hwnds.add(show_hwnd)
                same_screen = bool(screen[0] and screen[0] == target_screen)
                # A different monitor alone cannot associate two distinct
                # windows. Presenter View can still identify its deck by title.
                if hwnd and screen[0] and target_screen and not same_screen and not matches_name:
                    note_window_candidate(show_hwnd, app_hwnd, "other_screen")
                    continue
                if not show_hwnd:
                    # NativeOM and the file ROT can expose separate wrappers for
                    # the same COM show while its HWND property is unavailable.
                    # Access wrapper storage directly. Asking a COM object for
                    # a private Python attribute can raise DISP_E_UNKNOWNNAME.
                    try:
                        attrs = object.__getattribute__(window, "__dict__")
                    except AttributeError:
                        attrs = {}
                    dispatch = attrs.get("_oleobj_")
                    if dispatch is None:
                        dispatch = attrs.get("_comobj")
                    if dispatch is not None:
                        if dispatch in seen_dispatch:
                            continue
                        seen_dispatch.append(dispatch)
                rank = 1 + int(same_screen) + 2 * int(matches_name)
                if rank > best_rank:
                    best_rank, best_window, best_screen = rank, window, screen
                    note_window_candidate(show_hwnd, app_hwnd, evidence)
                elif rank == best_rank:
                    # Several shows with equal evidence: don't report another deck's page.
                    best_window = None
                    note_window_candidate(show_hwnd, app_hwnd, "ambiguous")
                    note_page_read("window", "ambiguous")
            except Exception as error:
                note_page_read("window", "error", error)
                continue
    return best_window, best_screen


def read_windows_pages(hwnd=0):
    """Read pages in the calling thread's COM apartment. hwnd is the detected show."""
    owner = _process_name(hwnd) if hwnd else ""
    window_class = ""
    if hwnd and sys.platform == "win32":
        import win32gui
        try:
            window_class = win32gui.GetClassName(hwnd)
        except (OSError, win32gui.error):
            pass
    with capture_page_read(hwnd, owner, window_class) as trace:
        if hwnd:
            note_page_read("process", "ok" if owner else "unavailable")
        result = _read_windows_page_source(hwnd, owner)
        if not hwnd and result.hwnd:
            trace.identify(result.hwnd, _process_name(result.hwnd))
        closed = False
        if hwnd and not result.page_text and not result.ended:
            try:
                import win32gui
                closed = not win32gui.IsWindowVisible(hwnd)
            except Exception:
                pass
        trace.finish(result, target_closed=closed)
        return result


def _read_windows_page_source(hwnd, owner):
    if owner in {"acrobat.exe", "acrord32.exe"}:
        _release_other_show(hwnd, _pid(hwnd))
        return _acrobat_snapshot(hwnd)
    if not hwnd:
        window, screen = windows_show_window()
        show = _office_hwnd(window)
        if not show or not _confirmed_fullscreen(show):
            return SlideSnapshot()
        return _window_snapshot(window, screen, show)
    result = SlideSnapshot()
    for window, screen in _windows_show_candidates(hwnd):
        if window is None:
            continue
        note_page_read("window", "ok")
        # The COM object supplies the page. The audience window selected by
        # Windows owns the overlay even if Office returned a child/editor HWND.
        audience_screen = _screen_info(hwnd)
        candidate = _window_snapshot(window, audience_screen if audience_screen[0] else screen, hwnd)
        if candidate.ended or _page_quality(candidate) == 3:
            return candidate
        if not result.active or _page_quality(candidate) > _page_quality(result):
            result = candidate
    if not result.active:
        note_page_read("window", "empty")
    if owner in {"powerpnt.exe", "pptview.exe", "wpp.exe", "wps.exe"} and not (result.current and result.total or result.ended):
        return _wps_window_pages(hwnd, result)
    return result


def _confirmed_fullscreen(hwnd):
    # A native caption/class detector may miss a newer player. Only a genuine
    # full-monitor window obtained from Office's SlideShowWindow is sufficient
    # to start auto timing on its own, never the normal maximized editor.
    import win32api
    import win32con
    import win32gui
    try:
        if not win32gui.IsWindowVisible(hwnd) or win32gui.IsIconic(hwnd):
            return False
        if _process_name(hwnd) not in {"powerpnt.exe", "pptview.exe", "wpp.exe", "wps.exe"}:
            return False
        if win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE) & win32con.WS_CAPTION:
            return False
        monitor = win32api.MonitorFromWindow(hwnd, win32con.MONITOR_DEFAULTTONULL)
        bounds = win32api.GetMonitorInfo(monitor)["Monitor"]
        return all(abs(a-b) <= 4 for a,b in zip(win32gui.GetWindowRect(hwnd), bounds))
    except Exception:
        return False


_PAGE_LABEL = re.compile(
    r"(?:幻灯片|页码|当前页|\b(?:slide|page)s?\b)\s*[:：]?\s*(?:第\s*)?(\d{1,6})\s*(?:张|页)?\s*"
    r"(?:[/／]|\bof\b|[,，]?\s*共)\s*(?:共\s*)?(\d{1,6})(?!\d)"
    r"|第\s*(\d{1,6})\s*[页张]\s*[,，]?\s*(?:共|[/／])\s*(\d{1,6})\s*[页张]", re.I)
_BASIC_FRACTION = re.compile(r"\s*(\d{1,6})\s*[/／]\s*(\d{1,6})\s*")
_PAGE_WIDGET = re.compile(r"页码|当前页|幻灯片|slide.?num|page.?num|page.?index|page.?status", re.I)
_UIA = threading.local()


def _page_labels(text, allow_fraction=False):
    found = set()
    text = (text or "")[:512]
    for match in _PAGE_LABEL.finditer(text):
        a, b, c, d = match.groups()
        found.add((int(a or c), int(b or d)))
    if allow_fraction:
        bare = _BASIC_FRACTION.fullmatch(text)
        if bare:
            found.add(tuple(map(int, bare.groups())))
    return {(a,b) for a,b in found if 1 <= a <= b}


def _wps_accessible_labels(hwnd, pid):
    """Read at most 50 bounded descendants. Skip document/slide content subtrees."""
    from comtypes.client import CreateObject
    try:
        from comtypes.gen import UIAutomationClient as ui
    except ImportError:
        from comtypes.client import GetModule
        ui = GetModule("UIAutomationCore.dll")
    if not hasattr(_UIA, "client"):
        _UIA.client = CreateObject(ui.CUIAutomation, interface=ui.IUIAutomation)
        try:
            options = _UIA.client.QueryInterface(ui.IUIAutomation2)
            options.ConnectionTimeout = 400
            options.TransactionTimeout = 400
        except Exception:
            pass  # UIAutomation 1.0 has no configurable transaction deadline.
    client = _UIA.client
    root = client.ElementFromHandle(hwnd)
    walker = client.ControlViewWalker
    # Walk only frame, pane, toolbar, status, and group containers. Some Qt
    # controls use Custom for a container; their document surface is never read.
    containers = {ui.UIA_WindowControlTypeId, ui.UIA_PaneControlTypeId,
                  ui.UIA_GroupControlTypeId, ui.UIA_ToolBarControlTypeId, ui.UIA_StatusBarControlTypeId}
    leaves = {ui.UIA_TextControlTypeId, ui.UIA_EditControlTypeId,
              ui.UIA_SpinnerControlTypeId, ui.UIA_CustomControlTypeId}
    pending = [(root, 0, False, False)]
    labels = set()
    inspected = 0
    deadline = time.monotonic() + 0.35
    while pending and inspected < 50 and time.monotonic() < deadline:
        parent, depth, in_chrome, page_context = pending.pop(0)
        child = walker.GetFirstChildElement(parent)
        while child and inspected < 50 and time.monotonic() < deadline:
            inspected += 1
            try:
                kind = child.CurrentControlType
                if child.CurrentProcessId == pid and (kind in containers or kind in leaves):
                    bar = kind in {ui.UIA_StatusBarControlTypeId, ui.UIA_ToolBarControlTypeId}
                    chrome = in_chrome or bar
                    specific = page_context or bool(_PAGE_WIDGET.search(child.CurrentAutomationId or ""))
                    if chrome or specific:
                        name = (child.CurrentName or "")[:512]
                        specific = specific or bool(_PAGE_WIDGET.search(name))
                        labels.update(_page_labels(name, specific))
                        if specific and kind in {ui.UIA_EditControlTypeId, ui.UIA_SpinnerControlTypeId}:
                            value = child.GetCurrentPattern(ui.UIA_ValuePatternId)
                            if value:
                                labels.update(_page_labels(value.QueryInterface(ui.IUIAutomationValuePattern).CurrentValue, True))
                    if depth < 4 and kind in containers:
                        pending.append((child, depth+1, chrome, specific))
            except Exception:
                pass  # Qt may have replaced this individual control after a page turn.
            child = walker.GetNextSiblingElement(child)
    return labels


def _wps_window_pages(hwnd, result):
    import win32api
    import win32con
    import win32gui
    note_page_read("accessibility", "checking")
    try:
        if not win32gui.IsWindowVisible(hwnd):
            note_page_read("accessibility", "skipped")
            return result
        cls = win32gui.GetClassName(hwnd).casefold()
        if cls not in {"kslideshow", "wps_slideshow", "wpp_slideshow_class", "screenclass"}:
            monitor = win32api.MonitorFromWindow(hwnd, win32con.MONITOR_DEFAULTTONULL)
            if not monitor or win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE) & win32con.WS_CAPTION:
                note_page_read("accessibility", "skipped")
                return result
            bounds = win32api.GetMonitorInfo(monitor)["Monitor"]
            if any(abs(a-b) > 4 for a,b in zip(win32gui.GetWindowRect(hwnd), bounds)):
                note_page_read("accessibility", "skipped")
                return result
        # A slide count in the actual filename is document content.
        caption = win32gui.GetWindowText(hwnd)
        caption = re.sub(r"\[[^\]]*\.(?:pptx?|ppsx?|dps)\]", "", caption, flags=re.I)
        caption = re.split(r"\.(?:pptx?|ppsx?|dps)\b", caption, flags=re.I)[-1]
        labels = _page_labels(caption)
        if not labels:
            labels = _wps_accessible_labels(hwnd, _pid(hwnd))
        if len(labels) == 1:
            current, total = labels.pop()
            name, point = _screen_info(hwnd)
            note_page_read("accessibility", "ok")
            return replace(result, active=True, current=current, total=total, hwnd=hwnd,
                           screen_name=name, screen_point=point)
        note_page_read("accessibility", "ambiguous" if labels else "empty")
    except Exception as error:
        note_page_read("accessibility", "error", error)
    return result


def _physical_slide_index(slide, deck):
    # Read this independently of the total. A view can expose a slide while its
    # document collection is busy or restricted.
    try:
        return int(slide.SlideIndex)
    except Exception as exc:
        note_page_read("current", "error", exc)
    try:
        return int(deck.Slides.FindBySlideID(slide.SlideID).SlideIndex)
    except Exception:
        pass
    try:
        return int(slide.SlideNumber) - int(deck.PageSetup.FirstSlideNumber) + 1
    except Exception:
        return None


def _physical_show_position(view, deck, position):
    # CurrentShowPosition is relative to the show. Translate a named show's
    # ordered IDs; only use an ordinal directly in a complete, unhidden deck.
    named = int(view.IsNamedShow)
    settings = deck.SlideShowSettings
    if named == -1:
        ids = settings.NamedSlideShows.Item(view.SlideShowName).SlideIDs
        return int(deck.Slides.FindBySlideID(ids[position - 1]).SlideIndex)
    if named == 0 and int(settings.RangeType) == 1:  # ppShowAll
        # The enum on a SlideRange is msoTriStateMixed if slides disagree.
        if int(deck.Slides.Range().SlideShowTransition.Hidden) == 0:
            return position
    return None


def _window_snapshot(window, screen, hwnd=0):
    info = SlideSnapshot(True, screen_name=screen[0], screen_point=screen[1],
                         document=_window_document_name(window), hwnd=hwnd)
    try:
        view = window.View
    except Exception as exc:
        note_page_read("current", "error", exc)
        return info
    try:
        if int(view.State) == 5:  # ppSlideShowDone, after the final slide.
            return replace(info, ended=True)
    except Exception:
        pass
    try:
        slide = view.Slide
    except Exception as exc:
        note_page_read("current", "error", exc)
        slide = None
    try:
        deck = window.Presentation
    except Exception:
        deck = None
    if deck is None and slide is not None:
        try:
            deck = slide.Parent
            info = replace(info, document=str(deck.FullName or ""))
        except Exception:
            pass
    total = None
    try:
        count = int(deck.Slides.Count)
        if count > 0:
            total = count
            note_page_read("total", "ok", value=count)
        else:
            note_page_read("total", "invalid", value=count)
    except Exception as exc:
        note_page_read("total", "error", exc)
        if slide is not None:
            try:
                alternate = slide.Parent
                count = int(alternate.Slides.Count)
                if count > 0:
                    deck, total = alternate, count
                    note_page_read("total", "ok", value=count)
            except Exception:
                pass

    current = _physical_slide_index(slide, deck) if slide is not None else None
    if current is not None and current > 0 and (total is None or current <= total):
        note_page_read("current", "ok", value=current)
        return replace(info, current=current, total=total)
    if current is not None:
        note_page_read("current", "invalid", value=current)
    try:
        position = int(view.CurrentShowPosition)
    except Exception as exc:
        note_page_read("current", "error", exc)
        position = 0
    if position > 0:
        if total is not None and position > total:
            # A named show can list the same slide more than once. In an
            # ordinary show a position beyond the deck can only be transitional.
            try:
                if int(view.IsNamedShow) != -1:
                    return info
            except Exception:
                return info
        try:
            current = _physical_show_position(view, deck, position)
        except Exception:
            current = None
        if current is not None and current > 0 and (total is None or current <= total):
            note_page_read("current", "ok", value=current)
            return replace(info, current=current, total=total)
        # Keep show ordinals out of physical slide tracking and timing reports.
        return replace(info, show_position=position)
    return info


def _process_name(hwnd):
    import ctypes
    from ctypes import wintypes
    pid = _pid(hwnd)
    if not pid:
        return ""
    kernel = ctypes.windll.kernel32
    kernel.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    kernel.OpenProcess.restype = wintypes.HANDLE
    kernel.QueryFullProcessImageNameW.argtypes = [
        wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD)]
    kernel.CloseHandle.argtypes = [wintypes.HANDLE]
    handle = kernel.OpenProcess(0x1000, False, pid)
    if not handle:
        return ""
    try:
        size = wintypes.DWORD(32768)
        name = ctypes.create_unicode_buffer(size.value)
        if kernel.QueryFullProcessImageNameW(handle, 0, name, ctypes.byref(size)):
            return name.value.rsplit("\\", 1)[-1].lower()
    finally:
        kernel.CloseHandle(handle)
    return ""


def is_adobe_show(hwnd):
    import win32api
    import win32con
    import win32gui
    if not hwnd or not win32gui.IsWindowVisible(hwnd):
        return False
    # Acrobat's regular maximized editor has a frame. Only count an actual
    # frame-free, monitor-sized presentation as a show.
    if win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE) & win32con.WS_CAPTION:
        return False
    monitor = win32api.MonitorFromWindow(hwnd, win32con.MONITOR_DEFAULTTONULL)
    if not monitor:
        return False
    bounds = win32api.GetMonitorInfo(monitor)["Monitor"]
    if any(abs(a-b) > 4 for a, b in zip(win32gui.GetWindowRect(hwnd), bounds)):
        return False
    return _process_name(hwnd) in {"acrobat.exe", "acrord32.exe"}


def _find_adobe_show():
    import win32gui
    foreground = win32gui.GetForegroundWindow()
    if is_adobe_show(foreground):
        return foreground
    matches = []
    def visit(hwnd, _):
        try:
            if is_adobe_show(hwnd):
                matches.append(hwnd)
        except Exception:
            pass  # The window may disappear during enumeration.
        return True
    win32gui.EnumWindows(visit, None)
    return matches[0] if len(matches) == 1 else 0


def _acrobat_snapshot(hwnd):
    import pythoncom
    import win32gui
    from win32com.client import dynamic
    screen, point = _screen_info(hwnd)
    info = SlideSnapshot(True, screen_name=screen, screen_point=point, hwnd=hwnd)
    stage = "active_object"
    try:
        note_page_read(stage, "checking")
        raw = pythoncom.GetActiveObject("AcroExch.App").QueryInterface(pythoncom.IID_IDispatch)
        note_page_read(stage, "ok")
        stage = "window"
        avdoc = dynamic.Dispatch(raw).GetActiveDoc()
        if not avdoc:
            note_page_read(stage, "empty")
            return info
        # Acrobat's automation object follows its active AVDoc. Reject a
        # background PDF window belonging to a different Reader instance.
        if _pid(win32gui.GetForegroundWindow()) != _pid(hwnd):
            note_page_read(stage, "skipped")
            return info
        note_page_read(stage, "ok")
        stage = "total"
        pddoc = avdoc.GetPDDoc()
        count = int(pddoc.GetNumPages())
        note_page_read(stage, "ok" if count > 0 else "invalid", value=count)
        stage = "current"
        number = int(avdoc.GetAVPageView().GetPageNum()) + 1
        note_page_read(stage, "ok" if 1 <= number <= count else "invalid", value=number)
        if 1 <= number <= count:
            info = replace(info, current=number, total=count)
        try:
            return replace(info, document=str(pddoc.GetFileName() or ""))
        except Exception:
            return info
    except Exception as error:
        note_page_read(stage, "error", error)
        return info


class _WindowsPages:
    """At most one COM call in flight. A busy Office process cannot stall detection."""

    def __init__(self):
        self.condition = threading.Condition()
        self.target = None
        self.target_pid = 0
        self.next_idle_probe = 0.0
        self.generation = 0
        self.requested = False
        self.closed = False
        self.result = SlideSnapshot()
        self.updated = 0.0
        threading.Thread(target=self._run, name="Office page reader", daemon=True).start()

    def sample(self, detected):
        # Older boolean detectors can still supply a show with no target HWND.
        target = int(detected) if detected and not isinstance(detected, bool) else 0
        pid = _pid(target) if target else 0
        now = time.monotonic()
        with self.condition:
            if self.target != target or self.target_pid != pid:
                note_detected_window(target)
                self.target = target
                self.target_pid = pid
                self.generation += 1
                self.result = SlideSnapshot()
                self.next_idle_probe = 0.0
            # A running PowerPoint instance may publish a real SlideShowWindow
            # even if Windows supplied no recognized foreground/window class.
            if detected or self.result.active or now >= self.next_idle_probe:
                self.requested = True
                self.next_idle_probe = now + 2.4
                self.condition.notify()
            result = self.result
            fresh = now - self.updated < 1.8
        if fresh and result.active and (not target or result.screen_name):
            return result
        if not detected:
            return SlideSnapshot()
        screen = _screen_info(target)
        if fresh and result.active:
            return replace(result, screen_name=screen[0], screen_point=screen[1], hwnd=target)
        return SlideSnapshot(True, screen_name=screen[0], screen_point=screen[1], hwnd=target)

    def _run(self):
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception as error:
            note_reader_failure(error)
            return
        try:
            while True:
                with self.condition:
                    self.condition.wait_for(lambda: self.closed or self.requested)
                    if self.closed:
                        return
                    self.requested = False
                    target, generation = self.target, self.generation
                try:
                    result = read_windows_pages(target)
                except Exception:
                    result = SlideSnapshot()
                with self.condition:
                    if generation == self.generation:
                        # The live page is read after binding to the show. A
                        # slow first bind must not expire a newly read result.
                        self.result, self.updated = result, time.monotonic()
        finally:
            for key in ("cache", "factory"):
                if hasattr(_APP_DISCOVERY, key):
                    delattr(_APP_DISCOVERY, key)
            pythoncom.CoUninitialize()

    def close(self):
        with self.condition:
            self.closed = True
            self.condition.notify()


_MAC_SCRIPT = '''
if application id "com.microsoft.Powerpoint" is not running then return "idle"
tell application id "com.microsoft.Powerpoint"
    if (count of slide show windows) is 0 then return "idle"
    set showWindow to slide show window 1
    set documentPath to ""
    try
        set documentPath to full name of presentation of showWindow as text
    end try
    set pageInfo to ""
    try
        set currentPage to slide index of slide of slide show view of showWindow
        set totalPages to count of slides of presentation of showWindow
        set pageInfo to (currentPage as text) & "/" & (totalPages as text)
    end try
    set showCenter to ""
    try
        set {x1, y1, x2, y2} to bounds of showWindow
        set showCenter to (((x1 + x2) div 2) as text) & "," & (((y1 + y2) div 2) as text)
    end try
    return "show|" & pageInfo & "|" & showCenter & "|" & documentPath
end tell
'''


def _mac_snapshot(output):
    fields = output.strip().split("|", 3)
    if len(fields) < 3 or fields[0] != "show":
        return SlideSnapshot()
    try:
        x, y = map(int, fields[2].split(","))
        point = (x, y)
    except ValueError:
        point = None
    info = SlideSnapshot(True, screen_point=point, document=fields[3] if len(fields) == 4 else "")
    try:
        current, total = map(int, fields[1].split("/"))
        if 1 <= current <= total:
            return replace(info, current=current, total=total)
    except ValueError:
        pass
    return info


def _read_mac():
    # Check the running executable before AppleScript even compiles the Office dictionary.
    running = subprocess.run(
        ["/usr/bin/pgrep", "-f", r"/Microsoft PowerPoint\.app/Contents/MacOS/Microsoft PowerPoint( |$)"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=1,
    )
    if running.returncode:
        return SlideSnapshot(), 1.5
    result = subprocess.run(
        ["/usr/bin/osascript", "-e", _MAC_SCRIPT], capture_output=True, text=True, timeout=4,
    )
    if result.returncode:
        # Give the presenter time to change Automation permission after a denial.
        return SlideSnapshot(), 30 if "-1743" in result.stderr else 2
    return _mac_snapshot(result.stdout), 0.75


class SlideMonitor(QObject):
    changed = Signal(object)
    _ready = Signal()

    def __init__(self, fallback_detector, parent=None):
        super().__init__(parent)
        self._detect = fallback_detector
        self._lock = threading.Lock()
        self._stop = None
        self._latest = None
        self._ready.connect(self._deliver, Qt.ConnectionType.QueuedConnection)

    def start(self):
        with self._lock:
            if self._stop is not None:
                return
            stop = self._stop = threading.Event()
        threading.Thread(target=self._watch, args=(stop,), name="Slideshow monitor", daemon=True).start()

    def stop(self):
        with self._lock:
            if self._stop:
                self._stop.set()
            self._stop = self._latest = None

    def _watch(self, stop):
        pages = _WindowsPages() if sys.platform == "win32" else None
        try:
            while not stop.is_set():
                delay = 0.60
                try:
                    if pages is not None:
                        snapshot = pages.sample(self._detect() or _find_adobe_show())
                    elif sys.platform == "darwin":
                        snapshot, delay = _read_mac()
                    else:
                        snapshot, delay = SlideSnapshot(), 2
                except Exception:
                    snapshot = SlideSnapshot()
                    delay = 2
                with self._lock:
                    if stop.is_set():
                        return
                    notify = self._latest is None
                    self._latest = snapshot
                    if notify:
                        self._ready.emit()
                stop.wait(delay)
        finally:
            if pages:
                pages.close()

    @Slot()
    def _deliver(self):
        # Coalesce pending samples if the GUI was busy. Always deliver on Qt's thread.
        with self._lock:
            snapshot, self._latest = self._latest, None
        if snapshot is not None:
            self.changed.emit(snapshot)
