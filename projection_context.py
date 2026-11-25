"""Route right-clicks on a Windows click-through projection strip."""

import ctypes
from ctypes import wintypes
import threading

import pywintypes
import win32api
import win32con
import win32event
import win32gui
from PySide6.QtCore import QPointF


_callback_type = ctypes.WINFUNCTYPE(ctypes.c_ssize_t, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM)
_user = ctypes.WinDLL("user32", use_last_error=True)
_user.SetWindowsHookExW.argtypes = [ctypes.c_int, _callback_type, wintypes.HINSTANCE, wintypes.DWORD]
_user.SetWindowsHookExW.restype = wintypes.HANDLE
_user.UnhookWindowsHookEx.argtypes = [wintypes.HANDLE]
_user.UnhookWindowsHookEx.restype = wintypes.BOOL
_user.CallNextHookEx.argtypes = [wintypes.HANDLE, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM]
_user.CallNextHookEx.restype = ctypes.c_ssize_t


def _hit_strip(targets, x, y):
    for hwnd, outline in targets:
        if not win32gui.IsWindowVisible(hwnd):
            continue
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        if not (left <= x < right and top <= y < bottom):
            continue
        relative_x, relative_y = x - left, y - top
        if not outline.contains(QPointF(relative_x / (right - left), relative_y / (bottom - top))):
            continue
        # WindowFromPoint skips the transparent strip. A menu or dialog above
        # the strip must still receive its right-clicks normally.
        receiver = win32gui.WindowFromPoint((x, y))
        if not receiver:
            continue
        receiver = win32gui.GetAncestor(receiver, win32con.GA_ROOT)
        seen = set()
        while receiver and receiver != hwnd and receiver not in seen:
            seen.add(receiver)
            receiver = win32gui.GetWindow(receiver, win32con.GW_HWNDPREV)
        if receiver == hwnd:
            return hwnd, relative_x, relative_y
    return None


class _StripClick:
    def __init__(self, show_menu):
        self.show_menu = show_menu
        self.pressed = None

    def handle(self, message, targets, point):
        try:
            hit = _hit_strip(targets, point.x, point.y)
        except pywintypes.error:
            hit = None  # A screen or underlying window just disappeared.
        if message == win32con.WM_RBUTTONDOWN:
            modifiers = (win32con.VK_SHIFT, win32con.VK_CONTROL, win32con.VK_MENU,
                         win32con.VK_LWIN, win32con.VK_RWIN)
            if hit and any(win32api.GetAsyncKeyState(key) & 0x8000 for key in modifiers):
                hit = None
            self.pressed = hit
            return bool(hit)
        previous, self.pressed = self.pressed, None
        if previous and hit and hit[0] == previous[0]:
            self.show_menu(*hit)
        return bool(previous)


class ProjectionContext:
    """Own one short-lived mouse hook while one or more strips are visible."""

    def __init__(self, show_menu):
        self.show_menu = show_menu
        self.targets = ()
        self._stop = None

    def watch(self, targets):
        self.targets = tuple(targets)
        if self.targets and self._stop is None:
            self._stop = win32event.CreateEvent(None, True, False, None)
            threading.Thread(target=self._listen, args=(self._stop,), daemon=True,
                             name="Projection context menu").start()
        elif not self.targets and self._stop is not None:
            win32event.SetEvent(self._stop)
            self._stop = None

    def _listen(self, stop):
        click = _StripClick(self.show_menu)

        @_callback_type
        def handle(code, message, data):
            if code == 0 and message in (win32con.WM_RBUTTONDOWN, win32con.WM_RBUTTONUP):
                # MSLLHOOKSTRUCT starts with a per-monitor physical POINT.
                pt = ctypes.cast(data, ctypes.POINTER(wintypes.POINT)).contents
                if click.handle(message, self.targets if self._stop is stop else (), pt):
                    return 1
            return _user.CallNextHookEx(None, code, message, data)

        module = win32api.GetModuleHandle(None)
        while True:
            hook = _user.SetWindowsHookExW(win32con.WH_MOUSE_LL, handle, module, 0)
            if hook:
                break
            if win32event.WaitForSingleObject(stop, 2500) == win32event.WAIT_OBJECT_0:
                return
        try:
            # Keep global mouse delivery independent of modal dialogs or a busy UI.
            while win32event.MsgWaitForMultipleObjects((stop,), False, win32event.INFINITE,
                                                        win32event.QS_ALLINPUT) == win32event.WAIT_OBJECT_0 + 1:
                if win32gui.PumpWaitingMessages():
                    break
        finally:
            _user.UnhookWindowsHookEx(hook)
