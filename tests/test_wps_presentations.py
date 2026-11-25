from contextlib import ExitStack
from types import SimpleNamespace as Obj
import unittest
from unittest.mock import patch

import slideshow as slides


def collection(*items):
    return Obj(Count=len(items), Item=lambda index: items[index - 1])


def presentation(name="Quarterly.pptx", page=2):
    deck = Obj(Name=name, Slides=collection(*range(6)))
    return Obj(HWND=0, Presentation=deck, View=Obj(State=1, Slide=Obj(SlideIndex=page)))


class WpsMatchTests(unittest.TestCase):
    def match(self, app, title="", sole=True, pids=None):
        with ExitStack() as stack:
            stack.enter_context(patch.object(slides, "_screen_info", return_value=("Projector", (1, 1))))
            stack.enter_context(patch.object(slides, "_native_root", return_value=0))
            stack.enter_context(patch.object(slides, "_pid", side_effect=lambda h: (pids or {}).get(h, 0)))
            confirm = stack.enter_context(patch.object(slides, "_unique_wps_show", return_value=sole))
            window, _ = slides._match_show([app], 44, 90, title, wps_target=True)
            return window, confirm

    def test_handleless_wps_from_registry_or_running_file(self):
        window = presentation()
        raw = Obj(HWND=0, Name="WPS 演示", SlideShowWindows=collection(window))
        for target in (raw, slides._DocumentShow(window, 0, "Quarterly.pptx", raw),
                       slides._WpsApplication(Obj(SlideShowWindows=collection(window)))):
            with self.subTest(source=type(target).__name__):
                found, confirm = self.match(target)
                self.assertIs(found, window)
                self.assertEqual(slides._window_snapshot(found, ("Projector", (1, 1)), 44).page_text, "2/6")
                confirm.assert_called_once_with(44, 90, check_processes=True, document="Quarterly.pptx")
        powerpoint = Obj(Name="Microsoft PowerPoint", SlideShowWindows=collection(window))
        self.assertIsNone(self.match(powerpoint)[0])
        self.assertIsNone(self.match(slides._DocumentShow(window, 0, "Quarterly.pptx", powerpoint),
                                     "WPS演示 - [Quarterly.pptx]")[0])

    def test_document_anchor_and_rejections(self):
        window = presentation()
        app = Obj(Name="WPS Office", SlideShowWindows=collection(window))
        self.assertIsNone(self.match(app, "WPS Presentation - [Another.pptx]")[0])
        self.assertIsNone(self.match(app, sole=False)[0])
        self.assertIsNone(self.match(Obj(Name="WPS Office", SlideShowWindows=collection(window, presentation("Other.dps"))))[0])
        window.Presentation.Windows = collection(Obj(HWND=55))
        found, confirm = self.match(app, pids={55: 90})
        self.assertIs(found, window)
        confirm.assert_called_once_with(44, 90, check_processes=False, document="Quarterly.pptx")
        self.assertIsNone(self.match(app, pids={55: 20})[0])

    def test_unique_process_and_audience_checks_are_independent(self):
        import sys
        def collect(callback, data):
            for w in [40, 41, 42]:
                callback(w, data)
        gui = Obj(EnumWindows=collect, error=OSError)
        inventory = [(1, 90, "WPP.EXE", None), (1, 701, "wps.exe", None)]
        ts = Obj(WTSEnumerateProcesses=lambda: inventory, ProcessIdToSessionId=lambda pid: 1)
        visible = {40}
        with (patch.dict(sys.modules, {"win32ts": ts, "win32gui": gui}), patch.object(sys, "platform", "win32"),
              patch.object(slides, "_process_name", return_value="wpp.exe"),
              patch.object(slides, "_pid", side_effect=lambda h: 90 if h in {40,41} else 701),
              patch.object(slides, "_is_wps_show_surface", side_effect=lambda h:h in visible)):
            self.assertTrue(slides._unique_wps_show(40, 90, check_processes=True))
            visible.add(42)
            self.assertFalse(slides._unique_wps_show(40, 90, check_processes=True))
            self.assertTrue(slides._unique_wps_show(40, 90, check_processes=False))
            visible.remove(42)
            inventory.append((1, 200, "wpp.exe", None))
            self.assertTrue(slides._unique_wps_show(40, 90, check_processes=True))
            self.assertTrue(slides._unique_wps_show(40, 90, check_processes=False))
            visible.add(41)
            self.assertFalse(slides._unique_wps_show(40, 90, check_processes=False))

    def test_generic_qt_fullscreen_and_presenter(self):
        import sys
        gui = Obj(IsWindowVisible=lambda w: True, IsIconic=lambda w: False,
                  GetClassName=lambda w: "Qt515QWindowIcon")
        with (patch.dict(sys.modules, {"win32gui": gui}), patch.object(slides, "_wps_fullscreen_surface", return_value=True) as full,
              patch.object(slides, "_window_caption", return_value="") as caption):
            self.assertTrue(slides._is_wps_show_surface(40))
            caption.return_value = "WPS 演讲者视图"
            self.assertFalse(slides._is_wps_show_surface(40))
            full.return_value = False
            caption.return_value = "放映 - [Talk.pptx]"
            self.assertTrue(slides._is_wps_show_surface(40))
            caption.return_value = "放映.pptx - WPS Office"
            self.assertFalse(slides._is_wps_show_surface(40))
            full.return_value = True
            caption.return_value = "Meeting_notes.docx - WPS Office"
            self.assertFalse(slides._is_wps_show_surface(40))
            caption.return_value = "放映控制台"
            self.assertFalse(slides._is_wps_show_surface(40))

    def test_borderless_maximized_projector_and_framed_editor(self):
        import sys
        area = (0, 0, 1920, 1080)
        con = Obj(GWL_STYLE=-16, WS_MAXIMIZE=0x1000000, WS_CAPTION=0xc00000,
                  WS_THICKFRAME=0x40000, MONITOR_DEFAULTTONULL=0)
        style = con.WS_MAXIMIZE | 0x80000000  # WS_POPUP
        gui = Obj(GetWindowLong=lambda w, i: style, GetWindowRect=lambda w: area, error=OSError)
        api = Obj(MonitorFromWindow=lambda w, flags: 1, GetMonitorInfo=lambda m: {"Monitor": area})
        with patch.object(sys, "platform", "win32"), patch.dict(sys.modules, {"win32gui": gui, "win32api": api, "win32con": con}):
            self.assertTrue(slides._wps_fullscreen_surface(40))
            style |= con.WS_THICKFRAME
            self.assertFalse(slides._wps_fullscreen_surface(40))
            style = con.WS_MAXIMIZE | con.WS_CAPTION
            self.assertFalse(slides._wps_fullscreen_surface(40))

    def test_owned_transparent_controls_and_vendor_window_classes(self):
        import sys
        style = 0x080000A0
        gui = Obj(GetWindowLong=lambda h, i: style, GetWindow=lambda h, kind: 40)
        with (patch.object(sys, "platform", "win32"), patch.dict(sys.modules, win32gui=gui),
              patch.object(slides, "_pid", return_value=90),
              patch.object(slides, "_wps_fullscreen_surface", return_value=True)):
            self.assertEqual(slides.wps_show_rank(41, "qt5qwindowicon", ""), 0)
            style = 0
            self.assertEqual(slides.wps_show_rank(40, "kxwpppresenterwindow", ""), 0)
            self.assertEqual(slides.wps_show_rank(40, "kxwppmainwindow", ""), 0)
            self.assertGreater(slides.wps_show_rank(40, "kxwppslideview", ""), 0)
            self.assertGreater(slides.wps_show_rank(40, "qt5qwindowicon", "Talk.pptx - WPS Office"), 0)

