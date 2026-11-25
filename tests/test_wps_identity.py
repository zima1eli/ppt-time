from contextlib import ExitStack, nullcontext
from types import SimpleNamespace as Obj
import sys
import unittest
from unittest.mock import patch

import slideshow as slides


class WpsCompatibilityTests(unittest.TestCase):
    def setUp(self):
        self.inventory = [(1, 90, "wpp.exe", None)]
        self.windows = {40: (90, "Qt5QWindowIcon", "")}
        self.doc = Obj(Name="Quarterly.pptx", Slides=Obj(Count=8))
        self.window = Obj(HWND=0, Presentation=self.doc, View=Obj(Slide=Obj(SlideIndex=3)))
        self.app = Obj(Name="Microsoft PowerPoint",
                       SlideShowWindows=Obj(Count=1, Item=lambda n: self.window))
        self.screens = {40: "Projector A"}
        self.audiences = {40}
        stack = self.enterContext(ExitStack())
        stack.enter_context(patch.object(sys, "platform", "win32"))
        gui = Obj(EnumWindows=lambda f, data: [f(h, data) for h in self.windows],
                  IsWindowVisible=lambda h: True, GetClassName=lambda h: self.windows[h][1], error=OSError)
        ts = Obj(WTSEnumerateProcesses=lambda: self.inventory, ProcessIdToSessionId=lambda p: 1)
        stack.enter_context(patch.dict(sys.modules, {"win32gui": gui, "win32ts": ts}))
        stack.enter_context(patch.object(slides, "_pid", side_effect=lambda h: self.windows.get(h, (0,))[0]))
        stack.enter_context(patch.object(slides, "_screen_info", side_effect=lambda h:(self.screens.get(h, ""), (1, 1))))
        stack.enter_context(patch.object(slides, "_native_root", return_value=0))
        stack.enter_context(patch.object(slides, "_process_name", return_value="wpp.exe"))
        stack.enter_context(patch.object(slides, "_window_caption", side_effect=lambda h: self.windows[h][2]))
        stack.enter_context(patch.object(slides, "_is_wps_show_surface", side_effect=lambda h: h in self.audiences))

    def match(self, obj=None):
        candidate, _ = slides._match_show([obj or self.app], 40, 90, self.windows[40][2], wps_target=True)
        return candidate

    def test_unbranded_wps_and_powerpoint_compatible_name(self):
        self.inventory.append((2, 700, "POWERPNT.EXE", None))
        self.assertIs(self.match(), self.window)
        del self.app.Name
        self.assertIs(self.match(), self.window)
        rot = slides._DocumentShow(self.window, 0, self.doc.Name, self.app)
        self.assertIs(self.match(rot), self.window)
        self.inventory.append((1, 300, "wpp.exe", None))
        self.assertIs(self.match(rot), self.window)

    def test_multiple_wps_processes_require_physical_audiences(self):
        self.inventory.extend([(1, 300, "wpp.exe", None), (1, 301, "wpp.exe", None)])
        self.windows[49] = (300, "Qt5QWindowIcon", "Budget.pptx - WPS Office")
        with patch.object(slides, "note_wps_audience") as note:
            self.assertIs(self.match(), self.window)
            note.assert_called_with("unique", processes=3, audiences=1, on_target_screen=1)
            self.windows[45] = (301, "Qt5QWindowIcon", "")
            self.audiences.add(45)
            self.screens[45] = "Projector B"
            self.assertIsNone(self.match())
            note.assert_called_with("multiple", processes=3, audiences=2, on_target_screen=1)
            self.windows[40] = (90, "screenClass", "放映 - [Quarterly.pptx]")
            self.assertIsNone(self.match())  # The other projection remains anonymous.
            self.windows[45] = (301, "screenClass", "放映 - [Quarterly.pptx]")
            self.assertIsNone(self.match())
            self.windows[45] = (301, "screenClass", "放映 - [Budget.pptx]")
            self.assertIs(self.match(), self.window)
            note.assert_called_with("document", processes=3, audiences=2, on_target_screen=1)
            self.audiences.remove(40)
            self.assertIsNone(self.match())
            note.assert_called_with("target_not_show", processes=3, audiences=1, on_target_screen=0)

    def test_coexisting_editor_needs_a_wps_window_showing_the_deck(self):
        self.inventory.append((1, 200, "POWERPNT.EXE", None))
        self.windows[52] = (200, "PPTFrameClass", "Other.pptx - PowerPoint")
        self.assertIsNone(self.match())
        self.windows[41] = (90, "Qt5QWindowIcon", "Quarterly.pptx - WPS Office")
        self.assertIs(self.match(), self.window)
        self.windows[52] = (200, "PPTFrameClass", "Quarterly.pptx - PowerPoint")
        self.assertIsNone(self.match())
        self.windows[52] = (200, "PPTFrameClass", "Other.pptx - PowerPoint")
        self.windows[53] = (200, "screenClass", "")
        self.assertIsNone(self.match())
        self.doc.Windows = Obj(Count=1, Item=lambda n: Obj(HWND=41))
        self.assertIs(self.match(), self.window)

    def test_positive_vendor_and_process_evidence_override_compatible_name(self):
        self.inventory.append((1, 200, "POWERPNT.EXE", None))
        self.app.Path = r"C:\Program Files\Kingsoft\WPS Office\office6"
        self.assertIs(self.match(), self.window)
        self.app.FullName = r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
        self.assertIsNone(self.match(slides._WpsApplication(self.app)))
        del self.app.FullName, self.app.Path
        self.assertIsNone(self.match(slides._WpsApplication(self.app, shared_class=True)))
        self.assertIs(self.match(slides._WpsApplication(self.app)), self.window)
        typelib = Obj(GetDocumentation=lambda i: ("WPP", "Kingsoft Presentation 12.0 Object Library", 0, ""))
        self.app._oleobj_ = Obj(GetTypeInfo=lambda: Obj(GetContainingTypeLib=lambda: (typelib, 0)))
        self.assertIs(self.match(), self.window)

    def test_native_ole_window_without_an_automation_hwnd(self):
        self.inventory.append((1, 200, "POWERPNT.EXE", None))
        self.windows[41] = (90, "Qt5QWindowIcon", "")
        self.window._oleobj_ = Obj(QueryInterface=lambda iid: Obj(GetWindow=lambda: 41))
        self.assertIs(self.match(), self.window)
        self.windows[41] = (200, "PPTFrameClass", "")
        self.assertIsNone(self.match())

    def test_native_show_handle_identifies_one_of_two_audiences_in_same_process(self):
        self.windows[41] = (90, "Qt5QWindowIcon", "")
        self.windows[42] = (90, "Qt5QWindowIcon", "")
        self.audiences.add(41)
        self.app.HWND = 42  # The editor is common to both shows.
        with patch.object(slides, "_native_root", side_effect=lambda h: h):
            for native, expected in ((40, self.window), (41, None)):
                self.window._oleobj_ = Obj(QueryInterface=lambda iid: Obj(GetWindow=lambda: native))
                self.assertIs(self.match(), expected)

    def test_foreign_presentation_server_does_not_use_anonymous_fallback(self):
        self.app.Name = "LibreOffice Impress"
        self.assertIsNone(self.match())
        del self.app.Name
        self.app.FullName = r"C:\Program Files\Other\slideshow.exe"
        self.assertIsNone(self.match())

    def test_registered_server_executable_in_the_progid_registry_view(self):
        registry = {}
        def read(key, value):
            if value not in registry:
                raise FileNotFoundError
            return registry[value], 1
        reg = Obj(HKEY_CLASSES_ROOT=1, KEY_READ=2, OpenKey=lambda *args: nullcontext(args),
                  QueryValueEx=read)
        with patch.dict(sys.modules, {"winreg": reg}):
            for command in [r'"C:\Program Files\WPS Office\wpp.exe" /Automation',
                            r'C:\Program Files\WPS Office\wpp.exe -Embedding']:
                registry[None] = command
                self.assertTrue(slides._registered_wps_server("{A}", 512))
            registry[None] = r'"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE" -Embedding'
            self.assertFalse(slides._registered_wps_server("{A}", 512))
            registry[None] = r'launcher.exe "C:\WPS Office\wpp.exe"'
            self.assertFalse(slides._registered_wps_server("{A}", 512))
            registry["ServerExecutable"] = r"C:\Program Files\WPS Office\wps.exe"
            self.assertTrue(slides._registered_wps_server("{A}", 512))

    def test_read_all_registered_views_before_querying_the_live_class(self):
        import pythoncom
        from win32com.client import dynamic

        def opened(root, name, reserved, access):
            if name != r"PowerPoint.Application\CLSID":
                raise FileNotFoundError
            return nullcontext(None)
        reg = Obj(HKEY_CLASSES_ROOT=1, KEY_READ=2, KEY_WOW64_32KEY=512, KEY_WOW64_64KEY=256,
                  OpenKey=opened, QueryValueEx=lambda key, name: ("{A}", 1))
        active = Obj(QueryInterface=lambda iid: self.app)
        with (patch.dict(sys.modules, {"winreg": reg}), patch.object(dynamic, "Dispatch", side_effect=lambda raw: raw),
              patch.object(pythoncom, "GetActiveObject", return_value=active) as query,
              patch.object(slides, "_registered_wps_server", side_effect=lambda clsid, view: view == 512)):
            self.assertEqual(list(slides._running_office_apps(identify_wps=True)), [slides._WpsApplication(self.app)])
            query.assert_called_once_with("{a}")
