import os
from pathlib import Path
import sys
import threading
import unittest
from unittest.mock import patch
from types import SimpleNamespace as Obj
import uuid

import slideshow as slides


class Collection:
    def __init__(self, *items):
        self.items = items
        self.Count = len(items)

    def Item(self, index):
        item = self.items[index - 1]
        if isinstance(item, Exception):
            raise item
        return item


def show(handle, name, index=3):
    return Obj(HWND=handle,
               Presentation=Obj(FullName=name, Slides=Obj(Count=12)),
               View=Obj(State=1, Slide=Obj(SlideIndex=index)))


class CaptureTests(unittest.TestCase):
    @unittest.skipUnless(os.name == "nt", "Windows window ownership")
    def test_pid_of_destroyed_native_window_is_zero(self):
        import win32gui
        hwnd = win32gui.CreateWindow('STATIC', '', 0, 0, 0, 8, 8, 0, 0, 0, None)
        try:
            self.assertEqual(slides._pid(hwnd), os.getpid())
        finally:
            win32gui.DestroyWindow(hwnd)
        self.assertEqual(slides._pid(hwnd), 0)

    def test_broken_collection_member_does_not_hide_next_show(self):
        audience = show(40, "talk.pptx")
        app = Obj(SlideShowWindows=Collection(OSError("protected window"), audience))
        with patch.object(slides, "_screen_info", return_value=("projector", (1, 1))):
            self.assertIs(slides._match_show([app], 40, 8, "")[0], audience)

    def test_document_slideshow_window_without_collection(self):
        audience = show(40, "talk.pptx")
        app = Obj(ActivePresentation=Obj(SlideShowWindow=audience))
        self.assertEqual(list(slides._app_show_windows(app)), [audience])
        self.assertEqual(list(slides._app_show_windows(Obj(ActivePresentation=Obj()))), [])

    def test_another_open_document_runs_while_active_document_is_being_edited(self):
        audience = show(40, "speaker.pptx")
        editing = Obj(Name="schedule.pptx")
        app = Obj(ActivePresentation=editing, SlideShowWindows=Collection(),
                  Presentations=Collection(editing, Obj(SlideShowWindow=audience)))
        self.assertEqual(list(slides._app_show_windows(app)), [audience])

    def test_running_document_improves_partial_global_view_without_factory(self):
        partial = show(40, "speaker.pptx")
        partial.Presentation.Slides = Obj()
        complete = show(40, "speaker.pptx", index=6)
        direct = slides._DocumentShow(complete, 0)
        app = Obj(SlideShowWindows=Collection(partial))
        with patch.object(slides, "_running_office_apps", return_value=[app]), \
             patch.object(slides, "_secondary_apps", return_value=[app, direct]), \
             patch.object(slides, "_process_name", return_value="powerpnt.exe"), \
             patch.object(slides, "_window_caption", return_value="speaker.pptx"), \
             patch.object(slides, "_screen_info", return_value=("Projector", (1, 1))), \
             patch.object(slides, "_pid", return_value=8), \
             patch.object(slides, "_powerpoint_factory_window") as factory:
            result = slides.read_windows_pages(40)
            self.assertEqual((result.hwnd, result.page_text), (40, "6/12"))
            factory.assert_not_called()

    @unittest.skipUnless(os.name == "nt", "Windows PowerPoint class factory")
    def test_factory_logon_mismatch_repeats_in_the_cached_diagnostic(self):
        from page_diagnostics import PageReadHistory
        history = PageReadHistory()
        with patch.object(slides, "_APP_DISCOVERY", threading.local()), \
             patch.object(slides, "_powerpoint_target", return_value=True), \
             patch.object(slides, "_same_com_logon", return_value=False) as check, \
             patch("pythoncom.CoCreateInstance") as activation:
            for _ in range(2):
                with history.capture(40, 'powerpnt.exe') as record:
                    self.assertIsNone(slides._powerpoint_factory_window(40, 8, '')[0])
                    record.finish(slides.SlideSnapshot(True, hwnd=40))
                self.assertEqual(record.stages['com_identity'][0], 'different')
                self.assertEqual(record.stages['powerpoint_factory'][0], 'skipped')
            check.assert_called_once()
            activation.assert_not_called()

    @unittest.skipUnless(os.name == "nt", "Windows PowerPoint class factory")
    def test_factory_matches_live_audience_and_releases_on_exit(self):
        audience = show(40, "speaker.pptx", index=7)
        app = Obj(SlideShowWindows=Collection(audience))
        with patch.object(slides, "_APP_DISCOVERY", threading.local()), \
             patch.object(slides, "_powerpoint_target", return_value=True) as target, \
             patch.object(slides, "_same_com_logon", return_value=True), \
             patch.object(slides, "_screen_info", return_value=("Projector", (1, 1))), \
             patch.object(slides, "_pid", return_value=8), \
             patch("pythoncom.CoCreateInstance") as activate, \
             patch("win32com.client.dynamic.Dispatch", return_value=app):
            for _ in range(2):
                self.assertIs(slides._powerpoint_factory_window(40, 8, "")[0], audience)
            activate.assert_called_once()
            slides._release_other_show(0, 0)
            self.assertFalse(hasattr(slides._APP_DISCOVERY, "factory"))
            # A different live audience on the same monitor and in the same
            # process cannot supply the target's pages.
            audience.HWND = 41
            self.assertIsNone(slides._powerpoint_factory_window(40, 8, "")[0])
            self.assertIsNone(slides._APP_DISCOVERY.factory[1])
            slides._release_other_show(0, 0)
            target.return_value = False
            with patch.object(slides, "_running_office_apps", return_value=[]):
                self.assertIsNone(slides.windows_show_window(0)[0])
            self.assertIsNone(slides._powerpoint_factory_window(40, 8, "")[0])
            self.assertEqual(activate.call_count, 2)

    def test_wrong_deck_is_rejected_even_on_same_monitor(self):
        wrong = show(41, r"C:\Meetings\SmallTalk.pptx")
        right = show(42, r"C:\Meetings\Talk.pptx")
        with patch.object(slides, "_screen_info", return_value=("projector", (0, 1))), \
             patch.object(slides, "_pid", return_value=8):
            title = "演示者视图 - [Talk.pptx]"
            self.assertIsNone(slides._match_show([Obj(SlideShowWindows=Collection(wrong))], 43, 8, title)[0])
            both = Obj(SlideShowWindows=Collection(wrong, right))
            self.assertIs(slides._match_show([both], 43, 8, title)[0], right)
            self.assertIsNone(slides._match_show([both], 43, 8, "")[0])

    def test_invalid_window_does_not_claim_unrelated_open_deck(self):
        with patch.object(slides, "_pid", return_value=0), patch.object(slides, "_running_office_apps") as apps:
            self.assertIsNone(slides.windows_show_window(88)[0])
            apps.assert_not_called()

    def test_missing_index_uses_document_numbering_not_custom_show_ordinal(self):
        deck = Obj(Name="Talk.pptx", Slides=Obj(Count=14), PageSetup=Obj(FirstSlideNumber=31))
        view = Obj(State=1, Slide=Obj(SlideNumber=38), CurrentShowPosition=2)
        snapshot = slides._window_snapshot(Obj(Presentation=deck, View=view), ("", None))
        self.assertEqual((snapshot.document, snapshot.page_text), ("Talk.pptx", "8/14"))
        del deck.PageSetup
        ordinal = slides._window_snapshot(Obj(Presentation=deck, View=view), ("", None))
        self.assertEqual(ordinal.page_text, "放映第 2 页")
        self.assertIsNone(ordinal.current)
        deck.Slides.FindBySlideID = lambda identity: Obj(SlideIndex={99:11}[identity])
        view.Slide = Obj(Parent=deck, SlideID=99)
        self.assertEqual(slides._window_snapshot(Obj(View=view), ("", None)).page_text, "11/14")

    def test_com_only_discovery_requires_a_confirmed_fullscreen_window(self):
        audience = show(70, "talk.pptx")
        with patch.object(slides, "windows_show_window", return_value=(audience, ("projector", (1, 1)))):
            with patch.object(slides, "_confirmed_fullscreen", return_value=False):
                self.assertFalse(slides.read_windows_pages().active)
            with patch.object(slides, "_confirmed_fullscreen", return_value=True):
                actual = slides.read_windows_pages()
                self.assertEqual((actual.hwnd, actual.page_text), (70, "3/12"))

    def test_worker_discards_reused_handle_and_uses_idle_probe(self):
        with patch.object(threading.Thread, "start"), patch.object(slides, "_screen_info", return_value=("", None)), \
             patch.object(slides.time, "monotonic", return_value=10), patch.object(slides, "_pid", return_value=1) as pid:
            worker = slides._WindowsPages()
            self.assertFalse(worker.sample(False).active)
            self.assertTrue(worker.requested)
            worker.requested = False
            self.assertFalse(worker.sample(False).active)
            self.assertFalse(worker.requested)
            worker.result = slides.SlideSnapshot(True, 5, 6, "Monitor", hwnd=90)
            worker.updated = 10
            self.assertEqual(worker.sample(False).page_text, "5/6")
            worker.sample(90)
            worker.result = slides.SlideSnapshot(True, 2, 8, "Monitor", hwnd=90)
            pid.return_value = 2
            self.assertEqual(worker.sample(90).page_text, "")

    def test_slow_first_bind_is_fresh_on_arrival_then_expires(self):
        clock = Obj(now=100.0)
        with patch.object(threading.Thread, "start"), patch.object(slides, "_pid", return_value=7), \
             patch.object(slides, "_screen_info", return_value=("Screen", (1, 1))), \
             patch.object(slides.time, "monotonic", side_effect=lambda:clock.now), \
             patch.dict(sys.modules, {"pythoncom":Obj(CoInitialize=lambda:None, CoUninitialize=lambda:None)}):
            worker = slides._WindowsPages()
            worker.sample(42)
            def slow_read(_):
                clock.now += 3
                return slides.SlideSnapshot(True, 4, 12, "Screen", hwnd=42)
            def ready(_):
                if not worker.requested:
                    worker.closed = True
            with patch.object(slides, "read_windows_pages", side_effect=slow_read), \
                 patch.object(worker.condition, "wait_for", side_effect=ready):
                worker._run()
            self.assertEqual(worker.sample(42).page_text, "4/12")
            clock.now += 2
            self.assertEqual(worker.sample(42).page_text, "")

    @unittest.skipUnless(os.name == "nt", "Windows running object table")
    def test_actual_running_file_moniker_without_opening_file(self):
        import pythoncom
        from win32com.server.util import wrap

        class App:
            _public_methods_ = []
            _public_attrs_ = ["Name", "HWND"]
            Name = "already-running PowerPoint fixture"
            HWND = 0

        class Doc:
            _public_methods_ = []
            _public_attrs_ = ["Application"]
            Application = wrap(App())

        path = Path.cwd() / ("fixture-" + uuid.uuid4().hex + ".pptx")
        rot = pythoncom.GetRunningObjectTable()
        file_ref = pythoncom.CreateFileMoniker(str(path))
        cookie = rot.Register(pythoncom.ROTFLAGS_REGISTRATIONKEEPSALIVE, wrap(Doc()), file_ref)
        self.addCleanup(rot.Revoke, cookie)
        apps = slides._running_file_apps(os.getpid(), f"PowerPoint Slide Show - [{path.name}]")
        self.assertEqual([app.Name for app in apps], [App.Name])
        self.assertFalse(path.exists())

    @unittest.skipUnless(os.name == "nt", "Windows NativeOM")
    def test_real_native_pane_exposes_its_document_without_an_application(self):
        import ctypes
        from ctypes import wintypes
        import comtypes
        from comtypes.automation import IDispatch
        import pythoncom
        import win32api
        import win32con
        import win32gui
        from win32com.server.util import wrap

        class Fields:
            _public_methods_ = []

            def __init__(self, **values):
                self._public_attrs_ = list(values)
                self.__dict__.update(values)

        instance = win32api.GetModuleHandle(None)
        create = win32gui.CreateWindowEx
        audience = create(0, 'STATIC', '', win32con.WS_POPUP, 0, 0, 10, 10, 0, 0, instance, None)
        self.addCleanup(win32gui.DestroyWindow, audience)
        deck = wrap(Fields(FullName='test.pptx', Slides=wrap(Fields(Count=11))))
        view = wrap(Fields(State=1, Slide=wrap(Fields(SlideIndex=7))))
        show = wrap(Fields(HWND=audience, Presentation=deck, View=view))
        pane = wrap(Fields(Presentation=wrap(Fields(SlideShowWindow=show))))
        # Transfer a test-only interface from pywin32 to ctypes, then revoke the
        # private bootstrap CLSID before exercising the actual HWND handshake.
        iid = str(pythoncom.CreateGuid())
        cookie = pythoncom.RegisterActiveObject(pane, iid, pythoncom.ACTIVEOBJECT_STRONG)
        try:
            pointer = comtypes.GetActiveObject(comtypes.GUID(iid), interface=IDispatch)
        finally:
            pythoncom.RevokeActiveObject(cookie)
        publish = ctypes.windll.oleacc.LresultFromObject
        publish.argtypes = (ctypes.POINTER(comtypes.GUID), wintypes.WPARAM, ctypes.POINTER(IDispatch))
        publish.restype = ctypes.c_ssize_t
        observed = []

        def procedure(h, message, wparam, lparam):
            if message == win32con.WM_GETOBJECT and lparam & 0xffffffff == 0xfffffff0:
                observed.append((wparam, lparam))
                return publish(ctypes.byref(IDispatch._iid_), wparam, pointer)
            return win32gui.DefWindowProc(h, message, wparam, lparam)

        wc = win32gui.WNDCLASS()
        wc.lpszClassName, wc.hInstance, wc.lpfnWndProc = 'paneClassDC', instance, procedure
        atom = win32gui.RegisterClass(wc)
        self.addCleanup(win32gui.UnregisterClass, atom, instance)
        child = create(0, atom, '', win32con.WS_CHILD, 0, 0, 5, 5, audience, 0, instance, None)
        self.addCleanup(win32gui.DestroyWindow, child)
        self.assertEqual(slides._pid(child), os.getpid())
        self.assertTrue(slides._process_name(child))
        sources = slides._document_window_apps(audience, os.getpid())
        actual, screen = slides._match_show(sources, audience, os.getpid(), '')
        self.assertEqual(slides._window_snapshot(actual, screen, audience).page_text, '7/11')
        self.assertTrue(observed)

    @unittest.skipUnless(os.name == "nt", "Windows running object table")
    def test_actual_file_moniker_exposes_only_its_own_show(self):
        import pythoncom
        from win32com.server.util import wrap

        class Fields:
            _public_methods_ = []

            def __init__(self, **values):
                self._public_attrs_ = list(values)
                self.__dict__.update(values)

        path = Path.cwd() / ("rot-only-" + uuid.uuid4().hex + ".pptx")
        deck = Fields(FullName=str(path), Slides=wrap(Fields(Count=3)))
        view = Fields(State=1, Slide=wrap(Fields(SlideIndex=2)))
        window = wrap(Fields(HWND=40, Presentation=wrap(deck), View=wrap(view)))
        # There is no application-level object available from this file moniker.
        document = wrap(Fields(SlideShowWindow=window))
        rot = pythoncom.GetRunningObjectTable()
        file_ref = pythoncom.CreateFileMoniker(str(path))
        cookie = rot.Register(pythoncom.ROTFLAGS_REGISTRATIONKEEPSALIVE, document, file_ref)
        self.addCleanup(rot.Revoke, cookie)
        with patch.object(slides, "_screen_info", return_value=("Projector", (1, 1))):
            sources = slides._running_file_apps(os.getpid(), path.name)
            actual, screen = slides._match_show(sources, 40, os.getpid(), path.name)
            self.assertEqual(slides._window_snapshot(actual, screen, 40).page_text, "2/3")
        self.assertFalse(path.exists())
