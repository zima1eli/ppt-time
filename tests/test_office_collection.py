import os
import unittest
from types import SimpleNamespace
from unittest.mock import patch

import slideshow


class OfficeCollectionTests(unittest.TestCase):
    def test_wps_partial_collection_retries_only_the_matched_document(self):
        hwnd = 40
        name = r"C:\meeting\talk.pptx"
        for mode in ("direct", "active", "other_document", "complete"):
            with self.subTest(mode=mode):
                selected = SimpleNamespace(FullName=name, Slides=SimpleNamespace(Count=3))
                good_deck = SimpleNamespace(FullName=r"D:\other\talk.pptx" if mode == "other_document" else name,
                                            Slides=SimpleNamespace(Count=3))
                good_view = SimpleNamespace(State=1, Slide=SimpleNamespace(SlideIndex=2))
                good = SimpleNamespace(HWND=hwnd, Presentation=good_deck, View=good_view)
                partial = SimpleNamespace(HWND=hwnd, Presentation=selected,
                                          View=good_view if mode == "complete" else SimpleNamespace())
                if mode == "direct":
                    selected.SlideShowWindow = good
                good_deck.SlideShowWindow = good
                app = SimpleNamespace(SlideShowWindows=SimpleNamespace(Count=1, Item=lambda _: partial),
                                      ActivePresentation=good_deck)
                with (patch.object(slideshow, "_pid", return_value=10),
                      patch.object(slideshow, "_screen_info", return_value=("Display", (1, 1))),
                      patch.object(slideshow, "_native_root", return_value=hwnd),
                      patch.object(slideshow, "_process_name", return_value="wpp.exe"),
                      patch.object(slideshow, "_window_caption", return_value=""),
                      patch.object(slideshow, "_running_office_apps", return_value=[slideshow._WpsApplication(app)]) as apps,
                      patch.object(slideshow, "_secondary_apps", return_value=[]),
                      patch.object(slideshow, "_wps_window_pages", side_effect=lambda _, result: result)):
                    page = slideshow._read_windows_page_source(hwnd, "wpp.exe").page_text
                self.assertEqual(page, "" if mode == "other_document" else "2/3")
                self.assertEqual(apps.call_count, 1 if mode in ("direct", "complete") else 2)

    def test_null_item_falls_back_to_running_document(self):
        show = SimpleNamespace(View=SimpleNamespace(State=1))
        windows = SimpleNamespace(Count=1, Item=lambda index: None)
        app = SimpleNamespace(SlideShowWindows=windows,
                              ActivePresentation=SimpleNamespace(SlideShowWindow=show))
        self.assertEqual(list(slideshow._app_show_windows(app)), [show])

    def test_uses_default_member_without_guessing_zero_index(self):
        show = SimpleNamespace(View=SimpleNamespace(State=1))
        indexes = []

        class Windows:
            Count = 1

            def Item(self, index):
                raise OSError("named property cannot be read")

            def __call__(self, index):
                indexes.append(index)
                return show

        self.assertEqual(list(slideshow._app_show_windows(SimpleNamespace(SlideShowWindows=Windows()))), [show])
        self.assertEqual(indexes, [1])

    @unittest.skipUnless(os.name == "nt", "Windows Automation")
    def test_real_dispatch_enumerator_after_wps_shaped_item_errors(self):
        import pythoncom
        import winerror
        from win32com.client import dynamic
        from win32com.server.exception import COMException
        from win32com.server.util import NewEnum, wrap

        def server(**members):
            return wrap(SimpleNamespace(_public_methods_=[], _public_attrs_=list(members), **members))

        page = SimpleNamespace(_public_methods_=[], _public_attrs_=["SlideIndex"], SlideIndex=1)
        deck = server(Name="fixture.dps", Slides=server(Count=3))
        window = server(HWND=0, Presentation=deck, View=server(State=1, Slide=wrap(page)))

        class Collection:
            _public_methods_ = ["Item"]
            _public_attrs_ = ["Count"]
            Count = 1

            def Item(self, *args):
                raise COMException(scode=winerror.E_FAIL)

            _value_ = Item

            def _NewEnum(self):
                return NewEnum([window])

        app = dynamic.Dispatch(server(SlideShowWindows=wrap(Collection())))
        with self.assertRaises(pythoncom.com_error) as error:
            app.SlideShowWindows.Item(1)
        self.assertEqual(error.exception.hresult, winerror.DISP_E_EXCEPTION)
        self.assertEqual(error.exception.excepinfo[5], winerror.E_FAIL)
        for expected in (1, 3, 2):
            page.SlideIndex = expected
            (current_show,) = slideshow._app_show_windows(app)
            self.assertEqual(slideshow._window_snapshot(current_show, ("", None)).page_text, f"{expected}/3")
