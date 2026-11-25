import os
import unittest
from types import SimpleNamespace as Obj
from unittest.mock import patch

import slideshow


def read(view, deck=None):
    return slideshow._window_snapshot(Obj(View=view, Presentation=deck), ("", None))


def deck_with_range(hidden=0, range_type=1):
    group = Obj(SlideShowTransition=Obj(Hidden=hidden))
    return Obj(Slides=Obj(Count=8, Range=lambda: group), SlideShowSettings=Obj(RangeType=range_type))


class PageFallbackTests(unittest.TestCase):
    def test_physical_page_survives_missing_total(self):
        result = read(Obj(Slide=Obj(SlideIndex=4), CurrentShowPosition=1))
        self.assertEqual((result.current, result.total, result.page_text), (4, None, "第 4 页"))

    def test_recover_total_from_slide_parent(self):
        view = Obj(Slide=Obj(SlideIndex=4, Parent=deck_with_range()))
        self.assertEqual(read(view, Obj(Slides=Obj())).page_text, "4/8")

    def test_full_show_maps_verified_ordinal_to_physical_page(self):
        view = Obj(IsNamedShow=0, CurrentShowPosition=5)
        self.assertEqual(read(view, deck_with_range()).page_text, "5/8")

    def test_hidden_or_range_show_distinguishes_show_position(self):
        view = Obj(IsNamedShow=0, CurrentShowPosition=3)
        for deck in (deck_with_range(-2), deck_with_range(-1), deck_with_range(0, 2)):
            result = read(view, deck)
            self.assertEqual((result.page_text, result.current, result.show_position), ("放映第 3 页", None, 3))

    def test_reordered_named_show_uses_ordered_slide_ids(self):
        view = Obj(IsNamedShow=-1, SlideShowName="Short", CurrentShowPosition=2)
        deck = deck_with_range(range_type=3)
        deck.SlideShowSettings.NamedSlideShows = Obj(Item=lambda name: Obj(SlideIDs=(421, 963, 321)))
        deck.Slides.FindBySlideID = lambda ident: Obj(SlideIndex={421: 3, 963: 7, 321: 1}[ident])
        self.assertEqual(read(view, deck).page_text, "7/8")
        del deck.SlideShowSettings.NamedSlideShows
        self.assertEqual(read(view, deck).page_text, "放映第 2 页")
        view.State = 5
        self.assertEqual(read(view, deck).page_text, "")

    def test_unpublished_or_stale_hwnd_can_match_one_live_deck(self):
        app = Obj(HWND=83)
        dispatch = object()
        window = Obj(HWND=0, _oleobj_=dispatch, Presentation=Obj(Name="Talk.pptx"))
        alias = Obj(HWND=0, _oleobj_=dispatch, Presentation=window.Presentation)
        with patch.object(slideshow, "_pid", side_effect=lambda hwnd: 7 if hwnd in {15,83} else 0), \
             patch.object(slideshow, "_screen_info", return_value=("", None)), \
             patch.object(slideshow, "_app_show_windows", return_value=[window, alias]):
            self.assertIs(slideshow._match_show([app], 15, 7, "[Talk.pptx]")[0], window)
            alias._oleobj_ = object()
            self.assertIsNone(slideshow._match_show([app], 15, 7, "[Talk.pptx]")[0])
        window.HWND = 99
        with patch.object(slideshow, "_pid", side_effect=lambda hwnd: 7 if hwnd == 83 else 0), \
             patch.object(slideshow, "_screen_info", return_value=("", None)), \
             patch.object(slideshow, "_app_show_windows", return_value=[window]):
            self.assertIs(slideshow._match_show([app], 15, 7, "[Talk.pptx]")[0], window)


@unittest.skipUnless(os.name == "nt", "Windows COM")
class PowerPointDispatchTests(unittest.TestCase):
    def test_pywin32_reads_show_through_installed_powerpoint_type_info(self):
        import pythoncom
        from win32com.client import dynamic
        from win32com.server.policy import DesignatedWrapPolicy
        from win32com.server.util import wrap

        try:
            library = pythoncom.LoadRegTypeLib("{91493440-5A91-11CF-8700-00AA0060263B}", 2, 0, 0)
        except pythoncom.com_error:
            self.skipTest("PowerPoint type library is not registered")
        typeinfos = {library.GetDocumentation(i)[0]: library.GetTypeInfo(i)
                     for i in range(library.GetTypeInfoCount())}

        class OfficePolicy(DesignatedWrapPolicy):
            def _build_typeinfos_(self):
                return [self._obj_.typeinfo]

        def dispatch(name, **members):
            info = typeinfos[name]
            getters, methods = {}, {}
            for index in range(info.GetTypeAttr().cFuncs):
                entry = info.GetFuncDesc(index)
                field = info.GetNames(entry.memid)[0]
                if field in members:
                    if entry.invkind == pythoncom.INVOKE_PROPERTYGET:
                        getters[entry.memid] = field
                    if entry.invkind == pythoncom.INVOKE_FUNC:
                        methods[entry.memid] = field
            server = Obj(typeinfo=info, _public_methods_=[], _dispid_to_func_=methods,
                         _dispid_to_get_=getters, **members)
            return wrap(server, usePolicy=OfficePolicy)

        # The server runs no Office application. It exposes Office's registered
        # ITypeInfo, member DISPIDs, property/method flags and return types to
        # pywin32, including the typed Item(Index) call and nested properties.
        page = dispatch("_Slide", SlideIndex=4)
        view = dispatch("SlideShowView", State=1, Slide=page)
        deck = dispatch("_Presentation", Name="TypeInfo.pptx", Slides=dispatch("Slides", Count=9))
        window = dispatch("SlideShowWindow", HWND=90, View=view, Presentation=deck)
        collection = dispatch("SlideShowWindows", Count=1, Item=lambda index: window if index == 1 else None)
        app = dynamic.Dispatch(dispatch("_Application", SlideShowWindows=collection))
        with patch.object(slideshow, "_screen_info", return_value=("", None)):
            matched, screen = slideshow._match_show([app], 90, 0, "")
        self.assertIsNotNone(matched)
        self.assertEqual(slideshow._window_snapshot(matched, screen).page_text, "4/9")

        # Exercise the same property-get call with a missing Slide and a valid
        # CurrentShowPosition. Neither an editor view nor key input is involved.
        restricted = dynamic.Dispatch(dispatch("SlideShowView", State=1, CurrentShowPosition=3))
        self.assertEqual(read(restricted, deck_with_range()).page_text, "放映第 3 页")
