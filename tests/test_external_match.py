import sys
import unittest
from types import SimpleNamespace
from unittest.mock import patch

import slideshow as slides


def presentation_app(name, windows):
    collection = SimpleNamespace(Count=len(windows), Item=lambda index: windows[index - 1])
    return SimpleNamespace(Name=name, SlideShowWindows=collection)


class ExternalShowMatchTests(unittest.TestCase):
    def test_com_child_matches_detected_root(self):
        window = SimpleNamespace(HWND=101, Presentation=SimpleNamespace(Name="Talk.pptx"))
        app = presentation_app("Microsoft PowerPoint", [window])
        ancestors = {100: 100, 101: 100}
        with patch.object(slides, "_native_root", side_effect=ancestors.get), \
                patch.object(slides, "_screen_info", return_value=("audience screen", (1, 1))), \
                patch.object(slides, "_sole_powerpoint_audience") as sole:
            found, screen = slides._match_show([app], 100, 30, "PowerPoint Slide Show - [Talk.pptx]")
        self.assertIs(found, window)
        self.assertEqual(screen, ("audience screen", (1, 1)))
        sole.assert_not_called()

    def test_handleless_view_requires_matching_deck_and_a_single_powerpoint(self):
        title = "幻灯片放映 - 【会议 计划.PPTX（只读）】"
        name = "https://host.example/docs/会议%20计划.pptx"
        window = SimpleNamespace(Presentation=SimpleNamespace(FullName=name))
        for source, sole, doc, expected in (
                ("Microsoft PowerPoint", True, name, True),
                ("Microsoft PowerPoint", False, name, False),
                ("Microsoft PowerPoint", True, "其他.pptx", False),
                ("WPS 演示", True, name, False)):
            with self.subTest(source=source, sole=sole, doc=doc):
                window.Presentation.FullName = doc
                app = presentation_app(source, [window])
                with patch.object(slides, "_native_root", return_value=0), \
                        patch.object(slides, "_pid", return_value=0), \
                        patch.object(slides, "_screen_info", return_value=("", None)), \
                        patch.object(slides, "_sole_powerpoint_audience", return_value=sole):
                    found, _ = slides._match_show([app], 100, 30, title)
                self.assertIs(found, window if expected else None)

    def test_process_and_screen_enumeration_identifies_only_one_possible_audience(self):
        session = 1
        pid = 30
        target = 100
        for peers, handles, expected in (
                ([(session + 1, 31, "POWERPNT.EXE", None)], [target], True),
                ([(session, 31, "POWERPNT.EXE", None)], [target], False),
                ([], [target, 102], False)):
            with self.subTest(peers=peers, handles=handles):
                ts = SimpleNamespace(ProcessIdToSessionId=lambda _: session,
                                     WTSEnumerateProcesses=lambda: [(session, pid, "POWERPNT.EXE", None)] + peers)
                def enumerate_windows(callback, context):
                    for handle in handles:
                        callback(handle, context)
                gui = SimpleNamespace(EnumWindows=enumerate_windows, IsWindowVisible=lambda _: True,
                                      GetClassName=lambda _: "screenClass", error=OSError)
                with patch.object(sys, "platform", "win32"), \
                        patch.dict(sys.modules, {"win32ts": ts, "win32gui": gui}), \
                        patch.object(slides, "_pid", return_value=pid), \
                        patch.object(slides, "_powerpoint_target", return_value=True):
                    self.assertEqual(slides._sole_powerpoint_audience(target, pid), expected)

    def test_titles_match_complete_local_or_url_names_only(self):
        for caption, document, expected in (
                ("PowerPoint Slide Show - [tALK.PpTx]", r"C:\speakers\Talk.pptx", True),
                ("PowerPoint Slide Show - [SmallTalk.pptx]", r"C:\speakers\Talk.pptx", False),
                ("幻灯片放映 - 【开幕报告（只读）】", "开幕报告补充.pptx", False),
                ("幻灯片放映 - 【开幕报告】", "开幕报告.pptx", True),
                ("幻灯片放映 - 【开幕报告.PPTX（只读）】", "开幕报告.pptx", True),
                ("PowerPoint Slide Show - [Report One.pptx]", "https://ex.test/team/Report%20One.pptx", True),
                ("PowerPoint Slide Show - [https://ex.test/team/Report%20One.pptx]", "Report One.pptx", True)):
            with self.subTest(caption=caption, document=document):
                self.assertEqual(slides._deck_named_in(caption, document), expected)
