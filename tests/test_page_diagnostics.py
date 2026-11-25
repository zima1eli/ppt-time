import threading
import unittest
from types import SimpleNamespace
from unittest.mock import patch

import page_diagnostics as diag


def reading(hwnd, current=None, total=None, *, active=True, ended=False):
    return SimpleNamespace(active=active, ended=ended, hwnd=hwnd, current=current, total=total,
                           document=r"C:\private\private-client.pptx")


class ComFailure(Exception):
    hresult = -2147352567
    excepinfo = (0, "private client", "private slide text", None, 0, -2147024891)


class PageDiagnosticsTests(unittest.TestCase):
    def setUp(self):
        self.history = diag.PageReadHistory()
        patcher = patch.object(diag, "_HISTORY", self.history)
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_missing_properties_and_hresult_survive_exit_without_content(self):
        diag.note_detected_window(64)
        with diag.capture_page_read(64, "powerpnt.exe") as trace:
            diag.note_page_read("collection", "ok", value=1)
            diag.note_page_read("window", "ok")
            diag.note_page_read("total", "ok", value=14)
            diag.note_page_read("current", "error", ComFailure("private slide text"))
            diag.note_page_read("current", "error", AttributeError("private file"))
            trace.finish(reading(64, total=14))
        diag.note_detected_window(0)
        with diag.capture_page_read() as idle:
            idle.finish(reading(0, active=False))
        result = diag.native_page_report()
        for expected in ("最近一次外部放映窗口检测 HWND：未检测到", "0x40", "powerpnt.exe", "总页数属性：已取得 14",
                         "当前页属性：读取失败", "0x80020009", "0x80070005"):
            self.assertIn(expected, result)
        self.assertNotIn("private", result)

    def test_last_valid_page_is_scoped_to_the_show_and_not_the_current_result(self):
        with diag.capture_page_read(72, "wpp.exe") as trace:
            trace.finish(reading(72, 3, 5))
        with diag.capture_page_read(72, "wpp.exe") as trace:
            trace.finish(reading(72, active=False), target_closed=True)
        self.assertIn("外部读取结果：3/5", self.history.report())
        with diag.capture_page_read(72, "wpp.exe") as trace:
            trace.finish(reading(72))
        self.assertIn("该次放映上一次有效页码：3/5", self.history.report())
        self.assertIn("外部读取结果：未取得页码", self.history.report())
        with diag.capture_page_read(72, "wpp.exe") as trace:
            trace.finish(reading(72, ended=True))
        self.assertIn("该次放映上一次有效页码：3/5", self.history.report())
        with diag.capture_page_read() as trace:
            trace.finish(reading(0, active=False))
        with diag.capture_page_read(72, "private-client.exe", "private-document") as trace:
            trace.finish(reading(72))
        self.assertNotIn("3/5", self.history.report())
        self.assertNotIn("private", self.history.report())

    def test_wps_report_identifies_the_reader_after_exit(self):
        with diag.capture_page_read(42, "wpp.exe", "Qt5QWindowIcon") as trace:
            diag.note_page_read("collection_item", "error", ComFailure())
            diag.note_page_read("collection_enum", "ok")
            diag.note_window_candidate(0, 0, "wps_unique")
            trace.finish(reading(42, 4, 15))
        with diag.capture_page_read() as idle:
            idle.finish(reading(0, active=False))
        report = diag.native_page_report("等待 Microsoft PowerPoint 连接", "已连接 WPS 演示，等待放映。")
        for expected in ("使用 WPS", "qt5qwindowicon", "按序号读取集合成员：读取失败", "枚举集合成员：已取得", "已确认唯一的 WPS 放映窗口", "4/15"):
            self.assertIn(expected, report)
        self.assertNotIn("PowerPoint 加载项", report)
        self.assertIn("WPS 演示页码加载项：已连接 WPS 演示", report)
        with diag.capture_page_read(43, "powerpnt.exe") as trace:
            trace.finish(reading(43, 1, 5))
        self.assertIn("PowerPoint 加载项：已连接", diag.native_page_report("已连接"))

    def test_wps_audience_report_shows_per_read_counts_and_checked_reason(self):
        reasons = {
            "unique": "已确定当前放映窗口", "document": "已通过文稿匹配当前放映窗口",
            "multiple": "多个放映窗口仍无法区分", "target_gone": "目标窗口已关闭或已切换",
            "target_not_show": "目标窗口未被识别为放映窗口", "unavailable": "未取得核对所需的窗口信息",
        }
        for reason, label in reasons.items():
            with self.subTest(reason=reason), diag.capture_page_read(42, "wpp.exe") as trace:
                diag.note_wps_audience(reason, processes=3, audiences=2, on_target_screen=0)
                trace.finish(reading(42))
            self.assertIn(f"WPS 放映窗口核对：{label}（WPS 进程 3，放映窗口 2，目标屏上的放映窗口 0）",
                          self.history.report())
        with diag.capture_page_read(42, "wpp.exe") as trace:
            diag.note_wps_audience("unavailable", processes="private", audiences=-1, on_target_screen=False)
            diag.note_wps_audience("private title", processes=3, audiences=2, on_target_screen=1)
            trace.finish(reading(42))
        result = self.history.report()
        self.assertIn("WPS 进程 未取得，放映窗口 未取得，目标屏上的放映窗口 未取得", result)
        self.assertNotIn("private", result)
        with diag.capture_page_read(43, "powerpnt.exe") as trace:
            trace.finish(reading(43))
        self.assertNotIn("WPS 放映窗口核对", self.history.report())

    def test_physical_page_without_total_and_show_order_are_distinguishable(self):
        with diag.capture_page_read(0) as trace:
            trace.identify(88, "powerpnt.exe")
            trace.finish(reading(88, current=4))
        result = self.history.report()
        self.assertIn("0x58（由 Office 放映接口识别）", result)
        self.assertIn("第 4 页（总数未取得）", result)
        with diag.capture_page_read(88, "powerpnt.exe") as trace:
            snapshot = reading(88, total=25)
            snapshot.show_position = 2
            trace.finish(snapshot)
        self.assertIn("放映第 2 页（仅放映顺序）", self.history.report())
        self.assertNotIn("2/25", self.history.report())

    def test_discovery_cache_carries_real_evidence_and_error_precedence(self):
        with diag.capture_page_read(1) as trace:
            diag.note_page_read("native_object", "empty")
            diag.note_page_read("rot", "error", ComFailure())
            diag.note_page_read("rot", "empty")
            evidence = diag.discovery_notes()
            trace.finish(reading(1))
        with diag.capture_page_read(1) as trace:
            diag.reuse_discovery_notes(evidence)
            diag.note_page_read("active_object", "error", ComFailure())
            diag.note_page_read("active_object", "ok")
            diag.note_page_read("active_object", "empty")
            trace.finish(reading(1))
        result = self.history.report()
        self.assertIn("活动 COM 应用：已取得", result)
        self.assertIn("ROT）：读取失败（HRESULT", result)
        self.assertIn("窗口 Office 对象：没有结果", result)

    def test_busy_worker_and_initialization_failure_have_distinct_status(self):
        diag.note_detected_window(75)
        diag.note_reader_failure(ComFailure())
        self.assertIn("后台页码读取线程初始化失败（HRESULT", self.history.report())
        wait = threading.Event()
        ready = threading.Event()
        def read():
            with diag.capture_page_read(75, "powerpnt.exe") as trace:
                diag.note_page_read("rot", "checking")
                trace.started -= 4
                ready.set()
                wait.wait(4)
                trace.finish(reading(75, 1, 4))
        worker = threading.Thread(target=read)
        worker.start()
        try:
            self.assertTrue(ready.wait(2))
            result = self.history.report()
            self.assertIn("查询仍在进行", result)
            self.assertIn("ROT）：正在查询", result)
            self.assertNotIn("线程初始化失败", result)
        finally:
            wait.set()
            worker.join(4)
        self.assertIn("外部读取结果：1/4", self.history.report())

    def test_window_candidates_are_bounded_deduplicated_and_scoped_to_one_read(self):
        diag.note_window_candidate(800, 900, "exact")
        with diag.capture_page_read(200, "powerpnt.exe") as trace:
            for i in range(5):
                diag.note_window_candidate(100 + i, 150 + i, "unknown_owner")
            diag.note_window_candidate(101, 151, "related")
            diag.note_window_candidate(500, 600, "private title")
            trace.finish(reading(200))
        candidates = [line for line in self.history.report().splitlines() if "候选窗口 ·" in line]
        self.assertEqual(len(candidates), 4)
        self.assertNotIn("0x64", " ".join(candidates))
        self.assertIn("COM 放映 HWND 0x65 · 应用或窗格 HWND 0x97 · 属于同一组窗口", candidates[-1])
        self.assertNotIn("private", self.history.report())
        with diag.capture_page_read(200, "powerpnt.exe") as trace:
            diag.note_window_candidate(r"C:\private\talk.pptx", None, "unknown_owner")
            trace.finish(reading(200))
        fresh = self.history.report()
        self.assertIn("COM 放映 HWND 0x0 · 应用或窗格 HWND 0x0 · 未能确定所属进程", fresh)
        self.assertNotIn("private", fresh)
        self.assertNotIn("0x65", fresh)
