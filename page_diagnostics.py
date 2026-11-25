"""In-memory, read-only observations from the Windows slide-page reader."""
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import datetime
import re
import threading
import time


_STAGES = {
    "process": "进程识别", "active_object": "活动 COM 应用", "native_object": "窗口 Office 对象",
    "rot": "运行文稿注册表（ROT）", "com_identity": "放映进程与计时器的登录及权限",
    "powerpoint_factory": "正在放映的 PowerPoint 应用",
    "wps_source": "WPS 演示接口", "collection": "放映窗口集合",
    "collection_item": "按序号读取集合成员", "collection_enum": "枚举集合成员",
    "document_show": "文稿放映窗口", "window": "匹配放映窗口",
    "current": "当前页属性", "total": "总页数属性", "accessibility": "窗口标题或无障碍页码",
    "reader": "读取线程",
}
_OUTCOMES = {
    "checking": "正在查询", "ok": "已取得", "empty": "没有结果", "unavailable": "不可用",
    "ambiguous": "多个结果无法区分", "invalid": "数值无效", "skipped": "未使用", "error": "读取失败",
    "different": "不一致，请用同一个 Windows 账户和相同的权限级别运行两个程序",
}
_PROCESSES = {"powerpnt.exe", "pptview.exe", "wpp.exe", "wps.exe", "wpspdf.exe", "acrobat.exe", "acrord32.exe"}
_CLASSES = {"screenclass", "paneclassdc", "kslideshow", "wps_slideshow", "wpp_slideshow_class",
            "mso_subwindow_class", "acrobatwclass"}
_WINDOW_DECISIONS = {
    "wps_compat": "已将兼容接口与当前 WPS 放映窗口匹配",
    "exact": "与放映窗口句柄一致", "related": "属于同一组窗口",
    "same_process": "与放映窗口属于同一进程", "unique_instance": "已确认唯一的放映实例",
    "wrong_process": "属于其他进程", "unknown_owner": "未能确定所属进程",
    "different_deck": "文稿不匹配", "other_screen": "位于其他屏幕",
    "other_show": "属于其他放映窗口", "ambiguous": "多个候选窗口无法区分",
    "wps_unique": "已确认唯一的 WPS 放映窗口", "wps_document": "已匹配 WPS 文稿所属窗口",
    "wrong_suite": "来自其他演示软件", "wps_ambiguous": "未能确定当前 WPS 放映，详见放映窗口核对",
}
_WPS_AUDIENCE_REASONS = {
    "unique": "已确定当前放映窗口",
    "document": "已通过文稿匹配当前放映窗口",
    "multiple": "多个放映窗口仍无法区分",
    "target_gone": "目标窗口已关闭或已切换",
    "target_not_show": "目标窗口未被识别为放映窗口",
    "unavailable": "未取得核对所需的窗口信息",
}
_LOCAL = threading.local()


def _error_code(error):
    if error is None:
        return ""
    hr = getattr(error, "hresult", None)
    if type(hr) is int:
        codes = [f"HRESULT 0x{hr & 0xffffffff:08X}"]
        info = getattr(error, "excepinfo", None)
        if isinstance(info, tuple) and len(info) >= 6 and type(info[5]) is int and info[5] not in (0, hr):
            codes.append(f"0x{info[5] & 0xffffffff:08X}")
        return " / ".join(codes)
    winerror = getattr(error, "winerror", None)
    return f"Win32 {winerror}" if type(winerror) is int else ""


@dataclass
class _Reading:
    hwnd: int
    process: str
    window_class: str
    lock: object
    started: float = field(default_factory=time.monotonic)
    timestamp: str = field(default_factory=lambda: datetime.now().astimezone().strftime("%Y-%m-%d %H:%M:%S %z"))
    stages: dict = field(default_factory=dict)
    window_candidates: dict = field(default_factory=dict)
    wps_audience: tuple | None = None
    active: bool = False
    ended: bool = False
    page: str = ""
    detected_by_com: bool = False

    def identify(self, hwnd, process):
        with self.lock:
            self.detected_by_com = not self.hwnd
            self.hwnd = hwnd
            self.process = process if process in _PROCESSES else ("其他进程" if process else "未取得")
        note_page_read("process", "ok" if process else "unavailable")

    def finish(self, snapshot, *, target_closed=False):
        current, total = snapshot.current, snapshot.total
        page = ""
        if snapshot.active and not snapshot.ended:
            order = getattr(snapshot, "show_position", None)
            if type(current) is int and type(total) is int and 1 <= current <= total:
                page = f"{current}/{total}"
            elif type(current) is int and current > 0 and total is None:
                page = f"第 {current} 页（总数未取得）"
            elif type(order) is int and order > 0:
                page = f"放映第 {order} 页（仅放映顺序）"
        with self.lock:
            self.active = bool(snapshot.active)
            self.ended = bool(snapshot.ended or target_closed)
            if not self.hwnd and type(snapshot.hwnd) is int:
                self.hwnd = snapshot.hwnd
            self.page = page


class PageReadHistory:
    def __init__(self):
        self.lock = threading.Lock()
        self.pending = self.latest = self.idle = self.last_page = None
        self.detected = 0
        self.worker_failure = None

    @contextmanager
    def capture(self, hwnd=0, process="", window_class=""):
        owner = (process or "").lower()
        category = (window_class or "").lower()
        known_class = category in _CLASSES or re.fullmatch(r"qt\d{1,4}qwindow(?:icon|owndcicon|popupsavebits)?", category)
        record = _Reading(hwnd if type(hwnd) is int else 0,
                          owner if owner in _PROCESSES else ("其他进程" if owner else "未取得"),
                          category if known_class else ("其他窗口类" if category else ""), self.lock)
        previous = getattr(_LOCAL, "current", None)
        _LOCAL.current = self, record
        with self.lock:
            self.pending = record
            self.worker_failure = None
        try:
            yield record
        except Exception as error:
            note_page_read("reader", "error", error)
            raise
        finally:
            _LOCAL.current = previous
            with self.lock:
                if self.pending is record:
                    self.pending = None
                if record.hwnd or record.active:
                    if self.latest is None or record.started >= self.latest.started:
                        if self.latest is None or record.hwnd != self.latest.hwnd or self.idle is not None:
                            self.last_page = None
                        self.idle = None
                        # An Office end screen has no current slide.
                        if not record.ended or self.latest is None or record.hwnd != self.latest.hwnd:
                            self.latest = record
                        if record.page:
                            self.last_page = record
                elif self.idle is None or record.started >= self.idle.started:
                    self.idle = record

    def report(self, powerpoint_status=None, wps_status=None):
        with self.lock:
            pending, saved = self.pending, self.latest
            if pending is not None and pending.hwnd and time.monotonic() - pending.started >= 1:
                reading, running = pending, True
            else:
                reading, running = saved or self.idle or pending, saved is None and self.idle is None and pending is not None
            detector = f"0x{self.detected:X}" if self.detected else "未检测到"
            lines = [f"最近一次外部放映窗口检测 HWND：{detector}"]
            if reading and reading.process in {"wpp.exe", "wps.exe"}:
                lines.append("最近一次外部放映使用 WPS 演示。")
                if wps_status is not None:
                    lines.append("WPS 演示页码加载项：" + wps_status)
            elif powerpoint_status is not None and (not reading or reading.process in {"powerpnt.exe", "未取得"}):
                lines.append("PowerPoint 加载项：" + powerpoint_status)
            if self.worker_failure is not None:
                lines.append("后台页码读取线程初始化失败" + (f"（{self.worker_failure}）" if self.worker_failure else ""))
            if reading is None:
                return "\n".join(lines + ["Windows 外部放映页码：本次运行尚未收到后台读取记录"])
            target = f"0x{reading.hwnd:X}" if reading.hwnd else "未检测到"
            if reading.detected_by_com:
                target += "（由 Office 放映接口识别）"
            if reading.window_class:
                target += " · " + reading.window_class
            lines.extend([f"最近一次 Windows 外部放映读取：{reading.timestamp}",
                          f"读数对应 HWND：{target} · 进程：{reading.process}"])
            if running:
                status = f"查询仍在进行（{int(time.monotonic() - reading.started)} 秒）"
            elif reading.page:
                status = reading.page
            elif reading.ended:
                status = "放映结束画面"
            else:
                status = "未取得页码"
            lines.append(f"外部读取结果：{status}")
            if (not reading.page and self.last_page is not None and self.last_page.hwnd == reading.hwnd
                    and self.last_page is not reading):
                lines.append(f"该次放映上一次有效页码：{self.last_page.page}（{self.last_page.timestamp}）")
            for key, label in _STAGES.items():
                if key in reading.stages:
                    outcome, code, value = reading.stages[key]
                    result = "一致" if key == "com_identity" and outcome == "ok" else _OUTCOMES[outcome]
                    if value is not None:
                        result += f" {value}"
                    if code:
                        result += f"（{code}）"
                    lines.append(f"{label}：{result}")
            if reading.wps_audience is not None:
                reason, *counts = reading.wps_audience
                labels = ("WPS 进程", "放映窗口", "目标屏上的放映窗口")
                details = "，".join(f"{label} {count if count is not None else '未取得'}"
                                   for label, count in zip(labels, counts))
                lines.append(f"WPS 放映窗口核对：{_WPS_AUDIENCE_REASONS[reason]}（{details}）")
            for (show, owner), decision in reading.window_candidates.items():
                lines.append(f"候选窗口 · COM 放映 HWND 0x{show:X} · 应用或窗格 HWND 0x{owner:X}"
                             f" · {_WINDOW_DECISIONS[decision]}")
            return "\n".join(lines)


_HISTORY = PageReadHistory()


def capture_page_read(hwnd=0, process="", window_class=""):
    return _HISTORY.capture(hwnd, process, window_class)


def note_detected_window(hwnd):
    with _HISTORY.lock:
        _HISTORY.detected = hwnd if type(hwnd) is int else 0


def note_reader_failure(error):
    with _HISTORY.lock:
        _HISTORY.worker_failure = _error_code(error)


def note_window_candidate(show_hwnd, owner_hwnd, decision):
    context = getattr(_LOCAL, "current", None)
    if context is None or type(decision) is not str or decision not in _WINDOW_DECISIONS:
        return
    handles = tuple(h if type(h) is int and 0 < h <= 0xffffffffffffffff else 0
                    for h in (show_hwnd, owner_hwnd))
    history, record = context
    with history.lock:
        candidates = record.window_candidates
        candidates.pop(handles, None)
        if len(candidates) == 4:
            del candidates[next(iter(candidates))]
        candidates[handles] = decision


def note_wps_audience(reason, *, processes, audiences, on_target_screen):
    context = getattr(_LOCAL, "current", None)
    if context is None or type(reason) is not str or reason not in _WPS_AUDIENCE_REASONS:
        return
    history, record = context
    counts = tuple(n if type(n) is int and 0 <= n <= 0xffffffff else None
                   for n in (processes, audiences, on_target_screen))
    with history.lock:
        record.wps_audience = (reason, *counts)


def note_page_read(stage, outcome, error=None, *, value=None):
    context = getattr(_LOCAL, "current", None)
    if context is None or stage not in _STAGES or outcome not in _OUTCOMES:
        return
    history, record = context
    number = value if stage in {"current", "total", "collection"} and type(value) is int else None
    span = (outcome, _error_code(error), number)
    with history.lock:
        previous = record.stages.get(stage)
        if previous and previous[0] == "error" and previous[1] and outcome == "error" and not span[1]:
            return
        # Discovery can try several providers. Retain a successful provider,
        # or the specific failure ahead of a later empty provider.
        if previous is not None and stage not in {"current", "total"}:
            if previous[0] == "ok" and outcome != "unavailable":
                return
            if previous[0] in {"error", "ambiguous"} and outcome in {"empty", "checking"}:
                return
        record.stages[stage] = span


def discovery_notes(stages=("native_object", "rot")):
    context = getattr(_LOCAL, "current", None)
    if context is None:
        return ()
    history, record = context
    with history.lock:
        return tuple((key, record.stages[key]) for key in stages if key in record.stages)


def reuse_discovery_notes(notes):
    context = getattr(_LOCAL, "current", None)
    if context is not None:
        history, record = context
        with history.lock:
            record.stages.update(notes)


def native_page_report(powerpoint_status=None, wps_status=None):
    return _HISTORY.report(powerpoint_status, wps_status)
