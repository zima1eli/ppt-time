import http.client
import json
import os
import shutil
import tempfile
import unittest
import uuid
from pathlib import Path
from unittest.mock import patch

import powerpoint_bridge as pp
from slideshow import SlideSnapshot


class MemoryRegistry:
    """Two CLSID views, one shared Office and ProgID view, like 64-bit Windows."""
    views = (64, 32)

    def __init__(self):
        self.items = {}
        self.failed_key = ""

    def index(self, view, key, hive):
        return hive, view if key.startswith(pp.CLS_KEY) else 0, key

    def exists(self, view, key, hive="user"):
        return self.index(view, key, hive) in self.items

    def get(self, view, key, name, hive="user"):
        return self.items.get(self.index(view, key, hive), {}).get(name, pp._MISSING)

    def set(self, view, key, name, value):
        if key == self.failed_key:
            raise PermissionError("test fault")
        self.items.setdefault(self.index(view,key,"user"), {})[name] = value

    def remove(self, view, key, name):
        self.items.get(self.index(view,key,"user"), {}).pop(name,None)

    def prune(self, view, key):
        ident = self.index(view,key,"user")
        if not self.items.get(ident):
            self.items.pop(ident,None)


class PowerPointBridgeTest(unittest.TestCase):
    def setUp(self):
        self.root = Path(tempfile.gettempdir()) / ("ppt-time-ppt-" + uuid.uuid4().hex)
        self.root.mkdir()
        self.addCleanup(shutil.rmtree, self.root)
        self.reg = MemoryRegistry()
        self.bundle = self.root / "bundle"
        self.bundle.mkdir()
        (self.bundle / pp.DLL).write_bytes(b"fixture-dll")
        self.bridge = pp.PowerPointBridge(self.root, self.bundle, registry=self.reg,
                                         runtime_dir=self.root/"roaming", install_dir=self.root/"local")
        self.addCleanup(self.bridge.close)
        mock = patch("powerpoint_bridge.sys.platform", "win32")
        mock.start()
        self.addCleanup(mock.stop)

    def post(self, seq=1, **fields):
        body = {"key":self.bridge.key, "session":"live-1", "seq":seq, "pid":200, "hwnd":10,
                "active":True, "ended":False, "current":3, "total":8, "document":r"C:\deck\Talk.pptx"}
        body.update(fields)
        conn = http.client.HTTPConnection(*self.bridge.server.server_address, timeout=2)
        try:
            conn.request("POST", "/report", json.dumps(body))
            return conn.getresponse().status
        finally:
            conn.close()

    def test_install_real_http_restart_remove(self):
        self.reg.set(64, pp.OFFICE_KEY + ".unrelated", "LoadBehavior", 1)
        self.bridge.start()
        self.assertIsNone(self.bridge.server)
        self.bridge.install()
        for view in self.reg.views:
            self.assertEqual(self.reg.get(view, pp.CLS_KEY,pp.MARKER),pp.OWNER)
            self.assertIn(pp.DLL, self.reg.get(view,pp.INPROC,"CodeBase"))
            self.assertEqual(self.reg.get(view,pp.OFFICE_KEY,"LoadBehavior"),3)
        endpoint = json.loads(self.bridge.endpoint.read_text())
        self.assertEqual(endpoint["key"],self.bridge.key)
        self.assertTrue(endpoint["url"].startswith("http://127.0.0.1:"))
        self.assertEqual(self.post(key="wrong"),403)
        self.assertEqual(self.post(pid=False),400)
        self.assertEqual(self.post(),204)
        self.assertIn("3/8",self.bridge.status())
        self.assertEqual(self.post(2,active=False),204)
        self.assertEqual(self.post(),204)
        self.assertIn("等待放映",self.bridge.status())
        old = self.bridge.key
        self.bridge.close()
        self.assertFalse(self.bridge.endpoint.exists())
        self.bridge.start()
        self.assertNotEqual(self.bridge.key,old)
        self.assertEqual(self.post(key=old),403)
        self.reg.set(32, pp.OFFICE_KEY, "unrelated_ext", "preserve")
        self.bridge.uninstall()
        self.assertIs(self.reg.get(64,pp.OFFICE_KEY,pp.MARKER),pp._MISSING)
        self.assertEqual(self.reg.get(64,pp.OFFICE_KEY,"unrelated_ext"),"preserve")
        self.assertEqual(self.reg.get(64,pp.OFFICE_KEY+".unrelated","LoadBehavior"),1)
        self.assertFalse(any((self.root/"local").glob("*/"+pp.DLL)))

    def test_blocked_policy_or_foreign_owner_leaves_registry_alone(self):
        policy = r"Software\Policies\Microsoft\Office\16.0\PowerPoint\Resiliency\AddinList"
        self.reg.set(64,policy,pp.PROGID,"0")
        with self.assertRaisesRegex(ValueError,"策略"):
            self.bridge.install()
        self.assertFalse(self.bridge.endpoint.exists())
        self.reg.remove(64,policy,pp.PROGID)
        self.reg.set(32,pp.CLS_KEY,"","other registration")
        with self.assertRaisesRegex(ValueError,"同名"):
            self.bridge.install()
        self.assertEqual(self.reg.get(32,pp.CLS_KEY,""),"other registration")
        self.assertFalse(self.reg.exists(64,pp.CLS_KEY))

    def test_setup_registers_office_only_and_preserves_live_receiver_on_uninstall(self):
        self.assertIn("跳过", self.bridge.install_for_setup())
        self.assertFalse(self.reg.items)
        self.reg.set(64, pp.POWERPOINT_CLASS, "", pp.POWERPOINT_CLSID)
        self.assertIn("启用", self.bridge.install_for_setup())
        self.assertIsNone(self.bridge.server)
        self.assertFalse(self.bridge.endpoint.exists())
        self.assertEqual(self.reg.get(64, pp.PREF_KEY, pp.PREF_NAME), "setup")
        self.bridge.start()
        endpoint = self.bridge.endpoint.read_bytes()
        self.assertIn("保留", self.bridge.install_for_setup())
        self.assertIn("移除", self.bridge.remove_for_setup())
        self.assertEqual(self.bridge.endpoint.read_bytes(), endpoint)
        self.assertIsNotNone(self.bridge.server)
        self.assertFalse(self.reg.exists(64, pp.PREF_KEY))
        self.assertFalse(self.reg.exists(64, pp.OFFICE_KEY))

    def test_setup_respects_manual_old_registration_office_disable_and_opt_out(self):
        self.reg.set(64, pp.POWERPOINT_CLASS, "", pp.POWERPOINT_CLSID)
        self.bridge.install()
        self.bridge.close()
        self.assertIn("手动", self.bridge.install_for_setup())
        self.bridge.remove_for_setup()
        self.assertEqual(self.reg.get(64, pp.OFFICE_KEY, pp.MARKER), pp.OWNER)
        self.reg.remove(64, pp.PREF_KEY, pp.PREF_NAME)  # Manually installed before version 1.0.5.
        self.bridge.install_for_setup()
        self.assertIs(self.reg.get(64, pp.PREF_KEY, pp.PREF_NAME), pp._MISSING)
        self.reg.set(64, pp.OFFICE_KEY, "LoadBehavior", 2)
        self.assertIn("停用", self.bridge.install_for_setup())
        self.assertEqual(self.reg.get(32, pp.OFFICE_KEY, "LoadBehavior"), 2)
        self.bridge.start()
        self.assertIsNone(self.bridge.server)
        self.bridge.uninstall()
        self.assertIn("停用", self.bridge.install_for_setup())
        self.assertFalse(self.reg.exists(64, pp.OFFICE_KEY))
        self.bridge.install()
        self.assertEqual(self.reg.get(64, pp.PREF_KEY, pp.PREF_NAME), "manual")

    def test_setup_respects_policy_foreign_keys_and_removed_registration(self):
        self.reg.set(64, pp.POWERPOINT_CLASS, "", pp.POWERPOINT_CLSID)
        key = r"Software\Policies\Microsoft\Office\16.0\PowerPoint\Security"
        for name in ("RequireAddinSig", "DisableAllAddins"):
            self.reg.set(64, key, name, 1)
            self.assertIn("信任中心", self.bridge.install_for_setup())
            self.assertFalse(self.reg.exists(64, pp.PREF_KEY))
            self.reg.remove(64, key, name)
        self.reg.set(32, pp.CLS_KEY, "", "another program")
        self.assertIn("同名", self.bridge.install_for_setup())
        self.assertFalse(self.reg.exists(64, pp.OFFICE_KEY))
        self.reg.remove(32, pp.CLS_KEY, "")
        self.reg.prune(32, pp.CLS_KEY)
        self.bridge.install_for_setup()
        self.reg.items.pop(self.reg.index(64, pp.OFFICE_KEY, "user"))
        self.assertIn("移除", self.bridge.install_for_setup())
        self.assertFalse(self.reg.exists(64, pp.OFFICE_KEY))

    def test_failed_registration_does_not_leave_half_installed_keys(self):
        self.reg.failed_key = pp.INPROC
        with self.assertRaises(PermissionError):
            self.bridge.install()
        self.assertFalse(self.reg.items)
        self.assertFalse(self.bridge.endpoint.exists())
        self.assertIsNone(self.bridge.server)
        self.assertFalse(any(self.bridge.install_dir.glob("*/" + pp.DLL)))

    def test_physical_pages_and_custom_show_position_are_independent(self):
        self.bridge.install()
        cases = [
            ({"current":None, "total":20, "show_position":4}, (None,20,4,"放映第 4 页")),
            ({"current":15, "total":20, "show_position":4}, (15,20,4,"15/20")),
            ({"current":15, "total":False, "show_position":4}, (15,None,4,"第 15 页")),
            ({"current":21, "total":20, "show_position":4}, (None,20,4,"放映第 4 页")),
            ({"current":"8", "total":20, "show_position":True}, (None,20,None,"")),
            ({"active":False, "show_position":4}, (None,None,None,"")),
            ({"ended":True, "show_position":4}, (None,None,None,"")),
        ]
        for seq, (report, expected) in enumerate(cases,1):
            with self.subTest(report=report):
                self.assertEqual(self.post(seq, **report),204)
                snapshot = self.bridge._fresh()[0].snapshot
                self.assertEqual((snapshot.current,snapshot.total,snapshot.show_position,snapshot.page_text),expected)

    def test_merge_identity_quality_expiry_and_exit(self):
        self.bridge.install()
        observed = SlideSnapshot(True, screen_name="Primary", hwnd=10, document=r"C:\deck\Talk.pptx")
        with (patch("powerpoint_bridge._process_name",return_value="powerpnt.exe"),
              patch("powerpoint_bridge._valid_show",return_value=True),
              patch("powerpoint_bridge._pid",return_value=200),
              patch("powerpoint_bridge._show_windows",return_value=[10]),
              patch("powerpoint_bridge._window_caption",return_value="Slide Show - [Talk.pptx]"),
              patch("powerpoint_bridge._screen_info",return_value=("Projector",(2,2)))):
            self.post()
            snap = self.bridge.merge(observed)
            self.assertEqual((snap.current,snap.total,snap.screen_name),(3,8,"Projector"))
            self.assertEqual(self.bridge.merge(SlideSnapshot(True,hwnd=11)).hwnd,11)
            wrong = SlideSnapshot(True,document=r"D:\other\Talk.pptx",hwnd=10)
            self.assertIs(self.bridge.merge(wrong),wrong)
            different = SlideSnapshot(True,document="manual.pdf")
            self.assertIs(self.bridge.merge(different),different)
            self.post(2,hwnd=0,current=6)
            self.assertEqual(self.bridge.merge(observed).page_text,"6/8")
            with patch("powerpoint_bridge._show_windows",return_value=[10,11]):
                self.assertIs(self.bridge.merge(observed),observed)
            self.post(3,total=None)
            good = SlideSnapshot(True,3,8,hwnd=10)
            self.assertIs(self.bridge.merge(good),good)
            self.post(4,ended=True)
            self.assertTrue(self.bridge.merge(observed).ended)
            self.post(5,active=False)
            self.assertIs(self.bridge.merge(observed),observed)
            self.post(6)
            with patch("powerpoint_bridge.time.monotonic",return_value=1e15):
                self.assertIs(self.bridge.merge(observed),observed)

    @unittest.skipUnless(os.name == "nt", "native Windows window hierarchy")
    def test_child_stale_and_editor_handles_are_matched_to_the_native_audience(self):
        import win32api
        import win32con
        import win32gui

        self.bridge.install()
        kernel = win32api.GetModuleHandle(None)
        frames = set()
        def frame(title, parent=0):
            style = win32con.WS_VISIBLE | (win32con.WS_CHILD if parent else win32con.WS_POPUP)
            hwnd = win32gui.CreateWindowEx(win32con.WS_EX_NOACTIVATE | win32con.WS_EX_TOOLWINDOW,
                "Static", title, style, -15000, -15000, 8, 8, parent, 0, kernel, None)
            frames.add(hwnd)
            self.addCleanup(lambda: win32gui.IsWindow(hwnd) and win32gui.DestroyWindow(hwnd))
            return hwnd

        audience = frame("PowerPoint Slide Show - [Talk.pptx]")
        projection_child = frame("", audience)
        editor = frame("PowerPoint")
        stale = frame("Closed")
        win32gui.DestroyWindow(stale)
        pid = os.getpid()
        observed = SlideSnapshot(True, hwnd=audience)
        with (patch("powerpoint_bridge._process_name", side_effect=lambda h: "powerpnt.exe" if h in frames else ""),
              patch("powerpoint_bridge._screen_info", return_value=("Projector", (2, 2)))):
            for seq, raw in enumerate((projection_child, editor, stale, 0), start=1):
                self.post(seq, pid=pid, hwnd=raw)
                actual = self.bridge.merge(observed)
                self.assertEqual((actual.page_text, actual.hwnd, actual.screen_name), ("3/8", audience, "Projector"),
                                 f"Office handle case {seq}")
            self.post(5, pid=pid, hwnd=editor, document="Different.pptx")
            self.assertIs(self.bridge.merge(observed), observed)
            self.post(6, pid=pid, hwnd=editor, document="")
            self.assertIs(self.bridge.merge(observed), observed)
            self.post(7, pid=pid+1, hwnd=projection_child)
            self.assertIs(self.bridge.merge(observed), observed)
            other = frame("PowerPoint Slide Show - [Second.pptx]")
            other_child = frame("", other)
            for seq, raw in enumerate((other, other_child, editor, 0), start=8):
                self.post(seq, pid=pid, hwnd=raw)
                self.assertIs(self.bridge.merge(observed), observed)
