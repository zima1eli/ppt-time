from contextlib import ExitStack
import http.client
import json
import os
from pathlib import Path
import shutil
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch
import uuid
import xml.etree.ElementTree as ET

from slideshow import SlideSnapshot
import wps_presentation as wpp


class WpsPresentationBridgeTests(unittest.TestCase):
    def setUp(self):
        self.stack = ExitStack()
        self.addCleanup(self.stack.close)
        root = Path(tempfile.gettempdir()) / ("ppt-time-wpp-" + uuid.uuid4().hex)
        root.mkdir()
        self.stack.callback(shutil.rmtree, root)
        self.stack.enter_context(patch.object(sys, "platform", "win32"))
        self.preferences = {}
        self.registry = SimpleNamespace(get=lambda v, k, n: self.preferences.get((k,n)),
                                        set=lambda v, k, n, value: self.preferences.update({(k,n):value}),
                                        remove=lambda v, k, n: self.preferences.pop((k,n), None),
                                        prune=lambda *args:None)
        self.bridge = wpp.WpsPresentationBridge(root / "settings", root / "addons", registry=self.registry)
        self.stack.callback(self.bridge.close)
        self.report = dict(session="show", seq=1, active=True, current=1, total=5,
                           document=r"D:\decks\talk.pptx", hwnd=0)
        self.external = SlideSnapshot(True, hwnd=123)

    def post(self, message):
        port = self.bridge.server.server_address[1]
        conn = http.client.HTTPConnection("127.0.0.1", port, timeout=2)
        try:
            conn.request("POST", "/report", json.dumps(message), headers={"Content-Type": "text/plain"})
            return conn.getresponse().status
        finally:
            conn.close()

    def test_opt_in_install_restart_remove(self):
        bridge = self.bridge
        bridge.start()
        self.assertFalse(bridge.state_file.parent.exists())
        self.assertFalse(bridge.addin_root.exists())
        bridge.addin_root.mkdir()
        original = '<jsplugins><!--keep--><jsplugin name="ppt-time-pdf" type="pdf"/></jsplugins>'
        bridge.manifest.write_text(original)
        bridge.install()
        self.assertEqual(bridge.server.server_address[0], "127.0.0.1")
        manifest = bridge.manifest.read_bytes()
        entry = ET.fromstring(manifest).find('./jsplugin[@type="wpp"]')
        self.assertEqual(entry.get("url"), bridge.folder.resolve().as_uri())
        self.assertIsNone(entry.get("enable"))
        self.assertEqual(entry.get("version"), wpp.ADDIN_VERSION)
        self.assertEqual(set(wpp.ASSETS) | {"connection.js"}, {p.name for p in bridge.folder.iterdir()})
        config = (bridge.folder / "connection.js").read_text().removeprefix("window.pptTimeWppBridge = ").strip().rstrip(";")
        pair = json.loads(config)
        self.assertEqual(pair, dict(key=bridge.key, url=f"http://127.0.0.1:{bridge.server.server_port}/report"))
        bridge.install()
        self.assertEqual(manifest, bridge.manifest.read_bytes())
        bridge.close()
        bridge.start()
        self.assertEqual(pair["key"], bridge.key)
        self.assertEqual(pair["url"], f"http://127.0.0.1:{bridge.server.server_port}/report")
        unrelated = bridge.folder / "another.txt"
        unrelated.write_bytes(b"keep")
        bridge.uninstall()
        bridge.uninstall()
        self.assertIsNone(bridge.server)
        self.assertFalse(bridge.state_file.exists())
        self.assertEqual(unrelated.read_bytes(), b"keep")
        self.assertIn("<!--keep-->", bridge.manifest.read_text())
        self.assertEqual([n.get("name") for n in ET.parse(bridge.manifest).getroot()], ["ppt-time-pdf"])

    def test_installer_registers_without_listening_and_can_rollback_uninstall(self):
        bridge = self.bridge
        with patch.object(wpp, "_wps_installed", return_value=False):
            bridge.install_for_setup()
            self.assertFalse(bridge.state_file.parent.exists())
            self.assertFalse(bridge.manifest.exists())
            self.assertFalse(self.preferences)
        with patch.object(wpp, "_wps_installed", return_value=True):
            bridge.install_for_setup()
            key, port = bridge._state()
            self.assertIsNone(bridge.server)
            self.assertTrue(key and port)
            self.assertEqual(bridge._owner(), "setup")
            bridge.start()
            # A repair process reuses the endpoint even when another process already owns the port.
            repair = wpp.WpsPresentationBridge(bridge.state_file.parent, bridge.addin_root, registry=self.registry)
            repair.install_for_setup()
            self.assertIsNone(repair.server)
            self.assertEqual(repair._state(), (key, port))
            self.assertEqual(self.post(dict(self.report, key=key)), 204)
            bridge.close()
            repair.remove_for_setup()
            self.assertFalse(repair.state_file.exists())
            self.assertIsNone(repair._owner())
            repair.install_for_setup()
            self.assertEqual(repair._owner(), "setup")
            self.assertTrue(repair._state()[0])
            self.assertIn(wpp.ADDIN_NAME, bridge.manifest.read_text())
            bridge.uninstall()
            self.assertEqual(bridge._owner(), "disabled")
            repair.install_for_setup()
            self.assertFalse(bridge.state_file.exists())
            # A fresh explicit request in the tray belongs to the user.
            bridge.install()
            self.assertEqual(bridge._owner(), "manual")
            repair.remove_for_setup()
            self.assertTrue(bridge.state_file.exists())

    def test_custom_office_install_roots_require_a_real_presentation_executable(self):
        from contextlib import nullcontext
        root = self.bridge.addin_root.parent
        install = root / "custom location" / "12.1.0.29690" / "office6"
        install.mkdir(parents=True)
        exe = install / "wpp.exe"
        exe.write_bytes(b"fixture")
        values = {}
        uninstall = r"Software\Microsoft\Windows\CurrentVersion\Uninstall"
        def open_key(hive, path, reserved, flags):
            if path == uninstall or (hive, path, flags) in values:
                return nullcontext((hive, path, flags))
            raise FileNotFoundError()
        def query(handle, key):
            try:
                return values[handle][key], 1
            except KeyError as exc:
                raise FileNotFoundError() from exc
        reg = SimpleNamespace(KEY_READ=0, KEY_WOW64_64KEY=64, KEY_WOW64_32KEY=32, HKEY_CURRENT_USER=1,
                              HKEY_LOCAL_MACHINE=2, HKEY_CLASSES_ROOT=3, OpenKey=open_key, QueryValueEx=query,
                              EnumKey=lambda handle, i:"CustomWps", QueryInfoKey=lambda h:(1,0,0))
        ts = SimpleNamespace(ProcessIdToSessionId=lambda _:1, WTSEnumerateProcesses=lambda:[], error=OSError)
        with (patch.dict(sys.modules, winreg=reg, win32ts=ts),
              patch.dict(os.environ, {k:str(root) for k in ("LOCALAPPDATA","ProgramFiles","ProgramFiles(x86)")})):
            common = r"Software\Kingsoft\Office\6.0\Common"
            values[1,common,32] = {"InstallRoot":str(install.parent)}
            self.assertTrue(wpp._wps_installed())
            values.clear()
            values[2,uninstall + r"\CustomWps",64] = {"DisplayName":"WPS Office", "InstallLocation":str(install.parents[1])}
            self.assertTrue(wpp._wps_installed())
            exe.unlink()
            self.assertFalse(wpp._wps_installed())

    def test_start_refreshes_legacy_manual_scripts_without_repairing_user_opt_out(self):
        bridge = self.bridge
        with patch.object(wpp, "_wps_installed", return_value=True):
            bridge.install()
            bridge.close()
            self.preferences.clear()  # 1.0.7 had no owner marker.
            script = bridge.folder / "main.js"
            expected = script.read_bytes()
            real_settings = bridge.state_file.read_bytes()
            connection = (bridge.folder / "connection.js").read_bytes()
            script.write_bytes(b"older")
            bridge.start()
            self.assertEqual(script.read_bytes(), expected)
            bridge.close()
            script.write_bytes(b"another older release")
            bridge.install_for_setup()
            self.assertEqual(script.read_bytes(), expected)
            self.assertIsNone(bridge._owner())
            self.assertEqual(real_settings, bridge.state_file.read_bytes())
            self.assertEqual(connection, (bridge.folder / "connection.js").read_bytes())
            bridge.remove_for_setup()
            self.assertTrue(script.exists())
            bridge.manifest.write_text("<jsplugins/>")
            bridge.install_for_setup()
            self.assertEqual(bridge.manifest.read_text(), "<jsplugins/>")
            bridge.start()
            self.assertIsNone(bridge.server)

    def test_uninstall_rollback_skips_user_disabled_entries(self):
        bridge = self.bridge
        with patch.object(wpp, "_wps_installed", return_value=True):
            bridge.install_for_setup()
            xml = ET.parse(bridge.manifest)
            xml.getroot().find("jsplugin").set("enable", "disable")
            xml.write(bridge.manifest)
            bridge.start()
            self.assertIsNone(bridge.server)
            bridge.remove_for_setup()
            self.assertEqual(bridge._owner(), "disabled")
            bridge.install_for_setup()
            self.assertFalse(bridge.state_file.exists())

    def test_refresh_failure_retains_old_working_endpoint(self):
        bridge = self.bridge
        bridge.install()
        bridge.close()
        script = bridge.folder / "main.js"
        script.write_bytes(b"older")
        with patch.object(wpp, "_atomic_write", side_effect=PermissionError("WPS holds file")):
            bridge.start()
        self.assertIsNotNone(bridge.server)
        self.assertEqual(script.read_bytes(), b"older")
        self.assertIn("未能更新", bridge.status())

    def test_manifest_failures_preserve_prior_state(self):
        bridge = self.bridge
        bridge.addin_root.mkdir()
        for xml in ('<x>', '<wrong/>',
                    f'<jsplugins><jsplugin type="wpp" name="{wpp.ADDIN_NAME}" url="https://somewhere/"/></jsplugins>'):
            with self.subTest(manifest=xml):
                bridge.manifest.write_text(xml)
                with self.assertRaises(ValueError):
                    bridge.install()
                self.assertEqual(bridge.manifest.read_text(), xml)
                self.assertFalse(bridge.folder.exists())
                self.assertIsNone(bridge.server)
        bridge.manifest.write_text("<jsplugins/>")
        change = wpp._change
        def fail_registration(path, previous, data):
            if path == bridge.manifest:
                raise OSError("manifest locked")
            change(path, previous, data)
        with patch.object(wpp, "_change", side_effect=fail_registration), self.assertRaises(OSError):
            bridge.install()
        self.assertFalse(bridge.state_file.exists())
        self.assertEqual(list(bridge.folder.iterdir()), [])
        self.assertIsNone(bridge.server)

    def test_unregistered_directory_is_not_overwritten(self):
        bridge = self.bridge
        bridge.folder.mkdir(parents=True)
        main = bridge.folder / "main.js"
        main.write_bytes(b"another extension")
        with self.assertRaises(ValueError):
            bridge.install()
        bridge.uninstall()
        self.assertEqual(main.read_bytes(), b"another extension")
        self.assertIsNone(bridge.server)

    def test_manual_removal_and_cloud_identity(self):
        bridge = self.bridge
        bridge.install()
        bridge.close()
        bridge.manifest.write_text("<jsplugins/>")
        bridge.start()
        self.assertIsNone(bridge.server)
        self.assertIn("重新启用", bridge.status())
        self.assertFalse(wpp._same_document("https://office.local/teamA/test.pptx", "https://office.local/teamB/test.pptx"))
        self.assertTrue(wpp._same_document("test.pptx", "https://office.local/teamA/test.pptx"))

    def test_authenticated_jump_rewind_expiry_and_session_isolation(self):
        bridge = self.bridge
        bridge.install()
        report = dict(self.report, key=bridge.key)
        with (patch.object(wpp, "_audience", return_value=self.external.hwnd),
              patch.object(wpp, "_screen_info", return_value=("Projector", (1, 1)))):
            self.assertEqual(self.post(dict(report, key="incorrect")), 403)
            self.assertEqual(self.post(dict(report, hwnd="wrong")), 400)
            self.assertEqual(self.post(dict(report, owner_hwnd="wrong")), 400)
            for seq, page in enumerate((1, 5, 2), start=1):
                self.assertEqual(self.post(dict(report, seq=seq, current=page)), 204)
                merged = bridge.merge(self.external)
                self.assertEqual((merged.page_text, merged.hwnd, merged.screen_name), (f"{page}/5", 123, "Projector"))
            self.assertEqual(self.post(report), 204)
            self.assertEqual(bridge.merge(self.external).current, 2)
            self.assertIn("2/5", bridge.status())
            self.assertEqual(self.post(dict(report, session="other")), 204)
            self.assertIs(bridge.merge(self.external), self.external)
            self.assertEqual(self.post(dict(report, session="other", seq=2, active=False)), 204)
            self.assertEqual(self.post(dict(report, seq=4, total=None)), 204)
            self.assertEqual(bridge.merge(self.external).page_text, "第 1 页")
            self.assertEqual(self.post(dict(report, seq=5, current=None, show_position=4)), 204)
            self.assertEqual(bridge.merge(self.external).page_text, "放映第 4 页")
            with patch.object(wpp.time, "monotonic", return_value=1e20):
                self.assertIs(bridge.merge(self.external), self.external)
            self.assertEqual(self.post(dict(report, seq=6, ended=True)), 204)
            self.assertTrue(bridge.merge(self.external).ended)
            self.assertEqual(self.post(dict(report, seq=7, active=False)), 204)
            self.assertIs(bridge.merge(self.external), self.external)

    def test_match_zero_hwnd_to_only_dedicated_wps_show_and_reject_wrong_process(self):
        report = SlideSnapshot(True, 1, 5, document=self.report["document"])
        gui = SimpleNamespace(EnumWindows=lambda cb, param: cb(123, param), error=OSError)
        with (patch.dict(sys.modules, win32gui=gui), patch.object(wpp, "_unique_wps_show", return_value=True) as only,
              patch.object(wpp, "_process_name", return_value="wpp.exe") as proc,
              patch.object(wpp, "_window_caption", return_value="") as title,
              patch.object(wpp, "_pid", return_value=456), patch.object(wpp, "_is_wps_show_surface", return_value=True)):
            self.assertEqual(wpp._audience(report, 123, ""), 123)
            self.assertEqual(wpp._audience(report, 0, ""), 123)
            only.return_value = False
            self.assertEqual(wpp._audience(report, 123, ""), 0)
            only.return_value = True
            for other in ("powerpnt.exe", "wpspdf.exe", "wps.exe"):
                proc.return_value = other
                self.assertEqual(wpp._audience(report, 123, ""), 0)
            title.return_value = "放映 - [talk.pptx]"
            self.assertEqual(wpp._audience(report, 123, ""), 123)
            title.return_value = "放映 - [other.pptx]"
            self.assertEqual(wpp._audience(report, 123, ""), 0)
            with patch.object(wpp, "_native_root", return_value=999):
                wrong = SlideSnapshot(True, 4, 7, hwnd=999)
                self.assertEqual(wpp._audience(wrong, 123, ""), 0)

    def test_preserves_different_document_and_other_viewers(self):
        bridge = self.bridge
        bridge.install()
        self.assertEqual(self.post(dict(self.report, key=bridge.key)), 204)
        with patch.object(wpp, "_audience", side_effect=AssertionError("should not query native windows")):
            for ext in (SlideSnapshot(True, document=r"E:\decks\talk.pptx", hwnd=123),
                        SlideSnapshot(True, document="book.pdf", hwnd=123), SlideSnapshot(True),
                        SlideSnapshot(True, ended=True, hwnd=123)):
                self.assertIs(bridge.merge(ext), ext)

    def test_verified_editor_pid_anchors_hwndless_report_without_becoming_projection(self):
        bridge = self.bridge
        bridge.install()
        self.assertEqual(self.post(dict(self.report, key=bridge.key, owner_hwnd=555)), 204)
        report = bridge._fresh()[0]
        gui = SimpleNamespace(EnumWindows=lambda cb, arg: [cb(i,arg) for i in (123,555,999)], error=OSError)
        with (patch.dict(sys.modules, win32gui=gui), patch.object(wpp, "_unique_wps_show", return_value=True) as only,
              patch.object(wpp, "_process_name", return_value="wpp.exe"), patch.object(wpp, "_window_caption", return_value=""),
              patch.object(wpp, "_pid", side_effect=lambda h: 100 if h in (123,555) else 200),
              patch.object(wpp, "_native_root", side_effect=lambda h:h),
              patch.object(wpp, "_is_wps_show_surface", side_effect=lambda h:h in (123,999))):
            self.assertEqual(wpp._audience(report, 123, ""), 123)
            only.assert_called_with(123, 100, check_processes=False, document=report.document)
            self.assertEqual(wpp._audience(report, 0, ""), 123)
            self.assertEqual(wpp._audience(wpp.replace(report, hwnd=555), 123, ""), 123)
            self.assertEqual(wpp._audience(report, 999, ""), 0)
            self.assertEqual(wpp._audience(wpp.replace(report, hwnd=999), 123, ""), 0)
            only.return_value = False
            self.assertEqual(wpp._audience(report, 123, ""), 0)
