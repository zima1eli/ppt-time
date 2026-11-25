import http.client
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import uuid
import xml.etree.ElementTree as ET

from slideshow import SlideSnapshot
from wps_pdf import ADDIN_NAME, WpsPdfBridge


class WpsPdfTest(unittest.TestCase):
    def test_install_pair_report_expire_restart_and_remove(self):
        # mkdir inherits the user directory ACL on Windows.
        root = Path(tempfile.gettempdir()) / ("ppt-time-pdf-" + uuid.uuid4().hex)
        root.mkdir()
        settings, registry = root / "config", root / "wps-addins"
        registry.mkdir()
        manifest = registry / "publish.xml"
        manifest.write_text('<jsplugins><!--existing--><jsplugin name="unrelated" type="wpp"/></jsplugins>')
        self.addCleanup(manifest.unlink, missing_ok=True)
        import shutil
        self.addCleanup(shutil.rmtree, root)
        bridge = WpsPdfBridge(settings, registry)
        self.addCleanup(bridge.close)
        with patch("wps_pdf.sys.platform", "win32"):
            bridge.install()
        port = bridge.server.server_address[1]
        self.assertEqual(bridge.server.server_address[0], "127.0.0.1")
        self.assertEqual([node.get("name") for node in ET.parse(manifest).getroot()], ["unrelated", ADDIN_NAME])
        report = dict(key=bridge.key, session="testTab", seq=1, active=True, current=3, total=7,
                      document="talk.pdf")

        def post(obj):
            conn = http.client.HTTPConnection("127.0.0.1", port, timeout=2)
            try:
                conn.request("POST", "/report", json.dumps(obj))
                return conn.getresponse().status
            finally:
                conn.close()

        self.assertEqual(post(dict(report, key="wrong")), 403)
        self.assertEqual(post(report), 204)
        with patch("wps_pdf._wps_pdf_window", return_value=0):
            self.assertEqual(bridge.merge(SlideSnapshot()).page_text, "3/7")
            self.assertEqual(post(dict(report, seq=2, current=4)), 204)
            self.assertEqual(post(report), 204)
            self.assertEqual(bridge.merge(SlideSnapshot()).page_text, "4/7")
            with patch("wps_pdf.time.monotonic", return_value=1e15):
                self.assertFalse(bridge.merge(SlideSnapshot()).active)
        self.assertEqual(post(dict(report, seq=3, active=False)), 204)
        self.assertFalse(bridge.merge(SlideSnapshot()).active)
        key = bridge.key
        bridge.close()
        with patch("wps_pdf.sys.platform", "win32"):
            bridge.start()
        self.assertEqual(bridge.server.server_address[1], port)
        self.assertEqual(bridge.key, key)
        bridge.uninstall()
        self.assertFalse(bridge.state_file.exists())
        self.assertEqual([n.get("name") for n in ET.parse(manifest).getroot()], ["unrelated"])
        self.assertIn("<!--existing-->", manifest.read_text())
