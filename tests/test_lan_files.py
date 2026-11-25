import io
from pathlib import Path
import shutil
import tempfile
import threading
import unittest
import uuid

from lan_control import MeetingFiles, MeetingServer, RequestError, paired_request


class MeetingArchiveTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(tempfile.gettempdir()) / ("ppt-time-files-" + uuid.uuid4().hex)
        self.root.mkdir()
        self.root = self.root.resolve()
        self.addCleanup(shutil.rmtree, self.root)
        self.active = threading.Event()
        self.active.set()
        self.files = MeetingFiles(self.root)

    def put_pdf(self, name, content):
        data = b"%PDF-1.7\n" + content + b"\n%%EOF"
        self.files.receive(name, len(data), io.BytesIO(data).read, self.active)
        return data

    def test_restore_preserves_folders_content_and_duplicate_submissions(self):
        name = "会场A/报告/第10讲.pdf"
        old = self.put_pdf(name, b"old")
        old_mtime = self.files.listing()[0]["modified"]
        self.files.archive(name)
        self.assertFalse((self.root / name).exists())
        with self.assertRaises(RequestError):
            self.files.for_open(name)
        self.assertEqual(self.files.total, len(old))
        reopened = MeetingFiles(self.root)
        catalog = reopened.catalog()
        self.assertEqual(catalog["files"], [])
        token = catalog["archived"][0]["id"]
        self.assertEqual(catalog["archived"][0]["name"], name)
        self.files = reopened
        new = self.put_pdf(name, b"new")
        restored = self.files.restore(token)
        self.assertEqual(restored, "会场A/报告/第10讲 (2).pdf")
        self.assertEqual((self.root / name).read_bytes(), new)
        self.assertEqual((self.root / restored).read_bytes(), old)
        self.assertEqual(self.files.catalog()["archived"], [])
        self.assertEqual(self.files.modified[restored], old_mtime)
        with self.assertRaises(RequestError):
            self.files.restore(token)

    def test_archive_and_restore_http_use_control_role_and_refresh_board(self):
        calls, received = [], []
        server = MeetingServer("127.0.0.1", 0, self.files, lambda a,v: calls.append((a,v)), received.append)
        worker = threading.Thread(target=server.serve_forever, kwargs={"poll_interval":.01}, daemon=True)
        worker.start()
        try:
            name = "speaker.pdf"
            data = self.put_pdf(name, b"file")
            for role in ("submit", "monitor"):
                with self.assertRaises(ValueError):
                    paired_request(server.link(role), "POST", "/api/files/archive", data={"name":name})
            self.assertEqual(self.files.for_open(name).read_bytes(), data)
            controller = server.link("control")
            result = paired_request(controller, "POST", "/api/files/archive", data={"name":name})
            self.assertEqual(result["name"], name)
            self.assertEqual(calls, [("inbox_archived", str(self.root / name))])
            token = paired_request(controller, "GET", "/api/files")["archived"][0]["id"]
            result = paired_request(controller, "POST", "/api/files/restore", data={"id":token})
            self.assertEqual(received, [str(self.root / name)])
            self.assertEqual(result["name"], name)
        finally:
            server.active.clear()
            server.shutdown()
            server.server_close()
            worker.join(2)
