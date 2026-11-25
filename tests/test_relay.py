import hashlib
import shutil
import tempfile
import threading
import unittest
import uuid
from pathlib import Path
from urllib.error import HTTPError
from urllib.parse import quote
from urllib.request import Request, urlopen
from http.server import ThreadingHTTPServer

from cloud_transfer import RoomClient
from relay_server import Handler, Relay


class RelayTransferTests(unittest.TestCase):
    def setUp(self):
        base = Path(tempfile.gettempdir()) / ("ppt-time-relay-" + uuid.uuid4().hex)
        base.mkdir()
        self.addCleanup(shutil.rmtree, base)
        self.store = Relay(base / "server")
        self.addCleanup(self.store.db.close)
        self.server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        self.server.store = self.store
        self.url = "http://127.0.0.1:" + str(self.server.server_port)
        thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        thread.start()
        self.addCleanup(self.server.server_close)
        self.addCleanup(self.server.shutdown)
        self.client = RoomClient(self.url, self.store.receive, base / "received")

    def submit(self, token, name, payload, slot=""):
        req = Request(self.url + "/v1/files", data=payload, method="POST", headers={
            "Authorization": "Bearer " + token, "X-File-Name": quote(name, safe=""), "X-Meeting-Slot": slot})
        with urlopen(req, timeout=3) as response:
            self.assertEqual(response.status, 201)

    def test_phone_to_receiver_and_verified_deletion(self):
        upload_token = self.client.get_json("/v1/pair")["submit"]
        self.submit(upload_token, "讲者 A.pdf", b"%PDF-1.4\nmeeting-example")
        self.submit(upload_token, "讲者 A.pdf", b"%PDF-1.4\ndifferent-talk")
        paths = []
        self.client.collect(threading.Event(), paths.append)
        self.assertEqual([Path(p).name for p in paths], ["讲者 A.pdf"] * 2)
        self.assertNotEqual(paths[0], paths[1])
        self.assertEqual(Path(paths[1]).read_bytes(), b"%PDF-1.4\ndifferent-talk")
        self.assertEqual(self.client.get_json("/v1/pending"), [])
        self.assertFalse(any(len(p.name) == 32 for p in self.store.path.iterdir()))

    def test_roles_validation_and_acknowledgement(self):
        with self.assertRaises(HTTPError) as denied:
            self.submit(self.store.receive, "sample.pdf", b"file")
        self.assertEqual(denied.exception.code, 403)
        with self.assertRaises(HTTPError) as unsafe:
            self.submit(self.store.submit, "../sample.pdf", b"file")
        self.assertEqual(unsafe.exception.code, 400)
        self.submit(self.store.submit, "sample.pptx", b"test")
        item, = self.client.get_json("/v1/pending")
        with self.assertRaises(HTTPError):
            self.client.request("/v1/ack/"+item["id"], "POST", **{"X-SHA256": "a"*64})
        self.assertTrue((self.store.path/item["id"]).exists())
        self.assertEqual(item["sha256"], hashlib.sha256(b"test").hexdigest())

    def test_speaker_selects_a_published_session(self):
        slot = uuid.uuid4().hex
        import json
        schedule = [{"id":slot,"label":"09:30 · 主会场 · 研讨"}]
        with self.client.request("/v1/agenda","POST",json.dumps(schedule).encode()):
            pass
        with urlopen(Request(self.url+"/v1/agenda",headers={"Authorization":"Bearer "+self.store.submit})) as response:
            self.assertEqual(json.load(response), schedule)
        self.submit(self.store.submit, "opening.pdf", b"%PDF-speaker", slot)
        received = []
        self.client.collect(threading.Event(), self.fail, lambda path,key:received.append((Path(path).name,key)))
        self.assertEqual(received,[("opening.pdf",slot)])
