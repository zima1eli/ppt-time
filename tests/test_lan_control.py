import http.client
import io
import json
import os
import shutil
import tempfile
import threading
import unittest
import uuid
from pathlib import Path
from urllib.parse import quote

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from lan_control import MeetingFiles, MeetingServer, RequestError, agenda_revision, media_name, paired_request, safe_agenda


def pdf(text=b"agenda"):
    return b"%PDF-1.7\n" + text + b"\n%%EOF"


class MeetingNetworkTests(unittest.TestCase):
    def setUp(self):
        self.scratch = Path(tempfile.gettempdir()) / ("ppt-time-net-" + uuid.uuid4().hex)
        self.scratch.mkdir()
        self.addCleanup(shutil.rmtree, self.scratch)
        self.files = MeetingFiles(self.scratch / "inbox")
        self.actions = []
        self.received = []
        self.server = MeetingServer("127.0.0.1", 0, self.files,
                                    lambda a, v: self.actions.append((a, v)), self.received.append)
        self.server.state = {"room": "room", "agenda": [{"title": "planning"}], "time": "04:05"}
        self.worker = threading.Thread(target=self.server.serve_forever, kwargs={"poll_interval": 0.01}, daemon=True)
        self.worker.start()
        self.addCleanup(self.stop)

    def stop(self):
        self.server.active.clear()
        self.server.shutdown()
        self.server.server_close()
        self.worker.join(2)

    def call(self, method, route, role=None, body=None, headers=None):
        con = http.client.HTTPConnection(*self.server.server_address, timeout=3)
        fields = {"Authorization": "Bearer " + self.server.tokens[role]} if role else {}
        fields.update(headers or {})
        try:
            con.request(method, route, body=body, headers=fields)
            reply = con.getresponse()
            data = reply.read()
            return reply.status, json.loads(data) if data and reply.getheader("Content-Type", "").startswith("application/json") else data
        finally:
            con.close()

    def test_role_scope_and_host(self):
        self.assertEqual(self.call("GET", "/")[0], 200)
        self.assertEqual(self.call("GET", "/api/state")[0], 401)
        self.assertEqual(self.call("GET", "/api/state", "monitor")[1]["state"]["time"], "04:05")
        self.assertNotIn("time", self.call("GET", "/api/state", "submit")[1]["state"])
        self.assertEqual(self.call("GET", "/api/files", "monitor")[0], 403)
        self.assertEqual(self.call("POST", "/api/action", "monitor", b'{"action":"next"}')[0], 403)
        self.assertEqual(self.call("POST", "/api/upload?name=x.pdf", "monitor", pdf())[0], 403)
        self.assertEqual(self.call("GET", "/api/state", "control", headers={"Host": "elsewhere.example"})[0], 403)
        self.assertEqual(self.call("POST", "/api/action", "control", b'{"action":"next"}',
                                   {"Origin":"https://elsewhere.example"})[0], 403)

    def test_folder_upload_duplicates_open_and_cleanup(self):
        name = "会场 A/报告.pdf"
        path = "/api/upload?name=" + quote(name, safe="")
        self.assertEqual(self.call("POST", path, "submit", pdf()), (201, {"name": name}))
        self.assertEqual(self.call("POST", path, "control", pdf(b"second"))[1]["name"], "会场 A/报告 (2).pdf")
        self.assertEqual(len(self.call("GET", "/api/files", "control")[1]["files"]), 2)
        self.assertEqual(len(self.received), 2)
        self.assertEqual(self.call("POST", "/api/action", "control", json.dumps({"action":"open","value":name}).encode())[0], 202)
        actual = self.files.root / "会场 A" / "报告.pdf"
        self.assertEqual(self.actions, [("open", str(actual))])
        self.assertEqual(actual.read_bytes(), pdf())
        self.assertEqual(self.call("POST", path, "control", b"MZ not a document")[0], 415)
        self.assertFalse(list(self.files.root.glob(".ppt-time-*")))
        self.assertEqual(self.files.in_flight, 0)
        self.assertEqual(self.files.reserved, 0)

    def test_name_and_interrupted_upload(self):
        for name in ("../../test.pdf", "/etc/a.pdf", "C:/bad.pdf", "a.pdf:stream", ".secret/a.pdf",
                     "CON.pdf", "NUL/part.pdf", "x/c.py", "folder\\x.pdf", "x. /a.pdf"):
            with self.subTest(name=name), self.assertRaises(RequestError):
                media_name(name)
        self.assertEqual(self.call("POST","/api/upload?name=%2E%2E%2Fbad.pdf","control",pdf())[0], 400)
        with self.assertRaises(RequestError):
            self.files.receive("a.pdf", 100, io.BytesIO(pdf()).read, self.server.active)
        self.assertEqual(self.files.listing(), [])
        self.assertFalse(list(self.files.root.iterdir()))

    def test_bounded_commands_and_preview(self):
        cases = [({"action": "set_seconds", "value": 600}, 202), ({"action": "set_seconds", "value": True}, 400),
                 ({"action": "next", "value": "a"}, 400), ({"action": "open", "value": "missing.pdf"}, 404),
                 ({"action": "interstitial","value": 0}, 202), ({"action": "unknown"}, 400)]
        for payload, status in cases:
            with self.subTest(payload=payload):
                self.assertEqual(self.call("POST", "/api/action", "control", json.dumps(payload).encode())[0], status)
        self.assertEqual(self.call("GET", "/api/preview", "monitor")[0], 403)
        self.server.preview_allowed=True
        self.assertEqual(self.call("GET", "/api/preview", "monitor")[0], 204)
        self.server.preview = b"fixture pixels"
        self.assertEqual(self.call("GET", "/api/preview", "control"), (200, b"fixture pixels"))
        self.assertEqual(self.call("GET", "/api/preview", "submit")[0], 403)

    def test_room_sync(self):
        link=self.server.link("control")
        rows=[{"title":"Report","start":"09:00","speaker":"S","seconds":500,"file":"private/name.pdf"}]
        self.assertTrue(paired_request(link,"POST","/api/agenda",data=rows)["accepted"])
        self.assertEqual(self.actions, [("import_agenda",safe_agenda(rows))])
        self.assertEqual(self.actions[0][1][0]["file"], "name.pdf")
        self.assertEqual(paired_request(self.server.link("monitor"),"GET","/api/state")["state"]["room"], "room")
        with self.assertRaises(ValueError):
            paired_request(self.server.link("submit"),"POST","/api/agenda",data=rows)
        with self.assertRaises(ValueError):
            paired_request("https://example.com/test#token","GET","/api/state")
        testfile=self.scratch / "sender.pdf"
        testfile.write_bytes(pdf())
        self.assertEqual(paired_request(self.server.link("submit"),"POST","/api/upload?name=sender.pdf",file=testfile)["name"],"sender.pdf")

    def test_remote_role_boundaries_and_assets(self):
        self.server.state.update(system={"cpu_percent": 35}, appearance={"lcd": False}, desktop_controls=False)
        for path in ("/agenda", "/agenda.js", "/agenda.css", "/obs", "/obs.js", "/obs.css", "/desk.css", "/desk.js"):
            self.assertEqual(self.call("GET", path)[0], 200)
        self.assertIn("/obs#", self.server.link("obs"))
        self.assertNotIn("system", self.call("GET", "/api/state", "obs")[1]["state"])
        self.assertNotIn("appearance", self.call("GET", "/api/state", "monitor")[1]["state"])
        act = lambda role, name, value=None: self.call("POST", "/api/action", role, json.dumps({"action": name, "value": value}).encode())[0]
        for action in ("start_show", "first", "last"):
            self.assertEqual(act("control", action), 202)
        self.assertEqual(act("obs", "first"), 403)
        self.assertEqual(act("control", "desk_show"), 403)
        self.server.state["desktop_controls"] = True
        self.assertEqual(act("control", "desk_show"), 202)
        self.assertEqual(act("monitor", "desk_show"), 403)
        self.assertEqual(act("control", "set_appearance", {"lcd": True,"time_size": 32}), 202)
        self.assertEqual(act("control", "set_appearance", {"size_mode": "160"}), 202)
        self.assertEqual(act("control", "set_appearance", {"size_mode": 160}), 400)
        self.assertEqual(act("control", "set_appearance", {"size_mode": "900"}), 400)
        self.assertEqual(act("control", "set_appearance", {"time_size": 600}), 400)
        self.assertEqual(act("submit", "set_appearance", {"lcd": True}), 403)
        self.assertEqual(act("control", "set_timing_options", {"pages_only": False, "overtime": True}), 202)
        self.assertEqual(act("control", "set_timing_options", {"pages_only": 1}), 400)
        self.assertEqual(act("control", "set_timing_options", {"end_mode": "anything"}), 400)
        self.assertEqual(act("monitor", "set_timing_options", {"overtime": True}), 403)

    def test_agenda_version_and_metadata(self):
        current = self.server.state["agenda"]
        self.server.state["agenda_revision"] = agenda_revision(current)
        new = [{"title":"Morning","uid":"test-uid","meeting":"M","room":"North","date":"2026-11-24",
                "status":"进行中","organization":"A","speaker":"B","seconds":120,"file":"slides.pdf"}]
        raw = json.dumps(new).encode()
        self.assertEqual(self.call("POST", "/api/agenda", "control", raw, {"If-Match": "old"})[0], 409)
        headers = {"If-Match": self.server.state["agenda_revision"]}
        self.assertEqual(self.call("POST", "/api/agenda", "control", raw, headers)[0], 202)
        self.assertEqual(self.call("POST", "/api/agenda", "control", raw, headers)[0], 409)
        self.assertEqual(self.actions[-1][1][0]["uid"], "test-uid")
        self.assertEqual(self.actions[-1][1][0]["room"], "North")
        for field,bad in (("date","2026-02-30"),("status","invalid"),("uid","../bad")):
            with self.subTest(field=field), self.assertRaises(RequestError):
                safe_agenda([{**new[0],field:bad}])
