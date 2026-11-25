"""ppt time meeting-file relay. Run python relay_server.py --help."""
import argparse
import hashlib
import hmac
import json
import os
import re
import secrets
import sqlite3
import ssl
import threading
import time
import unicodedata
import uuid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlsplit


MAX_FILE = 256 * 1024 * 1024
MAX_STORAGE = 4 * 1024 * 1024 * 1024
FORMATS = {".ppt", ".pptx", ".pps", ".ppsx", ".dps", ".pdf"}
WEB = """<!doctype html><html lang="zh-CN"><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>ppt time · 会前收件</title>
<style>*{box-sizing:border-box}body{margin:0;background:#f4f6f7;color:#162a32;font:16px system-ui,sans-serif}main{max-width:520px;margin:11vh auto;padding:30px}h1{font-size:36px}label{display:block;margin:24px 0 10px}button{padding:12px 20px;border:0;border-radius:8px;background:#125e57;color:white;font-weight:600;margin-top:20px}output{display:block;white-space:pre-wrap;margin-top:18px}</style>
<main><b>ppt time</b><h1>发送会议文件</h1><p>支持 PowerPoint、WPS 演示和 PDF。讲台电脑确认收件后，服务端自动删除对应文件。</p>
<div id="slotbox" hidden><label for="slot">对应议程</label><select id="slot"><option value="">由会务人员分配</option></select></div>
<label for="file">选择文件</label><input type="file" id="file" accept=".ppt,.pptx,.pps,.ppsx,.dps,.pdf" multiple><br>
<button id="send">发送到会场</button><output id="result" aria-live="polite"></output></main>
<script>const key=location.hash.substring(1);history.replaceState(null,'',location.pathname);
const msg=document.querySelector('#result'),btn=document.querySelector('#send');
const slot=document.querySelector('#slot');
if(key)fetch('/v1/agenda',{headers:{Authorization:'Bearer '+key}}).then(r=>r.ok?r.json():[]).then(rows=>{
for(const entry of rows)slot.add(new Option(entry.label,entry.id));document.querySelector('#slotbox').hidden=!rows.length}).catch(()=>{});
btn.onclick=async()=>{const files=document.querySelector('#file').files;if(!key||!files.length){msg.textContent='请使用会场收件链接，并选择会议文件';return}
btn.disabled=true;msg.textContent='';try{for(const f of files){const res=await fetch('/v1/files',{method:'POST',headers:{'Authorization':'Bearer '+key,'X-File-Name':encodeURIComponent(f.name),'X-Meeting-Slot':slot.value},body:f});if(!res.ok)throw Error(f.name+' · '+res.status);msg.textContent+=f.name+' 已提交\\n'}}catch(e){msg.textContent+='发送失败 '+e.message}finally{btn.disabled=false}};</script></html>"""


def file_name(value):
    name = unicodedata.normalize("NFC", value).strip()
    if (not name or len(name) > 150 or name.endswith((".", " ")) or
            re.search(r'[\\/<>:"|?*\x00-\x1f]', name) or
            name.split(".")[0].upper() in {"CON","AUX","PRN","NUL",*(f"COM{i}" for i in range(10)),*(f"LPT{i}" for i in range(10))} or
            Path(name).suffix.lower() not in FORMATS):
        raise ValueError("不支持的文件名或格式")
    return name


class Relay:
    def __init__(self, directory):
        self.path = Path(directory)
        self.path.mkdir(parents=True, exist_ok=True)
        self.db = sqlite3.connect(self.path / "room.db", check_same_thread=False)
        self.lock = threading.Lock()
        self.reserved = 0
        self.db.execute("CREATE TABLE IF NOT EXISTS files (id TEXT PRIMARY KEY, name TEXT, size INTEGER, sha TEXT, created REAL)")
        if "slot" not in {row[1] for row in self.db.execute("PRAGMA table_info(files)")}:
            self.db.execute("ALTER TABLE files ADD COLUMN slot TEXT DEFAULT ''")
        self.db.execute("CREATE TABLE IF NOT EXISTS agenda (id TEXT PRIMARY KEY, label TEXT)")
        self.db.execute("CREATE TABLE IF NOT EXISTS room (id INTEGER PRIMARY KEY, submit TEXT, receive TEXT)")
        self.db.execute("INSERT OR IGNORE INTO room VALUES (1,?,?)", (secrets.token_urlsafe(32), secrets.token_urlsafe(32)))
        self.db.commit()
        self.submit, self.receive = self.db.execute("SELECT submit,receive FROM room WHERE id=1").fetchone()
        if os.name == "posix":
            (self.path / "room.db").chmod(0o600)

    def reserve(self, size):
        with self.lock:
            count, used = self.db.execute("SELECT count(*), coalesce(sum(size),0) FROM files").fetchone()
            if count >= 2000 or used + self.reserved + size > MAX_STORAGE:
                return False
            self.reserved += size
            return True

    def set_agenda(self, rows):
        if not isinstance(rows, list) or len(rows) > 2000:
            raise ValueError
        cleaned = []
        for row in rows:
            if (not isinstance(row, dict) or not re.fullmatch("[a-f0-9]{32}", str(row.get("id","")))
                    or not isinstance(row.get("label"), str) or not 1 <= len(row["label"]) <= 320):
                raise ValueError
            cleaned.append((row["id"],row["label"]))
        if len({key for key, _ in cleaned}) != len(cleaned):
            raise ValueError
        with self.lock, self.db:
            self.db.execute("DELETE FROM agenda")
            self.db.executemany("INSERT INTO agenda VALUES (?,?)", cleaned)

    def public_agenda(self):
        with self.lock:
            return [{"id":row[0],"label":row[1]} for row in self.db.execute("SELECT id,label FROM agenda ORDER BY rowid")]

    def accept(self, stream, name, size, slot=""):
        name = file_name(name)
        if not 0 < size <= MAX_FILE:
            raise ValueError("无效的文件大小")
        if slot:
            with self.lock:
                if not self.db.execute("SELECT 1 FROM agenda WHERE id=?", (slot,)).fetchone():
                    raise ValueError("议程已调整")
        if not self.reserve(size):
            raise OverflowError
        key = uuid.uuid4().hex
        target = self.path / key
        try:
            digest = hashlib.sha256()
            with target.open("xb") as output:
                left = size
                while left:
                    block = stream.read(min(left, 1024 * 1024))
                    if not block:
                        raise ValueError("传输中断")
                    output.write(block)
                    digest.update(block)
                    left -= len(block)
            with self.lock, self.db:
                self.db.execute("INSERT INTO files VALUES (?,?,?,?,?,?)", (key, name, size, digest.hexdigest(), time.time(), slot))
            return key
        except BaseException:
            target.unlink(missing_ok=True)
            raise
        finally:
            with self.lock:
                self.reserved -= size

    def pending(self):
        with self.lock:
            rows = self.db.execute("SELECT id,name,size,sha,slot FROM files ORDER BY created LIMIT 50").fetchall()
        return [dict(zip(("id","name","size","sha256","slot"), row)) for row in rows]

    def ack(self, key, sha):
        with self.lock, self.db:
            row = self.db.execute("SELECT sha FROM files WHERE id=?", (key,)).fetchone()
            if not row or not hmac.compare_digest(row[0], sha):
                return False
            (self.path / key).unlink(missing_ok=True)
            self.db.execute("DELETE FROM files WHERE id=?", (key,))
            return True


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *_):
        pass

    def setup(self):
        super().setup()
        self.connection.settimeout(45)

    def allowed(self, token):
        supplied = self.headers.get("Authorization", "").encode()
        return hmac.compare_digest(supplied, ("Bearer " + token).encode())

    def reply(self, status, data=None):
        payload = json.dumps(data, ensure_ascii=False).encode() if data is not None else b""
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def do_GET(self):
        store = self.server.store
        path = urlsplit(self.path).path
        if path == "/":
            page = WEB.encode()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Referrer-Policy", "no-referrer")
            self.send_header("Cache-Control", "no-store")
            self.send_header("Content-Length", str(len(page)))
            self.end_headers()
            self.wfile.write(page)
            return
        if path == "/v1/agenda" and (self.allowed(store.submit) or self.allowed(store.receive)):
            return self.reply(200,store.public_agenda())
        if not self.allowed(store.receive):
            return self.reply(403)
        if path == "/v1/pending":
            return self.reply(200, store.pending())
        if path == "/v1/pair":
            return self.reply(200, {"submit":store.submit})
        match = re.fullmatch(r"/v1/files/([a-f0-9]{32})", path)
        if not match:
            return self.reply(404)
        with store.lock:
            key = match[1]
            row = store.db.execute("SELECT size FROM files WHERE id=?", (key,)).fetchone()
            if row is None or not (store.path / key).is_file():
                return self.reply(404)
            file = (store.path / key).open("rb")
        with file:
            self.send_response(200)
            self.send_header("Content-Type","application/octet-stream")
            self.send_header("Content-Length",str(row[0]))
            self.send_header("Cache-Control","no-store")
            self.end_headers()
            while block := file.read(1024 * 1024):
                self.wfile.write(block)

    def do_POST(self):
        store = self.server.store
        path = urlsplit(self.path).path
        if path == "/v1/files":
            if not self.allowed(store.submit):
                return self.reply(403)
            try:
                size = int(self.headers.get("Content-Length", "0"))
                if not 0 < size <= MAX_FILE:
                    return self.reply(413)
                key = store.accept(self.rfile, unquote(self.headers.get("X-File-Name","")), size,
                                   self.headers.get("X-Meeting-Slot",""))
            except (ValueError, TimeoutError):
                return self.reply(400)
            except OverflowError:
                return self.reply(507)
            return self.reply(201, {"id": key})
        if path == "/v1/agenda" and self.allowed(store.receive):
            try:
                n = int(self.headers.get("Content-Length", "0"))
                if not 1 <= n <= 1024 * 1024:
                    return self.reply(413)
                store.set_agenda(json.loads(self.rfile.read(n)))
            except (ValueError, UnicodeError):
                return self.reply(400)
            return self.reply(204)
        match = re.fullmatch(r"/v1/ack/([a-f0-9]{32})", path)
        if match and self.allowed(store.receive):
            sha = self.headers.get("X-SHA256", "")
            if re.fullmatch(r"[a-f0-9]{64}",sha) and store.ack(match[1], sha):
                return self.reply(204)
            return self.reply(404)
        self.reply(403)


def main():
    parser = argparse.ArgumentParser(description="ppt time 会议文件中转服务")
    parser.add_argument("--directory", required=True)
    parser.add_argument("--bind", default="127.0.0.1")
    parser.add_argument("--port", default=8766, type=int)
    parser.add_argument("--cert")
    parser.add_argument("--key")
    parser.add_argument("--show-receiver-key", action="store_true", help="在服务端终端显示讲台接收密钥")
    args = parser.parse_args()
    relay = Relay(args.directory)
    if args.show_receiver_key:
        print(relay.receive)
        return
    if args.bind not in ("localhost", "127.0.0.1") and not (args.cert and args.key):
        parser.error("非本机监听需配置 --cert 和 --key，或在本机监听前配置 HTTPS 反向代理")
    with ThreadingHTTPServer((args.bind, args.port), Handler) as server:
        server.store = relay
        if args.cert or args.key:
            ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
            ctx.load_cert_chain(args.cert,args.key)
            server.socket = ctx.wrap_socket(server.socket, server_side=True)
        try:
            server.serve_forever()
        except KeyboardInterrupt:
            pass
        finally:
            relay.db.close()


if __name__ == "__main__":
    main()
