"""Opt-in, paired meeting desk for a trusted local network."""

import hashlib
import hmac
import http.client
import io
import ipaddress
import json
import os
import re
import secrets
import shutil
import socket
import stat
import tempfile
import threading
import time
import unicodedata
import zipfile
from datetime import date
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, urlsplit

from PySide6.QtCore import QBuffer, QIODevice, QObject, QSettings, QStandardPaths, Qt, QTimer, Signal, Slot
from PySide6.QtGui import QGuiApplication, QPixmap
from PySide6.QtNetwork import QNetworkInterface
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QDialog, QFileDialog, QFormLayout,
    QHBoxLayout, QLabel, QLineEdit, QMessageBox, QPushButton, QSpinBox, QVBoxLayout,
)


FORMATS = {".ppt", ".pps", ".pptx", ".ppsx", ".dps", ".pdf", ".mp4", ".mov", ".m4v",
           ".webm", ".mkv", ".avi", ".wmv", ".mpg", ".mpeg"}
MAX_FILE = 512 * 1024 * 1024
MAX_TOTAL = 8 * 1024 * 1024 * 1024
MAX_FILES = 2000
LOCAL_NETS = tuple(ipaddress.ip_network(n) for n in
                   ("10.0.0.0/8", "172.16.0.0/12", "192.168.0.0/16", "127.0.0.0/8", "169.254.0.0/16"))
ROLES = ("control", "monitor", "submit", "obs")
SIMPLE_ACTIONS = {"next", "previous", "stop", "first", "last", "start_show", "home", "interlude_next", "interlude_previous",
                  "board", "board_previous", "board_next", "hide", "toggle_timer", "reset_timer"}
DESKTOP_ACTIONS = {"desk_show", "desk_files", "desk_tasks", "desk_lock", "desk_close", "desk_switch", "desk_project"}
ROLE_NAMES = {"control": "控制", "monitor": "监看", "submit": "投稿", "obs": "OBS 投影画面"}


def agenda_revision(rows):
    raw = json.dumps(rows, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode()
    return hashlib.sha256(raw).hexdigest()


def private_address(value):
    try:
        address = ipaddress.ip_address(value)
        return address.version == 4 and any(address in network for network in LOCAL_NETS)
    except ValueError:
        return False


def safe_agenda(rows):
    if not isinstance(rows, list) or len(rows) > 2000:
        raise RequestError(400, "每次可同步至多 2000 项日程")
    clean = []
    seen = set()
    for row in rows:
        if not isinstance(row, dict):
            raise RequestError(400, "日程格式不正确")
        record = {}
        for key in ("title", "speaker", "start", "file", "uid", "date", "room", "meeting", "organization", "status"):
            value = row.get(key, "")
            if not isinstance(value, str) or len(value) > 1024 or any(unicodedata.category(c) == "Cc" for c in value):
                raise RequestError(400, "日程字段无效")
            if not value and key not in {"title", "speaker", "start", "file"}:
                continue
            if key == "uid":
                if not re.fullmatch(r"[\w-]{1,100}", value, re.ASCII) or value in seen:
                    raise RequestError(400, "日程编号无效或重复")
                seen.add(value)
            if key == "status" and value not in {"待开始", "进行中", "已完成"}:
                raise RequestError(400, "日程状态无效")
            if key == "date":
                try:
                    if date.fromisoformat(value).isoformat() != value:
                        raise ValueError
                except ValueError:
                    raise RequestError(400, "日期格式应为 YYYY-MM-DD") from None
            record[key] = value.replace("\\", "/").split("/")[-1] if key == "file" else value
        seconds = row.get("seconds", 0)
        if type(seconds) is not int or not 0 <= seconds <= 86400:
            raise RequestError(400, "日程时长无效")
        record["seconds"] = seconds
        clean.append(record)
    return clean


def available_addresses():
    result = {a.toString() for a in QNetworkInterface.allAddresses() if private_address(a.toString())}
    result.add("127.0.0.1")
    return sorted(result, key=lambda x: (x.startswith("127."), x.startswith("169.254."), x))


class RequestError(Exception):
    def __init__(self, status, message):
        super().__init__(message)
        self.status = status


def media_name(name):
    if not isinstance(name, str) or len(name) > 512:
        raise RequestError(400, "文件名过长或为空")
    parts = name.split("/")
    reserved = {"CON", "CONIN$", "CONOUT$", "PRN", "AUX", "NUL", "CLOCK$"}
    for part in parts:
        device = unicodedata.normalize("NFKC", part.split(".")[0]).upper().rstrip()
        if (not part or part.startswith(".") or part.endswith((" ", ".")) or len(part) > 160
                or any(c in '<>:"\\|?*' or unicodedata.category(c) in {"Cc", "Cf"} for c in part)
                or device in reserved or re.fullmatch(r"(COM|LPT)[0-9]", device)):
            raise RequestError(400, "文件夹或文件名无效")
    if len(parts) > 12 or Path(parts[-1]).suffix.lower() not in FORMATS:
        raise RequestError(415, "仅支持演示文稿、PDF 和会议视频")
    return parts


def check_media(path, suffix):
    with path.open("rb") as stream:
        header = stream.read(1024)
    if suffix in {".pptx", ".ppsx"} or suffix == ".dps" and header.startswith(b"PK"):
        try:
            with zipfile.ZipFile(path) as archive:
                entries = set(archive.namelist())
                if not {"[Content_Types].xml", "ppt/presentation.xml"} <= entries:
                    raise ValueError
        except (zipfile.BadZipFile, ValueError, OSError):
            raise RequestError(415, "文件内容与演示文稿格式不符") from None
        return
    if suffix not in FORMATS:
        raise RequestError(415, "不支持的媒体格式")
    ole = header.startswith(bytes.fromhex("D0CF11E0A1B11AE1"))
    ebml = header.startswith(bytes.fromhex("1A45DFA3"))
    mpeg = header.startswith((bytes.fromhex("000001BA"), bytes.fromhex("000001B3")))
    signatures = {
        ".pdf": header.lstrip().startswith(b"%PDF-"),
        ".ppt": ole, ".pps": ole, ".dps": ole, ".webm": ebml, ".mkv": ebml,
        ".avi": header.startswith(b"RIFF") and header[8:12] == b"AVI ",
        ".wmv": header.startswith(bytes.fromhex("3026B2758E66CF11A6D900AA0062CE6C")),
        ".mpg": mpeg, ".mpeg": mpeg, ".mov": header[4:8] in {b"ftyp", b"moov", b"mdat", b"wide", b"free"},
    }
    if not signatures.get(suffix, header[4:8] == b"ftyp"):
        raise RequestError(415, "文件内容与扩展名不符")


def redirected(path):
    try:
        info = path.lstat()
    except FileNotFoundError:
        return False
    return stat.S_ISLNK(info.st_mode) or bool(getattr(info, "st_file_attributes", 0) & 0x400)


class MeetingFiles:
    """One operator-selected room inbox. Transfers become visible at commit."""
    def __init__(self, root):
        root = Path(root)
        if redirected(root):
            raise ValueError("收件箱必须是普通文件夹")
        root.mkdir(parents=True, exist_ok=True)
        self.root = root.resolve()
        self.lock = threading.Lock()
        self.reserved = 0
        self.in_flight = 0
        self.items = {}
        self.modified = {}
        self.archived = {}
        self.archive_root = self.root / ".ppt-time-archive"
        for base, dirs, files in os.walk(self.root, followlinks=False):
            base = Path(base)
            dirs[:] = [d for d in dirs if not d.startswith(".") and not redirected(base / d)]
            for filename in files:
                path = base / filename
                name = path.relative_to(self.root).as_posix()
                try:
                    media_name(name)
                    if not redirected(path) and path.is_file():
                        info = path.stat()
                        self.items[name] = info.st_size
                        self.modified[name] = int(info.st_mtime)
                except (RequestError, OSError):
                    continue
        if self.archive_root.is_dir() and not redirected(self.archive_root):
            for entry in self.archive_root.iterdir():
                if not re.fullmatch(r"[0-9a-f]{32}", entry.name) or redirected(entry) or not entry.is_dir():
                    continue
                for base, dirs, files in os.walk(entry, followlinks=False):
                    base = Path(base)
                    dirs[:] = [d for d in dirs if not redirected(base / d) and not d.startswith(".")]
                    for filename in files:
                        path = base / filename
                        name = path.relative_to(entry).as_posix()
                        try:
                            media_name(name)
                            if not redirected(path) and path.is_file():
                                info = path.stat()
                                self.archived[entry.name] = (name, info.st_size, int(info.st_mtime))
                        except (RequestError, OSError):
                            continue
        self.total = sum(self.items.values()) + sum(size for _, size, _ in self.archived.values())

    def path_for(self, name):
        path = self.root
        for part in media_name(name):
            path = path / part
            if redirected(path):
                raise RequestError(400, "收件箱中的链接不可用")
        if not path.resolve().is_relative_to(self.root):
            raise RequestError(400, "无效的文件位置")
        return path

    def listing(self):
        return self.catalog()["files"]

    def catalog(self):
        with self.lock:
            return {
                "files": [{"name": name, "size": size, "modified": self.modified[name]}
                          for name, size in sorted(self.items.items())],
                "archived": [{"id": token, "name": name, "size": size, "modified": modified}
                             for token, (name, size, modified) in self.archived.items()],
            }

    def _archived_path(self, token, name):
        if not isinstance(token, str) or not re.fullmatch(r"[0-9a-f]{32}", token):
            raise RequestError(400, "已移除文件编号无效")
        path = self.root
        for part in (self.archive_root.name, token, *media_name(name)):
            path = path / part
            if redirected(path):
                raise RequestError(400, "收件箱中的链接不可用")
        if not path.resolve().is_relative_to(self.root):
            raise RequestError(400, "无效的文件位置")
        return path

    def archive(self, name):
        if not isinstance(name, str):
            raise RequestError(400, "请选择收件箱内的文件")
        with self.lock:
            source = self.path_for(name)
            if name not in self.items or not source.is_file():
                raise RequestError(404, "文件已移走")
            token = secrets.token_hex(16)
            target = self._archived_path(token, name)
            self.archive_root.mkdir(exist_ok=True)
            (self.archive_root / token).mkdir()
            target.parent.mkdir(parents=True, exist_ok=True)
            try:
                source.rename(target)
            except OSError:
                self._clean_archive_folders(target.parent)
                raise
            self.archived[token] = (name, self.items.pop(name), self.modified.pop(name))
            return source

    def _clean_archive_folders(self, parent):
        while parent != self.archive_root:
            try:
                parent.rmdir()
            except OSError:
                break
            parent = parent.parent

    def restore(self, token):
        if not isinstance(token, str):
            raise RequestError(400, "请选择需要还原的文件")
        with self.lock:
            if token not in self.archived:
                raise RequestError(404, "该文件已还原或移走")
            name, size, modified = self.archived[token]
            source = self._archived_path(token, name)
            if not source.is_file():
                raise RequestError(404, "已移除文件已在本机移走")
            target = self.path_for(name)
            target.parent.mkdir(parents=True, exist_ok=True)
            final = self._commit_file(source, target)
            result = final.relative_to(self.root).as_posix()
            self.items[result] = size
            self.modified[result] = modified
            del self.archived[token]
            self._clean_archive_folders(source.parent)
            return result

    @staticmethod
    def _commit_file(source, target):
        number = 1
        while True:
            final = target if number == 1 else target.with_name(f"{target.stem} ({number}){target.suffix}")
            try:
                fd = os.open(final, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
            except FileExistsError:
                number += 1
                continue
            os.close(fd)
            try:
                os.replace(source, final)
            except OSError:
                final.unlink(missing_ok=True)
                raise
            return final

    def for_open(self, name):
        with self.lock:
            if name not in self.items:
                raise RequestError(404, "文件不在会场收件箱")
            path = self.path_for(name)
        if not path.is_file():
            raise RequestError(404, "文件已移走")
        check_media(path, path.suffix.lower())
        return path

    def receive(self, name, length, read, active):
        target = self.path_for(name)
        if not 0 < length <= MAX_FILE:
            raise RequestError(413, "单个文件须大于零且不超过 512 MB")
        with self.lock:
            if self.in_flight >= 2:
                raise RequestError(429, "正在接收其他文件，请稍后重试")
            if (self.total + self.reserved + length > MAX_TOTAL or len(self.items) + len(self.archived) + self.in_flight >= MAX_FILES
                    or shutil.disk_usage(self.root).free < self.reserved + length + 16 * 1024 * 1024):
                raise RequestError(507, "会场收件箱容量不足，请在讲台端整理文件")
            self.reserved += length
            self.in_flight += 1
        temp = None
        try:
            with tempfile.NamedTemporaryFile(dir=self.root, prefix=".ppt-time-", delete=False) as output:
                temp = Path(output.name)
                remaining = length
                while remaining > 0:
                    if not active.is_set():
                        raise RequestError(409, "会场已停止接收")
                    chunk = read(min(256 * 1024, remaining))
                    if not chunk:
                        raise RequestError(400, "文件传输中断")
                    output.write(chunk)
                    remaining -= len(chunk)
            check_media(temp, target.suffix.lower())
            with self.lock:
                if not active.is_set():
                    raise RequestError(409, "会场已停止接收")
                target = self.path_for(name)
                target.parent.mkdir(parents=True, exist_ok=True)
                final = self._commit_file(temp, target)
                result = final.relative_to(self.root).as_posix()
                self.items[result] = length
                self.modified[result] = int(time.time())
                self.total += length
                return result
        finally:
            if temp is not None:
                temp.unlink(missing_ok=True)
            with self.lock:
                self.in_flight -= 1
                self.reserved -= length


class MeetingServer(ThreadingHTTPServer):
    daemon_threads = True
    request_queue_size = 12
    allow_reuse_address = False

    def __init__(self, host, port, files, dispatch, received=None):
        if not private_address(host):
            raise ValueError("请选择本机的局域网 IPv4 地址")
        self.gate = threading.BoundedSemaphore(16)
        self.active = threading.Event()
        self.active.set()
        self.files = files
        self.dispatch = dispatch
        self.received = received
        self.tokens = {role: secrets.token_urlsafe(32) for role in ROLES}
        self.state_lock = threading.Lock()
        self.state = {}
        self.preview = b""
        self.preview_allowed = False
        self.preview_requested_at = 0.0
        self.agenda_pending_until = 0.0
        self.agenda_expected = ""
        super().__init__((host, port), MeetingHandler)

    def verify_request(self, request, client_address):
        return private_address(client_address[0]) and self.active.is_set()

    def process_request(self, request, client_address):
        if not self.gate.acquire(blocking=False):
            self.shutdown_request(request)
            return
        try:
            super().process_request(request, client_address)
        except BaseException:
            self.gate.release()
            raise

    def process_request_thread(self, request, client_address):
        try:
            super().process_request_thread(request, client_address)
        finally:
            self.gate.release()

    def link(self, role):
        host, port = self.server_address
        page = "obs" if role == "obs" else ""
        return f"http://{host}:{port}/{page}#{self.tokens[role]}"


class MeetingHandler(BaseHTTPRequestHandler):
    server_version = "ppt-time"
    sys_version = ""

    def setup(self):
        super().setup()
        self.connection.settimeout(15)

    def log_message(self, *args):
        pass

    def read_body(self, size):
        data = self.rfile.read(size)
        self.body_read += len(data)
        return data

    def finish_error(self, status, message):
        self.reply(status, {"error": message})
        # Closing a socket with an unread small POST body can reset the connection
        # on Windows and discard the response before the browser displays it.
        if self.command == "POST":
            try:
                self.connection.shutdown(socket.SHUT_WR)
                remaining = int(self.headers.get("Content-Length", 0)) - self.body_read
                if 0 < remaining <= 65536:
                    self.connection.settimeout(0.25)
                    self.rfile.read(remaining)
            except (OSError, ValueError):
                pass

    def send_data(self, code, body, mime):
        self.send_response(code)
        self.send_header("Content-Type", mime)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("Referrer-Policy", "no-referrer")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("X-Frame-Options", "DENY")
        self.send_header("Content-Security-Policy", "default-src 'none'; script-src 'self'; style-src 'self'; img-src 'self' blob:; connect-src 'self'; frame-ancestors 'none'; base-uri 'none'; form-action 'none'")
        self.end_headers()
        self.wfile.write(body)

    def reply(self, code, data):
        self.send_data(code, json.dumps(data, ensure_ascii=False, allow_nan=False).encode(), "application/json; charset=utf-8")

    def role(self):
        value = self.headers.get("Authorization", "")
        candidate = value[7:] if value.startswith("Bearer ") else ""
        for role, token in self.server.tokens.items():
            if candidate.isascii() and hmac.compare_digest(candidate, token):
                return role
        raise RequestError(401, "请使用讲台端分享的本场连接")

    def serve(self):
        server = self.server
        endpoint = "%s:%s" % server.server_address
        if self.headers.get("Host") != endpoint or self.headers.get("Origin", "http://" + endpoint) != "http://" + endpoint:
            raise RequestError(403, "仅接受本场工作台连接")
        if not server.active.is_set():
            raise RequestError(409, "本场工作台已经关闭")
        route = urlsplit(self.path)
        if self.command == "GET" and route.path in ASSETS:
            mime, body = ASSETS[route.path]
            return self.send_data(200, body, mime)
        role = self.role()
        if self.command == "GET" and route.path == "/api/state":
            with server.state_lock:
                if role == "obs":
                    state = {k: server.state.get(k, "") for k in ("room", "time", "page", "stage")}
                elif role == "submit":
                    state = {"room": server.state.get("room", "")}
                else:
                    state = dict(server.state)
                    if role != "control":
                        state.pop("desktop_controls", None)
                        state.pop("appearance", None)
                state["preview_allowed"] = role != "submit" and server.preview_allowed
            return self.reply(200, {"role": role, "state": state})
        if self.command == "GET" and route.path == "/api/preview":
            with server.state_lock:
                if role == "submit" or not server.preview_allowed:
                    raise RequestError(403, "讲台未共享投影画面")
                server.preview_requested_at = time.monotonic()
                image = server.preview
            return self.send_data(200 if image else 204, image, "image/jpeg")
        if self.command == "GET" and route.path == "/api/files" and role == "control":
            return self.reply(200, server.files.catalog())
        if self.command == "POST" and route.path in {"/api/files/archive", "/api/files/restore"} and role == "control":
            try:
                length = int(self.headers.get("Content-Length", ""))
                if not 0 < length <= 4096:
                    raise ValueError
                value = json.loads(self.read_body(length))
                if not isinstance(value, dict):
                    raise ValueError
            except (ValueError, UnicodeError):
                raise RequestError(400, "文件操作格式无效") from None
            if route.path.endswith("/archive"):
                removed = server.files.archive(value.get("name"))
                server.dispatch("inbox_archived", str(removed))
                return self.reply(200, {"name": value["name"]})
            restored = server.files.restore(value.get("id"))
            if server.received:
                server.received(str(server.files.path_for(restored)))
            return self.reply(200, {"name": restored})
        if self.command == "POST" and route.path == "/api/upload" and role in {"submit", "control"}:
            try:
                length = int(self.headers.get("Content-Length", ""))
            except ValueError:
                raise RequestError(411, "上传需要文件大小") from None
            name = parse_qs(route.query, max_num_fields=4).get("name", [""])[0]
            saved = server.files.receive(name, length, self.read_body, server.active)
            if server.received:
                server.received(str(server.files.path_for(saved)))
            return self.reply(201, {"name": saved})
        if self.command == "POST" and route.path == "/api/agenda" and role == "control":
            try:
                length = int(self.headers.get("Content-Length", ""))
                if not 0 < length <= 4 * 1024 * 1024:
                    raise ValueError
                agenda = safe_agenda(json.loads(self.read_body(length)))
            except (ValueError, UnicodeError):
                raise RequestError(400, "日程同步内容无效") from None
            expected = self.headers.get("If-Match")
            if expected:
                with server.state_lock:
                    if (time.monotonic() < server.agenda_pending_until or
                            expected.strip('"') != server.state.get("agenda_revision", agenda_revision([]))):
                        raise RequestError(409, "日程已发生变化，请重新读取")
                    server.agenda_expected = agenda_revision(agenda)
                    server.agenda_pending_until = time.monotonic() + 5
            server.dispatch("import_agenda", agenda)
            return self.reply(202, {"accepted": True})
        if self.command == "POST" and route.path == "/api/action" and role == "control":
            try:
                length = int(self.headers.get("Content-Length", ""))
                if not 0 < length <= 4096:
                    raise ValueError
                data = json.loads(self.read_body(length))
                if not isinstance(data, dict) or not isinstance(data.get("action"), str):
                    raise ValueError
            except (ValueError, UnicodeError):
                raise RequestError(400, "控制请求格式无效") from None
            action, value = data["action"], data.get("value")
            if action == "open" and isinstance(value, str):
                value = str(server.files.for_open(value))
            elif action == "set_seconds" and type(value) is int and 1 <= value <= 359999:
                pass
            elif action in {"interstitial", "agenda", "file"} and type(value) is int and 0 <= value < MAX_FILES:
                pass
            elif action == "set_appearance" and isinstance(value, dict) and value:
                from overlay_options import OverlayOptions
                try:
                    OverlayOptions.from_mapping(value, strict=True)
                except ValueError as exc:
                    raise RequestError(400, str(exc)) from None
            elif action == "set_timing_options" and isinstance(value, dict) and value:
                if value.keys() - {"pages_only", "overtime"} or any(type(v) is not bool for v in value.values()):
                    raise RequestError(400, "计时模式设置无效")
            elif action in DESKTOP_ACTIONS and value is None:
                with server.state_lock:
                    if server.state.get("desktop_controls") is not True:
                        raise RequestError(403, "讲台端未开启桌面控制")
            elif action in SIMPLE_ACTIONS and value is None:
                pass
            else:
                raise RequestError(400, "不支持的会场操作")
            server.dispatch(action, value)
            return self.reply(202, {"accepted": True})
        if route.path in {"/api/files", "/api/files/archive", "/api/files/restore", "/api/action", "/api/upload", "/api/agenda"}:
            raise RequestError(403, "此连接没有该操作权限")
        raise RequestError(404, "没有此页面")

    def handle_request(self):
        self.body_read = 0
        try:
            self.serve()
        except RequestError as err:
            self.finish_error(err.status, str(err))
        except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError, socket.timeout):
            return
        except (OSError, ValueError):
            self.finish_error(400, "操作未完成，请检查文件和收件箱")

    do_GET = handle_request
    do_POST = handle_request


HTML = """<!doctype html><html lang="zh-CN"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>ppt time · 会场工作台</title>
<link rel="stylesheet" href="/desk.css"><script defer src="/desk.js"></script></head><body>
<header><strong>ppt time</strong><span id="room">会场工作台</span><span id="role"></span></header>
<main>
<section id="login" hidden><h1>连接会场</h1><p>粘贴讲台端提供的连接或配对密钥。</p>
<form id="pair"><input id="key" type="password" autocomplete="off" aria-label="配对连接或密钥" required>
<button>连接</button></form></section>
<div id="workspace" hidden>
<section id="live" draggable="true"><h2>放映监看</h2><div class="clock"><strong id="time">--:--</strong>
<span id="page"></span></div><div id="stage"></div><div id="scene"></div><p id="hardware"></p><img id="preview" alt="讲台共享的投影画面" hidden></section>
<section id="controls" draggable="true"><h2>放映控制</h2>
<div class="actions"><button data-act="start_show">开始放映</button><button data-act="first">首页</button>
<button data-act="previous">上一页</button><button data-act="next" class="primary">下一页</button><button data-act="last">末页</button>
<button data-act="stop">结束放映</button></div>
<div class="actions"><button data-act="toggle_timer">开始 / 暂停计时</button><button data-act="reset_timer">重置计时</button></div>
<form id="duration"><label>倒计时 <input id="minutes" type="number" value="5" step="1" min="0" max="5999" required> 分
<input id="seconds" type="number" value="0" step="1" min="0" max="59" required> 秒</label><button>设置</button></form>
<div class="actions" aria-label="快捷时长"><button data-min="5">5 分钟</button><button data-min="10">10 分钟</button>
<button data-min="15">15 分钟</button><button data-min="20">20 分钟</button><button data-min="30">30 分钟</button></div>
<div class="actions" id="timing-options" hidden>
<label><input id="pages-only" type="checkbox" data-opt="pages_only">仅显示页码</label>
<label><input id="overtime" type="checkbox" data-opt="overtime">超时后正计时</label></div>
<div class="actions"><button data-act="home">大会主画面</button><button data-act="board">日程看板</button>
<button data-act="board_previous">日程上一屏</button><button data-act="board_next">日程下一屏</button>
<button data-act="hide">收起会场画面</button></div><div class="actions">
<button data-act="interlude_previous">上一张串场图</button><button data-act="interlude_next">下一张串场图</button></div>
<div id="cues" class="actions"></div>
<details id="desktop-tools" hidden><summary>讲台桌面控制</summary><div class="actions">
<button data-act="desk_show">显示桌面</button><button data-act="desk_files">文件管理</button>
<button data-act="desk_tasks">任务管理</button><button data-act="desk_switch">切换窗口</button>
<button data-act="desk_project">投影模式</button><button data-act="desk_close">关闭当前窗口</button>
<button data-act="desk_lock">锁定讲台</button></div></details>__APPEARANCE_HTML__</section>
<section id="agenda" draggable="true"><div class="heading"><h2>会议日程</h2><a href="/agenda">打开日程看板</a></div><ol id="sessions"></ol></section>
<section id="inbox" draggable="true"><div class="heading"><h2>会场文件</h2><div id="view">
<button id="layout">图标视图</button><label id="column-label">列数 <select id="columns">
<option>1</option><option>2</option><option>3</option><option>4</option><option>5</option></select></label></div></div>
<div class="actions"><label class="button">上传文件<input type="file" multiple id="files" class="picker"></label>
<label class="button">上传文件夹<input type="file" multiple webkitdirectory id="folder" class="picker"></label></div>
<progress id="progress" max="100" value="0" hidden></progress><p id="upload-message" role="status"></p>
<div id="file-tools" hidden><div class="actions">
<button id="file-up">上一级</button><strong id="file-path">收件箱</strong></div>
<div class="actions"><label>排序 <select id="file-sort"><option value="natural">数字顺序</option>
<option value="name">名称</option><option value="date">修改日期</option></select></label>
<button id="file-direction">升序</button><label><input id="file-edit" type="checkbox"> 整理文件</label>
<button id="file-archive">已移除</button></div></div>
<div id="library"></div><div id="desk-files" class="actions"></div></section></div><p id="feedback" role="status"></p>
</main></body></html>"""

CSS = """
:root{color-scheme:light;font-family:"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif;color:#183536;background:#f2f5f2}
*{box-sizing:border-box}[hidden]{display:none!important}body{margin:0}header{padding:16px max(5vw,20px);
background:#143f40;color:white;display:flex;align-items:center;gap:22px}header strong{font-size:22px;letter-spacing:-1px}
#role{margin-left:auto;font-size:13px}main{max-width:1120px;margin:auto;padding:22px}
#workspace{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px;align-items:start}
section{background:white;border:1px solid #d5dedb;border-radius:12px;padding:18px;min-width:0}
h1,h2{margin:0 0 12px}h2{font-size:17px}.clock{display:flex;align-items:baseline;gap:22px}
#time{font-size:clamp(42px,7vw,64px);font-variant-numeric:tabular-nums;letter-spacing:-2px}
#page{font-size:20px}#scene,#stage,#feedback,#upload-message{color:#426463}#scene,#stage{margin-top:5px}
#preview{width:100%;margin-top:12px;border-radius:5px;background:#102a2c}.actions,.heading,form,#view{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.heading{justify-content:space-between;margin-bottom:8px}.heading h2{margin:0}.actions,form{margin:10px 0}
button,.button,select,input{font:inherit;border:1px solid #c5d4d0;border-radius:7px;background:white;color:#183536;padding:7px 11px}
button,.button{cursor:pointer}button:focus-visible,.button:focus-within,input:focus{outline:2px solid #008475}
.primary{background:#126455;color:white;border-color:#126455}input[type=number]{width:66px;padding:6px}
.picker{position:absolute;width:1px;height:1px;opacity:0}#library{display:grid;gap:5px}
#library{grid-template-columns:repeat(var(--cols,1),minmax(0,1fr))}.file{display:flex;align-items:center;gap:9px;text-align:left;min-width:0}
.file span{overflow-wrap:anywhere}.file i{font-style:normal;font-weight:700;background:#e0f0e7;color:#185648;padding:6px;border-radius:5px}
.grid .file{flex-direction:column;justify-content:flex-start;align-items:flex-start}.file small{color:#57716d}
.file-entry{display:flex;align-items:stretch;gap:4px;min-width:0}.file-entry .file{flex:1;width:100%}
.file-entry .manage{padding:4px 8px;color:#814a35}.grid .file-entry{flex-direction:column}
#file-path{overflow-wrap:anywhere}#file-tools label{white-space:nowrap}
ol{padding-left:22px;margin:0}li{padding:5px 0;overflow-wrap:anywhere}progress{width:100%}
#feedback{min-height:1.5em}#feedback.error,#upload-message.error{color:#9c331f}
@media(max-width:730px){#workspace{grid-template-columns:1fr}main{padding:12px}section{padding:14px}
header{padding:12px 16px;gap:12px}.actions button,.actions .button{flex-grow:1;text-align:center}}
"""

JS = r"""
'use strict';
const el = id => document.getElementById(id);
const put = (id, value) => { el(id).textContent = value || ''; };
const getPref = (key, def) => { try { return localStorage.getItem(key) || def; } catch { return def; } };
const setPref = (key, val) => { try { localStorage.setItem(key, val); } catch {} };
let token = location.hash.slice(1);
try { token = token || sessionStorage.getItem('meeting-link') || ''; } catch {}
if (location.hash) history.replaceState(null, '', location.pathname);
let role = '', rows = [], grid = getPref('view','list') === 'grid', previousImage = '', busy = false;
let lastFiles = 0, lastRoomResult = '', lastSchedule = '', lastButtons='', dragged = null;
let fileFolder = '', archivedFiles = [], archiveView = false;
const fileNatural = new Intl.Collator('zh-CN', {numeric:true,sensitivity:'base'});
const fileAlpha = new Intl.Collator('zh-CN', {numeric:false,sensitivity:'base'});
let sortDescending = getPref('file-direction','asc') === 'desc';
function showMessage(msg, bad=false) {
  put('feedback', msg); el('feedback').classList.toggle('error',bad);
}
async function api(path, options={}) {
  const response = await fetch(path, {...options, headers:{...options.headers, Authorization:'Bearer '+token},cache:'no-store'});
  const data = await response.json();
  if (!response.ok) { if(response.status===401){el('login').hidden=false;el('workspace').hidden=true;role='';}
    throw Error(data.error || '网络连接失败'); }
  return data;
}
function assignRole(value) {
  role=value; put('role', {control:'控制端',monitor:'监看端',submit:'投稿端'}[role]);
  el('workspace').hidden=false;el('login').hidden=true;el('controls').hidden=role!=='control';
  el('live').hidden=el('agenda').hidden=role==='submit'; el('inbox').hidden=role==='monitor';
  el('view').hidden=el('library').hidden=role!=='control';
  el('file-tools').hidden=role!=='control';
  try {sessionStorage.setItem('meeting-link',token);} catch {}
}
function renderFiles() {
  const box = el('library');
  box.replaceChildren();
  box.classList.toggle('grid',grid);
  box.style.setProperty('--cols',el('columns').value);
  el('file-up').disabled = archiveView || !fileFolder;
  el('file-edit').disabled = archiveView;
  put('layout',grid?'列表视图':'图标视图');
  put('file-path',archiveView ? '已移除文件' : '收件箱' + (fileFolder ? ' / ' + fileFolder : ''));
  put('file-archive',archiveView ? '返回收件箱' : '已移除 ('+archivedFiles.length+')');
  put('file-direction',sortDescending ? '降序' : '升序');
  const entries = new Map();
  const prefix = fileFolder ? fileFolder+'/' : '';
  for(const row of archiveView ? archivedFiles : rows) {
    if(!archiveView && !row.name.startsWith(prefix))continue;
    const relative = archiveView ? row.name : row.name.slice(prefix.length);
    const slash = archiveView ? -1 : relative.indexOf('/');
    if(slash < 0){
      entries.set(archiveView ? row.id : row.name, {...row,label:relative});
      continue;
    }
    const folder = relative.slice(0,slash);
    const key = prefix+folder+'/';
    const entry = entries.get(key) || {folder:true,label:folder,name:prefix+folder,count:0,modified:0};
    entry.count++;
    entry.modified=Math.max(entry.modified,row.modified || 0);
    entries.set(key,entry);
  }
  const ordered = Array.from(entries.values()).sort((a,b)=>{
    if(Boolean(a.folder)!==Boolean(b.folder))return a.folder ? -1 : 1;
    const mode=el('file-sort').value;
    const byName=(mode==='name'?fileAlpha:fileNatural).compare(a.label,b.label);
    const difference=mode==='date' ? (a.modified-b.modified || byName) : byName;
    return sortDescending ? -difference : difference;
  });
  if(!ordered.length){box.textContent=archiveView?'没有已移除文件':'此目录暂无文件';return;}
  for(const row of ordered) {
    const wrap=document.createElement('div');wrap.className='file-entry';
    const b=document.createElement('button');b.className='file';
    b.title=(archiveView?'还原 ':row.folder?'进入 ':'打开 ')+row.name;
    if(grid){
      const icon=document.createElement('i');
      icon.textContent=row.folder ? '目录' : Array.from(row.label.split('/').pop())[0];
      b.append(icon);
    }
    const span=document.createElement('span');span.textContent=row.label;
    const details=document.createElement('small');
    details.textContent = row.folder ? row.count+' 个文件' : (row.size/1048576).toFixed(1)+' MB';
    if(el('file-sort').value==='date' && row.modified)
      details.textContent+=' · '+new Date(row.modified*1000).toLocaleDateString();
    b.append(span,details);
    b.onclick=()=>{
      if(archiveView){changeFile('restore',row);return;}
      if(row.folder){fileFolder=row.name;renderFiles();return;}
      command('open',row.name);
    };
    wrap.append(b);
    if(!row.folder && !archiveView && el('file-edit').checked){
      const remove=document.createElement('button');
      remove.className='manage';remove.textContent='移除';remove.title='移除 '+row.name;
      remove.onclick=()=>changeFile('archive',row);wrap.append(remove);
    }
    box.append(wrap);
  }
}
async function changeFile(action,row){
  if(action==='archive' && !confirm('从收件箱移除“'+row.name+'”？稍后可在“已移除”中还原。'))return;
  try{
    const field=action==='archive'?{name:row.name}:{id:row.id};
    const result=await api('/api/files/'+action,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(field)});
    await loadFiles();
    showMessage((action==='archive'?'已移除 ':'已还原 ')+result.name);
  }catch(e){showMessage(e.message,true);}
}
async function loadFiles() {
  if(role!=='control')return;
  const result=await api('/api/files');
  if(JSON.stringify(rows)!==JSON.stringify(result.files) || JSON.stringify(archivedFiles)!==JSON.stringify(result.archived) || !el('library').childNodes.length) {
    rows=result.files;archivedFiles=result.archived || [];renderFiles();
  }
  lastFiles=Date.now();
}
async function preview(allowed) {
  if(!allowed) {el('preview').hidden=true;if(previousImage){URL.revokeObjectURL(previousImage);previousImage='';}return;}
  const res=await fetch('/api/preview',{headers:{Authorization:'Bearer '+token},cache:'no-store'});
  if(res.status!==200){el('preview').hidden=true;return;}
  const url=URL.createObjectURL(await res.blob());el('preview').src=url;el('preview').hidden=false;
  if(previousImage)URL.revokeObjectURL(previousImage);previousImage=url;
}
async function refresh() {
  if(busy || !token || document.hidden)return;busy=true;
  try {
    const start=performance.now();const {role: nextRole, state}=await api('/api/state');if(role!==nextRole)assignRole(nextRole);
    if(role==='obs'){location.replace('/obs#'+token);return;}
    const sys=state.system||{}, stats=[];
    for(const [key,title] of [['cpu_percent','CPU'],['memory_percent','内存'],['battery_percent','电量']]){
      if(typeof sys[key]==='number')stats.push(title+' '+sys[key]+'%');}
    if(sys.power)stats.push(sys.power);if(sys.graphics)stats.push(sys.graphics);
    stats.push('往返 '+Math.round(performance.now()-start)+' ms');put('hardware',stats.join(' · '));
    el('desktop-tools').hidden=role!=='control'||!state.desktop_controls;
    el('timing-options').hidden=role!=='control'||!state.timing_options;
    if(role==='control' && state.timing_options){
      for(const box of document.querySelectorAll('[data-opt]')){
        box.checked=state.timing_options[box.dataset.opt]===true;
      }
    }
    if(role==='control' && typeof loadAppearance==='function')loadAppearance(state.appearance);
    put('room',state.room || '会场工作台');put('time',state.time || '--:--');put('stage',state.stage);
    put('page',state.page);put('scene',state.scene);
    const agenda=Array.isArray(state.schedule)?state.schedule:[];const representation=role+JSON.stringify(agenda);
    if(lastSchedule!==representation) {const list=el('sessions');list.replaceChildren();
      for(const [i,entry] of agenda.entries()) {const li=document.createElement('li');
        const title=typeof entry==='string'?entry:[entry.start || entry.time,entry.title,entry.speaker].filter(Boolean).join(' · ');
        if(role==='control'){const b=document.createElement('button');b.textContent=title||'打开日程';b.onclick=()=>command('agenda',i);li.append(b);}
        else li.textContent=title;list.append(li);}
      if(!agenda.length)list.textContent='暂无日程';lastSchedule=representation;
    }
    const buttons=JSON.stringify([role,state.interstitials,state.files]);
    if(lastButtons!==buttons){for(const [id,items,act] of [['cues',state.interstitials,'interstitial'],['desk-files',state.files,'file']]){
      const box=el(id);box.replaceChildren();if(role==='control' && Array.isArray(items)){
        for(const [i,title] of items.entries()){const b=document.createElement('button');b.textContent=title;
          b.onclick=()=>command(act,i);box.append(b);}}}lastButtons=buttons;}
    if(state.result && state.result!==lastRoomResult){lastRoomResult=state.result;showMessage(state.result);}
    if(Date.now()-lastFiles>5000)await loadFiles();if(role!=='submit')await preview(state.preview_allowed);
  }catch(e) {showMessage(e.message,true);} finally {busy=false;}
}
async function command(action,value=null) {
  try{await api('/api/action',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({action,value})});
    showMessage('指令已送达讲台');}catch(e){showMessage(e.message,true);}
}
function sendFile(file, name) {
  return new Promise((resolve,reject)=>{
    const x=new XMLHttpRequest();x.open('POST','/api/upload?name='+encodeURIComponent(name));
    x.setRequestHeader('Authorization','Bearer '+token);
    x.upload.onprogress=e=>{if(e.lengthComputable)el('progress').value=100*e.loaded/e.total;};
    x.onerror=()=>reject(Error('网络中断'));x.onload=()=>{
      try{const result=JSON.parse(x.responseText);if(x.status!==201)throw Error(result.error || '上传失败');resolve(result.name);}
      catch(err){reject(err);}
    };x.send(file);
  });
}
let uploading=false;
async function submitFiles(input) {
  if(uploading)return;uploading=true;
  const files=Array.from(input.files).filter(f=>el('files').accept.split(',').includes('.'+f.name.toLowerCase().split('.').pop()));
  const skipped=input.files.length-files.length, failures=[];let done=0;
  el('files').disabled=el('folder').disabled=true;el('progress').hidden=false;
  try {
    for(const f of files) {const name=f.webkitRelativePath||f.name;el('progress').value=0;
      put('upload-message','正在上传 '+name);try{await sendFile(f,name);done++;}catch(e){failures.push(name+'（'+e.message+'）');}}
    put('upload-message','已接收 '+done+' 个文件'+(skipped?'，已跳过 '+skipped+' 个其他格式':'')+
      (failures.length?'，失败 '+failures.join('、'):''));
    el('upload-message').classList.toggle('error',failures.length>0);
    await loadFiles();
  }catch(e){showMessage(e.message,true);}finally{input.value='';uploading=false;el('progress').hidden=true;
    el('files').disabled=el('folder').disabled=false;}
}
el('files').accept='__MEDIA_FORMATS__';
for(const id of ['folder','files'])el(id).addEventListener('change',()=>submitFiles(el(id)));
for(const b of document.querySelectorAll('[data-act]'))b.addEventListener('click',()=>{
  if(['desk_lock','desk_close'].includes(b.dataset.act) && !confirm('确认'+b.textContent+'？'))return;
  command(b.dataset.act);
});
for(const b of document.querySelectorAll('[data-min]'))b.onclick=()=>{el('minutes').value=b.dataset.min;el('seconds').value=0;
  command('set_seconds',Number(b.dataset.min)*60);};
for(const box of document.querySelectorAll('[data-opt]'))box.onchange=()=>command('set_timing_options',{[box.dataset.opt]:box.checked});
function restoreColumns(){
  const fallback=grid?'3':'1';
  const preferred=getPref(grid?'columns':'list-columns',fallback);
  el('columns').value=['1','2','3','4','5'].includes(preferred)?preferred:fallback;
}
el('layout').onclick=()=>{grid=!grid;setPref('view',grid?'grid':'list');restoreColumns();renderFiles();};
el('file-up').onclick=()=>{fileFolder=fileFolder.split('/').slice(0,-1).join('/');renderFiles();};
el('file-sort').value=['name','date'].includes(getPref('file-sort','natural'))?getPref('file-sort','natural'):'natural';
el('file-sort').onchange=()=>{setPref('file-sort',el('file-sort').value);renderFiles();};
el('file-direction').onclick=()=>{sortDescending=!sortDescending;setPref('file-direction',sortDescending?'desc':'asc');renderFiles();};
el('file-edit').onchange=renderFiles;
el('file-archive').onclick=()=>{archiveView=!archiveView;renderFiles();};
restoreColumns();
el('columns').onchange=()=>{setPref(grid?'columns':'list-columns',el('columns').value);renderFiles();};
el('duration').onsubmit=e=>{e.preventDefault();const sec=Number(el('minutes').value)*60+Number(el('seconds').value);
  if(sec<1 || sec>359999)return showMessage('请输入 1 秒至 99:59:59 的时间',true);command('set_seconds',sec);};
el('pair').onsubmit=e=>{e.preventDefault();const raw=el('key').value.trim();token=raw.includes('#')?raw.split('#').pop():raw;
  el('key').value='';refresh();};
const panelIds=['live','controls','agenda','inbox'];
for(const id of getPref('panels','').split(','))if(panelIds.includes(id))el('workspace').append(el(id));
for(const id of panelIds) {const panel=el(id);
  panel.ondragstart=e=>{if(e.target!==panel)return;dragged=panel;e.dataTransfer.setData('text/plain',id);};
  panel.ondragover=e=>{if(dragged)e.preventDefault();};panel.ondrop=e=>{if(!dragged || dragged===panel)return;
    e.preventDefault();el('workspace').insertBefore(dragged,panel);setPref('panels',Array.from(el('workspace').children).map(x=>x.id).join(','));};
  panel.ondragend=()=>{dragged=null;};
}
if(!token)el('login').hidden=false;refresh();setInterval(refresh,1200);
"""

OBS_HTML = """<!doctype html><html lang="zh-CN"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>ppt time · OBS 投影</title>
<link rel="stylesheet" href="/obs.css"><script defer src="/obs.js"></script></head>
<body><img id="frame" alt="投影画面" hidden><output id="obs-time" hidden></output><p id="status">正在读取投影画面…</p></body></html>"""
OBS_CSS = """html,body{margin:0;width:100%;height:100%;overflow:hidden;font-family:system-ui,sans-serif;background:transparent}
#frame{width:100%;height:100%;object-fit:contain;background:#000}#status{position:absolute;left:50%;top:50%;
transform:translate(-50%,-50%);color:white;background:#182321dd;border-radius:.5em;padding:.8em 1em}
#obs-time{position:absolute;top:0;left:50%;transform:translateX(-50%);color:white;background:#1e2225d9;
font-size:clamp(18px,3.6vw,68px);font-variant-numeric:tabular-nums;font-weight:650;padding:.15em .6em;border-radius:0 0 .35em .35em}
[hidden]{display:none!important}"""
OBS_JS = r"""
'use strict';
let token=location.hash.slice(1),previous='',busy=false;
try{token=token||sessionStorage.getItem('obs-key')||'';if(token)sessionStorage.setItem('obs-key',token);}catch{}
if(location.hash)history.replaceState(null,'',location.pathname+location.search);
const label=document.getElementById('status'), frame=document.getElementById('frame'), timer=document.getElementById('obs-time');
const onlyTime=new URLSearchParams(location.search).get('view')==='timer';
async function poll(){
  if(busy)return;busy=true;
  try{
    const options={headers:{Authorization:'Bearer '+token},cache:'no-store'};
    const stateReply=await fetch('/api/state',options);if(!stateReply.ok)throw Error('投影连接已关闭或失效');
    const {role,state}=await stateReply.json();if(role!=='obs')throw Error('请使用讲台生成的 OBS 连接');
    if(onlyTime){timer.hidden=false;timer.textContent=[state.time,state.page].filter(Boolean).join('  |  ');label.hidden=true;return;}
    if(!state.preview_allowed)throw Error('请在讲台开启共享投影画面');
    const image=await fetch('/api/preview',options);
    if(image.status!==200)throw Error('等待投影画面…');
    const next=URL.createObjectURL(await image.blob());frame.src=next;frame.hidden=false;label.hidden=true;
    if(previous)URL.revokeObjectURL(previous);previous=next;
  }catch(e){frame.hidden=true;timer.hidden=true;label.textContent=e.message;label.hidden=false;if(previous){URL.revokeObjectURL(previous);previous='';}}
  finally{busy=false;}
}
poll();setInterval(poll,1000);
"""

from remote_appearance import APPEARANCE_CSS, APPEARANCE_HTML, APPEARANCE_JS
from remote_agenda import AGENDA_CSS, AGENDA_HTML, AGENDA_JS

ASSETS = {path: (mime + "; charset=utf-8", source.encode()) for path, mime, source in (
    ("/", "text/html", HTML.replace("__APPEARANCE_HTML__", APPEARANCE_HTML)),
    ("/desk.css", "text/css", CSS + APPEARANCE_CSS),
    ("/desk.js", "application/javascript", JS.replace("__MEDIA_FORMATS__", ",".join(sorted(FORMATS))) + "\n" + APPEARANCE_JS),
    ("/obs", "text/html", OBS_HTML), ("/obs.css", "text/css", OBS_CSS), ("/obs.js", "application/javascript", OBS_JS),
    ("/agenda", "text/html", AGENDA_HTML), ("/agenda.css", "text/css", AGENDA_CSS), ("/agenda.js", "application/javascript", AGENDA_JS),
)}


def paired_request(link, method, endpoint, *, data=None, file=None):
    """Room-to-room client, initiated by the operator. Never follow redirects."""
    parsed = urlsplit(link.strip())
    if (parsed.scheme != "http" or not private_address(parsed.hostname or "") or
            parsed.username is not None or parsed.password is not None or parsed.query or
            parsed.path not in {"", "/"} or not parsed.port or not re.fullmatch(r"[\w-]{43}", parsed.fragment, re.ASCII)):
        raise ValueError("请粘贴另一会场生成的完整局域网连接")
    headers = {"Authorization": "Bearer " + parsed.fragment}
    client = http.client.HTTPConnection(parsed.hostname, parsed.port, timeout=20)
    try:
        if file:
            check_media(file, file.suffix.lower())
            length = file.stat().st_size
            if not 0 < length <= MAX_FILE:
                raise ValueError("单个文件不能超过 512 MB")
            headers.update({"Content-Type": "application/octet-stream", "Content-Length": str(length)})
            with file.open("rb") as stream:
                client.request(method, endpoint, body=stream, headers=headers)
        else:
            body = None if data is None else json.dumps(data, ensure_ascii=False).encode()
            client.request(method, endpoint, body=body, headers=headers)
        res = client.getresponse()
        payload = res.read(1024 * 1024 + 1)
        if len(payload) > 1024 * 1024:
            raise ValueError("对方返回数据过大")
        reply = json.loads(payload)
        if not isinstance(reply, dict) or res.status not in (200, 201, 202):
            raise ValueError(reply.get("error", "目标会场未接受请求") if isinstance(reply, dict) else "目标会场响应无效")
        return reply
    finally:
        client.close()


class LanControl(QObject):
    """Network callbacks are queued onto the GUI thread for application actions."""
    command = Signal(str, object)
    received = Signal(str)
    _queued = Signal(object, str, object)
    _transfer_finished = Signal(str, object)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.service = None
        self._state = {}
        self.agenda_provider = None
        self.dialog = None
        self.preferences = QSettings("ppt time", "meeting-lan")
        self.grab_timer = QTimer(self)
        self.grab_timer.setInterval(1000)
        self.grab_timer.timeout.connect(self._grab_preview)
        self._queued.connect(self._deliver, Qt.ConnectionType.QueuedConnection)
        self._transfer_finished.connect(self._finish_transfer, Qt.ConnectionType.QueuedConnection)
        self.transfer_running = False
        self._closed = False

    def set_state(self, state):
        if not isinstance(state, dict):
            return
        visible = {"room": self.preferences.value("room", "主会场")}
        for key in ("time", "page", "stage", "scene", "room", "active_session", "title","board_page","board_pages"):
            if state.get(key) is not None and isinstance(state[key], (int, str)):
                visible[key] = str(state[key])[:250]
        for key in ("files", "interstitials", "meetings"):
            if isinstance(state.get(key), list):
                visible[key] = [str(item).replace("\\", "/").split("/")[-1][:250] for item in state[key][:MAX_FILES]]
        agenda = state.get("agenda", state.get("schedule", []))
        try:
            visible["agenda"] = safe_agenda(agenda)
            visible["schedule"] = visible["agenda"]
            visible["agenda_revision"] = agenda_revision(visible["agenda"])
        except RequestError:
            visible["schedule"] = [str(entry)[:250] for entry in agenda[:500] if isinstance(entry, str)] if isinstance(agenda, list) else []
        if "result" in self._state:
            visible["result"] = self._state["result"]
        visible["desktop_controls"] = state.get("desktop_controls") is True
        visible["running"] = state.get("running") is True
        if isinstance(state.get("timing_options"), dict):
            visible["timing_options"] = {k: state["timing_options"][k] for k in ("pages_only", "overtime")
                                         if type(state["timing_options"].get(k)) is bool}
        if isinstance(state.get("system"), dict):
            visible["system"] = {k: v for k, v in state["system"].items()
                                 if k in {"cpu_percent", "memory_percent", "battery_percent", "graphics", "power", "os"}
                                 and isinstance(v, (str, int, float))}
        if isinstance(state.get("appearance"), dict):
            visible["appearance"] = state["appearance"].copy()
        self._state = visible
        if self.service:
            with self.service.state_lock:
                self.service.state = visible.copy()
                if visible.get("agenda_revision") == self.service.agenda_expected:
                    self.service.agenda_pending_until = 0

    def set_result(self, message):
        self._state["result"] = str(message)[:360]
        if self.service:
            with self.service.state_lock:
                self.service.state["result"] = self._state["result"]

    @Slot(object, str, object)
    def _deliver(self, generation, action, value):
        if self.service is None or self.service.active is not generation or not generation.is_set():
            return
        if action == "_received":
            self.received.emit(value)
        else:
            self.command.emit(action, value)

    def _start(self):
        try:
            service = MeetingServer(self.address.currentText(), self.port.value(), MeetingFiles(self.inbox.text()),
                                    lambda a, v: self._queued.emit(service.active, a, v),
                                    lambda path: self._queued.emit(service.active, "_received", path))
        except (OSError, ValueError) as exc:
            QMessageBox.warning(self.dialog, "无法开启局域网", str(exc))
            return
        self.service = service
        for key, val in (("room", self.room.text().strip() or "主会场"), ("address", self.address.currentText()),
                         ("port", self.port.value()), ("inbox", self.inbox.text()), ("screen", self.screens.currentText())):
            self.preferences.setValue(key, val)
        self._state["room"] = self.preferences.value("room")
        service.state = self._state.copy()
        self._set_preview()
        threading.Thread(target=service.serve_forever, kwargs={"poll_interval": 0.15}, daemon=True,
                         name="ppt time LAN meeting desk").start()
        self._refresh_dialog()

    def _stop(self):
        svc, self.service = self.service, None
        self.grab_timer.stop()
        if svc:
            svc.active.clear()
            with svc.state_lock:
                svc.preview_allowed = False
                svc.preview = b""
            svc.shutdown()
            svc.server_close()
        if self.dialog is not None:
            self._refresh_dialog()

    def close(self):
        self._closed = True
        self._stop()
        if self.dialog is not None:
            self.dialog.close()

    def _set_preview(self):
        if self.dialog:
            self.screens.setEnabled(self.preview_box.isChecked())
        if not self.service:
            return
        with self.service.state_lock:
            self.service.preview_allowed = self.preview_box.isChecked()
            self.service.preview = b""
        if self.preview_box.isChecked():
            self.grab_timer.start()
        else:
            self.grab_timer.stop()

    def _grab_preview(self):
        srv = self.service
        if not srv or not srv.preview_allowed or time.monotonic() - srv.preview_requested_at > 4:
            return
        name = self.screens.currentData()
        screen = next((screen for screen in QGuiApplication.screens() if screen.name() == name), None)
        if not screen:
            with srv.state_lock:
                srv.preview = b""
            return
        shot = screen.grabWindow(0)
        if shot.isNull():
            with srv.state_lock:
                srv.preview = b""
            return
        pixmap = shot.scaled(750, 425, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
        buf = QBuffer()
        buf.open(QIODevice.OpenModeFlag.WriteOnly)
        if pixmap.save(buf, "JPEG", 60):
            with srv.state_lock:
                srv.preview = bytes(buf.data())

    def _refresh_dialog(self):
        running = self.service is not None
        for w in (self.room, self.port, self.address, self.inbox_button):
            w.setEnabled(not running)
        self.switch.setText("关闭局域网连接" if running else "开启局域网连接")
        self.role.setEnabled(running)
        self.copy_button.setEnabled(running)
        self.qr_button.setEnabled(running)
        self.screens.setEnabled(self.preview_box.isChecked())
        role = self.role.currentData()
        url = self.service.link("obs" if role == "obs_timer" else role) if running else ""
        self.link.setText(url.replace("/obs#", "/obs?view=timer#") if role == "obs_timer" else url)
        host = self.address.currentText()
        self.subtitle.setText("仅本机可连接" if running and host.startswith("127.") else
                              ("已开启，同一局域网内可用对应连接访问" if running else "局域网连接已关闭"))

    def _pick_inbox(self):
        base = QFileDialog.getExistingDirectory(self.dialog, "为会场收件箱选择上级文件夹", str(Path(self.inbox.text()).parent))
        if base:
            self.inbox.setText(str(Path(base) / "ppt time 收件箱"))

    def _show_qr(self):
        try:
            import segno
        except ImportError:
            QMessageBox.information(self.dialog, "二维码", "当前环境缺少二维码组件，可点「复制连接」发送到手机。")
            return
        data = io.BytesIO()
        segno.make_qr(self.link.text(), error="m").save(data, kind="png", scale=5, border=4)
        pixmap = QPixmap()
        pixmap.loadFromData(data.getvalue(), "PNG")
        qr = QDialog(self.dialog)
        qr.setWindowTitle("ppt time · " + self.role.currentText() + "连接")
        layout = QVBoxLayout(qr)
        image = QLabel()
        image.setPixmap(pixmap)
        layout.addWidget(image)
        qr.exec()

    def _transfer(self, mode):
        if self.transfer_running:
            return
        peer = self.peer.text().strip()
        if not peer:
            QMessageBox.information(self.dialog, "连接另一会场", "请先粘贴目标会场分享的完整连接。")
            return
        names = []
        if mode == "files":
            chosen, _ = QFileDialog.getOpenFileNames(self.dialog, "选择发送至分会场的文件", self.inbox.text(),
                          "会议文件 (" + " ".join("*" + s for s in sorted(FORMATS)) + ")")
            names = [Path(f) for f in chosen]
            if not names:
                return
        try:
            agenda = []
            if mode == "push":
                rows = self.agenda_provider() if self.agenda_provider else self._state.get("agenda")
                if rows is None:
                    raise RequestError(400, "本会场日程暂不可用")
                agenda = safe_agenda(rows)
        except RequestError as exc:
            QMessageBox.warning(self.dialog, "日程无效", str(exc))
            return
        if mode == "push" and QMessageBox.question(self.dialog, "同步日程", "用本会场日程替换目标会场的日程？") != QMessageBox.StandardButton.Yes:
            return
        if mode == "pull" and QMessageBox.question(self.dialog, "同步日程", "从目标会场读取日程并替换本会场的日程？") != QMessageBox.StandardButton.Yes:
            return
        self.transfer_running = True
        self.transfer_status.setText("正在连接分会场…")
        def run():
            done = 0
            result = None
            try:
                if mode == "files":
                    for file in names:
                        paired_request(peer, "POST", "/api/upload?name=" + quote(file.name, safe=""), file=file)
                        done += 1
                    msg = f"分会场已接收 {done} 个文件"
                elif mode == "push":
                    paired_request(peer, "POST", "/api/agenda", data=agenda)
                    msg = f"已发送 {len(agenda)} 项日程"
                else:
                    state = paired_request(peer, "GET", "/api/state")
                    if state.get("role") not in {"control", "monitor"} or "agenda" not in state.get("state", {}):
                        raise ValueError("请使用目标会场的监看或控制连接")
                    result = safe_agenda(state["state"]["agenda"])
                    msg = f"已接收 {len(result)} 项日程"
            except (OSError, http.client.HTTPException, ValueError, RequestError) as exc:
                msg = (f"已发送 {done} 个文件。" if done else "") + "连接未完成：" + str(exc)
            self._transfer_finished.emit(msg, result)
        threading.Thread(target=run, name="ppt time room transfer", daemon=True).start()

    @Slot(str, object)
    def _finish_transfer(self, message, result):
        self.transfer_running = False
        if self._closed:
            return
        if result is not None:
            self.command.emit("import_agenda", result)
        if self.dialog:
            self.transfer_status.setText(message)

    def open_settings(self, *_):
        if self.dialog:
            self.dialog.show()
            self.dialog.raise_()
            self.dialog.activateWindow()
            return
        dlg = self.dialog = QDialog()
        dlg.setWindowTitle("ppt time · 局域网工作台")
        dlg.setMinimumWidth(540)
        layout = QVBoxLayout(dlg)
        self.subtitle = QLabel()
        layout.addWidget(self.subtitle)
        form = QFormLayout()
        layout.addLayout(form)
        self.room = QLineEdit(str(self.preferences.value("room", "主会场")))
        self.room.setMaxLength(80)
        form.addRow("会场名称", self.room)
        self.address = QComboBox()
        self.address.addItems(available_addresses())
        saved = self.address.findText(str(self.preferences.value("address", "")))
        self.address.setCurrentIndex(max(saved, 0))
        self.port = QSpinBox()
        self.port.setRange(1024, 65535)
        try:
            self.port.setValue(int(self.preferences.value("port", 48731)))
        except (TypeError, ValueError):
            self.port.setValue(48731)
        addr = QHBoxLayout()
        addr.addWidget(self.address, 1)
        addr.addWidget(QLabel("端口"))
        addr.addWidget(self.port)
        form.addRow("本机地址", addr)
        location = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DocumentsLocation) or str(Path.home())
        self.inbox = QLineEdit(str(self.preferences.value("inbox", str(Path(location) / "ppt time 收件箱"))))
        self.inbox.setReadOnly(True)
        self.inbox_button = QPushButton("更换")
        self.inbox_button.clicked.connect(self._pick_inbox)
        paths = QHBoxLayout()
        paths.addWidget(self.inbox)
        paths.addWidget(self.inbox_button)
        form.addRow("会场收件箱", paths)
        self.preview_box = QCheckBox("向监看端和控制端共享所选屏幕画面")
        form.addRow("实况画面", self.preview_box)
        self.screens = QComboBox()
        for i, screen in enumerate(QGuiApplication.screens()):
            self.screens.addItem(f"屏幕 {i+1} · {screen.name()}", screen.name())
        screen_index = self.screens.findText(str(self.preferences.value("screen", "")))
        self.screens.setCurrentIndex(max(0, screen_index))
        form.addRow("共享屏幕", self.screens)
        self.preview_box.toggled.connect(self._set_preview)
        self.screens.currentIndexChanged.connect(self._set_preview)
        self.switch = QPushButton()
        self.switch.clicked.connect(lambda: self._stop() if self.service else self._start())
        layout.addWidget(self.switch)
        pair = QHBoxLayout()
        self.role = QComboBox()
        for r in ("submit", "monitor", "control", "obs"):
            self.role.addItem(ROLE_NAMES[r], r)
        self.role.addItem("OBS 透明计时条", "obs_timer")
        self.copy_button = QPushButton("复制连接")
        self.qr_button = QPushButton("手机扫码")
        self.link = QLineEdit()
        self.link.setReadOnly(True)
        self.link.setEchoMode(QLineEdit.EchoMode.Password)
        self.link.setPlaceholderText("开启后可分享投稿、监看或控制连接")
        self.copy_button.clicked.connect(lambda: QApplication.clipboard().setText(self.link.text()))
        self.qr_button.clicked.connect(self._show_qr)
        self.role.currentIndexChanged.connect(self._refresh_dialog)
        for w in (self.role, self.link, self.copy_button, self.qr_button):
            pair.addWidget(w)
        layout.addLayout(pair)
        desc = QLabel("投稿连接只收文件，监看连接可查看计时和日程，控制连接可开启放映和翻页。连接仅供受信的会场局域网使用，关闭后失效。")
        desc.setWordWrap(True)
        layout.addWidget(desc)
        layout.addSpacing(12)
        layout.addWidget(QLabel("分会场传输"))
        self.peer = QLineEdit()
        self.peer.setEchoMode(QLineEdit.EchoMode.Password)
        self.peer.setPlaceholderText("粘贴另一会场的连接")
        layout.addWidget(self.peer)
        actions = QHBoxLayout()
        for caption, mode in (("发送会议文件", "files"), ("发送本会场日程", "push"), ("接收对方日程", "pull")):
            b = QPushButton(caption)
            b.clicked.connect(lambda checked=False, mode=mode: self._transfer(mode))
            actions.addWidget(b)
        layout.addLayout(actions)
        self.transfer_status = QLabel()
        self.transfer_status.setWordWrap(True)
        layout.addWidget(self.transfer_status)
        self._refresh_dialog()
        dlg.show()
