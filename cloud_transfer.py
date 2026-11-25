"""Operator-paired receiver for a ppt time file relay."""
import hashlib
import json
import os
import re
import tempfile
import threading
from pathlib import Path
from urllib.parse import urlsplit
from urllib.request import build_opener, HTTPRedirectHandler, Request

from PySide6.QtCore import QObject, Signal, Qt
from PySide6.QtGui import QGuiApplication, QImage, QPixmap
from PySide6.QtWidgets import (QCheckBox, QDialog, QFileDialog, QFormLayout, QHBoxLayout, QLabel,
                               QLineEdit, QPushButton, QVBoxLayout)
from relay_server import MAX_FILE, file_name


class NoRedirect(HTTPRedirectHandler):
    def redirect_request(self, *args):
        return None


def relay_url(value):
    value = value.strip().rstrip("/")
    url = urlsplit(value)
    if (not url.hostname or url.username or url.password or url.path or url.query or url.fragment or
            not (url.scheme == "https" or (url.scheme == "http" and url.hostname in ("localhost", "127.0.0.1", "::1")))):
        raise ValueError("请输入 HTTPS 服务地址，本机测试也可使用 http://127.0.0.1:8766")
    return value


def file_hash(path):
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        while block := stream.read(1024 * 1024):
            digest.update(block)
    return digest.hexdigest()


class RoomClient:
    def __init__(self, url, token, inbox):
        self.url = relay_url(url)
        self.token = token.strip()
        if not re.fullmatch(r"[A-Za-z0-9_-]{24,160}", self.token):
            raise ValueError("请输入服务器生成的讲台接收密钥")
        self.inbox = Path(inbox)
        self.opener = build_opener(NoRedirect())

    def request(self, path, method="GET", data=None, **headers):
        headers["Authorization"] = "Bearer " + self.token
        request = Request(self.url + path, method=method, headers=headers, data=data)
        return self.opener.open(request, timeout=12)

    def get_json(self, path):
        with self.request(path) as response:
            raw = response.read(1024 * 1024 + 1)
        if len(raw) > 1024 * 1024:
            raise ValueError("中转响应超出大小限制")
        return json.loads(raw)

    def collect(self, stop, on_file, on_assigned=None):
        rows = self.get_json("/v1/pending")
        if not isinstance(rows, list) or len(rows) > 50:
            raise ValueError("无效的中转队列")
        for row in rows:
            if stop.is_set():
                break
            key, name, size, sha = row["id"], file_name(row["name"]), row["size"], row["sha256"]
            if (not re.fullmatch("[a-f0-9]{32}", key) or not re.fullmatch("[a-f0-9]{64}", sha)
                    or type(size) is not int or not 0 < size <= MAX_FILE):
                raise ValueError("中转清单无效")
            directory = self.inbox / key
            directory.mkdir(parents=True, exist_ok=True)
            path = directory / name
            if not path.is_file() or file_hash(path) != sha:
                if path.exists():
                    raise ValueError("本地同名收件已经修改，请移走后重试")
                with tempfile.NamedTemporaryFile(dir=directory, delete=False) as sink:
                    temp = Path(sink.name)
                try:
                    digest = hashlib.sha256()
                    count = 0
                    with self.request("/v1/files/" + key) as response, temp.open("wb") as sink:
                        while block := response.read(1024 * 1024):
                            count += len(block)
                            if count > size or stop.is_set():
                                raise ValueError("本次下载已中止")
                            digest.update(block)
                            sink.write(block)
                        sink.flush()
                        os.fsync(sink.fileno())
                    if count != size or digest.hexdigest() != sha:
                        raise ValueError("文件校验未通过，中转站保留原件")
                    os.replace(temp, path)
                finally:
                    temp.unlink(missing_ok=True)
            with self.request("/v1/ack/" + key, "POST", **{"X-SHA256":sha}):
                pass
            slot = row.get("slot", "")
            if on_assigned and slot and re.fullmatch("[a-f0-9]{32}", slot):
                on_assigned(str(path), slot)
            else:
                on_file(str(path))


def qr_image(link):
    import segno
    matrix = tuple(segno.make(link, micro=False, error="m").matrix_iter(scale=1, border=4))
    pixels = b"".join(bytes(0 if dark else 255 for dark in row) for row in matrix)
    width = len(matrix)
    image = QImage(pixels, width, width, width, QImage.Format.Format_Grayscale8).copy()
    size = width * max(2, min(6, 270 // width))
    return QPixmap.fromImage(image).scaled(size, size, Qt.AspectRatioMode.KeepAspectRatio,
                                          Qt.TransformationMode.FastTransformation)


class CloudReceiver(QObject):
    received = Signal(str)
    assigned = Signal(str, str)
    _event = Signal(object, str, str)

    def __init__(self, directory, parent=None):
        super().__init__(parent)
        self.directory = Path(directory)
        self.inbox = self.directory / "云端收件"
        self.url = ""
        self.upload_link = ""
        self.status = "未连接"
        self.session = None
        self.dialog = None
        self.agenda_json = b"[]"
        self.publish = False
        self._event.connect(self._apply_event)

    def set_agenda(self, sessions):
        rows = [{"id":s.uid, "label":" · ".join(filter(None, (s.date, s.start, s.room, s.title, s.speaker)))[:320]}
                for s in sessions[:2000]]
        self.agenda_json = json.dumps(rows,ensure_ascii=False).encode()

    def _start(self):
        try:
            client = RoomClient(self.url_field.text(), self.key_field.text(), self.path_field.text())
        except ValueError as exc:
            self.status_label.setText(str(exc))
            return
        self.close_connection()
        self.inbox, self.url = client.inbox, client.url
        stop = self.session = threading.Event()
        self.status = "正在连接…"
        self._refresh()
        def work():
            paired = False
            last_agenda = None
            while not stop.is_set():
                try:
                    if not paired:
                        submit = client.get_json("/v1/pair")["submit"]
                        if not re.fullmatch(r"[a-zA-Z0-9_-]{24,160}", submit):
                            raise ValueError("中转返回了无效的投递链接")
                        self._event.emit(stop, "link", client.url + "/#" + submit)
                        paired = True
                    agenda = self.agenda_json if self.publish else b"[]"
                    if agenda != last_agenda:
                        with client.request("/v1/agenda","POST",agenda, **{"Content-Type":"application/json"}):
                            pass
                        last_agenda = agenda
                    client.collect(stop, lambda path: self._event.emit(stop, "file", path),
                                   lambda path, slot: self._event.emit(stop, "assigned", json.dumps([path,slot])))
                    self._event.emit(stop, "status", "已连接，正在等待会议文件")
                except Exception as exc:
                    self._event.emit(stop, "status", "连接或收件失败：" + str(exc))
                stop.wait(5)
        threading.Thread(target=work, name="ppt time file receiver", daemon=True).start()

    def _apply_event(self, session, kind, text):
        if session is not self.session:
            return
        if kind == "file":
            self.received.emit(text)
        elif kind == "assigned":
            path, slot = json.loads(text)
            self.assigned.emit(path,slot)
        elif kind == "link":
            self.upload_link = text
        else:
            self.status = text
        self._refresh()

    def close_connection(self):
        if self.session is not None:
            self.session.set()
            self.session = None
        self.upload_link = ""
        self.status = "未连接"
        self._refresh()

    def _refresh(self):
        if self.dialog is None:
            return
        self.status_label.setText(self.status)
        if self._last_link != self.upload_link:
            self._last_link = self.upload_link
            self.code.clear()
            if self.upload_link:
                self.code.setPixmap(qr_image(self.upload_link))
        self.copy.setEnabled(bool(self.upload_link))
        self.disconnect.setEnabled(self.session is not None)

    def open_settings(self):
        if self.dialog is not None:
            self.dialog.show()
            self.dialog.raise_()
            self.dialog.activateWindow()
            return
        dlg = self.dialog = QDialog()
        dlg.setWindowTitle("ppt time · 云端收件")
        dlg.setMinimumWidth(450)
        layout = QVBoxLayout(dlg)
        help_text = QLabel("连接会务方部署的 ppt time 中转服务后，手机可扫码上传。讲台电脑校验文件并加入会议控制台后即可打开放映。")
        help_text.setWordWrap(True)
        layout.addWidget(help_text)
        form = QFormLayout()
        layout.addLayout(form)
        self.url_field = QLineEdit(self.url)
        self.url_field.setPlaceholderText("https://会议收件域名")
        self.key_field = QLineEdit()
        self.key_field.setEchoMode(QLineEdit.EchoMode.Password)
        self.path_field = QLineEdit(str(self.inbox))
        choose = QPushButton("选择…")
        def select_directory():
            path = QFileDialog.getExistingDirectory(dlg, "选择收件目录", self.path_field.text())
            if path:
                self.path_field.setText(path)
        choose.clicked.connect(select_directory)
        paths = QHBoxLayout()
        paths.addWidget(self.path_field)
        paths.addWidget(choose)
        form.addRow("服务地址", self.url_field)
        form.addRow("接收密钥", self.key_field)
        form.addRow("收件目录", paths)
        publish = QCheckBox("在手机投递页显示会议议程，允许讲者选择对应场次")
        publish.setChecked(self.publish)
        publish.toggled.connect(lambda checked: setattr(self, "publish", checked))
        layout.addWidget(publish)
        self.status_label = QLabel()
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)
        self.code = QLabel()
        self.code.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.code)
        buttons = QHBoxLayout()
        layout.addLayout(buttons)
        start = QPushButton("连接并收件")
        start.clicked.connect(self._start)
        self.disconnect = QPushButton("断开")
        self.disconnect.clicked.connect(self.close_connection)
        self.copy = QPushButton("复制手机投递链接")
        self.copy.clicked.connect(lambda: QGuiApplication.clipboard().setText(self.upload_link))
        for button in (start, self.disconnect, self.copy):
            buttons.addWidget(button)
        self._last_link = None
        self._refresh()
        dlg.show()

    def close(self):
        self.close_connection()
        if self.dialog:
            self.dialog.close()
            self.dialog.deleteLater()
            self.dialog = None
