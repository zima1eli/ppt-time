"""Original synthesized meeting cues and local audio playback."""
import io
import math
import struct
import wave
from functools import lru_cache
from pathlib import Path
from PySide6.QtCore import QByteArray, QBuffer, QIODevice, QUrl
from PySide6.QtMultimedia import QAudioOutput, QMediaPlayer
from PySide6.QtWidgets import QApplication


CUES = {
    "内置 · 轻敲": ((660, .14),),
    "内置 · 双音": ((660, .14), (880, .21)),
    "内置 · 三声提示": ((540, .12), (540, .12), (720, .2)),
    "内置 · 柔和钟声": ((523, .58),),
    "内置 · 环节转换": ((494, .18), (622, .2), (740, .32)),
    "内置 · 时间到": ((784, .28), (659, .25), (523, .55)),
}
_player = None
_source = None


@lru_cache(maxsize=6)
def wav_cue(name):
    sample_rate = 22050
    samples = bytearray()
    for frequency, duration in CUES[name]:
        for i in range(int(sample_rate * duration)):
            t = i / sample_rate
            angle = 2 * math.pi * frequency * t
            envelope = min(1, t / .009, (duration - t) / .04) * math.exp(-3 * t / duration)
            level = (math.sin(angle) + .22 * math.sin(2 * angle)) * envelope
            samples.extend(struct.pack("<h", int(15000 * level)))
        samples.extend(bytes(2 * int(sample_rate * .055)))
    file = io.BytesIO()
    with wave.open(file, "wb") as stream:
        stream.setparams((1, 2, sample_rate, 0, "NONE", "not compressed"))
        stream.writeframes(samples)
    return file.getvalue()


def play_custom(name):
    global _player, _source
    if name not in CUES and not (name and Path(name).is_file()):
        return False
    if _player is None:
        _player = QMediaPlayer(QApplication.instance())
        _player.setAudioOutput(QAudioOutput(_player))
    _player.stop()
    _player.setSource(QUrl())
    if _source is not None:
        _source.close()
        _source.deleteLater()
        _source = None
    if name in CUES:
        _source = QBuffer(_player)
        _source.setData(QByteArray(wav_cue(name)))
        _source.open(QIODevice.OpenModeFlag.ReadOnly)
        _player.setSourceDevice(_source, QUrl("file:///meeting-cue.wav"))
    else:
        _player.setSource(QUrl.fromLocalFile(str(Path(name).resolve())))
    _player.play()
    return True
