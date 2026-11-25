import os
import unittest
import wave
from io import BytesIO
from unittest.mock import patch
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
from PySide6.QtWidgets import QApplication
from operator_tools import OperatorTools
from timer_audio import CUES, wav_cue


class OperatorTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_commands_need_explicit_session_opt_in(self):
        tools = OperatorTools()
        self.addCleanup(tools.close)
        with patch("operator_tools.subprocess.Popen") as launch:
            self.assertFalse(tools.execute("desk_files"))
            self.assertFalse(tools.execute("unregistered"))
            launch.assert_not_called()
        self.assertIn("os",tools.metrics())

    def test_audio_presets_are_distinct_playable_files(self):
        cues = [wav_cue(label) for label in CUES]
        self.assertEqual(len(set(cues)), 6)
        for data in cues:
            with wave.open(BytesIO(data)) as audio:
                self.assertGreater(audio.getnframes(), 2000)
                self.assertEqual(audio.getnchannels(),1)
