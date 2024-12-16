# pylint: disable=E0611,E1101
# (Disables "no-name-in-module" and "no-member" warnings for PySide6 & Mido)

import sys
import json
import os
import re
import socket
import threading
import time
from collections import deque

# PySide6 imports
from PySide6.QtCore import Qt, QTimer, QEvent, Signal
from PySide6.QtGui import QIcon, QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QLabel,
    QPushButton,
    QComboBox,
    QCheckBox,
    QLineEdit,
    QPlainTextEdit,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QFileDialog,
    QSlider,
    QMessageBox,
    QSpinBox,
    QFrame,
    QGroupBox,
)

# Third-party imports
import mido
from mido import Message, MidiFile
from pythonosc import dispatcher, osc_server, udp_client

# Explicitly set a Mido backend (optional, but helps with certain environments):
try:
    mido.set_backend('mido.backends.rtmidi')  # or another valid backend
except Exception as e:
    print("Warning: Could not set Mido backend:", e)


class PianoRoll(QWidget):
    """
    A simple piano roll UI with on-screen white and black keys,
    sending OSC note_on and note_off to /chXnote and /chXnoff respectively.
    """
    def __init__(self, osc_client=None, channel=1, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Piano Roll")
        self.resize(600, 200)

        self.osc_client = osc_client
        self.channel = channel

        self.white_keys = ["C", "D", "E", "F", "G", "A", "B"]
        self.black_keys = ["C#", "D#", "E#", "F#", "G#", "A#"]

        self.main_layout = QGridLayout(self)
        self.buttons = []  # Keep references to QPushButtons

        # Create white keys
        for i, key_label in enumerate(self.white_keys):
            btn = QPushButton(key_label)
            btn.setStyleSheet("background-color: white; color: black;")
            btn.setFixedSize(60, 150)
            btn.setCheckable(False)  # Not toggleable
            btn.pressed.connect(lambda k=key_label: self.send_note_on(k))
            btn.released.connect(lambda k=key_label: self.send_note_off(k))
            self.main_layout.addWidget(btn, 0, i * 2)
            self.buttons.append(btn)

        # Create black keys
        for key_label in self.black_keys:
            btn = QPushButton(key_label)
            btn.setStyleSheet("background-color: black; color: white;")
            btn.setFixedSize(40, 100)
            btn.setCheckable(False)  # Not toggleable
            btn.pressed.connect(lambda k=key_label: self.send_note_on(k))
            btn.released.connect(lambda k=key_label: self.send_note_off(k))

            # Determine the correct column alignment for black keys
            # e.g. "C#" -> put it above the gap between "C" and "D"
            if key_label in ["C#", "D#"]:
                column = self.white_keys.index(key_label[:-1]) * 2 + 1
            elif key_label == "E#":  # E# is F
                column = self.white_keys.index("E") * 2 + 1
            else:
                # F#, G#, A#
                column = self.white_keys.index(key_label[:-1]) * 2 + 1

            self.main_layout.addWidget(btn, 0, column)
            self.buttons.append(btn)

    def send_note_on(self, key_label):
        """Send OSC note_on message."""
        if not self.osc_client:
            print("[PianoRoll] No OSC client connected yet.")
            return

        key_to_note = {
            "C": 60,  "C#": 61,  "D": 62,  "D#": 63, "E": 64,  "E#": 65,  # E# = F
            "F": 65,  "F#": 66,  "G": 67,  "G#": 68, "A": 69,  "A#": 70,
            "B": 71,  "B#": 72
        }

        note_number = key_to_note.get(key_label, 60)

        osc_address_note = f"/ch{self.channel}note"
        self.osc_client.send_message(osc_address_note, note_number)
        print(f"[PianoRoll] Key pressed: {key_label} -> MIDI note {note_number}")
        self.update_button_color(key_label, True)

    def send_note_off(self, key_label):
        """Send OSC note_off message."""
        if not self.osc_client:
            print("[PianoRoll] No OSC client connected yet.")
            return

        key_to_note = {
            "C": 60,  "C#": 61,  "D": 62,  "D#": 63, "E": 64,  "E#": 65,  # E# = F
            "F": 65,  "F#": 66,  "G": 67,  "G#": 68, "A": 69,  "A#": 70,
            "B": 71,  "B#": 72
        }

        note_number = key_to_note.get(key_label, 60)

        osc_address_noff = f"/ch{self.channel}noff"
        self.osc_client.send_message(osc_address_noff, note_number)
        print(f"[PianoRoll] Key released: {key_label} -> MIDI note {note_number}")
        self.update_button_color(key_label, False)

    def update_button_color(self, key_label, pressed):
        for btn in self.buttons:
            if btn.text() == key_label:
                if pressed:
                    btn.setStyleSheet("background-color: lightblue;")
                else:
                    if "#" in key_label:
                        btn.setStyleSheet("background-color: black; color: white;")
                    else:
                        btn.setStyleSheet("background-color: white; color: black;")
                break


class DnDFrame(QFrame):
    """
    A QFrame that accepts dragged files (MIDI) via Qt's drag-and-drop system.
    """
    files_dropped = Signal(list)  # emits a list of file paths

    def __init__(self, text="", parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        layout = QVBoxLayout(self)
        label = QLabel(text)
        label.setAlignment(Qt.AlignCenter)
        layout.addWidget(label)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            paths = [url.toLocalFile() for url in event.mimeData().urls()]
            self.files_dropped.emit(paths)
            event.acceptProposedAction()
        else:
            event.ignore()


class OSCMIDIApp(QMainWindow):
    """
    Main application class for OSC2MIDI.
    Uses a logSignal to update UI logs from worker threads safely.
    """
    CONFIG_FILE = "config.json"
    MAX_LOG_LINES = 30

    logSignal = Signal(str)  # For thread-safe log updates

    def __init__(self):
        super().__init__()
        self.setWindowTitle("OSC2MIDI - Multi MIDI + Burger Menu")
        self.resize(800, 600)

        self.playing = False
        self.looping = False
        self.tempo = 120
        self.tempo_microseconds_per_beat = 500000
        self.paused = False
        self.request_stop = False

        self.midi_playlist = []
        self.current_playlist_index = 0
        self.sent_messages = deque(maxlen=100)

        self.osc_client = None
        self.osc_server = None
        self.dispatcher = None
        self.osc_server_thread = None

        self.midi_in = None
        self.midi_out = None
        self.midi_input_thread = None

        self.lock = threading.Lock()

        # Config
        self.saved_port = "5550"
        self.saved_midi_port = ""
        self.saved_midi_out_port = ""
        self.saved_out_ip = ""
        self.saved_out_port = ""
        self.osc_out_channels = {}

        self.piano_roll_window = None
        self.piano_roll_visible = False

        self.load_config()

        # Build UI
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        self.main_layout = QHBoxLayout(main_widget)

        # Left panel layout
        self.left_panel = QVBoxLayout()
        self.main_layout.addLayout(self.left_panel, stretch=0)

        # Right side is a vertical layout for logs / MIDI player
        self.right_panel = QVBoxLayout()
        self.main_layout.addLayout(self.right_panel, stretch=1)

        # Top row: "☰ Piano Roll" and "☰ MIDI Player" buttons
        top_button_layout = QHBoxLayout()
        self.left_panel.addLayout(top_button_layout)

        self.piano_burger_button = QPushButton("☰ Piano Roll")
        self.piano_burger_button.clicked.connect(self.toggle_piano_roll)
        top_button_layout.addWidget(self.piano_burger_button)

        self.midi_burger_button = QPushButton("☰ MIDI Player")
        self.midi_burger_button.clicked.connect(self.toggle_midi_player)
        top_button_layout.addWidget(self.midi_burger_button)

        # Input group
        self.input_group = QGroupBox("OSC Input Settings")
        self.left_panel.addWidget(self.input_group)
        input_layout = QGridLayout(self.input_group)

        ip_label = QLabel("IP for OSC In:")
        input_layout.addWidget(ip_label, 0, 0)
        self.ip_entry = QLineEdit(self.get_local_ip())
        self.ip_entry.setReadOnly(True)
        input_layout.addWidget(self.ip_entry, 0, 1)

        port_label = QLabel("Port for OSC In:")
        input_layout.addWidget(port_label, 1, 0)
        self.port_entry = QLineEdit(str(self.saved_port))
        input_layout.addWidget(self.port_entry, 1, 1)

        # Output group
        self.output_group = QGroupBox("OSC Output Settings")
        self.left_panel.addWidget(self.output_group)
        output_layout = QGridLayout(self.output_group)

        output_ip_label = QLabel("IP for OSC Out:")
        output_layout.addWidget(output_ip_label, 0, 0)
        self.output_ip_entry = QLineEdit(str(self.saved_out_ip))
        output_layout.addWidget(self.output_ip_entry, 0, 1)

        output_port_label = QLabel("Port for OSC Out:")
        output_layout.addWidget(output_port_label, 1, 0)
        self.output_port_entry = QLineEdit(str(self.saved_out_port))
        output_layout.addWidget(self.output_port_entry, 1, 1)

        # Channel selection for single-channel mode
        channel_label = QLabel("OSC Out Channel:")
        output_layout.addWidget(channel_label, 2, 0)
        self.channel_menu = QComboBox()
        self.channel_menu.addItems([str(i) for i in range(1, 17)])
        self.channel_menu.setCurrentIndex(0)
        self.channel_menu.currentIndexChanged.connect(self.update_piano_roll_channel)
        output_layout.addWidget(self.channel_menu, 2, 1)

        # 16-channel checkbox
        self.ch16_checkbox = QCheckBox("16 Ch")
        self.ch16_checkbox.stateChanged.connect(self.toggle_16_channels)
        output_layout.addWidget(self.ch16_checkbox, 2, 2)

        # MIDI group
        self.midi_group = QGroupBox("MIDI Settings")
        self.left_panel.addWidget(self.midi_group)
        midi_layout = QGridLayout(self.midi_group)

        midi_label = QLabel("Select MIDI Input:")
        midi_layout.addWidget(midi_label, 0, 0)

        self.midi_input_combo = QComboBox()
        # Safely get Mido input names
        try:
            midi_input_names = mido.get_input_names()
        except Exception as e:
            midi_input_names = []
            print("Error retrieving MIDI input names:", e)
        self.midi_input_combo.addItems(midi_input_names)
        if self.saved_midi_port and self.saved_midi_port in midi_input_names:
            self.midi_input_combo.setCurrentText(self.saved_midi_port)
        self.midi_input_combo.currentIndexChanged.connect(self.change_midi_input)
        midi_layout.addWidget(self.midi_input_combo, 0, 1)

        midi_out_label = QLabel("Select MIDI Output:")
        midi_layout.addWidget(midi_out_label, 1, 0)

        self.midi_out_combo = QComboBox()
        # Safely get Mido output names
        try:
            midi_output_names = mido.get_output_names()
        except Exception as e:
            midi_output_names = []
            print("Error retrieving MIDI output names:", e)
        self.midi_out_combo.addItems(midi_output_names)
        if self.saved_midi_out_port and self.saved_midi_out_port in midi_output_names:
            self.midi_out_combo.setCurrentText(self.saved_midi_out_port)
        self.midi_out_combo.currentIndexChanged.connect(self.change_midi_output)
        midi_layout.addWidget(self.midi_out_combo, 1, 1)

        # Server control button
        self.start_button = QPushButton("Start Server")
        self.start_button.clicked.connect(self.start_server)
        self.left_panel.addWidget(self.start_button)

        # Left panel stretch
        self.left_panel.addStretch(1)

        # Logs + MIDI Player side by side in the right panel
        self.log_text = QPlainTextEdit()
        self.log_text.setReadOnly(True)

        self.midi_player_frame = QWidget()
        midi_player_layout = QVBoxLayout(self.midi_player_frame)

        # MIDI Player log
        self.midi_player_log_text = QPlainTextEdit()
        self.midi_player_log_text.setReadOnly(True)
        midi_player_layout.addWidget(QLabel("MIDI Player Log:"))
        midi_player_layout.addWidget(self.midi_player_log_text)

        # MIDI player controls
        controls_widget = QWidget()
        controls_layout = QGridLayout(controls_widget)
        midi_player_layout.addWidget(controls_widget)

        self.play_button = QPushButton("Play")
        self.play_button.clicked.connect(self.play)
        controls_layout.addWidget(self.play_button, 0, 0)

        self.stop_button = QPushButton("Stop")
        self.stop_button.clicked.connect(self.stop)
        controls_layout.addWidget(self.stop_button, 0, 1)

        self.load_folder_button = QPushButton("Add MIDI Folder")
        self.load_folder_button.clicked.connect(self.add_folder_to_playlist)
        controls_layout.addWidget(self.load_folder_button, 0, 2)

        self.load_button = QPushButton("Add MIDI File(s)")
        self.load_button.clicked.connect(self.add_file_to_playlist)
        controls_layout.addWidget(self.load_button, 0, 3)

        self.send_note_button = QPushButton("Send Test Note")
        self.send_note_button.clicked.connect(self.send_test_note)
        controls_layout.addWidget(self.send_note_button, 1, 0)

        self.skip_button = QPushButton("Skip ▶")
        self.skip_button.clicked.connect(self.skip_forward)
        controls_layout.addWidget(self.skip_button, 1, 1)

        self.back_button = QPushButton("◀ Back")
        self.back_button.clicked.connect(self.skip_back)
        controls_layout.addWidget(self.back_button, 1, 2)

        self.pause_button = QPushButton("Pause")
        self.pause_button.clicked.connect(self.toggle_pause)
        controls_layout.addWidget(self.pause_button, 1, 3)

        # Tempo slider
        tempo_label = QLabel("Tempo Slider:")
        midi_player_layout.addWidget(tempo_label)

        self.tempo_slider = QSlider(Qt.Horizontal)
        self.tempo_slider.setRange(20, 420)
        self.tempo_slider.setValue(self.tempo)
        self.tempo_slider.valueChanged.connect(self.update_tempo_display)
        midi_player_layout.addWidget(self.tempo_slider)

        self.bpm_display_label = QLabel(f"{self.tempo} BPM")
        midi_player_layout.addWidget(self.bpm_display_label)

        self.reset_tempo_button = QPushButton("Reset Tempo")
        self.reset_tempo_button.clicked.connect(self.reset_tempo)
        midi_player_layout.addWidget(self.reset_tempo_button)

        self.looping_check = QCheckBox("Loop Playlist")
        self.looping_check.setChecked(self.looping)
        self.looping_check.stateChanged.connect(self.toggle_looping)
        midi_player_layout.addWidget(self.looping_check)

        self.playlist_text = QPlainTextEdit()
        self.playlist_text.setReadOnly(True)
        midi_player_layout.addWidget(self.playlist_text)

        drop_label_frame = DnDFrame("Drop MIDI Files Here", parent=self)
        drop_label_frame.files_dropped.connect(self.handle_dropped_files)
        drop_label_frame.setFixedHeight(80)
        drop_label_frame.setStyleSheet("background-color: #444; color: white;")
        midi_player_layout.addWidget(drop_label_frame)

        # Add the log text and MIDI player to the right panel
        self.right_panel.addWidget(QLabel("OSC / MIDI Log:"))
        self.right_panel.addWidget(self.log_text)
        self.midi_player_frame.hide()
        self.right_panel.addWidget(self.midi_player_frame)

        # Connect the logSignal to our slot that updates the UI
        self.logSignal.connect(self.on_log_signal)

        self.closeEvent = self._close_event
        self.update_playlist_label()

        # Timer for cleaning up sent messages
        self.cleanup_timer = QTimer()
        self.cleanup_timer.timeout.connect(self.cleanup_sent_messages)
        self.cleanup_timer.start(50)

    def _close_event(self, event):
        self.quit_app()
        event.accept()

    # SLOT for logSignal
    def on_log_signal(self, message: str):
        """Called in main thread; updates the appropriate log widget."""
        if self.midi_player_frame.isVisible():
            text_widget = self.midi_player_log_text
        else:
            text_widget = self.log_text

        text_widget.appendPlainText(message)

        # Cleanup old lines
        line_count = text_widget.blockCount()
        if line_count > self.MAX_LOG_LINES:
            text = text_widget.toPlainText().splitlines()
            trimmed = text[-self.MAX_LOG_LINES:]
            text_widget.setPlainText("\n".join(trimmed))

    def log_message(self, message: str):
        """
        Thread-safe logging: instead of direct UI calls, emit a signal.
        """
        self.logSignal.emit(message)

    def load_config(self):
        if os.path.exists(self.CONFIG_FILE):
            with open(self.CONFIG_FILE, 'r') as f:
                config = json.load(f)
                self.saved_port = str(config.get("osc_in_port", "5550"))
                self.saved_midi_port = config.get("midi_input_port", "")
                self.saved_midi_out_port = config.get("midi_output_port", "")
                self.saved_out_ip = config.get("osc_out_ip", self.get_local_ip())
                self.saved_out_port = str(config.get("osc_out_port", "3330"))
                self.osc_out_channels = config.get("osc_out_channels", {})
        else:
            self.saved_port = "5550"
            self.saved_midi_port = ""
            self.saved_midi_out_port = ""
            self.saved_out_ip = self.get_local_ip()
            self.saved_out_port = "3330"
            self.osc_out_channels = {}

    def save_config(self):
        config = {
            "osc_in_port": self.saved_port,
            "midi_input_port": self.midi_input_combo.currentText(),
            "midi_output_port": self.midi_out_combo.currentText(),
            "osc_out_ip": self.output_ip_entry.text(),
            "osc_out_port": self.output_port_entry.text(),
            "osc_out_channels": self.osc_out_channels,
        }
        with open(self.CONFIG_FILE, 'w') as f:
            json.dump(config, f, indent=4)

    def toggle_piano_roll(self):
        if self.piano_roll_visible:
            if self.piano_roll_window:
                self.piano_roll_window.close()
                self.piano_roll_window = None
            self.piano_roll_visible = False
            self.log_message("Piano Roll hidden.")
        else:
            channel = int(self.channel_menu.currentText())
            self.piano_roll_window = PianoRoll(osc_client=self.osc_client, channel=channel)
            self.piano_roll_window.show()
            self.piano_roll_visible = True
            self.log_message("Piano Roll shown.")

    def toggle_midi_player(self):
        if self.midi_player_frame.isVisible():
            self.midi_player_frame.hide()
            self.log_text.show()
            self.log_message("MIDI Player hidden. Main log shown.")
        else:
            self.midi_player_frame.show()
            self.log_text.hide()
            self.log_message("MIDI Player shown. Main log hidden.")

    def toggle_16_channels(self, state):
        checked = (state == Qt.Checked)
        if checked:
            self.log_message("16 Channel Mode Enabled.")
            self.channel_menu.setEnabled(False)
        else:
            self.log_message("16 Channel Mode Disabled.")
            self.channel_menu.setEnabled(True)

        # If server is running, re-start to re-map addresses
        if self.start_button.text() == "Stop Server":
            self.stop_server()
            # Adding a slight delay to ensure server has stopped
            QTimer.singleShot(100, self.start_server)

    def get_local_ip(self):
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(("8.8.8.8", 80))
            ip = s.getsockname()[0]
            s.close()
            return ip
        except Exception as e:
            print(f"Error getting local IP: {e}")
            return "127.0.0.1"

    def change_midi_input(self):
        selected_port = self.midi_input_combo.currentText()
        with self.lock:
            # Close previous port
            if self.midi_in:
                try:
                    self.midi_in.close()
                    self.log_message("Previous MIDI input closed.")
                except Exception as e:
                    self.log_message(f"Error closing previous MIDI input: {e}")

            # Try to open the new port if it’s not empty
            if selected_port:
                try:
                    self.midi_in = mido.open_input(selected_port)
                    self.log_message(f"MIDI Input changed to: {selected_port}")
                    # Update OSC channel if not in 16 Ch mode
                    if not self.ch16_checkbox.isChecked():
                        ch = int(self.channel_menu.currentText())
                        if self.piano_roll_window:
                            self.piano_roll_window.channel = ch
                        self.log_message(f"OSC channel set to {ch}")
                except (AttributeError, OSError) as err:
                    QMessageBox.critical(self, "Error", f"Could not open MIDI input port:\n{err}")
                    self.log_message(f"Error changing MIDI input: {err}")

    def change_midi_output(self):
        selected_port = self.midi_out_combo.currentText()
        with self.lock:
            # Close previous port
            if self.midi_out:
                try:
                    self.midi_out.close()
                    self.log_message("Previous MIDI output closed.")
                except Exception as e:
                    self.log_message(f"Error closing MIDI output: {e}")

            # Try to open new port
            if selected_port:
                try:
                    self.midi_out = mido.open_output(selected_port)
                    self.log_message(f"MIDI Output changed to: {selected_port}")
                except (AttributeError, OSError) as err:
                    QMessageBox.critical(self, "Error", f"Could not open MIDI output port:\n{err}")
                    self.log_message(f"Error changing MIDI output: {err}")

    def start_server(self):
        with self.lock:
            if self.start_button.text() == "Start Server":
                midi_input_port_name = self.midi_input_combo.currentText()
                midi_output_port_name = self.midi_out_combo.currentText()
                osc_out_ip = self.output_ip_entry.text()

                try:
                    osc_out_port = int(self.output_port_entry.text())
                    osc_in_port = int(self.port_entry.text())
                except ValueError:
                    QMessageBox.critical(self, "Error", "OSC ports must be integers.")
                    return

                if not midi_input_port_name or not midi_output_port_name:
                    QMessageBox.critical(self, "Error", "Select both MIDI Input and Output ports.")
                    return

                try:
                    self.dispatcher = dispatcher.Dispatcher()
                    if self.ch16_checkbox.isChecked():
                        # 16 Channel Mode: Map all 16 channels
                        for ch in range(1, 17):
                            self.dispatcher.map(f"/ch{ch}note", self.handle_osc_message)
                            self.dispatcher.map(f"/ch{ch}noff", self.handle_osc_message)
                            # Map specific CC addresses
                            for cc_num in range(128):  # MIDI CC numbers 0-127
                                self.dispatcher.map(f"/ch{ch}cc{cc_num}", self.handle_osc_message)
                    else:
                        # Single Channel Mode: Map only the selected channel
                        selected_ch = int(self.channel_menu.currentText())
                        self.dispatcher.map(f"/ch{selected_ch}note", self.handle_osc_message)
                        self.dispatcher.map(f"/ch{selected_ch}noff", self.handle_osc_message)
                        for cc_num in range(128):
                            self.dispatcher.map(f"/ch{selected_ch}cc{cc_num}", self.handle_osc_message)

                    if self.is_port_in_use(osc_in_port):
                        QMessageBox.critical(self, "Error", f"OSC In port {osc_in_port} is already in use.")
                        self.log_message(f"Error: OSC In port {osc_in_port} is already in use.")
                        return

                    self.osc_server = osc_server.ThreadingOSCUDPServer(("", osc_in_port), self.dispatcher)
                    self.osc_server_thread = threading.Thread(target=self.osc_server.serve_forever, daemon=True)
                    self.osc_server_thread.start()
                    self.log_message(f"OSC Server started on port {osc_in_port}")

                    self.osc_client = udp_client.SimpleUDPClient(osc_out_ip, osc_out_port)
                    self.log_message(f"OSC Client set to {osc_out_ip}:{osc_out_port}")

                    # Re-open MIDI Out if needed
                    if not self.midi_out:
                        try:
                            self.midi_out = mido.open_output(midi_output_port_name)
                            self.log_message(f"MIDI Out set to {midi_output_port_name}")
                        except Exception as e:
                            QMessageBox.critical(self, "Error", str(e))
                            self.log_message(f"Error opening MIDI output: {e}")
                            return

                    # Re-open MIDI In if needed
                    if not self.midi_in:
                        try:
                            self.midi_in = mido.open_input(midi_input_port_name)
                            self.log_message(f"MIDI In set to {midi_input_port_name}")
                        except Exception as e:
                            QMessageBox.critical(self, "Error", str(e))
                            self.log_message(f"Error opening MIDI input: {e}")
                            return

                    self.midi_input_thread = threading.Thread(target=self.run_midi_input, daemon=True)
                    self.midi_input_thread.start()
                    self.log_message("MIDI input thread started.")

                    if self.piano_roll_window and not self.ch16_checkbox.isChecked():
                        self.piano_roll_window.osc_client = self.osc_client
                        self.piano_roll_window.channel = int(self.channel_menu.currentText())

                    # Save config
                    self.saved_out_ip = osc_out_ip
                    self.saved_out_port = str(osc_out_port)
                    self.saved_port = str(osc_in_port)
                    self.saved_midi_port = midi_input_port_name
                    self.saved_midi_out_port = midi_output_port_name
                    if not self.ch16_checkbox.isChecked():
                        self.osc_out_channels[midi_input_port_name] = self.channel_menu.currentText()
                    self.save_config()

                    self.start_button.setText("Stop Server")

                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Failed to start server: {e}")
                    self.log_message(f"Error starting server: {e}")
            else:
                self.stop_server()

    def is_port_in_use(self, port):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("", port))
                return False
            except OSError:
                return True

    def run_midi_input(self):
        if not self.midi_in or not self.osc_client:
            self.log_message("MIDI In or OSC Client not initialized.")
            return

        # Continuously read from the MIDI input in the background thread
        for msg in self.midi_in:
            if self.start_button.text() != "Stop Server":
                break  # user requested stop

            if msg.type in ['note_on', 'note_off', 'control_change', 'aftertouch', 'pitchwheel']:
                channel = msg.channel + 1

                if msg.type == 'note_on' and msg.velocity > 0:
                    osc_address_note = f"/ch{channel}note"
                    self.osc_client.send_message(osc_address_note, msg.note)
                    self.log_message(f"MIDI->OSC note_on ch{channel} note{msg.note}")
                    self.sent_messages.append((osc_address_note, (msg.note,), "note_on", time.time()))

                elif msg.type == 'note_off' or (msg.type == 'note_on' and msg.velocity == 0):
                    osc_address_noff = f"/ch{channel}noff"
                    self.osc_client.send_message(osc_address_noff, msg.note)
                    self.log_message(f"MIDI->OSC note_off ch{channel} note{msg.note}")
                    self.sent_messages.append((osc_address_noff, (msg.note,), "note_off", time.time()))

                elif msg.type == 'control_change':
                    osc_address_cc = f"/ch{channel}cc{msg.control}"
                    self.osc_client.send_message(osc_address_cc, msg.value)
                    self.log_message(f"MIDI->OSC CC ch{channel} cc{msg.control} val{msg.value}")
                    self.sent_messages.append((osc_address_cc, (msg.value,), "control_change", time.time()))

                elif msg.type == 'aftertouch':
                    osc_address_pressure = f"/ch{channel}pressure"
                    self.osc_client.send_message(osc_address_pressure, msg.value)
                    self.log_message(f"MIDI->OSC aftertouch ch{channel} val{msg.value}")
                    self.sent_messages.append((osc_address_pressure, (msg.value,), "aftertouch", time.time()))

                elif msg.type == 'pitchwheel':
                    osc_address_pitch = f"/ch{channel}pitch"
                    self.osc_client.send_message(osc_address_pitch, msg.pitch)
                    self.log_message(f"MIDI->OSC pitchwheel ch{channel} pitch{msg.pitch}")
                    self.sent_messages.append((osc_address_pitch, (msg.pitch,), "pitchwheel", time.time()))

    def handle_osc_message(self, address, *args):
        """Dispatcher callback for OSC -> MIDI logic."""
        current_time = time.time()
        with self.lock:
            # Remove sent messages older than 0.1s to avoid feedback loops
            while self.sent_messages and (current_time - self.sent_messages[0][3] > 0.1):
                self.sent_messages.popleft()

        # Parse the OSC address
        match = re.match(r"/ch(\d+)(note|noff|cc(\d+)|pressure|pitch)", address)
        if match:
            channel = int(match.group(1)) - 1  # MIDI channels are 0-15
            command = match.group(2)
            cc_num = match.group(3)

            with self.lock:
                if self.paused:
                    return

                try:
                    if command == "note":
                        note = int(args[0])
                        midi_message = Message('note_on', channel=channel, note=note, velocity=100)
                        self.midi_out.send(midi_message)
                        self.log_message(f"OSC->MIDI {midi_message}")
                        self.sent_messages.append((address, args, "note_on", time.time()))

                    elif command == "noff":
                        note = int(args[0])
                        midi_message = Message('note_off', channel=channel, note=note, velocity=0)
                        self.midi_out.send(midi_message)
                        self.log_message(f"OSC->MIDI {midi_message}")
                        self.sent_messages.append((address, args, "note_off", time.time()))

                    elif command.startswith("cc"):
                        if cc_num is not None:
                            control = int(cc_num)
                            value = int(args[0])
                            midi_message = Message('control_change', channel=channel, control=control, value=value)
                            self.midi_out.send(midi_message)
                            self.log_message(f"OSC->MIDI {midi_message}")
                            self.sent_messages.append((address, args, "control_change", time.time()))
                        else:
                            self.log_message(f"Invalid CC address: {address}")

                    elif command == "pressure":
                        value = int(args[0])
                        midi_message = Message('aftertouch', channel=channel, value=value)
                        self.midi_out.send(midi_message)
                        self.log_message(f"OSC->MIDI {midi_message}")
                        self.sent_messages.append((address, args, "aftertouch", time.time()))

                    elif command == "pitch":
                        pitch = int(args[0])
                        midi_message = Message('pitchwheel', channel=channel, pitch=pitch)
                        self.midi_out.send(midi_message)
                        self.log_message(f"OSC->MIDI {midi_message}")
                        self.sent_messages.append((address, args, "pitchwheel", time.time()))

                    else:
                        self.log_message(f"Unhandled OSC command: {command}")

                except Exception as e:
                    self.log_message(f"Error handling OSC message {address}: {e}")
        else:
            self.log_message(f"Invalid OSC address: {address}")

    def add_file_to_playlist(self):
        file_paths, _ = QFileDialog.getOpenFileNames(self, "Open MIDI Files", filter="MIDI Files (*.mid)")
        if file_paths:
            for fp in file_paths:
                if os.path.isfile(fp) and fp.lower().endswith(".mid"):
                    self.midi_playlist.append(fp)
            self.update_playlist_label()

    def add_folder_to_playlist(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select MIDI Folder")
        if folder_path:
            mid_files = sorted([f for f in os.listdir(folder_path) if f.lower().endswith('.mid')])
            for mf in mid_files:
                full_path = os.path.join(folder_path, mf)
                self.midi_playlist.append(full_path)
            self.update_playlist_label()

    def send_test_note(self):
        if not self.osc_client:
            self.log_message("OSC client not initialized. Start the server first.")
            return

        test_note = 60
        chan = int(self.channel_menu.currentText())
        osc_address_note = f"/ch{chan}note"

        self.osc_client.send_message(osc_address_note, test_note)
        self.log_message(f"Test note_on -> {osc_address_note} {test_note}")
        self.sent_messages.append((osc_address_note, (test_note,), "note_on", time.time()))

    def toggle_pause(self):
        self.paused = not self.paused
        if self.paused:
            self.pause_button.setText("Resume")
            self.log_message("Playback paused.")
        else:
            self.pause_button.setText("Pause")
            self.log_message("Playback resumed.")

    def skip_forward(self):
        if not self.midi_playlist:
            self.log_message("No files to skip.")
            return
        self.request_stop = True
        self.playing = False
        self.paused = False

        if hasattr(self, 'thread') and self.thread.is_alive():
            self.thread.join(timeout=0.2)

        self.current_playlist_index = (self.current_playlist_index + 1) % len(self.midi_playlist)
        self.log_message(f"Skipping forward to index {self.current_playlist_index}")
        self.play()

    def skip_back(self):
        if not self.midi_playlist:
            self.log_message("No files to skip back.")
            return
        self.request_stop = True
        self.playing = False
        self.paused = False

        if hasattr(self, 'thread') and self.thread.is_alive():
            self.thread.join(timeout=0.2)

        self.current_playlist_index = (self.current_playlist_index - 1) % len(self.midi_playlist)
        self.log_message(f"Skipping back to index {self.current_playlist_index}")
        self.play()

    def update_playlist_label(self):
        text = ""
        if self.midi_playlist:
            text += "Playlist:\n"
            for idx, path in enumerate(self.midi_playlist):
                text += f"{idx+1}. {os.path.basename(path)}\n"
        else:
            text = "No playlist loaded."
        self.playlist_text.setPlainText(text)

    def cleanup_sent_messages(self):
        with self.lock:
            current_time = time.time()
            while self.sent_messages and (current_time - self.sent_messages[0][3] > 0.1):
                self.sent_messages.popleft()

    def stop(self):
        if self.playing:
            self.playing = False
            self.paused = False
            self.request_stop = True
            self.log_message("Playback stopped.")
        else:
            self.log_message("Playback not running.")

        self.midi_playlist.clear()
        self.current_playlist_index = 0
        self.update_playlist_label()

    def play(self):
        if not self.midi_playlist:
            self.log_message("No files in the playlist to play.")
            return

        if not self.playing:
            self.playing = True
            self.request_stop = False
            self.log_message("Playback started.")
            self.thread = threading.Thread(target=self._play_playlist, daemon=True)
            self.thread.start()
        else:
            self.log_message("Playback already in progress.")

    def _play_playlist(self):
        while self.playing and not self.request_stop:
            if self.current_playlist_index >= len(self.midi_playlist):
                if self.looping:
                    self.current_playlist_index = 0
                else:
                    self.log_message("Reached end of playlist.")
                    break

            current_file = self.midi_playlist[self.current_playlist_index]
            self._play_single_midi(current_file)
            self.current_playlist_index += 1

        self.log_message("Playlist playback ended.")
        self.current_playlist_index = 0
        self.playing = False
        self.paused = False

    def _play_single_midi(self, file_path):
        self.log_message(f"Playing: {os.path.basename(file_path)}")
        try:
            midi = MidiFile(file_path)
            ticks_per_beat = midi.ticks_per_beat

            messages = []
            for track in midi.tracks:
                abs_time = 0
                for msg in track:
                    abs_time += msg.time
                    messages.append((abs_time, msg))
            messages.sort(key=lambda x: x[0])

            previous_abs_time = 0

            for abs_time, msg in messages:
                if not self.playing or self.request_stop:
                    self.log_message("Playback aborted.")
                    return

                while self.paused and self.playing and not self.request_stop:
                    time.sleep(0.01)

                delta_ticks = abs_time - previous_abs_time
                msg_time = mido.tick2second(delta_ticks, ticks_per_beat, self.tempo_microseconds_per_beat)
                time.sleep(msg_time)
                previous_abs_time = abs_time

                if msg.type in ['note_on', 'note_off', 'control_change', 'aftertouch', 'pitchwheel']:
                    channel = msg.channel + 1

                    if msg.type == 'note_on' and msg.velocity > 0:
                        osc_address_note = f"/ch{channel}note"
                        self.osc_client.send_message(osc_address_note, msg.note)
                        self.log_message(f"Note ON (Ch {channel}): {msg.note}")
                        self.sent_messages.append((osc_address_note, (msg.note,), "note_on", time.time()))

                    elif msg.type == 'note_off' or (msg.type == 'note_on' and msg.velocity == 0):
                        osc_address_noff = f"/ch{channel}noff"
                        self.osc_client.send_message(osc_address_noff, msg.note)
                        self.log_message(f"Note OFF (Ch {channel}): {msg.note}")
                        self.sent_messages.append((osc_address_noff, (msg.note,), "note_off", time.time()))

                    elif msg.type == 'control_change':
                        osc_address_cc = f"/ch{channel}cc{msg.control}"
                        self.osc_client.send_message(osc_address_cc, msg.value)
                        self.log_message(f"CC (Ch {channel}): CC{msg.control} -> {msg.value}")
                        self.sent_messages.append((osc_address_cc, (msg.value,), "control_change", time.time()))

                    elif msg.type == 'aftertouch':
                        osc_address_pressure = f"/ch{channel}pressure"
                        self.osc_client.send_message(osc_address_pressure, msg.value)
                        self.log_message(f"Aftertouch (Ch {channel}): {msg.value}")
                        self.sent_messages.append((osc_address_pressure, (msg.value,), "aftertouch", time.time()))

                    elif msg.type == 'pitchwheel':
                        osc_address_pitch = f"/ch{channel}pitch"
                        self.osc_client.send_message(osc_address_pitch, msg.pitch)
                        self.log_message(f"Pitchwheel (Ch {channel}): {msg.pitch}")
                        self.sent_messages.append((osc_address_pitch, (msg.pitch,), "pitchwheel", time.time()))

        except Exception as e:
            self.log_message(f"Error playing {os.path.basename(file_path)}: {e}")
        finally:
            if not self.playing or self.request_stop:
                self.log_message(f"Finished playing: {os.path.basename(file_path)}")

    def update_tempo_display(self, value):
        new_tempo = int(value)
        self.bpm_display_label.setText(f"{new_tempo} BPM")
        self.change_tempo(new_tempo)

    def change_tempo(self, new_tempo):
        if new_tempo <= 0:
            self.log_message("Invalid tempo.")
            return
        self.tempo = new_tempo
        self.tempo_microseconds_per_beat = 60000000 / new_tempo
        self.log_message(f"Tempo changed to {new_tempo} BPM.")

    def reset_tempo(self):
        if self.midi_playlist:
            try:
                first_midi = MidiFile(self.midi_playlist[0])
                tempo_found = False
                for track in first_midi.tracks:
                    for msg in track:
                        if msg.type == 'set_tempo':
                            self.tempo_microseconds_per_beat = msg.tempo
                            self.tempo = int(60000000 / msg.tempo)
                            self.tempo_slider.setValue(self.tempo)
                            self.bpm_display_label.setText(f"{self.tempo} BPM")
                            self.log_message(f"Tempo reset to {self.tempo} BPM from {os.path.basename(self.midi_playlist[0])}.")
                            tempo_found = True
                            break
                    if tempo_found:
                        break
                if not tempo_found:
                    self.log_message("No tempo event found. Using default 120 BPM.")
                    self.tempo = 120
                    self.tempo_microseconds_per_beat = 500000
                    self.tempo_slider.setValue(self.tempo)
                    self.bpm_display_label.setText("120 BPM")
            except Exception as e:
                self.log_message(f"Error resetting tempo: {e}")
        else:
            self.log_message("No playlist loaded to reset tempo from.")

    def update_piano_roll_channel(self):
        if self.piano_roll_window and not self.ch16_checkbox.isChecked():
            channel = int(self.channel_menu.currentText())
            self.piano_roll_window.channel = channel
            self.log_message(f"PianoRoll channel updated to {channel}")

    def handle_dropped_files(self, file_list):
        for path in file_list:
            if path.lower().endswith(".mid"):
                self.midi_playlist.append(path)
                self.log_message(f"Added via drag-drop: {path}")
            else:
                self.log_message(f"Invalid file type: {path}")
        self.update_playlist_label()

    def stop_server(self):
        with self.lock:
            self.log_message("Stopping OSC server...")

            # Save references and reset them within the lock
            osc_server = self.osc_server
            osc_server_thread = self.osc_server_thread
            midi_input_thread = self.midi_input_thread
            midi_out = self.midi_out
            midi_in = self.midi_in

            self.osc_server = None
            self.osc_server_thread = None
            self.midi_input_thread = None
            self.midi_out = None
            self.midi_in = None
            self.osc_client = None

            if self.piano_roll_window:
                self.piano_roll_window.osc_client = None

            self.start_button.setText("Start Server")

        # Perform shutdown operations outside the lock to prevent deadlocks
        if osc_server:
            try:
                osc_server.shutdown()
                osc_server.server_close()
                self.log_message("OSC server stopped.")
            except Exception as e:
                self.log_message(f"Error stopping OSC server: {e}")

        if osc_server_thread and osc_server_thread.is_alive():
            osc_server_thread.join(timeout=1)
            self.log_message("OSC server thread joined.")

        if midi_input_thread and midi_input_thread.is_alive():
            midi_input_thread.join(timeout=1)
            self.log_message("MIDI input thread joined.")

        if midi_out:
            try:
                midi_out.close()
                self.log_message("MIDI output closed.")
            except Exception as e:
                self.log_message(f"Error closing MIDI output: {e}")

        if midi_in:
            try:
                midi_in.close()
                self.log_message("MIDI input closed.")
            except Exception as e:
                self.log_message(f"Error closing MIDI input: {e}")

    def toggle_looping(self, state):
        with self.lock:
            self.looping = (state == Qt.Checked)
            self.log_message(f"Looping = {self.looping}")

    def quit_app(self):
        try:
            self.stop_server()
            if self.midi_out:
                self.midi_out.close()
            if self.midi_in:
                self.midi_in.close()
        except Exception as e:
            print(f"Error closing resources: {e}")
        self.close()


def main():
    app = QApplication(sys.argv)
    window = OSCMIDIApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
