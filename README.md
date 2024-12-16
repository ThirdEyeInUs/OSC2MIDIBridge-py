# OSC2MIDIBridge-py
Turn Midi into OSC Messages (PatchWorld Project)

Install requirements pip install PySide6 mido python-osc python-rtmidi

The osc addresses are as follows...(X = channel 1-16)

/chXnote 0-127, /chXnoteoff 0-127, /chXnoffvalue 0-1 /chXpitch,-8200 to 8200 (sits at 0), /chX pressure 0-127 /chXcc 0-127, /chXccvalue 0-1
