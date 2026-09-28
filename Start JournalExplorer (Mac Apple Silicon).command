#!/bin/bash
# Double-click this to launch Journal Explorer on an Apple Silicon Mac (M1/M2/M3/M4 - use the other
# .command file for an older Intel Mac).
# First time only: macOS will refuse to open it ("unidentified developer") - right-click this file
# (or JournalExplorer-mac) and choose Open once, confirm, and it will run normally after that.
cd "$(dirname "$0")"
chmod +x ./JournalExplorer-mac
./JournalExplorer-mac
