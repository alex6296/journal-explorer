#!/bin/bash
# Double-click this to launch Journal Explorer on an Intel Mac (not Apple Silicon - use the other
# .command file for M1/M2/M3/M4 Macs).
# First time only: macOS will refuse to open it ("unidentified developer") - right-click this file
# (or JournalExplorer-mac-intel) and choose Open once, confirm, and it will run normally after that.
cd "$(dirname "$0")"
chmod +x ./JournalExplorer-mac-intel
./JournalExplorer-mac-intel
