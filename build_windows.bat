@echo off
REM Builds dist\JournalExplorer.exe (Windows). Needs Python installed on THIS machine only.
REM nicegui-pack is NiceGUI's wrapper around PyInstaller: it bundles the web UI files correctly.
pip install -r requirements.txt
nicegui-pack --onefile --name JournalExplorer journal_explorer.py
echo Done - the app is dist\JournalExplorer.exe
