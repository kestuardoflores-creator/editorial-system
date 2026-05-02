@echo off
cd /d "%~dp0"
python -m pip install -q -r app\assembler\requirements.txt
python app\launcher.py
