@echo off
cd /d %~dp0
"C:\project\automation\excel_generator.py" --input data --output output > log.txt 2>&1
pause