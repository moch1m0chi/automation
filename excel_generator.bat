@echo off
cd /d %~dp0
python excel_generator.py --input data --output output
pause