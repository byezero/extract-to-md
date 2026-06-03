@echo off
set "PATH=%~dp0;%PATH%"
python "%~dp0..\src\extract_to_md\cli.py" %*
