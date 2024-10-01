@echo off
setlocal

set python_executable=Python310\pythonw.exe
set python_script=GoWork\fanqiezhong.py


start %python_executable% %python_script%

exit