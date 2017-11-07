@echo off

setlocal ENABLEDELAYEDEXPANSION

echo UUID and computer name will be saved in list.txt file

set /p name=Enter computer name:

for /f "usebackq" %%i in (`wmic csproduct get uuid ^| find "-"`) do echo !name!;%%i >>list.txt

echo Done.

goto :eof
