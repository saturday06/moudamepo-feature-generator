:: -*- coding: shift_jis-dos -*-

cd /d "%~dp0"
%COMSPEC% /c ..\build\build.bat

cd /d "%~dp0"
%COMSPEC% /c ..\試験プログラム自動生成.bat "%~dp0input" "%~dp0got"

