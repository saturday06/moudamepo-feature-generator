:: -*- coding: shift_jis-dos -*-

cd /d "%~dp0.."
%COMSPEC% /c build\build.bat
rd /s /q deploy
git clone . deploy
cd deploy
git remote set-url origin git@github.com:saturday06/moudamepo-feature-generator.git
git fetch
git checkout deploy/windows/latest
copy ..\README.md README.txt
copy ..\Œ±ƒvƒƒOƒ‰ƒ€©“®¶¬.bat .
pause
