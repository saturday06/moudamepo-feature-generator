:: -*- coding: shift_jis-dos -*-

cd /d "%~dp0"

if "%1" == "" goto backup_self_and_restart

git checkout master
%COMSPEC% /c build\build.bat
copy README.md backup.txt
move 試験プログラム自動生成.bat backup.bat
git checkout deploy/windows/latest
move backup.bat 試験プログラム自動生成.bat
move backup.txt README.txt
pause
@goto :EOF

:backup_self_and_restart
copy "%~f0" ..\deploy_backup.bat
..\deploy_backup.bat restarted
