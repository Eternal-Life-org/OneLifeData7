@echo off
powershell -command "Remove-Item -Force -ErrorAction SilentlyContinue animations\*.fcz, categories\*.fcz, objects\*.fcz, sounds\*.fcz, sprites\*.fcz, transitions\*.fcz"
echo 已删除所有 .fcz 文件。
pause
