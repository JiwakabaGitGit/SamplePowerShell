@echo off
echo 処理を開始します
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& '.\readexcel.ps1'"
echo 完了しました！
pause > nul
exit