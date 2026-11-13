@echo off
chcp 65001 > nul
echo ========================================================
echo   正在推送程式碼至 Google Apps Script...
echo   專案 ID: 1BxuxNQMgRHFaZ3NNoW1XovwMUQDAFFMrVNtk9Y_leVvokG6-nAojY3U6
echo ========================================================
npx @google/clasp push -f
pause