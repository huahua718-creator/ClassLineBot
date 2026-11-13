@echo off
chcp 65001 > nul
echo ========================================================
echo   正在從 Google Apps Script 拉取最新程式碼...
echo   專案 ID: 1BxuxNQMgRHFaZ3NNoW1XovwMUQDAFFMrVNtk9Y_leVvokG6-nAojY3U6
echo ========================================================
npx @google/clasp pull
pause