@echo off
chcp 65001 >nul
cd /d "%~dp0"
rem 先取得線上新增的資料，避免本機與 GitHub 版本衝突
git pull --ff-only --quiet 2>nul || echo （無法從 GitHub 更新，將使用本機資料）
python server.py
if errorlevel 1 pause
