@echo off
rem 웹 대시보드 실행 — 브라우저에서 기간 지정 후 [분석 시작] 버튼으로 실행
chcp 65001 > nul
cd /d "%~dp0"
python dashboard_server.py
pause
