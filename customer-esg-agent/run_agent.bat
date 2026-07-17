@echo off
rem 고객 ESG 대응 Agent 진입점 (수동 실행 / 작업 스케줄러 공용)
chcp 65001 > nul
cd /d "%~dp0"
python main.py %*
if errorlevel 1 (
  echo.
  echo [오류] 실행이 비정상 종료되었습니다. logs\ 폴더의 오늘자 로그를 확인하세요.
  if "%1"=="" pause
)
