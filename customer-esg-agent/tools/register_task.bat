@echo off
rem Windows 작업 스케줄러 등록 (Phase 3) — 매일 08:30 자동 실행
rem 관리자 권한이 필요할 수 있습니다.
chcp 65001 > nul
set AGENT_DIR=%~dp0..
schtasks /Create /TN "CustomerESGAgent" /SC DAILY /ST 08:30 ^
  /TR "\"%AGENT_DIR%\run_agent.bat\" --no-browser" /F
if errorlevel 1 (
  echo [오류] 작업 등록에 실패했습니다. 관리자 권한으로 다시 실행해 보세요.
) else (
  echo 등록 완료: 매일 08:30에 CustomerESGAgent가 실행됩니다.
  echo 해제하려면: schtasks /Delete /TN "CustomerESGAgent" /F
)
pause
