@echo off
rem 사내 데이터플랫폼 자동 수집 (Phase C, opt-in)
rem 최초 1회: pip install -r requirements-connector.txt & python -m playwright install chromium
rem 그리고 platform_nav.example.json을 platform_nav.json으로 복사해 클릭 경로 설정
chcp 65001 > nul
cd /d "%~dp0"
python platform_connector.py %*
if errorlevel 1 (
  echo.
  echo [안내] 자동 수집에 실패했습니다. 플랫폼에서 받은 xlsx를
  echo        data\platform_drop\ 에 직접 넣고 아래를 실행해도 됩니다:
  echo        python tools\platform_ingest.py
  pause
)
