@echo off
rem 고객별 RE 할당 시뮬레이션 (목업) — 단독 실행 (서버 불필요)
chcp 65001 > nul
cd /d "%~dp0"
python re_simulator.py
