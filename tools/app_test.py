"""Streamlit AppTest — app.py가 예외 없이 렌더되고 Check 실행이 동작하는지 확인."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from streamlit.testing.v1 import AppTest

APP = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "app.py")

at = AppTest.from_file(APP, default_timeout=60).run()
assert not at.exception, f"초기 렌더 예외: {at.exception}"
print(f"[1] 초기 렌더 OK — 탭 {len(at.tabs)}개, 버튼 {len(at.button)}개, info {len(at.info)}개")

# Check 실행 버튼 누르기 (key=run_check)
btn = next(b for b in at.button if b.key == "run_check")
at = btn.click().run()
assert not at.exception, f"Check 실행 예외: {at.exception}"
metrics = [m.value for m in at.metric]
print(f"[2] Check 실행 OK — metric {len(at.metric)}개, dataframe {len(at.dataframe)}개")
print(f"    headline metrics: {metrics}")

# Site 배분 실행 (피벗·차트 코드 경로 확인)
at = next(b for b in at.button if b.key == "run_site").click().run()
assert not at.exception, f"Site 실행 예외: {at.exception}"
print(f"[3] Site 실행 OK — metric {len(at.metric)}개, dataframe {len(at.dataframe)}개")

# Greedy 최적화 실행 (느림 — 수십 초)
at = next(b for b in at.button if b.key == "run_greedy").click().run()
assert not at.exception, f"Greedy 실행 예외: {at.exception}"
print(f"[4] Greedy 실행 OK — dataframe {len(at.dataframe)}개")
print("✅ AppTest 통과 — 3개 탭 모두 정상")
