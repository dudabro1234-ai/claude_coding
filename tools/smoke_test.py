"""
엔진 스모크 테스트 — Python 설치 직후 전체가 돌아가는지 한 번에 확인한다.

실행:  python tools/smoke_test.py

순서:
  1) 샘플 데이터가 없으면 자동 생성
  2) load_data → run_check / run_site / run_greedy 를 차례로 호출
  3) 결과 형태와 주요 수치를 출력 (에러 없이 끝나면 통과)
"""

import os
import sys
import time

THIS_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.dirname(THIS_DIR)          # ppa-dashboard/
sys.path.insert(0, BASE_DIR)                  # core 패키지 import용
sys.path.insert(0, THIS_DIR)                  # generate_sample_data import용

from core.engine import load_data, run_check, run_site, run_greedy
from core.config import SimulationParams

SAMPLE_DIR = os.path.join(BASE_DIR, "data", "sample")


def ensure_sample():
    if not os.path.exists(os.path.join(SAMPLE_DIR, "Hourly_Data.csv")):
        print("샘플 데이터가 없어 생성합니다...\n")
        import generate_sample_data
        generate_sample_data.main()
        print()


def main():
    ensure_sample()

    print("=" * 60)
    print("데이터 로드 중...")
    t = time.time()
    bundle = load_data(SAMPLE_DIR)
    print(f"  로드 완료 ({time.time() - t:.1f}초) — {len(bundle.hourly)}개 연도\n")

    # --- Check ---
    print("=" * 60)
    print("[1/3] run_check — 포트폴리오 비용 검증")
    t = time.time()
    check = run_check(bundle, SimulationParams())
    last = check["cost"].iloc[-1]
    print(f"  완료 ({time.time() - t:.1f}초). 비용표 {check['cost'].shape}, RE표 {check['re'].shape}")
    print(f"  2050년 총비용: {last['총비용 (백만원)']:,.0f} 백만원\n")

    # --- Site ---
    print("=" * 60)
    print("[2/3] run_site — A(이천)/B(청주)/C(용인) 사업장 배분")
    t = time.time()
    site = run_site(bundle, SimulationParams())
    print(f"  완료 ({time.time() - t:.1f}초). 요약 {site['summary'].shape}, 배분 {site['allocation'].shape}")
    by_site = site["allocation"].groupby("Site")["MW"].sum()
    print(f"  사업장별 누적 배분(MW): {by_site.to_dict()}\n")

    # --- Greedy (작은 포트폴리오로 빠르게) ---
    print("=" * 60)
    print("[3/3] run_greedy — 포트폴리오 자동 최적화 (작은 초기값)")
    small = SimulationParams(fixed_ppas=[
        {"source": "태양광", "capacity": 100, "start_year": 2027},
        {"source": "수력", "capacity": 18, "start_year": 2026},
    ])
    t = time.time()
    greedy = run_greedy(bundle, small, progress=lambda s, m: print(f"   · {m}"))
    print(f"  완료 ({time.time() - t:.1f}초). 최종 계약 {len(greedy['portfolio'])}건, "
          f"추가 {len(greedy['optimization_log'])}건\n")

    print("=" * 60)
    print("✅ 스모크 테스트 통과 — 엔진이 정상 동작합니다.")


if __name__ == "__main__":
    main()
