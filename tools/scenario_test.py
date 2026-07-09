# -*- coding: utf-8 -*-
"""
기준정보 시나리오 분류 테스트

실행:  python tools/scenario_test.py

검증 항목:
  [1] 하위 호환 — 기존 4종 CSV(data/sample)만으로 로드·실행되고,
      기본 파라미터 결과가 시나리오 도입 이전 엔진과 동일(회귀 기준값 고정)
  [2] 시나리오 감지 — 시나리오 데이터가 없으면 각 축에 기본안만 노출
  [3] 시나리오 데이터(data/sample_scenario) — 2×2×3 = 12개 조합 감지
  [4] base/fixed/M 조합은 기존 데이터 결과와 완전 동일
  [5] 각 축이 실제로 결과를 바꾸는지 (worst 수요 → 비용↑ 등)
  [6] 발전원별 폴백 — V 컬럼이 없는 수력/SMR 단가는 *_M 과 동일
  [7] 없는 시나리오 요청 시 명확한 오류
"""

import os
import sys

THIS_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.dirname(THIS_DIR)
sys.path.insert(0, BASE_DIR)

from core.config import SimulationParams
from core.engine import load_data, run_check, company_total_cost

SAMPLE_DIR = os.path.join(BASE_DIR, "data", "sample")
SCENARIO_DIR = os.path.join(BASE_DIR, "data", "sample_scenario")

# 시나리오 도입 이전(원본 엔진) data/sample 기본 파라미터 25년 총비용 (백만원)
LEGACY_TOTAL_25Y = 11556160.534801122


def total_cost(bundle, **scn):
    params = SimulationParams(**scn)
    return company_total_cost(bundle, params, params.fixed_ppas)[0]


def main():
    ok = 0

    # ── [1] 하위 호환 회귀 ──
    bundle = load_data(SAMPLE_DIR)
    tc = total_cost(bundle)
    assert abs(tc - LEGACY_TOTAL_25Y) < 1e-3, f"회귀 실패: {tc} != {LEGACY_TOTAL_25Y}"
    run_check(bundle, SimulationParams())   # 예외 없이 전체 경로 실행
    ok += 1
    print(f"[1] 하위 호환 OK — 기존 입력 총비용 동일 ({tc:,.0f} 백만원)")

    # ── [2] 기존 데이터 = 기본 시나리오만 노출 ──
    opts = bundle.scenario_options()
    assert [o["code"] for o in opts["usage"]] == ["base"], opts
    assert [o["code"] for o in opts["ppa"]] == ["fixed"], opts
    assert [o["code"] for o in opts["smp"]] == ["M"], opts
    ok += 1
    print("[2] 시나리오 감지 OK — 기존 데이터는 base/fixed/M 만 노출")

    # ── [7] 없는 시나리오 요청 → 명확한 오류 ──
    for kw in ({"usage_scenario": "worst"}, {"ppa_scenario": "variable"}):
        try:
            total_cost(bundle, **kw)
            raise AssertionError(f"{kw} 가 오류 없이 실행됨")
        except ValueError as e:
            assert "시나리오" in str(e)
    ok += 1
    print("[7] 오류 처리 OK — 없는 시나리오는 안내 메시지와 함께 거부")

    # ── 시나리오 데모 데이터 (없으면 생성) ──
    if not os.path.exists(os.path.join(SCENARIO_DIR, "Hourly_Data.csv")):
        sys.path.insert(0, THIS_DIR)
        import generate_scenario_data
        generate_scenario_data.generate(SAMPLE_DIR, SCENARIO_DIR)

    sb = load_data(SCENARIO_DIR)

    # ── [3] 12개 조합 감지 ──
    o = sb.scenario_options()
    n = len(o["usage"]) * len(o["ppa"]) * len(o["smp"])
    assert n == 12, f"조합 수 {n} != 12: {o}"
    ok += 1
    print(f"[3] 조합 감지 OK — 2×2×3 = {n}개")

    # ── [4] 기본 조합은 기존 결과와 동일 ──
    tc_base = total_cost(sb)
    assert abs(tc_base - LEGACY_TOTAL_25Y) < 1e-3, tc_base
    ok += 1
    print("[4] 기본 조합 OK — base/fixed/M 결과가 기존과 동일")

    # ── [5] 각 축이 결과를 바꾸는지 ──
    tc_worst = total_cost(sb, usage_scenario="worst")
    assert tc_worst > tc_base, (tc_worst, tc_base)
    tc_var = total_cost(sb, ppa_scenario="variable")
    assert abs(tc_var - tc_base) > 1.0, (tc_var, tc_base)
    tc_h = total_cost(sb, smp_scenario="H")
    tc_l = total_cost(sb, smp_scenario="L")
    assert abs(tc_h - tc_l) > 1e-6, "SMP 시나리오가 결과에 반영되지 않음"
    ok += 1
    print(f"[5] 축별 반영 OK — worst {tc_worst:,.0f} > base {tc_base:,.0f} / "
          f"PPA변동 Δ{tc_var - tc_base:+,.0f} / SMP H-L Δ{tc_h - tc_l:+,.0f}")

    # ── [6] 발전원별 폴백 (수력/SMR 은 V 컬럼 없음 → M 단가) ──
    var_prices = sb.ppa_scenarios["variable"]
    fix_prices = sb.ppa_scenarios["fixed"]
    assert var_prices["수력"] == fix_prices["수력"]
    assert var_prices["SMR"] == fix_prices["SMR"]
    assert var_prices["태양광"] != fix_prices["태양광"]
    ok += 1
    print("[6] 폴백 OK — 수력/SMR 은 *_M 단가 사용, 태양광은 *_V 반영")

    print(f"\n✅ 시나리오 테스트 {ok}/7 통과")


if __name__ == "__main__":
    main()
