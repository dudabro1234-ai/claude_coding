# -*- coding: utf-8 -*-
"""
기준정보 시나리오 데모 데이터 생성기

data/sample(기존 4종 CSV)을 읽어 시나리오 축이 모두 채워진 데모 폴더
data/sample_scenario/ 를 만든다. 기존 샘플 파일은 절대 수정하지 않는다.

    python tools/generate_scenario_data.py            # data/sample → data/sample_scenario
    python tools/generate_scenario_data.py <입력폴더> <출력폴더>

생성 내용 (docs/DATA_SPEC.md 시나리오 확장 규칙과 동일):
  · Annual_Usage.csv        : 원본 복사 (base 시나리오)
  · Annual_Usage_worst.csv  : 사용량·피크를 연차별로 최대 +12%까지 증가시킨 Worst 수요안
  · Annual_PPA.csv          : 원본 + PV_V / On_W_V / Off_W_V (연도별 변동 단가)
                              — HP/SMR 은 V 컬럼을 두지 않아 '발전원별 *_M 폴백'을 시연
  · Annual_Rate.csv         : 원본 + SMP_H(고가) / SMP_L(저가) 배수 곡선
  · Hourly_Data.csv         : 원본 복사 (발전패턴·시간별 SMP 패턴은 시나리오 공통)

실행 후 서버를 이 폴더로 띄우면 2(사용량)×2(PPA)×3(SMP)=12개 조합을 테스트할 수 있다:
    PPA_DATA_DIR=data/sample_scenario python server.py
"""

import os
import shutil
import sys

import pandas as pd

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Worst 수요안: base 대비 증가율이 2026년 0%에서 2050년 +12%까지 선형 확대
WORST_USAGE_MAX_UPLIFT = 0.12
# 연도별 변동 PPA 단가: 체결연도가 늦을수록 단가 하락 (기술 학습곡선 가정)
VARIABLE_PPA_DECLINE = {"PV_V": 0.015, "On_W_V": 0.010, "Off_W_V": 0.020}   # 연 하락률
VARIABLE_PPA_BASE = {"PV_V": "PV_M", "On_W_V": "On_W_M", "Off_W_V": "Off_W_M"}
# SMP 고가/저가안: 기본 배수(SMP_M)에서 2050년까지 ±35% 선형 확대
SMP_SPREAD_2050 = 0.35


def _year_frac(years):
    y0, y1 = min(years), max(years)
    span = max(y1 - y0, 1)
    return [(y - y0) / span for y in years]


def generate(src_dir, out_dir):
    os.makedirs(out_dir, exist_ok=True)

    # 1) Hourly_Data.csv — 발전패턴 공통, 그대로 복사
    shutil.copyfile(os.path.join(src_dir, "Hourly_Data.csv"),
                    os.path.join(out_dir, "Hourly_Data.csv"))

    # 2) Annual_Usage.csv (base) + Annual_Usage_worst.csv
    df_u = pd.read_csv(os.path.join(src_dir, "Annual_Usage.csv"), thousands=",")
    df_u.columns = df_u.columns.str.strip()
    df_u.to_csv(os.path.join(out_dir, "Annual_Usage.csv"), index=False, encoding="utf-8-sig")

    worst = df_u.copy()
    fracs = _year_frac(worst["Year"].tolist())
    scale_cols = [c for c in worst.columns
                  if c != "Year" and not str(c).startswith("RE") and c != "SEC"]
    for i, frac in enumerate(fracs):
        factor = 1.0 + WORST_USAGE_MAX_UPLIFT * frac
        for c in scale_cols:
            worst.loc[worst.index[i], c] = round(float(df_u.iloc[i][c]) * factor, 2)
    worst.to_csv(os.path.join(out_dir, "Annual_Usage_worst.csv"), index=False, encoding="utf-8-sig")

    # 3) Annual_PPA.csv + 연도별 변동(V) 컬럼
    df_p = pd.read_csv(os.path.join(src_dir, "Annual_PPA.csv"))
    df_p.columns = df_p.columns.str.strip()
    fracs_p = [y - int(df_p["Year"].min()) for y in df_p["Year"]]
    for v_col, m_col in VARIABLE_PPA_BASE.items():
        decline = VARIABLE_PPA_DECLINE[v_col]
        base0 = float(df_p.iloc[0][m_col])
        df_p[v_col] = [round(base0 * ((1 - decline) ** n), 2) for n in fracs_p]
    df_p.to_csv(os.path.join(out_dir, "Annual_PPA.csv"), index=False, encoding="utf-8-sig")

    # 4) Annual_Rate.csv + SMP_H / SMP_L 배수 곡선
    df_r = pd.read_csv(os.path.join(src_dir, "Annual_Rate.csv"))
    df_r.columns = df_r.columns.str.strip()
    fracs_r = _year_frac(df_r["Year"].tolist())
    smp_m = df_r["SMP_M"] if "SMP_M" in df_r.columns else pd.Series([1.0] * len(df_r))
    df_r["SMP_H"] = [round(float(m) * (1 + SMP_SPREAD_2050 * f), 4) for m, f in zip(smp_m, fracs_r)]
    df_r["SMP_L"] = [round(float(m) * (1 - SMP_SPREAD_2050 * f), 4) for m, f in zip(smp_m, fracs_r)]
    df_r.to_csv(os.path.join(out_dir, "Annual_Rate.csv"), index=False, encoding="utf-8-sig")

    print(f"[완료] 시나리오 데모 데이터 생성 → {out_dir}")
    print("  · Annual_Usage_worst.csv  (Worst 수요안, 2050년 +12%)")
    print("  · Annual_PPA.csv          (+ PV_V/On_W_V/Off_W_V — HP/SMR은 *_M 폴백)")
    print("  · Annual_Rate.csv         (+ SMP_H/SMP_L 배수 곡선 ±35%)")
    print(f"실행:  PPA_DATA_DIR={os.path.relpath(out_dir, BASE_DIR)} python server.py")


if __name__ == "__main__":
    src = sys.argv[1] if len(sys.argv) > 1 else os.path.join(BASE_DIR, "data", "sample")
    out = sys.argv[2] if len(sys.argv) > 2 else os.path.join(BASE_DIR, "data", "sample_scenario")
    generate(src, out)
