"""
RE100 PPA 시스템 — 샘플 데이터 생성기

실제 데이터(회사 기밀)는 외부 반출이 어려우므로, 코드가 요구하는 것과
'동일한 컬럼 구조'를 가진 가짜 CSV 4종을 생성한다.
이 샘플로 대시보드를 개발/검증한 뒤, 사내에서는 같은 형식의 실제 데이터로 교체한다.

생성 파일 (모두 data/sample/ 에 저장):
  - Hourly_Data.csv   : 연도별 8,760시간 발전 패턴 / SMP
  - Annual_Usage.csv  : 연도별 사용량 / 피크 / RE 목표
  - Annual_PPA.csv    : 발전원별 연도별 PPA 단가
  - Annual_Rate.csv   : 한전 요금 단가 / SMP 기준

자세한 컬럼 정의는 docs/DATA_SPEC.md 참조.

실행:  python tools/generate_sample_data.py
"""

import os
import numpy as np
import pandas as pd

### 1. 기본 설정 ###

SEED = 42                      # 재현성 위한 난수 시드
YEARS = list(range(2026, 2051))  # 2026 ~ 2050 (25개 연도)
HOURS_PER_DAY = 24
DAYS_PER_YEAR = 365            # 8,760시간 (윤년 무시, 단순화)

# 출력 폴더: 이 스크립트 기준 ../data/sample/
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "..", "data", "sample")

rng = np.random.default_rng(SEED)


### 2. 시간별 데이터 생성 (Hourly_Data.csv) ###

def _days_in_month(month):
    # 윤년 무시한 단순 일수 (합계 365일)
    return [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month - 1]


def make_hourly_data():
    rows = []
    for year in YEARS:
        for month in range(1, 13):
            for day in range(1, _days_in_month(month) + 1):
                for hour in range(HOURS_PER_DAY):
                    # 태양광: 낮(6~18시) 종모양, 야간 0
                    if 6 <= hour <= 18:
                        pv = np.sin((hour - 6) / 12 * np.pi)        # 0 → 1 → 0
                        pv *= (0.7 + 0.3 * np.sin((month - 3) / 12 * 2 * np.pi))  # 계절성
                        pv = max(0.0, pv + rng.normal(0, 0.05))
                    else:
                        pv = 0.0

                    # 풍력: 평균 0.4 근처 변동, 겨울에 다소 높음
                    wt = 0.4 + 0.15 * np.sin((month - 10) / 12 * 2 * np.pi) + rng.normal(0, 0.12)
                    wt = float(np.clip(wt, 0.0, 1.0))

                    # SMP: 80~160 원/kWh, 낮에 높음
                    smp = 110 + 25 * np.sin((hour - 9) / 24 * 2 * np.pi) + rng.normal(0, 8)
                    smp = float(np.clip(smp, 60, 200))

                    # 요금 조정계수: 계절·시간대별 0.8~1.3
                    base_factor = 1.0 + 0.15 * np.sin((hour - 14) / 24 * 2 * np.pi)
                    season = 1.0 + 0.1 * np.sin((month - 7) / 12 * 2 * np.pi)

                    rows.append({
                        "Year": year,
                        "Month": month,
                        "Day": day,
                        "Hour": hour,
                        "PV_GENERATION_2": round(float(pv), 4),
                        "WT_GENERATION": round(wt, 4),
                        "SMP": round(smp, 2),
                        "Factor_ADJ": round(base_factor * season, 4),
                        "고압B 선택2": round(base_factor * season * 1.00, 4),  # IC(이천)
                        "고압B 선택3": round(base_factor * season * 1.03, 4),  # CJ(청주)
                        "고압C 선택2": round(base_factor * season * 0.98, 4),  # YI(용인)
                    })
    return pd.DataFrame(rows)


### 3. 연도별 사용량 데이터 생성 (Annual_Usage.csv) ###

def make_annual_usage():
    rows = []
    for i, year in enumerate(YEARS):
        # 전사 사용량: 약 200만 MWh, 매년 1% 소폭 증가
        skh = 2_000_000 * (1.01 ** i)

        # 사업장 IC/CJ/YI 분할 (합이 전사와 유사하도록)
        ic, cj, yi = skh * 0.40, skh * 0.35, skh * 0.25

        # RE 목표 비율: 2026년 33% → 2050년 100% 선형 증가
        progress = i / (len(YEARS) - 1)
        re33 = 0.33 + (1.00 - 0.33) * progress
        re44 = 0.44 + (1.00 - 0.44) * progress

        rows.append({
            "Year": year,
            "SKH": round(skh, 1),
            "SKH_Peak": round(skh / 8760 * 1.3, 2),       # 피크 (MW 스케일)
            "SKH_Peak_Min": round(skh / 8760 * 1.0, 2),   # 망이용 기준 + 시간당 사용량
            "A": round(ic, 1),
            "B": round(cj, 1),
            "C": round(yi, 1),
            "A_Peak": round(ic / 8760 * 1.3, 2),
            "B_Peak": round(cj / 8760 * 1.3, 2),
            "C_Peak": round(yi / 8760 * 1.3, 2),
            "A_Peak_Min": round(ic / 8760 * 1.0, 2),
            "B_Peak_Min": round(cj / 8760 * 1.0, 2),
            "C_Peak_Min": round(yi / 8760 * 1.0, 2),
            "SEC": round(skh * 0.25, 1),
            "RE33": round(re33 - 0.02, 4),       # SEC 가동중단 기준 (소폭 낮음)
            "RE33_SEC": round(re33, 4),          # SEC 계속가동 기준 (기본값)
            "RE44": round(re44 - 0.02, 4),
            "RE44_SEC": round(re44, 4),
        })
    return pd.DataFrame(rows)


### 4. PPA 단가 데이터 생성 (Annual_PPA.csv) ###

def make_annual_ppa():
    # 발전원별 2026년 기준단가(원/kWh)와 연 변동률
    base = {"PV_M": 140, "On_W_M": 160, "Off_W_M": 200, "HP_M": 90, "SMR_M": 120}
    trend = {"PV_M": -0.5, "On_W_M": -0.3, "Off_W_M": -0.8, "HP_M": 0.2, "SMR_M": 0.0}

    rows = []
    for i, year in enumerate(YEARS):
        row = {"Year": year}
        for col, b in base.items():
            row[col] = round(b + trend[col] * i + rng.normal(0, 1.5), 2)
        rows.append(row)
    return pd.DataFrame(rows)


### 5. 한전 요금 데이터 생성 (Annual_Rate.csv) ###

def make_annual_rate():
    rows = []
    for i, year in enumerate(YEARS):
        vc = 110 + 1.2 * i + rng.normal(0, 1)    # 전력량요금, 매년 상승
        rows.append({
            "Year": year,
            "VC_M": round(vc, 2),
            "FC_M": round(25 + 0.3 * i, 2),       # 기후환경 + 연료비조정
            "BC_M": round(7500 + 30 * i, 1),      # 기본요금 (원/kW)
            "EAC_M": round(50 + 0.8 * i, 2),      # EAC 단가
            "SMP_M": 1.0,                          # SMP 기준 배수
            "UC_IC": round(vc * 1.00, 2),         # 사업장별 전력량요금
            "UC_CJ": round(vc * 1.03, 2),
            "UC_YI": round(vc * 0.98, 2),
        })
    return pd.DataFrame(rows)


### 6. 메인 실행 ###

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    datasets = {
        "Hourly_Data.csv": make_hourly_data,
        "Annual_Usage.csv": make_annual_usage,
        "Annual_PPA.csv": make_annual_ppa,
        "Annual_Rate.csv": make_annual_rate,
    }

    for filename, builder in datasets.items():
        print(f"생성 중: {filename} ...", end=" ", flush=True)
        df = builder()
        path = os.path.join(OUTPUT_DIR, filename)
        df.to_csv(path, index=False, encoding="utf-8-sig")
        print(f"완료 ({len(df):,}행, {len(df.columns)}열)")

    print(f"\n✅ 샘플 데이터 4종을 생성했습니다.")
    print(f"   위치: {os.path.abspath(OUTPUT_DIR)}")


if __name__ == "__main__":
    main()
