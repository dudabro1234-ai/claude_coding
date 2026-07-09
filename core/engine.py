"""
RE100 PPA 시스템 — 공통 계산 엔진

기존 3개 코드(Greedy / Check / Site)의 로직을 통합한다.
설계 원칙: "발전량·단가 계산 코어는 공유, 비용 집계는 용도별로 분리"
  - compute_generation() : 세 모듈 공통 — 시간별 발전량/PPA비용/발전원별 집계
  - cost_company()       : Check/Greedy 모델 (전사 SKH 단일)
  - cost_site()          : Site 모델 (사업장 IC/CJ/YI 분리)

⚠️ 두 비용 모델은 원본에서 공식이 다르므로(망이용 ×12, 사용량 배분, SMP차감 등)
   기존 결과를 보존하기 위해 차이를 그대로 유지한다. 자세한 차이는 docs/DATA_SPEC.md 참조.

실행 진입점:
  - run_check(bundle, params)  → 연간 비용/RE 조달 상세
  - run_greedy(bundle, params) → 최적 포트폴리오 자동 탐색
  - run_site(bundle, params)   → IC/CJ/YI 사업장 배분
"""

import os
from dataclasses import dataclass

import numpy as np
import pandas as pd

from .config import (
    YEARS, SOURCES, GENERATION_COL_MAP, UTILIZATION_MAP, FLAT_SOURCES,
    SITES, SITE_CODES, COMPANY_USAGE_COL, GRID_USAGE_CHARGE,
    EFF_DECAY, PPA_REF_PERIOD, STEP_MW, MIN_PARTIAL_STEP, EPS,
    NEW_ENTRY_CAP, CAP_GRIDS, EAC_MIN_RATIO, ABBRS,
    USAGE_SCENARIOS, PPA_PRICE_SCENARIOS, SMP_SCENARIOS,
    DEFAULT_USAGE_SCENARIO, DEFAULT_PPA_SCENARIO, DEFAULT_SMP_SCENARIO,
    ppa_price_col, rate_col, SimulationParams, Site,
)


# ════════════════════════════════════════════════════════════════════
# 1. 유틸리티
# ════════════════════════════════════════════════════════════════════

def _normalize(s):
    return str(s).lower().replace(" ", "").replace("-", "")


def find_col(df_or_cols, key, required=True):
    """대소문자·공백·하이픈을 무시하고 컬럼명을 찾는다 (원본 get_colname 통합)."""
    if isinstance(df_or_cols, pd.DataFrame):
        cols = df_or_cols.columns
    elif isinstance(df_or_cols, pd.Series):
        cols = df_or_cols.index          # Series는 인덱스(라벨)에서 찾는다
    else:
        cols = df_or_cols
    norm = _normalize(key)
    for c in cols:
        if _normalize(c) == norm:
            return c
    if required:
        raise KeyError(f"'{key}' 컬럼을 찾을 수 없습니다.")
    return None


def get_ppa_reference_year(year, start_year):
    """20년 주기로 갱신되는 PPA 기준연도(vintage)."""
    if year < start_year:
        return None
    return start_year + ((year - start_year) // PPA_REF_PERIOD) * PPA_REF_PERIOD


# CSV 인코딩 자동 판별 순서 — 사내 실데이터는 엑셀 저장본(CP949)인 경우가 많다.
CSV_ENCODINGS = ("utf-8-sig", "cp949", "utf-8", "euc-kr")

# 엑셀 파일 매직 바이트 — 확장자가 아니라 내용으로 판별한다.
# (엑셀 통합문서를 이름만 *.csv 로 저장/변경한 경우도 자동으로 엑셀로 읽음)
_XLSX_MAGIC = b"PK\x03\x04"          # xlsx (zip 컨테이너)
_XLS_MAGIC = b"\xd0\xcf\x11\xe0"     # 구형 xls (OLE2)

# 입력 파일 탐색 확장자 우선순위 (find_input_file)
INPUT_EXTS = (".csv", ".xlsx", ".xls")


def _clean_columns(df):
    df.columns = [str(c).replace("﻿", "").strip() for c in df.columns]
    return df


def _read_excel_kr(path, engine):
    """엑셀 통합문서(첫 번째 시트)를 읽는다. openpyxl 미설치 시 설치 방법 안내."""
    try:
        df = pd.read_excel(path, sheet_name=0, engine=engine)
    except ImportError as e:
        pkg = "openpyxl" if engine == "openpyxl" else "xlrd"
        raise ImportError(
            f"'{os.path.basename(path)}' 는 엑셀 통합문서 형식입니다. "
            f"읽으려면 {pkg} 패키지가 필요합니다:  pip install {pkg}  (원인: {e})")
    return _clean_columns(df)


def read_csv_kr(path, **kwargs):
    """입력 파일 안전 로더 — 파일 '내용'을 보고 형식을 자동 판별한다.

    · 엑셀 통합문서(xlsx/xls): 확장자와 무관하게 매직 바이트로 감지해 첫 시트를 읽음
      (이름만 *.csv 인 엑셀 파일도 정상 로드)
    · CSV: UTF-8(BOM 포함)·CP949/EUC-KR(엑셀 'CSV' 저장본) 인코딩 자동 판별
    컬럼명의 BOM·앞뒤 공백도 함께 정리하므로, 어떤 방식으로 저장한 파일이든
    같은 컬럼명 규칙(DATA_SPEC.md)으로 매칭된다.
    """
    with open(path, "rb") as f:
        head = f.read(8)
    if head.startswith(_XLSX_MAGIC):
        return _read_excel_kr(path, "openpyxl")
    if head.startswith(_XLS_MAGIC):
        return _read_excel_kr(path, "xlrd")

    last_err = None
    for enc in CSV_ENCODINGS:
        try:
            df = pd.read_csv(path, encoding=enc, **kwargs)
        except (UnicodeDecodeError, UnicodeError) as e:
            last_err = e
            continue
        return _clean_columns(df)
    raise UnicodeError(
        f"'{os.path.basename(path)}' 파일의 문자 인코딩을 인식할 수 없습니다 "
        f"(시도한 인코딩: {', '.join(CSV_ENCODINGS)}). "
        f"엑셀에서 'CSV UTF-8(쉼표로 분리)' 형식으로 다시 저장한 뒤 시도해 보세요. 원인: {last_err}")


def find_input_file(data_dir, stem, required=True):
    """논리 이름(stem)의 실제 입력 파일 경로. Hourly_Data.csv → 없으면 .xlsx → .xls 순."""
    for ext in INPUT_EXTS:
        p = os.path.join(data_dir, stem + ext)
        if os.path.exists(p):
            return p
    if required:
        raise FileNotFoundError(
            f"'{data_dir}' 폴더에서 {stem}.csv (또는 {stem}.xlsx / {stem}.xls) 파일을 찾을 수 없습니다.")
    return None


# ════════════════════════════════════════════════════════════════════
# 2. 데이터 로드
# ════════════════════════════════════════════════════════════════════

@dataclass
class DataBundle:
    hourly: dict            # {year: DataFrame}
    usage_scenarios: dict   # {usage_scenario: {year: Series}}
    rate: dict              # {year: Series}
    ppa_scenarios: dict     # {ppa_scenario: {source: {ref_year: price}}}
    smp_rate_cols: dict     # {smp_scenario: Annual_Rate.csv 배수 컬럼명}
    smp_col: str
    factor_adj_col: str

    # ── 기본 시나리오 별칭 (기존 코드 호환) ──
    @property
    def usage(self):
        return self.usage_scenarios[DEFAULT_USAGE_SCENARIO]

    @property
    def price_dicts(self):
        return self.ppa_scenarios[DEFAULT_PPA_SCENARIO]

    # ── params 의 시나리오 선택 해석 (없는 시나리오는 명확한 오류) ──
    def usage_for(self, params):
        key = getattr(params, "usage_scenario", DEFAULT_USAGE_SCENARIO) or DEFAULT_USAGE_SCENARIO
        if key not in self.usage_scenarios:
            raise ValueError(
                f"사용량 시나리오 '{key}' 데이터가 없습니다. "
                f"(가능: {sorted(self.usage_scenarios)} — Annual_Usage_{key}.csv 를 데이터 폴더에 추가하세요)")
        return self.usage_scenarios[key]

    def prices_for(self, params):
        key = getattr(params, "ppa_scenario", DEFAULT_PPA_SCENARIO) or DEFAULT_PPA_SCENARIO
        if key not in self.ppa_scenarios:
            suffix = PPA_PRICE_SCENARIOS.get(key, {}).get("suffix", "?")
            raise ValueError(
                f"PPA 단가 시나리오 '{key}' 데이터가 없습니다. "
                f"(가능: {sorted(self.ppa_scenarios)} — Annual_PPA.csv 에 *_{suffix} 컬럼을 추가하세요)")
        return self.ppa_scenarios[key]

    def smp_rate_col_for(self, params):
        """선택한 SMP 시나리오의 Annual_Rate.csv 배수 컬럼명. (기본 M은 컬럼이 없어도 허용 → 배수 1)"""
        key = getattr(params, "smp_scenario", DEFAULT_SMP_SCENARIO) or DEFAULT_SMP_SCENARIO
        if key not in self.smp_rate_cols:
            raise ValueError(
                f"SMP 시나리오 '{key}' 데이터가 없습니다. "
                f"(가능: {sorted(self.smp_rate_cols)} — Annual_Rate.csv 에 SMP_{key} 컬럼을 추가하세요)")
        return self.smp_rate_cols[key]

    def scenario_options(self):
        """데이터에 실제 존재하는 시나리오 목록 (대시보드 노출용)."""
        return {
            "usage": [{"code": c, "label": USAGE_SCENARIOS.get(c, c)}
                      for c in USAGE_SCENARIOS if c in self.usage_scenarios],
            "ppa": [{"code": c, "label": PPA_PRICE_SCENARIOS[c]["label"]}
                    for c in PPA_PRICE_SCENARIOS if c in self.ppa_scenarios],
            "smp": [{"code": c, "label": SMP_SCENARIOS.get(c, c)}
                    for c in SMP_SCENARIOS if c in self.smp_rate_cols],
        }


def _usage_rows(df):
    return {int(r["Year"]): r for _, r in df.iterrows()}


def load_data(data_dir):
    """4종 CSV(+선택적 시나리오 데이터)를 읽어 DataBundle로 반환. (DATA_SPEC.md 형식 가정)

    시나리오 확장 규칙 — 모두 '있으면 로드, 없으면 생략(기본만)':
      · 사용량   : Annual_Usage_{시나리오}.csv (예: Annual_Usage_worst.csv)
      · PPA 단가 : Annual_PPA.csv 의 접미사 컬럼 (예: PV_V) — 없는 발전원은 *_M 폴백
      · SMP      : Annual_Rate.csv 의 SMP_H / SMP_L 배수 컬럼
    기존 4종 CSV만 있는 폴더는 이전 버전과 완전히 동일하게 동작한다.
    """
    # 형식(CSV/엑셀)·인코딩 자동 판별 로더 사용 — UTF-8/CP949 CSV, xlsx/xls 모두 그대로 로드.
    # 파일은 .csv → .xlsx → .xls 순으로 찾는다 (컬럼명 BOM·공백 정리는 read_csv_kr 일괄 처리).
    df_hourly = read_csv_kr(find_input_file(data_dir, "Hourly_Data"))
    df_usage = read_csv_kr(find_input_file(data_dir, "Annual_Usage"), thousands=",")
    df_ppa = read_csv_kr(find_input_file(data_dir, "Annual_PPA"), thousands=",")
    df_rate = read_csv_kr(find_input_file(data_dir, "Annual_Rate"), thousands=",")

    smp_col = find_col(df_hourly, "SMP")
    factor_adj_col = find_col(df_hourly, "Factor_ADJ")

    hourly = {int(y): sub.reset_index(drop=True) for y, sub in df_hourly.groupby("Year")}
    rate = {int(r["Year"]): r for _, r in df_rate.iterrows()}

    # ── 사용량 시나리오: base = Annual_Usage.csv, 그 외 = Annual_Usage_{코드}.csv ──
    usage_scenarios = {DEFAULT_USAGE_SCENARIO: _usage_rows(df_usage)}
    for code in USAGE_SCENARIOS:
        if code == DEFAULT_USAGE_SCENARIO:
            continue
        path = find_input_file(data_dir, f"Annual_Usage_{code}", required=False)
        if path:
            usage_scenarios[code] = _usage_rows(read_csv_kr(path, thousands=","))

    # ── PPA 단가 시나리오: 접미사 컬럼 (발전원별로 없으면 기본 *_M 폴백) ──
    ppa_scenarios = {}
    for code, spec in PPA_PRICE_SCENARIOS.items():
        prices, found_own = {}, False
        for src in SOURCES:
            col = find_col(df_ppa, f"{ABBRS[src]}_{spec['suffix']}", required=False)
            if col is not None and code != DEFAULT_PPA_SCENARIO:
                found_own = True
            if col is None:
                col = find_col(df_ppa, ppa_price_col(src))   # 기본(fixed, *_M) 폴백
            prices[src] = dict(zip(df_ppa["Year"], df_ppa[col]))
        if code == DEFAULT_PPA_SCENARIO or found_own:
            ppa_scenarios[code] = prices

    # ── SMP 시나리오: Annual_Rate.csv 의 SMP_{코드} 배수 컬럼 ──
    smp_rate_cols = {}
    for code in SMP_SCENARIOS:
        col = find_col(df_rate, f"SMP_{code}", required=False)
        if col is not None:
            smp_rate_cols[code] = col
    # 기본안(M)은 컬럼이 없어도 항상 선택 가능 (원본과 동일하게 배수 1 처리)
    smp_rate_cols.setdefault(DEFAULT_SMP_SCENARIO, rate_col("SMP"))

    return DataBundle(hourly, usage_scenarios, rate, ppa_scenarios,
                      smp_rate_cols, smp_col, factor_adj_col)


# ════════════════════════════════════════════════════════════════════
# 3. 발전량·PPA비용 계산 코어 (전사/사업장 공통)
# ════════════════════════════════════════════════════════════════════

@dataclass
class GenResult:
    gen_mwh: np.ndarray            # 시간별 총 재생에너지 발전량 (MWh)
    ppa_cost_won: np.ndarray       # 시간별 PPA 총비용 (원)
    flat_pv: np.ndarray            # 균등태양광 발전량 (MWh)
    flat_wt: np.ndarray            # 균등풍력 발전량 (MWh)
    total_capacity: float          # 유효 PPA 총 용량 (MW)
    gen_by_source: dict            # {source: 시간별 발전량 array}
    cap_by_source: dict            # {source: 용량 합}
    price_sum_by_source: dict      # {source: price*capacity 합} (가중평균단가용)


def compute_generation(df_h, portfolio, year, price_dicts, mode):
    """
    포트폴리오의 시간별 발전량과 PPA 총비용(원)을 계산한다.

    mode='company' : Check/Greedy — 계약은 'start_year' 사용, 종료계약은 기준연도 고정,
                     효율감소는 기준연도 기준 (20년 주기 리셋)
    mode='site'    : Site — 계약은 'vintage' 사용, 효율감소는 vintage 고정 (리셋 없음)
    """
    n = len(df_h)
    gen_mwh = np.zeros(n)
    ppa_cost_won = np.zeros(n)
    flat_pv = np.zeros(n)
    flat_wt = np.zeros(n)
    total_capacity = 0.0
    gen_by_source, cap_by_source, price_sum_by_source = {}, {}, {}

    for ppa in portfolio:
        src = ppa["source"]
        cap = ppa["capacity"]
        start = ppa.get("start_year", ppa.get("vintage"))
        vintage = ppa.get("vintage", start)
        end = ppa.get("end_year")

        if mode == "company":
            # 유효성: 시작연도 도래 + 종료연도 이내
            if not (start <= year and (end is None or year <= end)):
                continue
            ref = start if end is not None else get_ppa_reference_year(year, start)
            eff = EFF_DECAY ** (year - ref)
        else:  # site
            if end is not None and year > end:
                continue
            ref = get_ppa_reference_year(year, vintage)
            eff = EFF_DECAY ** (year - vintage)

        gen_col = GENERATION_COL_MAP.get(src)
        pattern = np.ones(n) if gen_col is None else df_h[gen_col].to_numpy()
        this_gen = cap * UTILIZATION_MAP[src] * pattern * eff
        price = price_dicts[src].get(ref, 0)

        gen_mwh += this_gen
        ppa_cost_won += this_gen * 1000.0 * price   # MWh→kWh × 원/kWh = 원
        total_capacity += cap
        if src == "균등태양광":
            flat_pv += this_gen
        if src == "균등풍력":
            flat_wt += this_gen

        gen_by_source[src] = gen_by_source.get(src, np.zeros(n)) + this_gen
        cap_by_source[src] = cap_by_source.get(src, 0) + cap
        price_sum_by_source[src] = price_sum_by_source.get(src, 0) + price * cap

    return GenResult(gen_mwh, ppa_cost_won, flat_pv, flat_wt,
                     total_capacity, gen_by_source, cap_by_source, price_sum_by_source)


# ════════════════════════════════════════════════════════════════════
# 4. 비용 집계 — 전사 모델 (Check / Greedy)
# ════════════════════════════════════════════════════════════════════

def _smp_multiplier(rate_row, params):
    """선택된 SMP 시나리오(M/H/L)의 해당 연도 배수.
    기본안(M)은 컬럼이 없어도 1로 처리(원본 동작 보존), 그 외 시나리오는 컬럼 필수."""
    key = getattr(params, "smp_scenario", DEFAULT_SMP_SCENARIO) or DEFAULT_SMP_SCENARIO
    col = find_col(rate_row, f"SMP_{key}", required=False)
    if col is not None:
        return rate_row[col]
    if key == DEFAULT_SMP_SCENARIO:
        return 1
    raise KeyError(f"Annual_Rate.csv 에 SMP_{key} 컬럼이 없습니다. (가능 시나리오: {list(SMP_SCENARIOS)})")


def cost_company(df_h, g, year, usage_row, rate_row, params, smp_col, factor_adj_col):
    """전사(SKH) 단일 비용 모델. 결과 단위: 백만원. (원본 Check/Greedy 보존)"""
    n = len(df_h)
    VC = rate_row[rate_col("VC")]
    FC = rate_row[rate_col("FC")]
    BC = rate_row[rate_col("BC")]
    EAC = rate_row[rate_col("EAC")]
    smp_base = _smp_multiplier(rate_row, params)

    annual_usage = float(usage_row[COMPANY_USAGE_COL])
    peak = float(usage_row[f"{COMPANY_USAGE_COL}_Peak"])
    peak_min = float(usage_row[f"{COMPANY_USAGE_COL}_Peak_Min"])

    usage = np.full(n, peak_min)                       # 원본: 시간당 = Peak_Min 상수
    re_goal = float(usage_row[params.re_goal_column])
    needed_re = int(round(annual_usage * re_goal))

    gen = g.gen_mwh
    min_gu = np.minimum(usage, gen)
    surplus = np.maximum(gen - usage, 0)
    smp_price = smp_base * df_h[smp_col].to_numpy()    # 원/kWh

    with np.errstate(divide="ignore", invalid="ignore"):
        ppa_cost_won_arr = np.nan_to_num(g.ppa_cost_won * min_gu / gen)
    ppa_cost = ppa_cost_won_arr.sum() / 1e6

    # 초과발전 SMP 차감 (균등발전원 제외)
    flat_total = g.flat_pv + g.flat_wt
    if params.flat_pv_mode or params.flat_wt_mode:
        with np.errstate(divide="ignore", invalid="ignore"):
            ratio = np.nan_to_num((gen - flat_total) / gen)
    else:
        ratio = 1.0
    smp_deduction_won = surplus * ratio * smp_price * 1000.0
    vppa_cost = (g.ppa_cost_won - ppa_cost_won_arr - smp_deduction_won).sum() / 1e6

    consumption = np.clip(usage - gen, 0, None)
    consumption_won = consumption * (VC * df_h[factor_adj_col].to_numpy() + FC) * 1000.0
    consumption_cost = consumption_won.sum() / 1e6

    basic_cost = peak * BC * 12 / 1000                 # 원본 단위 보존
    transmission_cost = max(0.0, g.total_capacity - peak_min) * GRID_USAGE_CHARGE / 1000

    re_sum = float(gen.sum())
    add_re = max(0.0, needed_re - re_sum)
    eac_cost = add_re * EAC / 1000

    total_cost = ppa_cost + consumption_cost + basic_cost + transmission_cost + vppa_cost + eac_cost

    return {
        "year": year,
        "annual_usage_sum": float(usage.sum()),
        "re_goal_percent": re_goal,
        "needed_re": needed_re,
        "ppa_generation": re_sum,
        "surplus_gen": float(surplus.sum()),
        "eac_needed": add_re,
        "add_re": add_re,
        "ppa_cost": ppa_cost,
        "basic_charge": basic_cost,
        "consumption_cost": consumption_cost,
        "transmission_cost": transmission_cost,
        "vppa_cost": vppa_cost,
        "eac_cost": eac_cost,
        "total_cost": total_cost,
    }


# ════════════════════════════════════════════════════════════════════
# 5. 비용 집계 — 사업장 모델 (Site)
# ════════════════════════════════════════════════════════════════════

def cost_site(df_h, g, year, site: Site, usage_row, rate_row, params, smp_col):
    """사업장(IC/CJ/YI) 단위 비용 모델. 결과 단위: 백만원. (원본 Site 보존)"""
    n = len(df_h)
    total_usage = float(usage_row[find_col(usage_row, site.code)])
    re_goal_pct = float(usage_row[params.re_goal_column])
    needed_re_mwh = total_usage * re_goal_pct

    usage = np.full(n, total_usage / n if n > 0 else 0.0)   # 원본: 연간/8760 균등
    gen = g.gen_mwh
    min_gu = np.minimum(usage, gen)
    surplus = np.maximum(gen - usage, 0)

    with np.errstate(divide="ignore", invalid="ignore"):
        direct_won = np.nan_to_num(g.ppa_cost_won * min_gu / gen)
    direct_ppa_cost = direct_won.sum() / 1e6

    smp_won_kwh = df_h[smp_col].to_numpy() * _smp_multiplier(rate_row, params)
    # 원본 Site는 균등발전원 구분 없이 전체 초과발전에 SMP 차감
    smp_deduction_won = float((surplus * 1000.0 * smp_won_kwh).sum())
    vppa_cost = ((g.ppa_cost_won - direct_won).sum() - smp_deduction_won) / 1e6

    consumption = np.clip(usage - gen, 0, None)
    uc_col = next((c for c in site.uc_cols if find_col(rate_row, c, required=False)), None)
    if uc_col is None:
        raise KeyError(f"Annual_Rate.csv에서 {site.code} 사업장 전력량요금 컬럼({site.uc_cols})을 찾을 수 없습니다.")
    uc = rate_row[find_col(rate_row, uc_col)]
    fc = rate_row[rate_col("FC")]
    consumption_cost = (consumption * (uc * df_h[site.factor_col].to_numpy() + fc)).sum() * 1000 / 1e6

    base_capacity = float(usage_row[find_col(usage_row, f"{site.code}_Peak")])
    basic_cost = base_capacity * site.basic_charge * 12 / 1e6

    # 망이용요금: 계약용량이 사업장 최소피크(Peak_Min)를 초과하는 부분에만 부과
    # [FIX] 단위 버그 수정 — 용량(MW→kW)과 Peak_Min(MW→kW)을 같은 단위로 비교.
    #        기존에는 Peak_Min을 kW값과 직접 비교해 사실상 0으로 취급 → 사업장 분산 유인이 소멸했음.
    base_capacity_min = float(usage_row[find_col(usage_row, f"{site.code}_Peak_Min")])
    excess_kw = max(0.0, (g.total_capacity - base_capacity_min) * 1000)   # MW→kW
    transmission_cost = excess_kw * GRID_USAGE_CHARGE * 12 / 1e6          # 원본: ×12

    total_site_cost = consumption_cost + direct_ppa_cost + basic_cost + transmission_cost + vppa_cost

    return {
        "site": site.code,
        "year": year,
        "total_site_cost": total_site_cost,
        "site_re_gen_total": float(gen.sum()),
        "company_needed_re_mwh": needed_re_mwh,
        "details": {
            "consumption_cost": consumption_cost,
            "direct_ppa_cost": direct_ppa_cost,
            "vppa_cost": vppa_cost,
            "basic_cost": basic_cost,
            "transmission_cost": transmission_cost,
            "total_ppa_capacity_mw": g.total_capacity,
            "base_capacity": base_capacity,
            "base_capacity_min": base_capacity_min,
        },
    }


# ════════════════════════════════════════════════════════════════════
# 6. 전사 비용 헬퍼 (Greedy/Check가 공유)
# ════════════════════════════════════════════════════════════════════

def _company_year(bundle, params, portfolio, year):
    df_h = bundle.hourly[year]
    g = compute_generation(df_h, portfolio, year, bundle.prices_for(params), "company")
    return g, cost_company(df_h, g, year, bundle.usage_for(params)[year], bundle.rate[year],
                           params, bundle.smp_col, bundle.factor_adj_col)


def company_total_cost(bundle, params, portfolio):
    """25년 총비용(백만원)과 연도별 (needed_re, add_re) 목록 반환. (원본 calculate_total_cost)"""
    total = 0.0
    eac_summary = []
    for year in YEARS:
        _, c = _company_year(bundle, params, portfolio, year)
        total += c["total_cost"]
        eac_summary.append((c["needed_re"], c["add_re"]))
    return total, eac_summary


# ════════════════════════════════════════════════════════════════════
# 7. 실행 진입점 — Check (포트폴리오 비용 검증)
# ════════════════════════════════════════════════════════════════════

def run_check(bundle, params):
    """
    확정 포트폴리오의 연도별 상세 비용과 RE 조달 현황을 계산한다.
    반환: {'cost': DataFrame, 're': DataFrame}
    """
    portfolio = params.fixed_ppas
    usage_map = bundle.usage_for(params)
    prices = bundle.prices_for(params)
    cost_rows, re_rows = [], []

    for year in YEARS:
        df_h = bundle.hourly[year]
        usage_row = usage_map[year]
        g = compute_generation(df_h, portfolio, year, prices, "company")
        c = cost_company(df_h, g, year, usage_row, bundle.rate[year],
                         params, bundle.smp_col, bundle.factor_adj_col)
        cost_rows.append(c)

        re_row = {
            "year": year,
            "전체 수요": float(usage_row[COMPANY_USAGE_COL]),
            "needed RE": c["needed_re"],
            "EAC 구매량": c["add_re"],
            "SEC 수전량": float(usage_row.get("SEC", 0)),
        }
        for src in SOURCES:
            arr = g.gen_by_source.get(src)
            re_row[src] = float(arr.sum()) if arr is not None else 0.0
        re_rows.append(re_row)

    # --- 연간비용분석 DataFrame ---
    cost_cols = {
        "year": "연도", "annual_usage_sum": "사용량 (MWh)", "re_goal_percent": "RE Goal (%)",
        "needed_re": "RE 목표량 (MWh)", "ppa_generation": "PPA 발전량 (MWh)",
        "surplus_gen": "초과발전량 (MWh)", "eac_needed": "추가 EAC (MWh)",
        "ppa_cost": "PPA비용", "basic_charge": "기본요금", "consumption_cost": "전력량요금",
        "transmission_cost": "망이용요금", "vppa_cost": "VPPA손익", "eac_cost": "EAC비용",
        "total_cost": "총비용 (백만원)",
    }
    df_cost = pd.DataFrame(cost_rows)[list(cost_cols.keys())].rename(columns=cost_cols)

    # --- RE 조달 현황 DataFrame ---
    re_order = ["태양광", "균등태양광", "육상풍력", "균등풍력", "해상풍력", "수력", "SMR",
                "SEC 수전량", "EAC 구매량", "needed RE", "전체 수요"]
    df_re = pd.DataFrame(re_rows).set_index("year")
    for col in re_order:
        if col not in df_re.columns:
            df_re[col] = 0
    df_re = df_re[re_order]

    return {"cost": df_cost, "re": df_re}


# ════════════════════════════════════════════════════════════════════
# 8. 실행 진입점 — Greedy (포트폴리오 자동 최적화)
# ════════════════════════════════════════════════════════════════════

def _eac_valid(eac_summary):
    """모든 연도에서 EAC 구매량 ≥ RE목표 × 30% 인지."""
    return all(add >= needed * EAC_MIN_RATIO
               for needed, add in eac_summary if needed > 0)


def run_greedy(bundle, params, progress=None):
    """
    기체결 PPA에서 출발해 비용을 줄이는 신규 계약을 하나씩 추가한다.
    반환: {'portfolio': [...], 'optimization_log': DataFrame, 'evaluation_log': DataFrame}
    progress: 선택적 콜백 progress(step, message)
    """
    fixed = params.fixed_ppas
    optimized = []
    optimization_log, evaluation_log = [], []
    eac_blacklist, cost_blacklist = set(), set()
    opt_sources = params.opt_sources()
    step = 0

    while True:
        step += 1
        current = fixed + optimized
        base_cost, _ = company_total_cost(bundle, params, current)
        if progress:
            progress(step, f"STEP {step}: 계약 {len(current)}건, 총비용 {base_cost:,.0f} 백만원")

        # 연도·발전원별 누적 용량 추적 (optimized만)
        cap_by_ys = {}
        for p in optimized:
            cap_by_ys[(p["start_year"], p["source"])] = \
                cap_by_ys.get((p["start_year"], p["source"]), 0) + p["capacity"]
            if params.flat_pv_mode and p["source"] in ("태양광", "균등태양광"):
                k = (p["start_year"], "태양광_TOTAL")
                cap_by_ys[k] = cap_by_ys.get(k, 0) + p["capacity"]
            if params.flat_wt_mode and p["source"] in ("육상풍력", "균등풍력"):
                k = (p["start_year"], "풍력_TOTAL")
                cap_by_ys[k] = cap_by_ys.get(k, 0) + p["capacity"]

        # 후보 생성
        candidates = []
        for src in opt_sources:
            if src == "SMR" and any(p["source"] == "SMR" for p in current):
                continue
            for year in YEARS:
                if src == "SMR" and year < 2035:
                    continue
                if params.flat_pv_mode and src in ("태양광", "균등태양광"):
                    limit_name, cur = "태양광", cap_by_ys.get((year, "태양광_TOTAL"), 0)
                elif params.flat_wt_mode and src in ("육상풍력", "균등풍력"):
                    limit_name, cur = "육상풍력", cap_by_ys.get((year, "풍력_TOTAL"), 0)
                else:
                    limit_name, cur = src, cap_by_ys.get((year, src), 0)
                cap_limit = NEW_ENTRY_CAP.get(limit_name, float("inf"))
                if cur >= cap_limit:
                    continue
                for capacity in CAP_GRIDS.get(src, []):
                    if cur + capacity > cap_limit:
                        continue
                    cand = {"source": src, "capacity": capacity, "start_year": year}
                    tup = (src, capacity, year)
                    if tup in eac_blacklist or tup in cost_blacklist:
                        continue
                    candidates.append(cand)

        if not candidates:
            if progress:
                progress(step, "생성 가능한 후보가 없어 종료")
            break

        # 후보 평가 (순차)
        results = []
        for cand in candidates:
            tc, eac_sum = company_total_cost(bundle, params, current + [cand])
            results.append({"candidate": cand, "tc": tc, "eac_valid": _eac_valid(eac_sum)})

        results.sort(key=lambda x: (x["tc"], x["candidate"]["source"], x["candidate"]["start_year"]))

        best, min_cost = None, float("inf")
        for r in results:
            cand, tc, valid = r["candidate"], r["tc"], r["eac_valid"]
            tup = (cand["source"], cand["capacity"], cand["start_year"])
            if not valid:
                status = "EAC규칙위반"
                eac_blacklist.add(tup)
            elif tc >= base_cost:
                status = "비용증가"
                cost_blacklist.add(tup)
            else:
                status = "비용절감"
            evaluation_log.append({
                "Step": step, "발전원": cand["source"], "시작연도": cand["start_year"],
                "설비용량(MW)": cand["capacity"], "25년 총비용(백만원)": round(tc),
                "비용 변화(백만원)": round(tc - base_cost), "상태": status,
            })
            if valid and tc < min_cost:
                min_cost, best = tc, cand

        if best and min_cost < base_cost:
            optimization_log.append({"Step": step, **best})
            optimized.append(best)
            if progress:
                progress(step, f"추가: {best['source']} {best['capacity']}MW ({best['start_year']}), "
                               f"절감 {base_cost - min_cost:,.0f} 백만원")
        else:
            if progress:
                progress(step, "더 이상 절감 가능한 계약 없음 — 종료")
            break

    return {
        "portfolio": fixed + optimized,
        "optimization_log": pd.DataFrame(optimization_log),
        "evaluation_log": pd.DataFrame(evaluation_log),
    }


# ════════════════════════════════════════════════════════════════════
# 9. 실행 진입점 — Site (IC/CJ/YI 사업장 배분)
# ════════════════════════════════════════════════════════════════════

def _site_year(bundle, params, site_code, year, portfolio):
    df_h = bundle.hourly[year]
    g = compute_generation(df_h, portfolio, year, bundle.prices_for(params), "site")
    return cost_site(df_h, g, year, SITES[site_code], bundle.usage_for(params)[year],
                     bundle.rate[year], params, bundle.smp_col)


def _company_cost_across_sites(bundle, params, port_by_site, year):
    """3개 사업장 합계 비용 + 회사 단위 EAC 비용(백만원)."""
    res = {s: _site_year(bundle, params, s, year, port_by_site[s]) for s in SITE_CODES}
    total_gen = sum(r["site_re_gen_total"] for r in res.values())
    # [FIX] 회사 RE 필요량 = 3개 사업장 필요량 합산 (Σ 사업장 사용량 × RE목표).
    #        기존 원본은 첫 사업장(A) 필요량만 사용해 EAC 비용이 과소평가되고 배분 판단이 왜곡됐음.
    needed = sum(r["company_needed_re_mwh"] for r in res.values())
    eac_rate = bundle.rate[year][find_col(bundle.rate[year], rate_col("EAC"))]
    eac_cost = max(0.0, needed - total_gen) * 1000 * eac_rate / 1e6
    total = sum(r["total_site_cost"] for r in res.values()) + eac_cost
    return total, res, needed, total_gen, eac_cost


def run_site(bundle, params, progress=None):
    """
    PPA 물량을 IC/CJ/YI 사업장에 회사 전체 비용이 최소가 되도록 배분한다.
    반환: {'summary': DataFrame, 'allocation': DataFrame}
    """
    fixed = params.fixed_ppas
    port_by_site = {s: [] for s in SITE_CODES}
    summary_rows, alloc_rows = [], []

    for year in YEARS:
        if params.reset_portfolio_annually:
            port_by_site = {s: [] for s in SITE_CODES}
            to_alloc = [p for p in fixed
                        if p["start_year"] <= year and (p.get("end_year") is None or p["end_year"] >= year)]
        else:
            to_alloc = [p for p in fixed if p["start_year"] == year]

        # 100MW 청크로 분할
        chunks = []
        for ppa in to_alloc:
            total_mw = ppa["capacity"]
            if total_mw < EPS:
                continue
            base = {"source": ppa["source"], "vintage": ppa["start_year"]}
            if "end_year" in ppa:
                base["end_year"] = ppa["end_year"]
            num = int(total_mw // STEP_MW)
            chunks.extend([{**base, "capacity": STEP_MW}] * num)
            rem = total_mw - num * STEP_MW
            if rem >= MIN_PARTIAL_STEP:
                chunks.append({**base, "capacity": rem})

        # 청크별 최적 사업장 탐색
        for i, chunk in enumerate(chunks):
            cost0, res0, _, _, _ = _company_cost_across_sites(bundle, params, port_by_site, year)
            savings, headroom = {}, {}
            for s in SITE_CODES:
                trial = {sc: port_by_site[sc] + [chunk] if sc == s else port_by_site[sc]
                         for sc in SITE_CODES}
                cost1, _, _, _, _ = _company_cost_across_sites(bundle, params, trial, year)
                savings[s] = cost0 - cost1
                # 수요 여유(미포화도) = 사업장 연간 사용량 − 현재 배분된 RE 발전량
                usage_row_s = bundle.usage_for(params)[year]
                site_usage = float(usage_row_s[find_col(usage_row_s, s)])
                headroom[s] = site_usage - res0[s]["site_re_gen_total"]
            # [FIX] 절감액이 사실상 동률(상대오차 1e-9 이내)이면 수요 여유가 큰 사업장 우선.
            #        기존에는 strict 비교로 항상 첫 사업장이 동률을 독식 → 특정 사업장 배분 0 발생.
            best = max(savings.values())
            tol = max(abs(best) * 1e-9, 1e-6)
            tied = [s for s in SITE_CODES if best - savings[s] <= tol]
            best_site = max(tied, key=lambda s: headroom[s])
            port_by_site[best_site].append(chunk)
            alloc_rows.append({
                "Year": year, "Site": best_site, "Source": chunk["source"],
                "Vintage": chunk["vintage"], "MW": chunk["capacity"],
                "End_Year": chunk.get("end_year"),
            })
        if progress:
            progress(year, f"{year}년: {len(chunks)}개 청크 배분 완료")

        # 연도별 요약 (baseline 대비)
        base_total, _, _, _, _ = _company_cost_across_sites(
            bundle, params, {s: [] for s in SITE_CODES}, year)
        opt_total, _, opt_needed, opt_gen, _ = _company_cost_across_sites(
            bundle, params, port_by_site, year)
        summary_rows.append({
            "Year": year,
            "Baseline_Total_Cost(Million_KRW)": base_total,
            "Optimized_Total_Cost(Million_KRW)": opt_total,
            "Total_Savings(Million_KRW)": base_total - opt_total,
            "Needed_RE_MWh": opt_needed,
            "PPA_Generation_MWh": opt_gen,
            "EAC_Shortfall_MWh": max(0.0, opt_needed - opt_gen),
        })

    return {
        "summary": pd.DataFrame(summary_rows),
        "allocation": pd.DataFrame(alloc_rows),
    }
