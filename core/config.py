"""
RE100 PPA 시스템 — 공통 설정 (config)

기존 3개 코드(Greedy / Check / Site)에 흩어지고 중복되던 상수·설정을 한 곳에 모았다.
- 잘 바뀌지 않는 값(이용률, 발전원 정의 등)은 모듈 상수
- 시나리오마다 바꾸는 값(목표·모드·포트폴리오)은 SimulationParams 로 분리 → 대시보드에서 주입
"""

from dataclasses import dataclass, field


### 1. 시뮬레이션 기간 ###

YEARS = list(range(2026, 2051))   # 2026 ~ 2050


### 2. 발전원 정의 ###

SOURCES = ['태양광', '육상풍력', '해상풍력', '수력', 'SMR', '균등태양광', '균등풍력']

# 시간별 발전 패턴 컬럼 (None = 균등발전, 상수 1.0 사용)
GENERATION_COL_MAP = {
    '태양광': 'PV_GENERATION_2',
    '육상풍력': 'WT_GENERATION',
    '해상풍력': 'WT_GENERATION',
    '수력': None, 'SMR': None, '균등태양광': None, '균등풍력': None,
}

# 발전원별 이용률
UTILIZATION_MAP = {
    '태양광': 0.15, '육상풍력': 0.25, '해상풍력': 0.35, '수력': 0.26,
    'SMR': 0.90, '균등태양광': 0.15, '균등풍력': 0.25,
}

# PPA 단가 컬럼 약어 (Annual_PPA.csv: f"{약어}_{시나리오}")
ABBRS = {
    '태양광': 'PV', '육상풍력': 'On_W', '해상풍력': 'Off_W', '수력': 'HP',
    'SMR': 'SMR', '균등태양광': 'PV', '균등풍력': 'On_W',
}

# 균등(flat) 발전원: 초과발전 시 SMP 차감을 적용하지 않음
FLAT_SOURCES = {'균등태양광', '균등풍력'}


### 3. 사이트 정의 (A 이천 / B 청주 / C 용인 — Portfolio_251020 데이터 형식) ###

@dataclass(frozen=True)
class Site:
    code: str            # 사이트 코드 (A, B, C) — Annual_Usage.csv 사용량 컬럼명
    name: str            # 사이트 이름 (이천, 청주, 용인)
    factor_col: str      # Hourly_Data.csv 요금 조정계수 컬럼
    uc_cols: tuple       # Annual_Rate.csv 전력량요금 컬럼 후보 (데이터 버전에 따라 UC_A 또는 UC_IC)
    basic_charge: float  # 기본요금 단가 (원/kW)

SITES = {
    "A": Site("A", "이천", "고압B 선택2", ("UC_A", "UC_IC"), 7380),
    "B": Site("B", "청주", "고압B 선택3", ("UC_B", "UC_CJ"), 8190),
    "C": Site("C", "용인", "고압C 선택2", ("UC_C", "UC_YI"), 7520),
}
SITE_CODES = list(SITES.keys())

# 전사(통합) — Greedy/Check가 사용하는 단일 사용량 컬럼
COMPANY_USAGE_COL = "SKH"


### 4. 요금 / 단가 상수 ###

GRID_USAGE_CHARGE = 1335.03   # 망 이용요금 단가 (원/kW)

# 시나리오 (M = 중간). 실제 컬럼명은 f"{항목}_{시나리오}" 형태
RATE_SCENARIO = "M"
PPA_SCENARIO = "M"

EFF_DECAY = 0.995       # 연 0.5% 발전 효율 감소
PPA_REF_PERIOD = 20     # 20년마다 PPA 기준연도(vintage) 갱신


### 5. 그리디 최적화 제약 ###

STEP_MW = 100           # 사이트 배분 시 청크 단위
MIN_PARTIAL_STEP = 1.0  # 잔여 물량 최소 단위
EPS = 1e-6

# 연도별 발전원 신규 계약 최대 용량 (MW)
NEW_ENTRY_CAP = {'태양광': 300, '육상풍력': 100, '해상풍력': 500, 'SMR': 700}

# 신규 계약 용량 단위 (MW)
CAP_GRIDS = {
    '태양광': [100], '육상풍력': [100], '해상풍력': [100],
    'SMR': [700], '균등태양광': [100], '균등풍력': [100],
}

EAC_MIN_RATIO = 0.30    # 모든 연도에서 EAC 구매량 ≥ RE목표 × 30%


### 6. 컬럼명 헬퍼 ###

def ppa_price_col(source):
    """Annual_PPA.csv 에서 해당 발전원의 단가 컬럼명."""
    return f"{ABBRS[source]}_{PPA_SCENARIO}"

def rate_col(item):
    """Annual_Rate.csv 요금 컬럼명. item: VC / FC / BC / EAC / SMP."""
    return f"{item}_{RATE_SCENARIO}"


### 7. 런타임 파라미터 (대시보드에서 변경하는 값) ###

@dataclass
class SimulationParams:
    re_goal_column: str = "RE33_SEC"        # RE 목표 시나리오 컬럼
    sec_stop: bool = False                   # 2043년 이후 SEC 가동 중단 여부
    smr_mode: bool = False                   # SMR 최적화 포함
    flat_pv_mode: bool = True                # 균등태양광 최적화 포함
    flat_wt_mode: bool = False               # 균등풍력 최적화 포함
    reset_portfolio_annually: bool = True    # (Site) 매년 포트폴리오 재배분
    detail_analysis_year: int = 2030         # (Check) 상세 분석 연도
    fixed_ppas: list = field(default_factory=lambda: list(DEFAULT_FIXED_PPAS))

    def opt_sources(self):
        """현재 모드에서 그리디 최적화 대상이 되는 발전원 목록."""
        srcs = ['태양광', '육상풍력', '해상풍력']
        if self.smr_mode:
            srcs.append('SMR')
        if self.flat_pv_mode:
            srcs.append('균등태양광')
        if self.flat_wt_mode:
            srcs.append('균등풍력')
        return srcs


### 8. 기본 기체결 PPA (Check / Site 기준 전체 목록) ###

DEFAULT_FIXED_PPAS = [
    {'source': '태양광', 'capacity': 300, 'start_year': 2026},
    {'source': '태양광', 'capacity': 400, 'start_year': 2027},
    {'source': '태양광', 'capacity': 300, 'start_year': 2028},
    {'source': '태양광', 'capacity': 370, 'start_year': 2029},
    {'source': '태양광', 'capacity': 300, 'start_year': 2030},
    {'source': '태양광', 'capacity': 300, 'start_year': 2031},
    {'source': '태양광', 'capacity': 300, 'start_year': 2032},
    {'source': '태양광', 'capacity': 300, 'start_year': 2033},
    {'source': '태양광', 'capacity': 300, 'start_year': 2034},
    {'source': '태양광', 'capacity': 300, 'start_year': 2035},
    {'source': '태양광', 'capacity': 300, 'start_year': 2036},
    {'source': '태양광', 'capacity': 300, 'start_year': 2037},
    {'source': '태양광', 'capacity': 300, 'start_year': 2038},
    {'source': '태양광', 'capacity': 300, 'start_year': 2039},
    {'source': '태양광', 'capacity': 300, 'start_year': 2040},
    {'source': '태양광', 'capacity': 300, 'start_year': 2041},
    {'source': '균등태양광', 'capacity': 300, 'start_year': 2042},
    {'source': '균등태양광', 'capacity': 300, 'start_year': 2043},
    {'source': '균등태양광', 'capacity': 300, 'start_year': 2044},
    {'source': '균등태양광', 'capacity': 300, 'start_year': 2045},
    {'source': '균등태양광', 'capacity': 300, 'start_year': 2046},
    {'source': '균등태양광', 'capacity': 300, 'start_year': 2047},
    {'source': '균등태양광', 'capacity': 300, 'start_year': 2048},
    {'source': '균등태양광', 'capacity': 200, 'start_year': 2049},
    {'source': '육상풍력', 'capacity': 40, 'start_year': 2026, 'end_year': 2027},
    {'source': '육상풍력', 'capacity': 100, 'start_year': 2030},
    {'source': '육상풍력', 'capacity': 100, 'start_year': 2031},
    {'source': '육상풍력', 'capacity': 100, 'start_year': 2032},
    {'source': '육상풍력', 'capacity': 100, 'start_year': 2033},
    {'source': '육상풍력', 'capacity': 100, 'start_year': 2034},
    {'source': '육상풍력', 'capacity': 100, 'start_year': 2035},
    {'source': '육상풍력', 'capacity': 100, 'start_year': 2036},
    {'source': '육상풍력', 'capacity': 100, 'start_year': 2037},
    {'source': '육상풍력', 'capacity': 100, 'start_year': 2038},
    {'source': '육상풍력', 'capacity': 100, 'start_year': 2039},
    {'source': '육상풍력', 'capacity': 100, 'start_year': 2040},
    {'source': '육상풍력', 'capacity': 100, 'start_year': 2041},
    {'source': '육상풍력', 'capacity': 100, 'start_year': 2042},
    {'source': '육상풍력', 'capacity': 100, 'start_year': 2043},
    {'source': '육상풍력', 'capacity': 100, 'start_year': 2044},
    {'source': '육상풍력', 'capacity': 100, 'start_year': 2045},
    {'source': '육상풍력', 'capacity': 100, 'start_year': 2046},
    {'source': '육상풍력', 'capacity': 100, 'start_year': 2047},
    {'source': '육상풍력', 'capacity': 100, 'start_year': 2048},
    {'source': '해상풍력', 'capacity': 200, 'start_year': 2045},
    {'source': '해상풍력', 'capacity': 500, 'start_year': 2046},
    {'source': '해상풍력', 'capacity': 500, 'start_year': 2047},
    {'source': '해상풍력', 'capacity': 500, 'start_year': 2048},
    {'source': '해상풍력', 'capacity': 500, 'start_year': 2049},
    {'source': '해상풍력', 'capacity': 300, 'start_year': 2050},
    {'source': '수력', 'capacity': 18, 'start_year': 2026},
]
