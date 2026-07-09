"""
RE100 PPA 시뮬레이터 — 백엔드 서버 (Python 표준 라이브러리만 사용)

실행:
    python server.py            # 기본 포트 8700, 샘플 데이터
    python server.py 8701       # 포트 지정
    환경변수 PPA_DATA_DIR 로 입력 폴더 지정 (사내: 실제 데이터 폴더)

역할:
    - web/ 정적 파일(대시보드 HTML, Chart.js) 서빙
    - /api/* 엔드포인트에서 검증된 엔진(core.engine)을 호출해 JSON 반환

추가 설치 의존성 없음(pandas/numpy는 엔진이 이미 사용). 사내 반입에 유리.
"""

import json
import os
import socket
import sys
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
WEB_DIR = os.path.join(BASE_DIR, "web")
sys.path.insert(0, BASE_DIR)

# 콘솔이 cp949여도 한글/기호 출력이 깨지거나 죽지 않도록 UTF-8로 (server.py 직접 더블클릭 대비)
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

from core.config import (
    YEARS, SOURCES, SITES, SITE_CODES, DEFAULT_FIXED_PPAS,
    COMPANY_USAGE_COL, SimulationParams, ppa_price_col,
)
from core.engine import (
    load_data, run_check, run_site, run_greedy, find_col, company_total_cost,
)
import agent

DATA_DIR = os.environ.get("PPA_DATA_DIR") or os.path.join(BASE_DIR, "data", "sample")

# 엔진이 SimulationParams로 받는(=실제 반영되는) 파라미터만 화이트리스트
ENGINE_PARAM_KEYS = {
    "re_goal_column", "sec_stop", "smr_mode",
    "flat_pv_mode", "flat_wt_mode", "reset_portfolio_annually",
}


_BUNDLE_CACHE = {}  # data_dir -> (mtime_signature, bundle)


def _data_signature(data_dir):
    """입력 CSV들의 최종 수정시각 서명. 파일 교체 시 값이 바뀌어 캐시가 무효화된다."""
    sig = []
    try:
        for name in sorted(os.listdir(data_dir)):
            if name.lower().endswith(".csv"):
                p = os.path.join(data_dir, name)
                sig.append((name, os.path.getmtime(p), os.path.getsize(p)))
    except OSError:
        pass
    return tuple(sig)


def get_bundle(data_dir):
    """데이터 로드(캐시). 서버 재시작 없이 CSV 파일 교체만으로 자동 반영."""
    sig = _data_signature(data_dir)
    cached = _BUNDLE_CACHE.get(data_dir)
    if cached and cached[0] == sig:
        return cached[1]
    bundle = load_data(data_dir)
    _BUNDLE_CACHE[data_dir] = (sig, bundle)
    return bundle


def make_params(payload):
    kwargs = {k: payload[k] for k in ENGINE_PARAM_KEYS if k in payload}
    fixed = payload.get("fixed_ppas")
    if fixed:
        kwargs["fixed_ppas"] = fixed
    return SimulationParams(**kwargs)


def capacity_by_source_by_year(fixed_ppas):
    """연도별·발전원별 누적 계약 용량(MW). 발전 믹스 차트용."""
    mix = {s: [0.0] * len(YEARS) for s in SOURCES}
    for p in fixed_ppas:
        src = p["source"]
        start = p.get("start_year", p.get("vintage"))
        end = p.get("end_year")
        for i, y in enumerate(YEARS):
            if start <= y and (end is None or y <= end):
                mix[src][i] += float(p["capacity"])
    # 값이 모두 0인 발전원은 제외
    return {s: arr for s, arr in mix.items() if any(arr)}


def f(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return 0.0


# ──────────────────────────────────────────────────────────────────
# API 핸들러
# ──────────────────────────────────────────────────────────────────
def api_meta():
    bundle = get_bundle(DATA_DIR)
    any_year = min(bundle.usage)
    re_cols = [c for c in bundle.usage[any_year].index if str(c).startswith("RE")]
    return {
        "years": YEARS,
        "sources": SOURCES,
        "sites": [{"code": c, "name": SITES[c].name} for c in SITE_CODES],
        "re_goal_columns": re_cols,
        "default_fixed_ppas": DEFAULT_FIXED_PPAS,
        "data_dir": DATA_DIR,
        "llm": "사내 LLM" if agent.llm_available() else "Mock",
    }


def api_inputs():
    """입력 페이지용 — 실제 CSV에서 파생한 시계열."""
    bundle = get_bundle(DATA_DIR)
    # 사업장별 사용량 (TWh)
    usage = {c: [] for c in SITE_CODES}
    for y in YEARS:
        row = bundle.usage[y]
        for c in SITE_CODES:
            col = find_col(row, c, required=False)
            usage[c].append(f(row[col]) / 1e6 if col else 0.0)  # MWh→TWh
    # 발전원별 PPA 단가
    ppa = {}
    for src in ["태양광", "육상풍력", "해상풍력", "SMR"]:
        d = bundle.price_dicts.get(src, {})
        ppa[src] = [f(d.get(y, 0)) for y in YEARS]
    # 연도별 평균 SMP
    smp = [round(f(bundle.hourly[y][bundle.smp_col].mean()), 1) for y in YEARS]
    # 요금 단가 (2026 기준)
    r0 = bundle.rate[YEARS[0]]
    def rget(col):
        c = find_col(r0, col, required=False)
        return f(r0[c]) if c else None
    rate = {k: rget(k) for k in ["VC_M", "FC_M", "BC_M", "EAC_M", "SMP_M"]}
    return {"years": YEARS, "usage": usage, "ppa": ppa, "smp": smp, "rate": rate}


def api_run(payload):
    return compute_run(payload)


def _cache_key(payload):
    """화이트리스트 파라미터 + fixed_ppas만으로 캐시 키 생성."""
    keyobj = {k: payload.get(k) for k in sorted(ENGINE_PARAM_KEYS)}
    keyobj["fixed_ppas"] = payload.get("fixed_ppas")
    return json.dumps(keyobj, ensure_ascii=False, sort_keys=True)


_RUN_CACHE = {}          # {key: result} — 같은 시나리오 재계산 방지 (Agent가 반복 조회)
_RUN_CACHE_MAX = 32


def compute_run(payload):
    """run_check(빠름) + 발전 믹스 → 대시보드 핵심 결과. (시나리오별 캐시)"""
    key = _cache_key(payload)
    if key in _RUN_CACHE:
        return _RUN_CACHE[key]
    result = _compute_run_uncached(payload)
    if len(_RUN_CACHE) >= _RUN_CACHE_MAX:
        _RUN_CACHE.pop(next(iter(_RUN_CACHE)))   # 가장 오래된 것 제거
    _RUN_CACHE[key] = result
    return result


def _compute_run_uncached(payload):
    bundle = get_bundle(DATA_DIR)
    params = make_params(payload)
    res = run_check(bundle, params)
    cost, re = res["cost"], res["re"]

    years = cost["연도"].tolist()
    # [FIX] RE 달성경로 분모는 실제 연간 수요(Annual_Usage의 SKH)를 사용.
    #        기존 "사용량 (MWh)"(=Peak_Min×8760, 비용모델용 프록시)를 분모로 쓰면
    #        실데이터에서 목표선이 100%를 초과하는 왜곡 발생. 비용 모델 자체는 원본 보존.
    usage_mwh = re["전체 수요"].tolist()
    target_mwh = cost["RE 목표량 (MWh)"].tolist()
    ppa_gen = cost["PPA 발전량 (MWh)"].tolist()
    eac_mwh = cost["추가 EAC (MWh)"].tolist()
    total_cost = cost["총비용 (백만원)"].tolist()

    def pct(num, den):
        return [round(f(n) / f(d) * 100, 1) if f(d) else 0.0 for n, d in zip(num, den)]

    re_path = {
        "years": years,
        "target": pct(target_mwh, usage_mwh),
        "ppa": pct(ppa_gen, usage_mwh),
        "eac": pct(eac_mwh, usage_mwh),
    }
    cost_breakdown = {
        "years": years,
        "ppa": [f(v) for v in cost["PPA비용"]],
        "utility": [f(v) for v in cost["전력량요금"]],
        "base": [f(v) for v in cost["기본요금"]],
        "eac": [f(v) for v in cost["EAC비용"]],
        "grid": [f(v) for v in cost["망이용요금"]],
    }
    mix = capacity_by_source_by_year(params.fixed_ppas)

    # KPI
    total25y = sum(f(v) for v in total_cost)              # 백만원
    last = cost.iloc[-1]
    target_last = f(last["RE 목표량 (MWh)"])
    covered_last = f(last["PPA 발전량 (MWh)"]) + f(last["추가 EAC (MWh)"])
    re_2050 = round(min(100.0, covered_last / target_last * 100), 1) if target_last else 0.0
    eac_ratios = [f(e) / f(t) for e, t in zip(eac_mwh, target_mwh) if f(t) > 0]
    eac_avg = round(sum(eac_ratios) / len(eac_ratios) * 100, 1) if eac_ratios else 0.0
    cap_2050 = round(sum(arr[-1] for arr in mix.values()))

    return {
        "kpi": {
            "total25y_trillion": round(total25y / 1e6, 2),
            "re_2050": re_2050,
            "cap_2050": cap_2050,
            "eac_avg": eac_avg,
        },
        "mix": mix,
        "re_path": re_path,
        "cost_breakdown": cost_breakdown,
    }


def api_site(payload):
    """run_site(수십초) → IC/CJ/YI 2050년 배분."""
    bundle = get_bundle(DATA_DIR)
    params = make_params(payload)
    res = run_site(bundle, params)
    alloc = res["allocation"]
    last_year = YEARS[-1]
    sub = alloc[alloc["Year"] == last_year]
    sites = {}
    for c in SITE_CODES:
        sites[c] = {"name": SITES[c].name, "mix": {}}
    for _, row in sub.iterrows():
        s, src, mw = row["Site"], row["Source"], f(row["MW"])
        sites[s]["mix"][src] = sites[s]["mix"].get(src, 0.0) + mw
    for c in sites:
        sites[c]["total"] = round(sum(sites[c]["mix"].values()))
        sites[c]["mix"] = {k: round(v) for k, v in sites[c]["mix"].items()}
    return {"sites": sites, "year": last_year}


def api_greedy(payload):
    """run_greedy(수십초) — SMR/균등 등 최적화 옵션이 실제 반영되는 자동 포트폴리오 탐색."""
    bundle = get_bundle(DATA_DIR)
    params = make_params(payload)
    base_cost, _ = company_total_cost(bundle, params, params.fixed_ppas)
    res = run_greedy(bundle, params)
    final_cost, _ = company_total_cost(bundle, params, res["portfolio"])
    added = res["optimization_log"].to_dict(orient="records") if len(res["optimization_log"]) else []
    # 최종 포트폴리오 발전원별 합계 (MW)
    by_source = {}
    for p in res["portfolio"]:
        by_source[p["source"]] = by_source.get(p["source"], 0.0) + float(p["capacity"])
    return {
        "base_cost_trillion": round(base_cost / 1e6, 2),
        "final_cost_trillion": round(final_cost / 1e6, 2),
        "saving_million": round(base_cost - final_cost),
        "added": added,                              # [{Step, source, capacity, start_year}]
        "portfolio_by_source": {k: round(v) for k, v in by_source.items()},
        "mix": capacity_by_source_by_year(res["portfolio"]),   # 최적화 반영 믹스 차트용
        "opt_sources": params.opt_sources(),
        "portfolio": res["portfolio"],               # 사업장 배분 파이프라인용 전체 계약 목록
    }


def api_llm_config(payload):
    """LLM 설정 조회/저장. {save:true, base_url, api_key, model} 이면 저장.
    api_key가 payload에 없으면 기존 키를 유지한다(화면에서 빈칸=미변경)."""
    if payload and payload.get("save"):
        cur = agent.get_config()
        api_key = payload["api_key"] if "api_key" in payload else cur["api_key"]
        agent.save_config(payload.get("base_url", ""), api_key, payload.get("model", ""))
    c = agent.get_config()
    return {"mode": "사내 LLM" if agent.llm_available() else "Mock",
            "base_url": c["base_url"], "model": c["model"], "has_key": bool(c["api_key"])}


def api_llm_test(payload):
    return agent.test_connection()


def api_export(payload):
    """run_check 결과표를 CSV(utf-8-sig, 엑셀 호환)로 반환. 추가 패키지 불필요 → 폐쇄망 OK."""
    bundle = get_bundle(DATA_DIR)
    params = make_params(payload or {})
    res = run_check(bundle, params)
    kind = (payload or {}).get("kind", "cost")
    if kind == "re":
        df, name = res["re"].reset_index(), "RE_조달현황.csv"
    else:
        df, name = res["cost"], "연간비용분석.csv"
    return {"_csv": df.to_csv(index=False), "_filename": name}


def api_agent(payload):
    """대화형 Agent — 결과 해석 답변 / 시나리오 변경 제안·확인·재실행."""
    params = payload.get("params") or {}
    history = payload.get("history") or []
    confirm = payload.get("confirm")
    if confirm:  # 사용자가 제안을 확인 → 실제 재실행
        patch = confirm.get("params_patch") or {}
        new_params = {**params, **patch}
        before = compute_run(params)["kpi"]
        after = compute_run(new_params)
        reply = agent.summarize(patch, before, after["kpi"])
        return {"reply": reply, "applied": True, "new_params": new_params, "ran": after["kpi"]}
    # 일반 질문/요청
    cur = compute_run(params)["kpi"]
    out = agent.respond(payload.get("message", ""), history, params, cur)
    resp = {"reply": out.get("reply", "")}
    if out.get("action") == "propose" and out.get("params_patch"):
        resp["proposal"] = {"params_patch": out["params_patch"], "label": out.get("label", "")}
    return resp


ROUTES = {
    "/api/meta": lambda p: api_meta(),
    "/api/inputs": lambda p: api_inputs(),
    "/api/run": api_run,
    "/api/site": api_site,
    "/api/greedy": api_greedy,
    "/api/export": api_export,
    "/api/agent": api_agent,
    "/api/llm/config": api_llm_config,
    "/api/llm/test": api_llm_test,
}


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        pass  # 조용히

    def _send_json(self, obj, status=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_file(self, path, content_type):
        with open(path, "rb") as fp:
            body = fp.read()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    @staticmethod
    def _route_of(path):
        """Ingress 경로 기반 배포(/ppa/ 등) 시 붙는 접두어를 무시하고 라우팅.
        예: /ppa/api/meta -> /api/meta,  /ppa/ -> /"""
        if "/api/" in path:
            return path[path.index("/api/"):]
        if path.endswith("/health"):
            return "/health"
        if path.endswith("/chart.umd.min.js"):
            return "/chart.umd.min.js"
        if path.endswith("/index.html") or path.endswith("/"):
            return "/"
        return path

    def do_GET(self):
        route = self._route_of(urlparse(self.path).path)
        if route in ("/", "/index.html"):
            return self._send_file(os.path.join(WEB_DIR, "index.html"), "text/html; charset=utf-8")
        if route == "/chart.umd.min.js":
            return self._send_file(os.path.join(WEB_DIR, "chart.umd.min.js"), "application/javascript")
        if route == "/health":
            try:
                get_bundle(DATA_DIR)
                return self._send_json({"status": "ok"})
            except Exception as e:
                return self._send_json({"status": "error", "detail": str(e)}, 500)
        if route in ("/api/meta", "/api/inputs"):
            try:
                return self._send_json(ROUTES[route](None))
            except Exception as e:
                return self._send_json({"error": str(e)}, 500)
        self._send_json({"error": "not found"}, 404)

    def do_POST(self):
        route = self._route_of(urlparse(self.path).path)
        if route not in ROUTES:
            return self._send_json({"error": "not found"}, 404)
        try:
            length = int(self.headers.get("Content-Length", 0))
            payload = json.loads(self.rfile.read(length) or b"{}")
            result = ROUTES[route](payload)
            if isinstance(result, dict) and "_csv" in result:   # CSV 다운로드 응답
                body = ("﻿" + result["_csv"]).encode("utf-8")   # BOM: 엑셀 한글 호환
                self.send_response(200)
                self.send_header("Content-Type", "text/csv; charset=utf-8")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
                return
            return self._send_json(result)
        except Exception as e:
            import traceback
            traceback.print_exc()
            return self._send_json({"error": str(e)}, 500)


def lan_ip():
    """사내망에서 동료가 접속할 때 쓸 이 PC의 IP."""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))   # 실제 전송 없음, 라우팅 인터페이스 IP만 확인
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        try:
            return socket.gethostbyname(socket.gethostname())
        except Exception:
            return "127.0.0.1"


def main():
    port = int(sys.argv[1]) if len(sys.argv) > 1 else int(os.environ.get("PPA_PORT", "8700"))
    if not os.path.exists(os.path.join(DATA_DIR, "Hourly_Data.csv")):
        print(f"[경고] 데이터 폴더에서 Hourly_Data.csv 를 찾을 수 없습니다: {DATA_DIR}")
        print("       PPA_DATA_DIR 환경변수로 올바른 폴더를 지정하세요.")
    else:
        print(f"[데이터] {DATA_DIR}")
    local_url = f"http://localhost:{port}/"
    ip = lan_ip()
    print("=" * 56)
    print(f"  내 PC:    {local_url}")
    print(f"  사내망:   http://{ip}:{port}/   ← 동료에게 알려줄 주소")
    print("=" * 56)
    print("  (처음 실행 시 Windows 방화벽 허용 창이 뜨면 '허용'을 누르세요)")
    print("  종료: 이 창에서 Ctrl+C")
    try:
        get_bundle(DATA_DIR)  # 미리 로드(첫 요청 지연 방지)
        print("[엔진] 데이터 로드 완료")
    except Exception as e:
        print(f"[엔진] 데이터 로드 실패: {e}")
    # 0.0.0.0 = 모든 네트워크 인터페이스 → 사내망의 동료가 위 '사내망' 주소로 접속 가능
    ThreadingHTTPServer(("0.0.0.0", port), Handler).serve_forever()


if __name__ == "__main__":
    main()
