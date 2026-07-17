# -*- coding: utf-8 -*-
"""플랫폼 인덱스 조회 헬퍼 (Phase A/B 공용).

data/platform_index.json을 로드해 지표를 검색·조회한다.
- analyzer(Phase B): factsheet 교차검증 후 최신·월별 값 인용에 사용
- chatbot: 담당자 내부 조회("이천 사이트 월별 용수 사용량")에 사용
"""
import json
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INDEX_PATH = os.path.join(BASE_DIR, "data", "platform_index.json")

_CACHE = {"mtime": None, "doc": None}


def load(index_path=None):
    """인덱스를 로드한다(파일 변경 시 자동 갱신). 없으면 빈 문서 반환."""
    path = index_path or INDEX_PATH
    try:
        mtime = os.path.getmtime(path)
    except OSError:
        return {"count": 0, "indicators": []}
    if _CACHE["doc"] is not None and _CACHE["mtime"] == mtime \
            and index_path is None:
        return _CACHE["doc"]
    with open(path, encoding="utf-8") as f:
        doc = json.load(f)
    if index_path is None:
        _CACHE.update(mtime=mtime, doc=doc)
    return doc


def get(platform_id, index_path=None):
    """platform_id(지표고유번호)로 지표를 찾는다. 여러 사이트면 첫 매칭."""
    for ind in load(index_path)["indicators"]:
        if ind["platform_id"] == platform_id or ind.get("uid") == platform_id:
            return ind
    return None


def get_all(platform_id, index_path=None):
    """같은 지표고유번호의 모든 사이트 레코드를 반환."""
    return [ind for ind in load(index_path)["indicators"]
            if ind["platform_id"] == platform_id]


def search(keywords, field=None, site=None, limit=10, index_path=None):
    """키워드(공백 구분)로 지표를 검색한다. 모든 키워드를 포함하는 지표 우선.

    반환: 매칭 점수 내림차순 지표 리스트.
    """
    terms = [t.lower() for t in str(keywords).split() if t.strip()]
    results = []
    for ind in load(index_path)["indicators"]:
        if field and ind.get("field") != field:
            continue
        if site and ind.get("site") != site:
            continue
        hay = ind.get("search", "")
        score = sum(1 for t in terms if t in hay)
        if score:
            results.append((score, ind))
    results.sort(key=lambda x: (-x[0], x[1]["name"]))
    return [ind for _, ind in results[:limit]]


def latest_value(ind):
    """지표의 대표값(가장 최근 연도값)과 연도를 반환. (값, 연도) 또는 (None, None)."""
    annual = ind.get("annual", {})
    if not annual:
        return None, None
    year = max(annual, key=lambda y: int(y) if y.isdigit() else -1)
    return annual[year], year


def monthly_series(ind, year=None):
    """월별값 리스트 [(월, 값)] 반환 (1~12 순서, 값 있는 것만)."""
    monthly = ind.get("monthly", {})
    out = []
    for m in range(1, 13):
        v = monthly.get(str(m))
        if v is not None:
            out.append((m, v))
    return out
