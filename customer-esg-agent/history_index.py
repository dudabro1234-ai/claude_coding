# -*- coding: utf-8 -*-
"""과거 고객대응이력 인덱스 조회 헬퍼.

data/history_index.json(과거 요청↔실발송 답변)을 로드해 유사 이력을 검색한다.
- analyzer : 신규 요구사항과 유사한 과거 이력 연계 → 초안 작성 참고
- reporter : 대시보드 '관련 과거 대응 이력' 섹션
- chatbot  : "작년에 A고객에 뭐라고 답했지?" 류 질문
"""
import json
import os
import re

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INDEX_PATH = os.path.join(BASE_DIR, "data", "history_index.json")

_CACHE = {"mtime": None, "doc": None}

_TOKEN_RE = re.compile(r"[0-9A-Za-z가-힣]{2,}")

# 검색 변별력이 없는 상투어
_STOPWORDS = {"데이터", "제출", "요청", "부탁", "바랍니다", "관련", "내용",
              "대한", "대해", "문의", "확인", "회신", "답변", "please", "the"}


def _tokens(text):
    return [t.lower() for t in _TOKEN_RE.findall(str(text))
            if t.lower() not in _STOPWORDS]


def load(index_path=None):
    path = index_path or INDEX_PATH
    try:
        mtime = os.path.getmtime(path)
    except OSError:
        return {"count": 0, "records": []}
    if _CACHE["doc"] is not None and _CACHE["mtime"] == mtime \
            and index_path is None:
        return _CACHE["doc"]
    with open(path, encoding="utf-8") as f:
        doc = json.load(f)
    if index_path is None:
        _CACHE.update(mtime=mtime, doc=doc)
    return doc


def available(index_path=None):
    return load(index_path).get("count", 0) > 0


def search(query, customer=None, limit=5, index_path=None):
    """키워드 유사 이력 검색. 같은 고객사 이력에 가중치를 준다.

    반환: 점수 내림차순 레코드 리스트 (score 필드 포함 사본).
    """
    terms = set(_tokens(query))
    if not terms and not customer:
        return []
    results = []
    for rec in load(index_path)["records"]:
        hay = rec.get("search", "")
        score = sum(1 for t in terms if t in hay)
        if customer and rec.get("customer") == customer:
            score += 2 if score else 1   # 동일 고객 보너스 (키워드 없으면 약하게)
        if score > 0:
            r = dict(rec)
            r["score"] = score
            results.append(r)
    results.sort(key=lambda r: (-r["score"], r.get("date", "")), reverse=False)
    return results[:limit]


def recent_for_customer(customer, limit=3, index_path=None):
    """해당 고객사의 최근 이력 (인덱스가 최신순 정렬돼 있음)."""
    if not customer:
        return []
    out = []
    for rec in load(index_path)["records"]:
        if rec.get("customer") == customer:
            out.append(dict(rec))
            if len(out) >= limit:
                break
    return out
