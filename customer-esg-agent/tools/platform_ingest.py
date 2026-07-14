# -*- coding: utf-8 -*-
"""사내 데이터플랫폼 export(xlsx) → data/platform_index.json 변환기 (Phase A).

플랫폼에서 받은 datasheet(31열)를 지표별로 정규화해 조회용 인덱스를 만든다.
- 조인 키: 지표고유번호 (platform_id)
- 값: 연간(2022~2025) + 월별(1~12) 시계열
- 이 인덱스는 실재무·월별 원천값(대외비 포함) → .gitignore 처리, LLM에 통째로
  넣지 않는다. 고객 답변 인용은 analyzer의 factsheet 교차검증 게이트(Phase B)가 통제.

사용법:
  python tools/platform_ingest.py [xlsx경로]
  (생략 시 data/platform_drop/ 안의 가장 최근 xlsx를 사용)
"""
import datetime
import glob
import json
import os
import sys

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DROP_DIR = os.path.join(BASE_DIR, "data", "platform_drop")
OUT_PATH = os.path.join(BASE_DIR, "data", "platform_index.json")

# 플랫폼 export 표준 헤더 (31열)
COL = {
    "no": "No.", "platform_id": "지표고유번호", "company": "회사", "site": "사이트",
    "prism": "PRISM", "field": "분야", "l1": "대분류", "l2": "중분류",
    "l3": "소분류", "l4": "세분류", "unit": "단위", "agg": "구분",
    "kind": "지표구분", "cycle": "수집주기", "public": "공시여부",
}
YEAR_COLS = ["2022년", "2023년", "2024년", "2025년"]
MONTH_COLS = [f"{m}월" for m in range(1, 13)]


def _num(v):
    """숫자면 그대로, 아니면 문자열/None 유지."""
    return v if v is not None else None


def _label(l1, l2, l3, l4):
    parts = [p for p in (l1, l2, l3, l4) if p and str(p).strip()]
    return " > ".join(str(p).strip() for p in parts)


def latest_drop():
    files = glob.glob(os.path.join(DROP_DIR, "*.xlsx"))
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    return max(files, key=os.path.getmtime) if files else None


def convert(xlsx_path=None, out_path=OUT_PATH):
    """플랫폼 export를 정규화해 platform_index.json으로 저장.

    반환: (지표 수, 출력 경로)
    """
    from openpyxl import load_workbook

    xlsx_path = xlsx_path or latest_drop()
    if not xlsx_path:
        raise FileNotFoundError(
            f"변환할 파일이 없습니다. 플랫폼 export를 {DROP_DIR}에 넣거나 "
            "경로를 인자로 지정하세요.")

    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    header = [str(h).strip() if h is not None else "" for h in next(rows)]

    idx = {}
    for key, name in COL.items():
        if name not in header:
            raise ValueError(f"'{name}' 열이 없습니다. 플랫폼 export 형식을 확인하세요. "
                             f"현재 헤더: {header[:15]}...")
        idx[key] = header.index(name)
    year_idx = {y: header.index(y) for y in YEAR_COLS if y in header}
    month_idx = {m: header.index(m) for m in MONTH_COLS if m in header}

    indicators = []
    seen_ids = set()
    for row in rows:
        if row is None or all(v is None for v in row):
            continue
        pid = row[idx["platform_id"]]
        if not pid:
            continue
        pid = str(pid).strip()
        # 지표고유번호가 중복되면(사이트/회사 상이) 사이트를 붙여 유일화
        base_id = pid
        if pid in seen_ids:
            pid = f"{pid}#{row[idx['site']] or len(indicators)}"
        seen_ids.add(pid)

        l1 = row[idx["l1"]]; l2 = row[idx["l2"]]
        l3 = row[idx["l3"]]; l4 = row[idx["l4"]]
        annual = {y.replace("년", ""): _num(row[i])
                  for y, i in year_idx.items() if row[i] is not None}
        monthly = {m.replace("월", ""): _num(row[i])
                   for m, i in month_idx.items() if row[i] is not None}

        rec = {
            "platform_id": base_id,
            "uid": pid,
            "company": row[idx["company"]],
            "site": row[idx["site"]],
            "field": row[idx["field"]],       # 분야 (환경/사회/경제·거버넌스)
            "l1": l1, "l2": l2, "l3": l3, "l4": l4,
            "name": _label(l1, l2, l3, l4),
            "unit": row[idx["unit"]],
            "aggregation": row[idx["agg"]],   # 합산/최신
            "cycle": row[idx["cycle"]],       # 분기/년
            "annual": annual,
            "monthly": monthly,
        }
        # 검색 키워드 (소문자, 공백 정규화)
        rec["search"] = " ".join(str(x) for x in
            (rec["company"], rec["site"], rec["field"], l1, l2, l3, l4)
            if x).lower()
        indicators.append(rec)

    wb.close()

    doc = {
        "generated_at": datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
        "source_file": os.path.basename(xlsx_path),
        "count": len(indicators),
        "indicators": indicators,
    }
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(doc, f, ensure_ascii=False, indent=1)
    return len(indicators), out_path


if __name__ == "__main__":
    src = sys.argv[1] if len(sys.argv) > 1 else None
    n, out = convert(src)
    print(f"플랫폼 인덱스 생성: {out}")
    print(f"  지표 {n}개 (원본: {os.path.basename(src or latest_drop())})")
    print("  ⚠️ 이 파일은 대외비 원천값을 포함하므로 외부 전송·커밋 금지 "
          "(.gitignore 등록됨).")
