# -*- coding: utf-8 -*-
"""과거 고객대응이력 엑셀 → data/history_index.json 변환기.

담당자가 보유한 과거 대응이력(고객 요청 ↔ 실제 발송한 답변)을 인덱싱해,
신규 요청 분석 시 유사 이력 연계·답변 초안 참고에 사용한다.

- 엑셀 양식이 팀마다 달라도 되도록 **헤더 별칭**을 자동 인식한다.
  필수: 일자 / 고객사 / 요청내용 / 답변내용 (각각 별칭 허용)
- 원본과 인덱스는 고객 커뮤니케이션 원문(대외비 소지)이므로 gitignore 처리.

사용법:
  python tools/history_ingest.py [xlsx경로]     # 생략 시 data/history_drop/ 최신 파일
  python tools/history_ingest.py --make-template  # 입력용 템플릿 생성
"""
import datetime
import glob
import json
import os
import sys

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DROP_DIR = os.path.join(BASE_DIR, "data", "history_drop")
OUT_PATH = os.path.join(BASE_DIR, "data", "history_index.json")
TEMPLATE_PATH = os.path.join(BASE_DIR, "data", "history_template.xlsx")

# 필드별 허용 헤더 별칭 (소문자 비교)
ALIASES = {
    "date":     ["일자", "날짜", "대응일", "답변일", "회신일", "수신일", "date"],
    "customer": ["고객사", "고객", "고객명", "customer"],
    "request":  ["요청내용", "요구사항", "요청사항", "질문", "문의내용",
                 "요구내용", "request"],
    "answer":   ["답변내용", "회신내용", "답변", "회신", "대응내용",
                 "answer", "response"],
    # 선택
    "request_type": ["요청유형", "유형", "구분"],
    "framework":    ["프레임워크", "framework", "양식"],
    "owner":        ["담당자", "담당", "owner"],
    "note":         ["비고", "결과", "메모", "처리결과", "note"],
}
REQUIRED = ["date", "customer", "request", "answer"]


def _norm_date(v):
    if v is None:
        return ""
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, datetime.date):
        return v.isoformat()
    s = str(v).strip().replace(".", "-").replace("/", "-")
    return s[:10]


def _map_header(header):
    """헤더 행에서 필드별 열 인덱스를 별칭으로 찾는다."""
    low = [str(h).strip().lower() if h is not None else "" for h in header]
    idx = {}
    for field, names in ALIASES.items():
        for i, h in enumerate(low):
            if h in [n.lower() for n in names]:
                idx[field] = i
                break
    missing = [f for f in REQUIRED if f not in idx]
    if missing:
        pretty = {f: "/".join(ALIASES[f][:3]) for f in missing}
        raise ValueError(
            f"필수 열을 찾지 못했습니다: {pretty}\n"
            f"현재 헤더: {[h for h in header if h]}\n"
            f"템플릿 생성: python tools/history_ingest.py --make-template")
    return idx


def latest_drop():
    files = [f for f in glob.glob(os.path.join(DROP_DIR, "*.xlsx"))
             if not os.path.basename(f).startswith(("~$", "history_template"))]
    return max(files, key=os.path.getmtime) if files else None


def convert(xlsx_path=None, out_path=OUT_PATH):
    """과거 이력 엑셀을 인덱싱한다. 반환: (레코드 수, 출력 경로)."""
    from openpyxl import load_workbook

    xlsx_path = xlsx_path or latest_drop()
    if not xlsx_path:
        raise FileNotFoundError(
            f"이력 파일이 없습니다. 과거 대응이력 xlsx를 {DROP_DIR}에 넣거나 "
            "경로를 인자로 지정하세요.")

    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    idx = _map_header(next(rows))

    records = []
    for row in rows:
        if row is None or all(v is None for v in row):
            continue

        def col(field, default=""):
            i = idx.get(field)
            if i is None or i >= len(row) or row[i] is None:
                return default
            return str(row[i]).strip()

        request, answer = col("request"), col("answer")
        if not request and not answer:
            continue
        # 안내문 행(템플릿 설명 등) 스킵: 날짜가 정상 파싱 안 되고 요청도 짧으면
        rec = {
            "date": _norm_date(row[idx["date"]]),
            "customer": col("customer"),
            "request_type": col("request_type"),
            "framework": col("framework"),
            "request": request,
            "answer": answer,
            "owner": col("owner"),
            "note": col("note"),
        }
        rec["search"] = " ".join(
            str(x) for x in (rec["customer"], rec["request_type"],
                             rec["framework"], rec["request"], rec["answer"])
            if x).lower()
        records.append(rec)
    wb.close()

    records.sort(key=lambda r: r["date"], reverse=True)
    doc = {
        "generated_at": datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
        "source_file": os.path.basename(xlsx_path),
        "count": len(records),
        "records": records,
    }
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(doc, f, ensure_ascii=False, indent=1)
    return len(records), out_path


def make_template(path=TEMPLATE_PATH):
    """과거 대응이력 입력용 템플릿을 생성한다."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "대응이력"
    HEADER = ["일자", "고객사", "요청유형", "프레임워크", "요청내용",
              "답변내용", "담당자", "비고"]
    DESC = ["★필수\nYYYY-MM-DD", "★필수", "(선택)\n설문/데이터제출 등",
            "(선택)\nCDP/EcoVadis 등", "★필수\n고객이 요청한 내용",
            "★필수\n실제 발송한 답변 전문 또는 요지", "(선택)", "(선택)\n결과·특이사항"]
    ws.append(HEADER)
    ws.append(DESC)
    ws.append(["2025-08-12", "예시고객A", "데이터제출", "CDP Supply Chain",
               "2024년 Scope 1/2 배출량 및 재생에너지 비율 제출 요청 (예시)",
               "안녕하세요. 당사 2024년 Scope 1 배출량은 ○○ tCO2eq, Scope 2는 "
               "○○ tCO2eq이며, 재생에너지 비율은 ○○%입니다. 상세 내역은 첨부 "
               "양식에 기재하였습니다. (예시 행 — 실데이터 입력 후 삭제)",
               "홍길동", "기한 내 제출 완료"])
    hf = PatternFill("solid", start_color="134E4A")
    df = PatternFill("solid", start_color="F1F5F9")
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = hf
        c.alignment = Alignment(horizontal="center")
    for c in ws[2]:
        c.font = Font(size=9, color="B91C1C" if "★" in str(c.value) else "64748B",
                      bold="★" in str(c.value))
        c.fill = df
        c.alignment = Alignment(wrap_text=True, vertical="top")
    for c in ws[3]:
        c.font = Font(color="94A3B8", italic=True)
    ws.row_dimensions[2].height = 44
    for i, w in enumerate([12, 14, 13, 16, 42, 52, 10, 18], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A3"
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)
    return path


if __name__ == "__main__":
    if "--make-template" in sys.argv:
        print("템플릿 생성:", make_template())
        sys.exit(0)
    src = sys.argv[1] if len(sys.argv) > 1 else None
    n, out = convert(src)
    print(f"이력 인덱스 생성: {out}")
    print(f"  레코드 {n}건")
    print("  ⚠️ 고객 커뮤니케이션 원문 포함 — 커밋·외부 전송 금지 (.gitignore 등록).")
