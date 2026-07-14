# -*- coding: utf-8 -*-
"""② 사내 LLM 분석기.

작업지시서 §3 ②, §5.2, §5.3:
- 메일 1건 = LLM 호출 1~3회 루프 (컨텍스트 제한 대응, §4.3)
    호출1: 요구사항 추출 (extract.txt)
    호출2: factsheet 매칭 (match.txt)
    호출3: 답변 초안 · 부서요청 초안 (draft.txt)
- 부분 실패 허용: 1건 실패 시 해당 건 status=ERROR 후 다음 건 진행
- 출력: work/analyzed/{mail_id}.json
"""
import datetime
import json
import logging
import os

import llm_client
import platform_index

log = logging.getLogger(__name__)


def platform_available():
    """사내 데이터플랫폼 인덱스가 존재하는지 여부."""
    return platform_index.load().get("count", 0) > 0

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INBOX_DIR = os.path.join(BASE_DIR, "work", "inbox")
ANALYZED_DIR = os.path.join(BASE_DIR, "work", "analyzed")
PROMPT_DIR = os.path.join(BASE_DIR, "prompts")
FACTSHEET_MD = os.path.join(BASE_DIR, "data", "factsheet.md")
FACTSHEET_JSON = os.path.join(BASE_DIR, "data", "factsheet.json")

VALID_STATUS = ("답변가능", "부분가능", "데이터없음")
VALID_SEVERITY = ("높음", "중간", "낮음")

DRAFT_BANNER = "⚠️ [AI 생성 초안 — 발송 전 반드시 검토 필요]"


def _load_prompt(name):
    with open(os.path.join(PROMPT_DIR, name), encoding="utf-8") as f:
        return f.read()


def _load_factsheet():
    try:
        with open(FACTSHEET_MD, encoding="utf-8") as f:
            return f.read()
    except FileNotFoundError:
        log.warning("data/factsheet.md가 없습니다. 매칭 단계에서 모든 항목이 "
                    "'데이터없음'으로 분류됩니다. tools/xlsx_to_md.py를 먼저 실행하세요.")
        return ""


def _load_factsheet_index():
    """factsheet.json(공개 Y행만)을 item_code → 상세 dict 로 로드한다."""
    try:
        with open(FACTSHEET_JSON, encoding="utf-8") as f:
            rows = json.load(f)
        return {r.get("item_code"): r for r in rows if r.get("item_code")}
    except Exception:
        return {}


def match_customer(sender, config):
    """발신 도메인으로 고객사를 코드 레벨에서 우선 매칭한다."""
    domain = sender.split("@")[-1].lower().strip() if "@" in sender else ""
    for cust in config.get("customers", []):
        for d in cust.get("domains", []):
            if domain == d.lower() or domain.endswith("." + d.lower()):
                return cust["name"]
    return None


def compute_urgency(deadline, received_at, config):
    """§5.2 추출 규칙: D-7 이내→긴급, D-30 이내→보통, 그 외/미상→낮음.

    LLM 출력에 의존하지 않고 코드에서 결정한다.
    """
    if not deadline:
        return "낮음"
    try:
        d = datetime.date.fromisoformat(str(deadline)[:10])
    except ValueError:
        return "낮음"
    base = datetime.date.fromisoformat(str(received_at)[:10]) \
        if received_at else datetime.date.today()
    today = max(base, datetime.date.today())
    days_left = (d - today).days
    ucfg = config.get("urgency", {})
    if days_left <= ucfg.get("urgent_days", 7):
        return "긴급"
    if days_left <= ucfg.get("normal_days", 30):
        return "보통"
    return "낮음"


def _truncate_body(body, config):
    max_chars = config.get("llm", {}).get("max_body_chars", 12000)
    if len(body) > max_chars:
        return body[:max_chars] + "\n...[본문이 길어 이하 생략됨]"
    return body


def _normalize_requirements(reqs):
    """LLM 출력 요구사항 목록을 스키마에 맞게 정규화한다."""
    out = []
    for i, r in enumerate(reqs or [], start=1):
        if not isinstance(r, dict):
            continue
        status = r.get("status", "데이터없음")
        if status not in VALID_STATUS:
            status = "데이터없음"
        out.append({
            "req_id": r.get("req_id") or f"R{i}",
            "content": str(r.get("content", "")).strip(),
            "matched_items": [str(m) for m in (r.get("matched_items") or [])],
            "status": status,
            "owner_dept": r.get("owner_dept") if status == "데이터없음" else None,
        })
    return out


def _normalize_risks(risks):
    """LLM 리스크 검토 출력을 정규화한다."""
    out = []
    for r in risks or []:
        if not isinstance(r, dict) or not str(r.get("description", "")).strip():
            continue
        sev = r.get("severity", "중간")
        out.append({
            "target": str(r.get("target") or "전체"),
            "severity": sev if sev in VALID_SEVERITY else "중간",
            "description": str(r.get("description", "")).strip(),
            "mitigation": str(r.get("mitigation", "")).strip(),
        })
    order = {"높음": 0, "중간": 1, "낮음": 2}
    out.sort(key=lambda r: order[r["severity"]])
    return out


def analyze_mail(mail, config, factsheet_md, prompts, factsheet_index=None):
    """메일 1건을 LLM 1~3회 호출로 분석해 §5.2 스키마 JSON을 반환한다."""
    factsheet_index = factsheet_index or {}
    body = _truncate_body(mail.get("body", ""), config)
    customer_names = [c["name"] for c in config.get("customers", [])]

    # ── 호출 1: 요구사항 추출 ────────────────────────────────
    extract_user = (
        f"[고객사 목록]\n{', '.join(customer_names) or '(없음)'}\n\n"
        f"[메일 정보]\n"
        f"- 발신자: {mail.get('sender_name', '')} <{mail.get('sender', '')}>\n"
        f"- 제목: {mail.get('subject', '')}\n"
        f"- 수신일시: {mail.get('received_at', '')}\n"
        f"- 첨부파일: {json.dumps(mail.get('attachments', []), ensure_ascii=False)}\n\n"
        f"[본문]\n{body}"
    )
    extracted = llm_client.call_llm_json([
        {"role": "system", "content": prompts["extract"]},
        {"role": "user", "content": extract_user},
    ])

    requirements = _normalize_requirements(extracted.get("requirements"))

    # ── 호출 2: factsheet 매칭 (요구사항이 있고 factsheet가 있을 때만) ──
    if requirements and factsheet_md:
        match_user = (
            f"[Factsheet]\n{factsheet_md}\n\n"
            f"[부서 목록]\n{', '.join(config.get('dept_contacts', {}).keys()) or '(없음)'}\n\n"
            f"[요구사항 목록]\n"
            f"{json.dumps(requirements, ensure_ascii=False, indent=2)}"
        )
        matched = llm_client.call_llm_json([
            {"role": "system", "content": prompts["match"]},
            {"role": "user", "content": match_user},
        ])
        matched_list = matched if isinstance(matched, list) \
            else matched.get("requirements", [])
        by_id = {r.get("req_id"): r for r in matched_list if isinstance(r, dict)}
        merged = []
        for r in requirements:
            m = by_id.get(r["req_id"], {})
            r["matched_items"] = [str(x) for x in (m.get("matched_items") or [])]
            status = m.get("status", "데이터없음")
            r["status"] = status if status in VALID_STATUS else "데이터없음"
            r["owner_dept"] = m.get("owner_dept") if r["status"] == "데이터없음" else None
            merged.append(r)
        requirements = merged
    else:
        for r in requirements:
            r["status"] = "데이터없음"

    # 보유정보 연결 + 사내플랫폼 교차검증 (Phase B).
    #  ① 매칭된 factsheet 공개항목(public_yn=Y)에 값·출처를 붙인다 (C4 유지).
    #  ② 그 항목에 platform_id가 있으면 플랫폼의 최신·월별값을 덧붙인다.
    #     factsheet가 '공개 가능'을 보증했으므로 인용해도 안전(교차검증 통과).
    #  ③ 데이터없음 요구사항은 플랫폼에서 유사지표를 검색해 '사내보유(미검증)'
    #     후보로만 표시한다(초안 자동 인용 금지 — 담당자 공개검토 유도).
    for r in requirements:
        md = []
        for c in r["matched_items"]:
            if c not in factsheet_index:
                continue
            entry = dict(factsheet_index[c])
            entry["source_type"] = "factsheet"
            pid = str(entry.get("platform_id") or "").strip()
            if pid and platform_available():
                ind = platform_index.get(pid)
                if ind:
                    entry["annual"] = ind.get("annual")
                    entry["monthly"] = ind.get("monthly")
                    entry["cross_validated"] = True
                    if not entry.get("site"):
                        entry["site"] = ind.get("site")
            md.append(entry)
        r["matched_data"] = md

        # ③ 플랫폼 미검증 후보 (데이터없음일 때만)
        r["platform_candidates"] = []
        if r["status"] == "데이터없음" and platform_available():
            for ind in platform_index.search(r.get("content", ""), limit=3):
                val, yr = platform_index.latest_value(ind)
                r["platform_candidates"].append({
                    "platform_id": ind["platform_id"], "name": ind["name"],
                    "site": ind.get("site"), "unit": ind.get("unit"),
                    "value": val, "year": yr,
                    "has_monthly": bool(ind.get("monthly")),
                })

    # ── 고객사/긴급도는 코드에서 확정 ────────────────────────
    customer = match_customer(mail.get("sender", ""), config) \
        or extracted.get("customer") or "미분류"
    if customer not in customer_names + ["미분류"]:
        customer = "미분류"
    deadline = extracted.get("deadline") or None

    result = {
        "mail_id": mail["mail_id"],
        "entry_id": mail.get("entry_id", ""),
        "store_id": mail.get("store_id", ""),
        "received_at": mail.get("received_at", ""),
        "customer": customer,
        "sender": mail.get("sender", ""),
        "subject": mail.get("subject", ""),
        "request_type": extracted.get("request_type", "기타"),
        "framework": extracted.get("framework", "해당없음"),
        "deadline": deadline,
        "urgency": compute_urgency(deadline, mail.get("received_at"), config),
        "summary": extracted.get("summary", ""),
        "attachments": mail.get("attachments", []),
        "requirements": requirements,
        "needs_human_review": True,   # §5.2: 항상 true
        "status": "OK",
        "reply_draft": None,
        "dept_requests": [],
        "risks": [],
    }

    # ── 호출 3: 초안 생성 ────────────────────────────────────
    answerable = [r for r in requirements if r["status"] in ("답변가능", "부분가능")]
    missing = [r for r in requirements if r["status"] == "데이터없음"]
    if answerable or missing:
        draft_user = (
            f"[Factsheet]\n{factsheet_md or '(없음)'}\n\n"
            f"[메일 요약]\n고객사: {customer} / 제목: {mail.get('subject', '')}\n"
            f"{result['summary']}\n\n"
            f"[답변 대상 요구사항 (답변가능/부분가능)]\n"
            f"{json.dumps(answerable, ensure_ascii=False, indent=2)}\n\n"
            f"[데이터 없음 요구사항 (부서요청 대상)]\n"
            f"{json.dumps(missing, ensure_ascii=False, indent=2)}"
        )
        drafts = llm_client.call_llm_json([
            {"role": "system", "content": prompts["draft"]},
            {"role": "user", "content": draft_user},
        ])
        reply = drafts.get("reply_draft")
        if reply and answerable:
            result["reply_draft"] = f"{DRAFT_BANNER}\n\n{reply}"
        dept_requests = []
        for dr in drafts.get("dept_requests", []) or []:
            if not isinstance(dr, dict):
                continue
            dept_requests.append({
                "owner_dept": dr.get("owner_dept", "미지정"),
                "body": f"{DRAFT_BANNER}\n\n{dr.get('body', '')}",
            })
        result["dept_requests"] = dept_requests
        result["risks"] = _normalize_risks(drafts.get("risks"))

    return result


def analyze_all(mail_ids, config, progress_cb=None):
    """수집된 메일들을 1건씩 분석한다. 실패 건은 ERROR로 표기하고 계속 진행.

    progress_cb(done, total, subject): 웹 대시보드 진행률 표시용 콜백(선택).
    """
    os.makedirs(ANALYZED_DIR, exist_ok=True)
    prompts = {
        "extract": _load_prompt("extract.txt"),
        "match": _load_prompt("match.txt"),
        "draft": _load_prompt("draft.txt"),
    }
    factsheet_md = _load_factsheet()
    factsheet_index = _load_factsheet_index()

    results = []
    for i, mail_id in enumerate(mail_ids):
        in_path = os.path.join(INBOX_DIR, f"{mail_id}.json")
        try:
            with open(in_path, encoding="utf-8") as f:
                mail = json.load(f)
            if progress_cb:
                progress_cb(i, len(mail_ids), mail.get("subject", ""))
            result = analyze_mail(mail, config, factsheet_md, prompts,
                                  factsheet_index)
            log.info("분석 완료: %s | %s | 요구사항 %d건",
                     mail_id, result["customer"], len(result["requirements"]))
        except Exception as e:
            log.error("분석 실패: %s | %s", mail_id, e)
            mail = {}
            try:
                with open(in_path, encoding="utf-8") as f:
                    mail = json.load(f)
            except Exception:
                pass
            result = {
                "mail_id": mail_id,
                "entry_id": mail.get("entry_id", ""),
                "store_id": mail.get("store_id", ""),
                "received_at": mail.get("received_at", ""),
                "customer": mail.get("sender_name") or "미분류",
                "sender": mail.get("sender", ""),
                "subject": mail.get("subject", ""),
                "request_type": "기타",
                "framework": "해당없음",
                "deadline": None,
                "urgency": "낮음",
                "summary": "",
                "attachments": mail.get("attachments", []),
                "requirements": [],
                "needs_human_review": True,
                "status": "ERROR",
                "error": str(e),
                "reply_draft": None,
                "dept_requests": [],
                "risks": [],
            }
        out_path = os.path.join(ANALYZED_DIR, f"{mail_id}.json")
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        results.append(result)
    if progress_cb:
        progress_cb(len(mail_ids), len(mail_ids), "")
    return results
