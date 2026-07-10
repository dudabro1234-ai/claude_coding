# -*- coding: utf-8 -*-
"""③ 산출물 생성기.

작업지시서 §3 ③, §6:
- output/report_YYYYMMDD.html : 단일 파일 검토 대시보드 (외부 CDN 미참조)
- output/tracker.xlsx         : 누적 관리대장 (append 전용, 수기 열 보존)
- Outlook 임시보관함 초안 저장 (C1: 자동 발송 금지 — 저장만 수행)
"""
import datetime
import html
import json
import logging
import os
import webbrowser

log = logging.getLogger(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

STATUS_COLORS = {
    "답변가능": "#1a7f37",
    "부분가능": "#bc4c00",
    "데이터없음": "#6e7781",
    "ERROR": "#cf222e",
}
URGENCY_COLORS = {"긴급": "#cf222e", "보통": "#bc4c00", "낮음": "#6e7781"}

TRACKER_COLUMNS = ["수신일", "고객사", "요청유형", "프레임워크", "마감일", "긴급도",
                   "req_id", "요구내용", "상태", "담당부서",
                   "처리상태(수기)", "비고(수기)", "mail_id(시스템)"]

DRAFT_SUBJECT_PREFIX = "[AI초안] "


# ──────────────────────────────────────────────────────────────
# 6.1 HTML 검토 대시보드
# ──────────────────────────────────────────────────────────────

def _esc(s):
    return html.escape(str(s if s is not None else ""))


def _kpi(results):
    reqs = [r for m in results for r in m.get("requirements", [])]
    return {
        "신규": len(results),
        "긴급": sum(1 for m in results if m.get("urgency") == "긴급"),
        "즉답가능": sum(1 for r in reqs if r.get("status") == "답변가능"),
        "부서협조필요": sum(1 for r in reqs if r.get("status") == "데이터없음"),
        "분석실패": sum(1 for m in results if m.get("status") == "ERROR"),
    }


def _mail_card(m):
    urgency = m.get("urgency", "낮음")
    ucolor = URGENCY_COLORS.get(urgency, "#6e7781")
    is_error = m.get("status") == "ERROR"

    rows = ""
    for r in m.get("requirements", []):
        scolor = STATUS_COLORS.get(r.get("status"), "#6e7781")
        matched = ", ".join(r.get("matched_items", [])) or "—"
        dept = r.get("owner_dept") or ""
        rows += (
            f"<tr><td>{_esc(r.get('req_id'))}</td>"
            f"<td>{_esc(r.get('content'))}"
            + (f"<div class='dept'>담당부서(추정): {_esc(dept)}</div>" if dept else "")
            + f"</td><td><span class='badge' style='background:{scolor}'>"
            f"{_esc(r.get('status'))}</span></td>"
            f"<td>{_esc(matched)}</td></tr>"
        )
    if not rows:
        note = _esc(m.get("error", "요구사항이 추출되지 않았습니다."))
        rows = f"<tr><td colspan='4' class='empty'>{note}</td></tr>"

    outlook_link = ""
    if m.get("entry_id"):
        outlook_link = (f"<a class='mail-link' href='outlook:{_esc(m['entry_id'])}'>"
                        f"원본 메일 열기</a>")

    drafts_html = ""
    if m.get("reply_draft"):
        drafts_html += _collapsible(
            f"답변 초안 ({_esc(m.get('customer'))})", m["reply_draft"],
            f"{m['mail_id']}-reply")
    for i, dr in enumerate(m.get("dept_requests", [])):
        drafts_html += _collapsible(
            f"부서요청 초안 → {_esc(dr.get('owner_dept', '미지정'))}",
            dr.get("body", ""), f"{m['mail_id']}-dept{i}")

    error_badge = ("<span class='badge' style='background:#cf222e'>분석실패</span>"
                   if is_error else "")
    return f"""
<div class="card{' card-error' if is_error else ''}">
  <div class="card-head">
    <div>
      <span class="customer">{_esc(m.get('customer'))}</span>
      <span class="subject">{_esc(m.get('subject'))}</span>
    </div>
    <div>
      {error_badge}
      <span class="badge" style="background:{ucolor}">{_esc(urgency)}</span>
      <span class="deadline">마감: {_esc(m.get('deadline') or '미상')}</span>
    </div>
  </div>
  <div class="meta">수신 {_esc(m.get('received_at'))} · {_esc(m.get('sender'))}
    · {_esc(m.get('request_type'))} · {_esc(m.get('framework'))} {outlook_link}</div>
  <div class="summary">{_esc(m.get('summary'))}</div>
  <table>
    <thead><tr><th>req_id</th><th>내용</th><th>상태</th><th>매칭항목</th></tr></thead>
    <tbody>{rows}</tbody>
  </table>
  {drafts_html}
</div>"""


def _collapsible(title, body, uid):
    return f"""
  <details class="draft">
    <summary>{title}
      <button class="copy-btn" data-target="txt-{_esc(uid)}"
              onclick="copyDraft(event, this)">복사</button>
    </summary>
    <pre id="txt-{_esc(uid)}">{_esc(body)}</pre>
  </details>"""


def build_html(results, run_date=None):
    """검토 대시보드 HTML 문자열을 생성한다 (단일 파일, 외부 참조 없음)."""
    run_date = run_date or datetime.date.today()
    kpi = _kpi(results)
    kpi_html = "".join(
        f"<div class='kpi'><div class='kpi-num'>{v}</div>"
        f"<div class='kpi-label'>{k}</div></div>"
        for k, v in kpi.items())

    # 고객사별 그룹 (긴급 우선 정렬)
    order = {"긴급": 0, "보통": 1, "낮음": 2}
    by_customer = {}
    for m in results:
        by_customer.setdefault(m.get("customer", "미분류"), []).append(m)

    sections = ""
    for customer in sorted(by_customer,
                           key=lambda c: min(order.get(m.get("urgency"), 2)
                                             for m in by_customer[c])):
        mails = sorted(by_customer[customer],
                       key=lambda m: order.get(m.get("urgency"), 2))
        cards = "".join(_mail_card(m) for m in mails)
        sections += (f"<h2>{_esc(customer)} "
                     f"<span class='count'>{len(mails)}건</span></h2>{cards}")

    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<title>ESG 요구사항 검토 대시보드 — {run_date:%Y-%m-%d}</title>
<style>
  * {{ box-sizing: border-box; }}
  body {{ font-family: 'Malgun Gothic', 'Segoe UI', sans-serif; margin: 0;
         background: #f6f8fa; color: #1f2328; }}
  header {{ background: #1f2328; color: #fff; padding: 16px 28px; }}
  header h1 {{ margin: 0; font-size: 20px; }}
  header .warn {{ color: #ffd33d; font-size: 13px; margin-top: 4px; }}
  main {{ max-width: 1100px; margin: 0 auto; padding: 20px 28px 60px; }}
  .kpi-bar {{ display: flex; gap: 12px; margin: 18px 0 6px; flex-wrap: wrap; }}
  .kpi {{ background: #fff; border: 1px solid #d0d7de; border-radius: 8px;
         padding: 12px 22px; text-align: center; min-width: 110px; }}
  .kpi-num {{ font-size: 26px; font-weight: 700; }}
  .kpi-label {{ font-size: 12px; color: #57606a; margin-top: 2px; }}
  h2 {{ margin: 26px 0 10px; font-size: 17px; }}
  h2 .count {{ font-size: 13px; color: #57606a; font-weight: 400; }}
  .card {{ background: #fff; border: 1px solid #d0d7de; border-radius: 8px;
          padding: 14px 18px; margin-bottom: 14px; }}
  .card-error {{ border-color: #cf222e; }}
  .card-head {{ display: flex; justify-content: space-between; gap: 10px;
               flex-wrap: wrap; align-items: baseline; }}
  .customer {{ font-weight: 700; font-size: 15px; }}
  .subject {{ color: #57606a; margin-left: 8px; font-size: 14px; }}
  .badge {{ color: #fff; border-radius: 10px; padding: 2px 10px;
           font-size: 12px; white-space: nowrap; }}
  .deadline {{ font-size: 12px; color: #57606a; margin-left: 6px; }}
  .meta {{ font-size: 12px; color: #57606a; margin: 6px 0; }}
  .mail-link {{ margin-left: 8px; }}
  .summary {{ font-size: 13px; margin: 8px 0; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 13px; margin-top: 6px; }}
  th, td {{ border: 1px solid #d0d7de; padding: 6px 9px; text-align: left;
           vertical-align: top; }}
  th {{ background: #f6f8fa; }}
  .dept {{ font-size: 12px; color: #57606a; margin-top: 3px; }}
  .empty {{ color: #6e7781; }}
  details.draft {{ margin-top: 10px; border: 1px dashed #d0d7de;
                  border-radius: 6px; padding: 6px 12px; }}
  details.draft summary {{ cursor: pointer; font-size: 13px; font-weight: 600; }}
  details.draft pre {{ white-space: pre-wrap; font-family: inherit;
                      font-size: 13px; background: #f6f8fa; padding: 12px;
                      border-radius: 6px; }}
  .copy-btn {{ margin-left: 10px; font-size: 12px; cursor: pointer;
              border: 1px solid #d0d7de; border-radius: 5px;
              background: #f6f8fa; padding: 2px 10px; }}
  .copy-btn:hover {{ background: #eaeef2; }}
</style>
</head>
<body>
<header>
  <h1>고객 ESG 요구사항 검토 대시보드 <small>{run_date:%Y-%m-%d}</small></h1>
  <div class="warn">⚠️ 모든 초안은 AI가 생성한 것입니다. 발송 전 반드시 검토하세요.
      (자동 발송 기능은 의도적으로 제공하지 않습니다)</div>
</header>
<main>
  <div class="kpi-bar">{kpi_html}</div>
  {sections or "<p>표시할 메일이 없습니다.</p>"}
</main>
<script>
function copyDraft(ev, btn) {{
  ev.preventDefault(); ev.stopPropagation();
  var text = document.getElementById(btn.dataset.target).innerText;
  function done() {{ btn.textContent = "복사됨!";
                     setTimeout(function() {{ btn.textContent = "복사"; }}, 1500); }}
  if (navigator.clipboard && navigator.clipboard.writeText) {{
    navigator.clipboard.writeText(text).then(done, function() {{ fallback(); }});
  }} else {{ fallback(); }}
  function fallback() {{
    var ta = document.createElement("textarea");
    ta.value = text; document.body.appendChild(ta); ta.select();
    try {{ document.execCommand("copy"); done(); }} catch (e) {{}}
    document.body.removeChild(ta);
  }}
}}
</script>
</body>
</html>"""


def write_html(results, open_browser=True):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    path = os.path.join(OUTPUT_DIR,
                        f"report_{datetime.date.today():%Y%m%d}.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(build_html(results))
    log.info("HTML 리포트 생성: %s", path)
    if open_browser:
        try:
            webbrowser.open("file://" + os.path.abspath(path))
        except Exception as e:
            log.warning("브라우저 자동 오픈 실패: %s", e)
    return path


# ──────────────────────────────────────────────────────────────
# 6.2 Excel 누적 트래커 (append 전용, 수기 열 보존)
# ──────────────────────────────────────────────────────────────

def append_tracker(results, tracker_path=None):
    """tracker.xlsx에 신규 행만 추가한다.

    - 기존 파일의 어떤 행도 수정/삭제하지 않는다 (C3).
    - 중복 판정 키: mail_id + req_id (§6.2) — 재실행 시 수기 입력값 보존.
    """
    from openpyxl import Workbook, load_workbook

    tracker_path = tracker_path or os.path.join(OUTPUT_DIR, "tracker.xlsx")
    os.makedirs(os.path.dirname(tracker_path), exist_ok=True)

    if os.path.exists(tracker_path):
        wb = load_workbook(tracker_path)
        ws = wb.active
        header = [c.value for c in ws[1]]
        try:
            mid_col = header.index("mail_id(시스템)")
            rid_col = header.index("req_id")
        except ValueError:
            raise RuntimeError(
                f"기존 {tracker_path}의 헤더가 예상과 다릅니다. "
                "파일을 백업 후 수동 확인하세요. (덮어쓰기 금지)")
        existing = set()
        for row in ws.iter_rows(min_row=2, values_only=True):
            existing.add((row[mid_col], row[rid_col]))
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "ESG요청관리대장"
        ws.append(TRACKER_COLUMNS)
        existing = set()

    added = 0
    for m in results:
        base = [str(m.get("received_at", ""))[:10], m.get("customer", ""),
                m.get("request_type", ""), m.get("framework", ""),
                m.get("deadline") or "", m.get("urgency", "")]
        reqs = m.get("requirements") or []
        if not reqs:
            # 요구사항이 없는 건(분석실패 포함)도 1행으로 기록해 추적한다.
            key = (m["mail_id"], "-")
            if key not in existing:
                status = "ERROR" if m.get("status") == "ERROR" else "데이터없음"
                ws.append(base + ["-", m.get("subject", ""), status, "",
                                  "", "", m["mail_id"]])
                existing.add(key)
                added += 1
            continue
        for r in reqs:
            key = (m["mail_id"], r["req_id"])
            if key in existing:
                continue
            ws.append(base + [r["req_id"], r["content"], r["status"],
                              r.get("owner_dept") or "", "", "", m["mail_id"]])
            existing.add(key)
            added += 1

    wb.save(tracker_path)
    log.info("tracker.xlsx 갱신: 신규 %d행 추가 (%s)", added, tracker_path)
    return added


# ──────────────────────────────────────────────────────────────
# 6.3 Outlook 임시보관함 초안 저장
#     C1: 자동 발송 절대 금지 — Save로 임시보관함 저장만 수행한다.
# ──────────────────────────────────────────────────────────────

def save_outlook_drafts(results, config):
    """답변/부서요청 초안을 Outlook 임시보관함(Drafts)에 저장한다.

    발송 동작은 어떤 경우에도 수행하지 않는다. 최종 검토·발송은 사용자 몫.
    """
    try:
        import win32com.client
    except ImportError:
        log.warning("pywin32 미설치 — Outlook 초안 저장을 건너뜁니다. "
                    "(초안은 HTML 리포트에서 복사 가능)")
        return 0

    dept_contacts = config.get("dept_contacts", {})
    outlook = win32com.client.Dispatch("Outlook.Application")
    saved = 0

    def _save_draft(subject, to_addr, body):
        nonlocal saved
        item = outlook.CreateItem(0)  # 0 = olMailItem
        item.Subject = DRAFT_SUBJECT_PREFIX + subject
        if to_addr:
            item.To = to_addr
        item.Body = body
        item.Save()  # 임시보관함 저장만. 발송 호출은 금지되어 있다 (C1).
        saved += 1

    for m in results:
        if m.get("status") == "ERROR":
            continue
        try:
            if m.get("reply_draft"):
                # 수신자는 비워둔다(§6.3) — 사용자가 검토 후 직접 지정.
                _save_draft(f"RE: {m.get('subject', '')} — {m.get('customer', '')} 답변",
                            "", m["reply_draft"])
            for dr in m.get("dept_requests", []):
                dept = dr.get("owner_dept", "미지정")
                to_addr = dept_contacts.get(dept, "")
                _save_draft(f"[데이터요청] {dept} — {m.get('customer', '')} ESG 요구 대응",
                            to_addr, dr.get("body", ""))
        except Exception as e:
            log.warning("초안 저장 실패 (mail_id=%s): %s", m.get("mail_id"), e)

    log.info("Outlook 임시보관함 초안 저장: %d건", saved)
    return saved
