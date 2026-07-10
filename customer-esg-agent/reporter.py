# -*- coding: utf-8 -*-
"""③ 산출물 생성기.

작업지시서 §3 ③, §6:
- output/report_YYYYMMDD.html : 단일 파일 검토 대시보드 (외부 CDN 미참조)
    요구사항별로 「어떤 요청이 언제 왔고 → 보유 정보는 무엇이며 →
    어떤 내용으로 답변하고 → 어떤 리스크가 있는지」 흐름으로 표시한다.
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

STATUS_STYLE = {
    "답변가능":  ("#065f46", "#d1fae5"),
    "부분가능":  ("#92400e", "#fef3c7"),
    "데이터없음": ("#374151", "#e5e7eb"),
    "ERROR":    ("#991b1b", "#fee2e2"),
}
URGENCY_STYLE = {
    "긴급": ("#991b1b", "#fee2e2"),
    "보통": ("#92400e", "#fef3c7"),
    "낮음": ("#374151", "#e5e7eb"),
}
SEVERITY_STYLE = {
    "높음": ("#991b1b", "#fee2e2"),
    "중간": ("#92400e", "#fef3c7"),
    "낮음": ("#065f46", "#d1fae5"),
}

TRACKER_COLUMNS = ["수신일", "고객사", "요청유형", "프레임워크", "마감일", "긴급도",
                   "req_id", "요구내용", "상태", "담당부서",
                   "처리상태(수기)", "비고(수기)", "mail_id(시스템)"]

DRAFT_SUBJECT_PREFIX = "[AI초안] "


# ──────────────────────────────────────────────────────────────
# 6.1 HTML 검토 대시보드
# ──────────────────────────────────────────────────────────────

def _esc(s):
    return html.escape(str(s if s is not None else ""))


def _pill(text, style_map, key=None):
    fg, bg = style_map.get(key or text, ("#374151", "#e5e7eb"))
    return (f"<span class='pill' style='color:{fg};background:{bg}'>"
            f"{_esc(text)}</span>")


def _dday(deadline):
    if not deadline:
        return ""
    try:
        d = datetime.date.fromisoformat(str(deadline)[:10])
    except ValueError:
        return ""
    n = (d - datetime.date.today()).days
    if n < 0:
        return f"<span class='dday over'>D+{-n} 경과</span>"
    return f"<span class='dday{' hot' if n <= 7 else ''}'>D-{n}</span>"


def _kpi(results):
    reqs = [r for m in results for r in m.get("requirements", [])]
    high_risk = sum(1 for m in results
                    for rk in m.get("risks", []) if rk["severity"] == "높음")
    return [
        ("신규 요청", len(results), "#0ea5e9"),
        ("긴급", sum(1 for m in results if m.get("urgency") == "긴급"), "#ef4444"),
        ("즉답 가능", sum(1 for r in reqs if r.get("status") == "답변가능"), "#10b981"),
        ("부서협조 필요", sum(1 for r in reqs if r.get("status") == "데이터없음"), "#f59e0b"),
        ("고위험 리스크", high_risk, "#a855f7"),
        ("분석 실패", sum(1 for m in results if m.get("status") == "ERROR"), "#6b7280"),
    ]


def _matched_data_html(req):
    """보유 정보 블록: 매칭된 factsheet 항목의 실제 값·출처."""
    data = req.get("matched_data") or []
    if not data:
        codes = req.get("matched_items") or []
        if codes:
            return ("<div class='nodata'>매칭 코드: " + _esc(", ".join(codes))
                    + " (상세값은 factsheet.json 재생성 필요)</div>")
        dept = req.get("owner_dept")
        return ("<div class='nodata'>보유 데이터 없음"
                + (f" → <b>{_esc(dept)}</b>에 데이터 요청 초안 생성됨" if dept else "")
                + "</div>")
    rows = ""
    for d in data:
        rows += (f"<tr><td class='code'>{_esc(d.get('item_code'))}</td>"
                 f"<td>{_esc(d.get('item_name'))}</td>"
                 f"<td class='val'>{_esc(d.get('value'))} {_esc(d.get('unit'))}</td>"
                 f"<td>{_esc(d.get('year'))} · {_esc(d.get('site'))}</td>"
                 f"<td class='src'>{_esc(d.get('source'))}</td></tr>")
    return ("<table class='fact'><thead><tr><th>코드</th><th>지표</th>"
            "<th>값</th><th>기준</th><th>출처</th></tr></thead>"
            f"<tbody>{rows}</tbody></table>")


def _requirement_block(req):
    return f"""
  <div class="req">
    <div class="req-head">
      <span class="req-id">{_esc(req.get('req_id'))}</span>
      {_pill(req.get('status'), STATUS_STYLE)}
    </div>
    <div class="req-grid">
      <div>
        <div class="label">① 요청 내용</div>
        <div class="req-content">{_esc(req.get('content'))}</div>
      </div>
      <div>
        <div class="label">② 보유 정보 (Factsheet · 공개가능 데이터만)</div>
        {_matched_data_html(req)}
      </div>
    </div>
  </div>"""


def _risks_html(m):
    risks = m.get("risks") or []
    if m.get("status") == "ERROR":
        return ""
    if not risks:
        return ("<div class='section'><div class='label'>④ 답변 리스크 검토</div>"
                "<div class='nodata'>식별된 리스크 없음 — 그래도 발송 전 검토는 "
                "필수입니다.</div></div>")
    items = ""
    for rk in risks:
        items += f"""
    <div class="risk">
      {_pill(rk['severity'], SEVERITY_STYLE)}
      <div class="risk-body">
        <div><span class="risk-target">[{_esc(rk['target'])}]</span>
             {_esc(rk['description'])}</div>
        {f"<div class='risk-mit'>↳ 권고: {_esc(rk['mitigation'])}</div>"
         if rk.get('mitigation') else ""}
      </div>
    </div>"""
    return (f"<div class='section'><div class='label'>④ 답변 리스크 검토 "
            f"({len(risks)}건)</div>{items}</div>")


def _collapsible(title, body, uid, open_=False):
    return f"""
  <details class="draft"{' open' if open_ else ''}>
    <summary>{title}
      <button class="copy-btn" data-target="txt-{_esc(uid)}"
              onclick="copyDraft(event, this)">복사</button>
    </summary>
    <pre id="txt-{_esc(uid)}">{_esc(body)}</pre>
  </details>"""


def _mail_card(m):
    is_error = m.get("status") == "ERROR"
    received = str(m.get("received_at", ""))[:16].replace("T", " ")

    # ① 요청/② 보유정보 (요구사항별)
    req_blocks = "".join(_requirement_block(r)
                         for r in m.get("requirements", []))
    if not req_blocks:
        note = _esc(m.get("error", "요구사항이 추출되지 않았습니다."))
        req_blocks = f"<div class='nodata'>{note}</div>"

    # ③ 초안
    drafts_html = ""
    if m.get("reply_draft"):
        drafts_html += _collapsible("고객 답변 초안", m["reply_draft"],
                                    f"{m['mail_id']}-reply")
    for i, dr in enumerate(m.get("dept_requests", [])):
        drafts_html += _collapsible(
            f"부서요청 초안 → {_esc(dr.get('owner_dept', '미지정'))}",
            dr.get("body", ""), f"{m['mail_id']}-dept{i}")
    if drafts_html:
        drafts_html = ("<div class='section'><div class='label'>③ 답변/요청 초안 "
                       "(발송 전 반드시 검토)</div>" + drafts_html + "</div>")

    outlook_link = ""
    if m.get("entry_id"):
        outlook_link = (f"<a class='mail-link' href='outlook:{_esc(m['entry_id'])}'>"
                        f"✉ 원본 메일 열기</a>")

    return f"""
<div class="card{' card-error' if is_error else ''}">
  <div class="card-head">
    <div class="card-title">
      <span class="customer">{_esc(m.get('customer'))}</span>
      {_pill(m.get('urgency'), URGENCY_STYLE)}
      {_pill('분석실패', SEVERITY_STYLE, '높음') if is_error else ''}
      {_dday(m.get('deadline'))}
    </div>
    <div class="deadline">마감 {_esc(m.get('deadline') or '미상')}</div>
  </div>
  <div class="subject">{_esc(m.get('subject'))}</div>
  <div class="meta">
    <span>🕐 수신 {_esc(received)}</span>
    <span>{_esc(m.get('sender'))}</span>
    <span>{_esc(m.get('request_type'))} · {_esc(m.get('framework'))}</span>
    {outlook_link}
  </div>
  {f"<div class='summary'>{_esc(m.get('summary'))}</div>" if m.get('summary') else ""}
  <div class="section"><div class="label">①·② 요구사항 및 보유 정보</div>
    {req_blocks}</div>
  {drafts_html}
  {_risks_html(m)}
</div>"""


def build_html(results, run_date=None):
    """검토 대시보드 HTML 문자열을 생성한다 (단일 파일, 외부 참조 없음)."""
    run_date = run_date or datetime.date.today()
    kpi_html = "".join(
        f"<div class='kpi' style='--accent:{color}'>"
        f"<div class='kpi-num'>{v}</div><div class='kpi-label'>{k}</div></div>"
        for k, v, color in _kpi(results))

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
<title>ESG 검토 대시보드 — {run_date:%Y-%m-%d}</title>
<style>
  :root {{
    --bg: #f1f5f9; --card: #ffffff; --line: #e2e8f0;
    --text: #0f172a; --sub: #64748b;
    --accent: #10b981; --accent2: #0ea5e9;
  }}
  * {{ box-sizing: border-box; }}
  body {{ margin: 0; font-family: 'Malgun Gothic', 'Segoe UI', sans-serif;
         background: var(--bg); color: var(--text); }}
  header {{ background: linear-gradient(120deg, #0f172a, #1e3a5f 60%, #134e4a);
           color: #e2e8f0; padding: 26px 32px; }}
  header h1 {{ margin: 0; font-size: 21px; letter-spacing: -0.3px; }}
  header h1 small {{ font-weight: 400; color: #94a3b8; margin-left: 8px; }}
  header .warn {{ display: inline-block; margin-top: 10px; font-size: 12px;
    color: #fbbf24; background: rgba(251,191,36,.12);
    border: 1px solid rgba(251,191,36,.4); border-radius: 8px;
    padding: 5px 12px; }}
  main {{ max-width: 1180px; margin: 0 auto; padding: 22px 32px 70px; }}
  .kpi-bar {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
             gap: 12px; margin: 18px 0 8px; }}
  .kpi {{ background: var(--card); border: 1px solid var(--line);
         border-radius: 14px; padding: 14px 18px; position: relative;
         overflow: hidden; box-shadow: 0 1px 3px rgba(15,23,42,.06); }}
  .kpi::before {{ content: ""; position: absolute; left: 0; top: 0; bottom: 0;
                 width: 4px; background: var(--accent); }}
  .kpi-num {{ font-size: 28px; font-weight: 800; color: var(--accent); }}
  .kpi-label {{ font-size: 12px; color: var(--sub); margin-top: 2px; }}
  h2 {{ margin: 30px 0 12px; font-size: 17px; }}
  h2 .count {{ font-size: 13px; color: var(--sub); font-weight: 400; }}
  .card {{ background: var(--card); border: 1px solid var(--line);
          border-radius: 16px; padding: 20px 24px; margin-bottom: 18px;
          box-shadow: 0 2px 8px rgba(15,23,42,.06); }}
  .card-error {{ border-color: #fca5a5; }}
  .card-head {{ display: flex; justify-content: space-between; gap: 10px;
               flex-wrap: wrap; align-items: center; }}
  .card-title {{ display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }}
  .customer {{ font-weight: 800; font-size: 16px; }}
  .subject {{ font-size: 14px; color: #334155; margin-top: 6px; font-weight: 600; }}
  .pill {{ border-radius: 999px; padding: 3px 12px; font-size: 12px;
          font-weight: 700; white-space: nowrap; }}
  .dday {{ font-size: 12px; font-weight: 800; color: #0369a1;
          background: #e0f2fe; border-radius: 999px; padding: 3px 12px; }}
  .dday.hot {{ color: #991b1b; background: #fee2e2; }}
  .dday.over {{ color: #fff; background: #991b1b; }}
  .deadline {{ font-size: 12px; color: var(--sub); }}
  .meta {{ display: flex; gap: 16px; flex-wrap: wrap; font-size: 12px;
          color: var(--sub); margin: 8px 0 4px; }}
  .mail-link {{ color: var(--accent2); text-decoration: none; font-weight: 600; }}
  .summary {{ font-size: 13px; background: #f8fafc; border-left: 3px solid
             var(--accent2); padding: 8px 14px; border-radius: 0 8px 8px 0;
             margin: 10px 0; }}
  .section {{ margin-top: 16px; }}
  .label {{ font-size: 11px; font-weight: 800; color: var(--sub);
           text-transform: uppercase; letter-spacing: .06em; margin-bottom: 8px; }}
  .req {{ border: 1px solid var(--line); border-radius: 12px;
         padding: 12px 16px; margin-bottom: 10px; background: #fcfdfe; }}
  .req-head {{ display: flex; gap: 8px; align-items: center; margin-bottom: 8px; }}
  .req-id {{ font-family: Consolas, monospace; font-size: 12px; font-weight: 700;
            background: #eef2ff; color: #3730a3; border-radius: 6px;
            padding: 2px 8px; }}
  .req-grid {{ display: grid; grid-template-columns: 1fr 1.2fr; gap: 16px; }}
  @media (max-width: 800px) {{ .req-grid {{ grid-template-columns: 1fr; }} }}
  .req-content {{ font-size: 13px; line-height: 1.6; }}
  table.fact {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
  table.fact th, table.fact td {{ border: 1px solid var(--line);
    padding: 5px 8px; text-align: left; vertical-align: top; }}
  table.fact th {{ background: #f8fafc; color: var(--sub); font-size: 11px; }}
  table.fact .code {{ font-family: Consolas, monospace; white-space: nowrap; }}
  table.fact .val {{ font-weight: 700; white-space: nowrap; }}
  table.fact .src {{ color: var(--sub); }}
  .nodata {{ font-size: 13px; color: var(--sub); background: #f8fafc;
            border: 1px dashed var(--line); border-radius: 8px;
            padding: 10px 14px; }}
  details.draft {{ border: 1px solid var(--line); border-radius: 10px;
                  padding: 8px 14px; margin-bottom: 8px; background: #fcfdfe; }}
  details.draft summary {{ cursor: pointer; font-size: 13px; font-weight: 700; }}
  details.draft pre {{ white-space: pre-wrap; font-family: inherit;
    font-size: 13px; background: #f8fafc; padding: 14px;
    border-radius: 8px; line-height: 1.7; }}
  .copy-btn {{ margin-left: 10px; font-size: 12px; cursor: pointer;
    border: 1px solid var(--line); border-radius: 6px; background: #fff;
    padding: 2px 12px; }}
  .copy-btn:hover {{ background: #f1f5f9; }}
  .risk {{ display: flex; gap: 10px; align-items: flex-start;
          border: 1px solid var(--line); border-radius: 10px;
          padding: 10px 14px; margin-bottom: 8px; background: #fcfdfe; }}
  .risk-body {{ font-size: 13px; line-height: 1.6; }}
  .risk-target {{ font-family: Consolas, monospace; font-size: 12px;
                 color: var(--sub); }}
  .risk-mit {{ color: #065f46; font-size: 12px; margin-top: 3px; }}
</style>
</head>
<body>
<header>
  <h1>고객 ESG 요구사항 검토 대시보드 <small>{run_date:%Y-%m-%d}</small></h1>
  <div class="warn">⚠️ 모든 초안·리스크 검토는 AI가 생성했습니다. 발송 전 반드시
      사람이 검토하세요. 자동 발송 기능은 제공하지 않습니다.</div>
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
  function fallback() {{
    var ta = document.createElement("textarea");
    ta.value = text; document.body.appendChild(ta); ta.select();
    try {{ document.execCommand("copy"); done(); }} catch (e) {{}}
    document.body.removeChild(ta);
  }}
  if (navigator.clipboard && navigator.clipboard.writeText) {{
    navigator.clipboard.writeText(text).then(done, function() {{ fallback(); }});
  }} else {{ fallback(); }}
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
