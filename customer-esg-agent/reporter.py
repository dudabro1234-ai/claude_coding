# -*- coding: utf-8 -*-
"""③ 산출물 생성기.

작업지시서 §3 ③, §6:
- output/report_YYYYMMDD.html : 단일 파일 검토 대시보드 (외부 CDN 미참조)
    첫 화면 = 조회기간 내 전체 요청건 표 (필터·정렬·검색).
    행 클릭 = 해당 건 상세 (요청내용 → 보유정보 → 초안 → 리스크 검토).
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
URGENCY_RANK = {"긴급": 0, "보통": 1, "낮음": 2}
SEVERITY_RANK = {"높음": 0, "중간": 1, "낮음": 2}

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
    """상세 화면 카드: ①·② 요구사항/보유정보 → ③ 초안 → ④ 리스크."""
    is_error = m.get("status") == "ERROR"
    received = str(m.get("received_at", ""))[:16].replace("T", " ")

    req_blocks = "".join(_requirement_block(r)
                         for r in m.get("requirements", []))
    if not req_blocks:
        note = _esc(m.get("error", "요구사항이 추출되지 않았습니다."))
        req_blocks = f"<div class='nodata'>{note}</div>"

    drafts_html = ""
    if m.get("reply_draft"):
        drafts_html += _collapsible("고객 답변 초안", m["reply_draft"],
                                    f"{m['mail_id']}-reply", open_=True)
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
<div class="card{' card-error' if is_error else ''}" id="card-{_esc(m['mail_id'])}">
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


def _status_chips(m):
    """대응현황 요약 칩: 즉답 n · 부서협조 n · ERROR."""
    if m.get("status") == "ERROR":
        return _pill("분석실패", STATUS_STYLE, "ERROR")
    reqs = m.get("requirements", [])
    parts = []
    n = sum(1 for r in reqs if r.get("status") == "답변가능")
    if n:
        parts.append(_pill(f"즉답 {n}", STATUS_STYLE, "답변가능"))
    n = sum(1 for r in reqs if r.get("status") == "부분가능")
    if n:
        parts.append(_pill(f"부분 {n}", STATUS_STYLE, "부분가능"))
    n = sum(1 for r in reqs if r.get("status") == "데이터없음")
    if n:
        parts.append(_pill(f"부서협조 {n}", STATUS_STYLE, "데이터없음"))
    return " ".join(parts) or "<span class='dim'>요구사항 없음</span>"


def _max_risk(m):
    risks = m.get("risks") or []
    if not risks:
        return None
    return min(risks, key=lambda r: SEVERITY_RANK[r["severity"]])["severity"]


def _overview_row(m):
    received = str(m.get("received_at", ""))[:10]
    deadline = m.get("deadline") or ""
    max_risk = _max_risk(m)
    searchable = " ".join(str(m.get(k, "")) for k in
                          ("customer", "subject", "summary", "sender",
                           "request_type", "framework")).lower()
    flags = []
    if m.get("status") == "ERROR":
        flags.append("error")
    for r in m.get("requirements", []):
        if r.get("status") == "답변가능":
            flags.append("answerable")
        elif r.get("status") == "데이터없음":
            flags.append("dept")
    return f"""
<tr class="ov-row" data-id="{_esc(m['mail_id'])}"
    onclick="showDetail(this.dataset.id)"
    data-customer="{_esc(m.get('customer'))}"
    data-urgency="{_esc(m.get('urgency'))}"
    data-flags="{','.join(sorted(set(flags)))}"
    data-text="{_esc(searchable)}">
  <td data-v="{_esc(received)}">{_esc(received)}</td>
  <td data-v="{_esc(m.get('customer'))}"><a class="cust-link"
      onclick="filterCustomer(event, this)">{_esc(m.get('customer'))}</a></td>
  <td class="ov-subject" data-v="{_esc(m.get('subject'))}">{_esc(m.get('subject'))}</td>
  <td data-v="{_esc(m.get('request_type'))}">{_esc(m.get('request_type'))}<br>
      <span class="dim">{_esc(m.get('framework'))}</span></td>
  <td data-v="{_esc(deadline or '9999-12-31')}">{_esc(deadline or '미상')}
      {_dday(deadline)}</td>
  <td data-v="{URGENCY_RANK.get(m.get('urgency'), 2)}">
      {_pill(m.get('urgency'), URGENCY_STYLE)}</td>
  <td data-v="{len(m.get('requirements', []))}">{_status_chips(m)}</td>
  <td data-v="{SEVERITY_RANK.get(max_risk, 3)}">
      {_pill(max_risk, SEVERITY_STYLE) if max_risk else "<span class='dim'>—</span>"}</td>
  <td class="ov-arrow">›</td>
</tr>"""


def build_html(results, run_date=None):
    """검토 대시보드 HTML 문자열을 생성한다 (단일 파일, 외부 참조 없음).

    구조: [목록 화면] 전체 요청건 표 (필터·정렬·검색)
          → 행 클릭 → [상세 화면] 요청내용·보유정보·초안·리스크 + 목록 복귀.
    """
    run_date = run_date or datetime.date.today()
    kpi_html = "".join(
        f"<div class='kpi' style='--accent:{color}'>"
        f"<div class='kpi-num'>{v}</div><div class='kpi-label'>{k}</div></div>"
        for k, v, color in _kpi(results))

    # 기본 정렬: 긴급도 → 마감 임박 → 최신 수신
    ordered = sorted(results, key=lambda m: (
        URGENCY_RANK.get(m.get("urgency"), 2),
        m.get("deadline") or "9999-12-31",
        m.get("received_at", "")))
    rows = "".join(_overview_row(m) for m in ordered)
    cards = "".join(_mail_card(m) for m in ordered)

    received_dates = [str(m.get("received_at", ""))[:10]
                      for m in results if m.get("received_at")]
    period = (f"{min(received_dates)} ~ {max(received_dates)}"
              if received_dates else "-")
    customers = sorted({m.get("customer", "미분류") for m in results})
    customer_opts = "".join(f"<option value='{_esc(c)}'>{_esc(c)}</option>"
                            for c in customers)

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
  html {{ background: var(--bg); }}
  body {{ margin: 0; font-family: 'Malgun Gothic', 'Segoe UI', sans-serif;
         background: var(--bg); color: var(--text); min-height: 100vh; }}
  header {{ background: linear-gradient(120deg, #0f172a, #1e3a5f 60%, #134e4a);
           color: #e2e8f0; padding: 24px 32px; }}
  header h1 {{ margin: 0; font-size: 21px; letter-spacing: -0.3px; }}
  header h1 small {{ font-weight: 400; color: #94a3b8; margin-left: 8px; }}
  header .warn {{ display: inline-block; margin-top: 10px; font-size: 12px;
    color: #fbbf24; background: rgba(251,191,36,.12);
    border: 1px solid rgba(251,191,36,.4); border-radius: 8px;
    padding: 5px 12px; }}
  main {{ max-width: 1240px; margin: 0 auto; padding: 22px 32px 70px; }}
  .kpi-bar {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
             gap: 12px; margin: 4px 0 18px; }}
  .kpi {{ background: var(--card); border: 1px solid var(--line);
         border-radius: 14px; padding: 14px 18px; position: relative;
         overflow: hidden; box-shadow: 0 1px 3px rgba(15,23,42,.06); }}
  .kpi::before {{ content: ""; position: absolute; left: 0; top: 0; bottom: 0;
                 width: 4px; background: var(--accent); }}
  .kpi-num {{ font-size: 28px; font-weight: 800; color: var(--accent); }}
  .kpi-label {{ font-size: 12px; color: var(--sub); margin-top: 2px; }}
  .dim {{ color: #94a3b8; font-size: 12px; }}

  /* ── 목록 화면 ── */
  .toolbar {{ display: flex; gap: 10px; flex-wrap: wrap; align-items: center;
             margin-bottom: 12px; }}
  .toolbar .title {{ font-size: 15px; font-weight: 800; margin-right: auto; }}
  .toolbar .title .period {{ font-weight: 400; font-size: 12px;
                            color: var(--sub); margin-left: 8px; }}
  .toolbar select, .toolbar input {{ padding: 8px 12px; border-radius: 10px;
    border: 1px solid var(--line); background: #fff; font-size: 13px;
    color: var(--text); font-family: inherit; }}
  .toolbar input {{ width: 200px; }}
  .table-wrap {{ background: var(--card); border: 1px solid var(--line);
                border-radius: 16px; overflow: auto;
                box-shadow: 0 2px 8px rgba(15,23,42,.06); }}
  table.ov {{ width: 100%; border-collapse: collapse; font-size: 13px;
             min-width: 900px; }}
  table.ov th {{ position: sticky; top: 0; background: #f8fafc;
    color: var(--sub); font-size: 11px; text-transform: uppercase;
    letter-spacing: .05em; text-align: left; padding: 10px 12px;
    border-bottom: 2px solid var(--line); cursor: pointer;
    user-select: none; white-space: nowrap; }}
  table.ov th:hover {{ color: var(--text); }}
  table.ov th .arrow {{ font-size: 10px; }}
  table.ov td {{ padding: 10px 12px; border-bottom: 1px solid var(--line);
                vertical-align: top; }}
  tr.ov-row {{ cursor: pointer; transition: background .12s; }}
  tr.ov-row:hover {{ background: #f0fdfa; }}
  .ov-subject {{ max-width: 340px; font-weight: 600; color: #1e293b; }}
  .ov-arrow {{ color: #94a3b8; font-size: 18px; font-weight: 700; }}
  .cust-link {{ color: var(--accent2); font-weight: 700; cursor: pointer; }}
  .cust-link:hover {{ text-decoration: underline; }}
  .empty-row td {{ text-align: center; color: var(--sub); padding: 30px; }}

  /* ── 상세 화면 ── */
  #view-detail {{ display: none; }}
  .back-bar {{ display: flex; align-items: center; gap: 12px; margin-bottom: 14px; }}
  .back-btn {{ border: 1px solid var(--line); background: #fff; cursor: pointer;
    border-radius: 10px; padding: 9px 18px; font-size: 13px; font-weight: 700;
    font-family: inherit; }}
  .back-btn:hover {{ background: #f1f5f9; }}
  .nav-btn {{ border: 1px solid var(--line); background: #fff; cursor: pointer;
    border-radius: 10px; padding: 9px 14px; font-size: 13px; font-family: inherit; }}
  .nav-btn:disabled {{ opacity: .4; cursor: default; }}
  .nav-pos {{ font-size: 12px; color: var(--sub); }}
  .card {{ display: none; background: var(--card); border: 1px solid var(--line);
          border-radius: 16px; padding: 20px 24px;
          box-shadow: 0 2px 8px rgba(15,23,42,.06); }}
  .card.active {{ display: block; }}
  .card-error {{ border-color: #fca5a5; }}
  .card-head {{ display: flex; justify-content: space-between; gap: 10px;
               flex-wrap: wrap; align-items: center; }}
  .card-title {{ display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }}
  .customer {{ font-weight: 800; font-size: 16px; }}
  .subject {{ font-size: 14px; color: #334155; margin-top: 6px; font-weight: 600; }}
  .pill {{ border-radius: 999px; padding: 3px 12px; font-size: 12px;
          font-weight: 700; white-space: nowrap; display: inline-block; }}
  .dday {{ font-size: 12px; font-weight: 800; color: #0369a1;
          background: #e0f2fe; border-radius: 999px; padding: 3px 12px;
          white-space: nowrap; display: inline-block; }}
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
  table.fact .val {{ font-weight: 700; min-width: 90px; }}  /* 서술형 값 줄바꿈 허용 */
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

  /* ── 우측 문답 사이드바 ── */
  #chat-fab {{ position: fixed; right: 24px; bottom: 24px; z-index: 60;
    width: 56px; height: 56px; border-radius: 50%; border: 0; cursor: pointer;
    font-size: 24px; color: #052e22;
    background: linear-gradient(135deg, var(--accent), var(--accent2));
    box-shadow: 0 8px 24px rgba(15,23,42,.28); transition: transform .12s; }}
  #chat-fab:hover {{ transform: scale(1.06); }}
  body.chat-open #chat-fab {{ display: none; }}
  #chat-panel {{ position: fixed; top: 0; right: 0; width: 400px; height: 100vh;
    max-width: 92vw; background: var(--card); border-left: 1px solid var(--line);
    box-shadow: -12px 0 40px rgba(15,23,42,.15); z-index: 55;
    display: flex; flex-direction: column;
    transform: translateX(100%); transition: transform .22s ease; }}
  body.chat-open #chat-panel {{ transform: translateX(0); }}
  @media (min-width: 1500px) {{
    body.chat-open main {{ margin-right: 400px; }}
  }}
  .chat-head {{ padding: 14px 18px; border-bottom: 1px solid var(--line);
    display: flex; align-items: center; gap: 10px; }}
  .chat-head b {{ font-size: 14px; }}
  .chat-head .chat-sub {{ font-size: 11px; color: var(--sub); }}
  .chat-close {{ margin-left: auto; border: 0; background: none;
    font-size: 20px; cursor: pointer; color: var(--sub); }}
  .chat-close:hover {{ color: var(--text); }}
  #chat-log {{ flex: 1; overflow-y: auto; padding: 14px; }}
  .chat-msg {{ display: flex; margin-bottom: 10px; }}
  .chat-msg.user {{ justify-content: flex-end; }}
  .chat-bubble {{ max-width: 85%; padding: 9px 13px; border-radius: 12px;
    font-size: 13px; line-height: 1.65; white-space: pre-wrap;
    word-break: break-word; }}
  .chat-msg.user .chat-bubble {{ background: #d1fae5; color: #064e3b;
    border-bottom-right-radius: 4px; }}
  .chat-msg.bot .chat-bubble {{ background: #f1f5f9;
    border: 1px solid var(--line); border-bottom-left-radius: 4px; }}
  .chat-msg.bot .chat-bubble.err {{ border-color: #fca5a5; color: #991b1b; }}
  .chat-hint {{ font-size: 11px; color: var(--sub); text-align: center;
    margin: 4px 0 12px; line-height: 1.6; }}
  .chat-chips {{ display: flex; flex-wrap: wrap; gap: 6px;
    justify-content: center; margin-bottom: 8px; }}
  .chat-chip {{ border: 1px solid var(--line); background: #fff;
    color: var(--sub); border-radius: 999px; padding: 5px 11px;
    font-size: 11px; cursor: pointer; font-family: inherit; }}
  .chat-chip:hover {{ color: var(--text); border-color: var(--accent); }}
  #chat-form {{ display: flex; gap: 8px; padding: 12px 14px;
    border-top: 1px solid var(--line); }}
  #chat-q {{ flex: 1; padding: 10px 13px; border-radius: 10px; font-size: 13px;
    border: 1px solid var(--line); font-family: inherit; }}
  #chat-q:focus {{ outline: none; border-color: var(--accent); }}
  #chat-send {{ padding: 10px 16px; border: 0; border-radius: 10px;
    font-size: 13px; font-weight: 700; cursor: pointer; font-family: inherit;
    background: linear-gradient(90deg, var(--accent), var(--accent2));
    color: #052e22; }}
  #chat-send:disabled {{ opacity: .45; cursor: not-allowed; }}
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

  <div id="view-list">
    <div class="toolbar">
      <div class="title">전체 요청 목록
        <span class="period">수신 {period} · {len(results)}건</span></div>
      <select id="f-customer" onchange="applyFilters()">
        <option value="">고객사: 전체</option>{customer_opts}
      </select>
      <select id="f-urgency" onchange="applyFilters()">
        <option value="">긴급도: 전체</option>
        <option value="긴급">긴급</option>
        <option value="보통">보통</option>
        <option value="낮음">낮음</option>
      </select>
      <select id="f-flag" onchange="applyFilters()">
        <option value="">현황: 전체</option>
        <option value="answerable">즉답 가능 포함</option>
        <option value="dept">부서협조 필요 포함</option>
        <option value="error">분석 실패</option>
      </select>
      <input id="f-text" type="search" placeholder="제목·내용 검색…"
             oninput="applyFilters()">
    </div>
    <div class="table-wrap">
      <table class="ov">
        <thead><tr>
          <th onclick="sortBy(this, 0)">수신일 <span class="arrow"></span></th>
          <th onclick="sortBy(this, 1)">고객사 <span class="arrow"></span></th>
          <th onclick="sortBy(this, 2)">제목 <span class="arrow"></span></th>
          <th onclick="sortBy(this, 3)">유형 <span class="arrow"></span></th>
          <th onclick="sortBy(this, 4)">마감 <span class="arrow"></span></th>
          <th onclick="sortBy(this, 5)">긴급도 <span class="arrow"></span></th>
          <th>대응현황</th>
          <th onclick="sortBy(this, 7)">리스크 <span class="arrow"></span></th>
          <th></th>
        </tr></thead>
        <tbody id="ov-body">{rows}
          <tr class="empty-row" style="display:none"><td colspan="9">
            조건에 맞는 요청이 없습니다.</td></tr>
        </tbody>
      </table>
    </div>
  </div>

  <div id="view-detail">
    <div class="back-bar">
      <button class="back-btn" onclick="showList()">← 전체 목록</button>
      <button class="nav-btn" id="nav-prev" onclick="navDetail(-1)">‹ 이전</button>
      <span class="nav-pos" id="nav-pos"></span>
      <button class="nav-btn" id="nav-next" onclick="navDetail(1)">다음 ›</button>
    </div>
    {cards or "<p>표시할 메일이 없습니다.</p>"}
  </div>
</main>

<button id="chat-fab" onclick="toggleChat()" title="대응 현황 문답">💬</button>
<aside id="chat-panel">
  <div class="chat-head">
    <b>💬 대응 현황 문답</b>
    <span class="chat-sub">사내 LLM · 조회 전용</span>
    <button class="chat-close" onclick="toggleChat()">✕</button>
  </div>
  <div id="chat-log">
    <div class="chat-hint">분석 이력·관리대장(수기 처리상태)·공개 Factsheet를
      근거로 답합니다.<br>데이터에 없는 내용은 답하지 않습니다.</div>
    <div class="chat-chips">
      <button class="chat-chip" onclick="askChip(this)">마감 임박 요청은?</button>
      <button class="chat-chip" onclick="askChip(this)">처리 안 된 요청 정리해줘</button>
      <button class="chat-chip" onclick="askChip(this)">고위험 리스크 건은?</button>
      <button class="chat-chip" onclick="askChip(this)">Scope 1 요청 이력은?</button>
    </div>
  </div>
  <form id="chat-form" onsubmit="return chatSubmit(event)">
    <input id="chat-q" placeholder="예: OO 고객 건 진행상황은?" autocomplete="off">
    <button id="chat-send" type="submit">전송</button>
  </form>
</aside>

<script>
var currentList = [];   // 필터 적용된 mail_id 순서
var currentIdx = -1;

function visibleRows() {{
  return Array.prototype.slice.call(
    document.querySelectorAll("#ov-body tr.ov-row"))
    .filter(function(tr) {{ return tr.style.display !== "none"; }});
}}

function applyFilters() {{
  var c = document.getElementById("f-customer").value;
  var u = document.getElementById("f-urgency").value;
  var fl = document.getElementById("f-flag").value;
  var q = document.getElementById("f-text").value.trim().toLowerCase();
  var shown = 0;
  document.querySelectorAll("#ov-body tr.ov-row").forEach(function(tr) {{
    var ok = (!c || tr.dataset.customer === c)
          && (!u || tr.dataset.urgency === u)
          && (!fl || tr.dataset.flags.split(",").indexOf(fl) !== -1)
          && (!q || tr.dataset.text.indexOf(q) !== -1);
    tr.style.display = ok ? "" : "none";
    if (ok) shown++;
  }});
  document.querySelector(".empty-row").style.display = shown ? "none" : "";
}}

function filterCustomer(ev, a) {{
  ev.stopPropagation();
  document.getElementById("f-customer").value = a.textContent.trim();
  applyFilters();
}}

var sortState = {{ col: -1, asc: true }};
function sortBy(th, col) {{
  sortState.asc = (sortState.col === col) ? !sortState.asc : true;
  sortState.col = col;
  var body = document.getElementById("ov-body");
  var rows = Array.prototype.slice.call(body.querySelectorAll("tr.ov-row"));
  rows.sort(function(a, b) {{
    var va = a.cells[col].dataset.v || "", vb = b.cells[col].dataset.v || "";
    var na = parseFloat(va), nb = parseFloat(vb);
    var cmp = (!isNaN(na) && !isNaN(nb)) ? na - nb : va.localeCompare(vb, "ko");
    return sortState.asc ? cmp : -cmp;
  }});
  rows.forEach(function(r) {{ body.insertBefore(r, body.lastElementChild); }});
  document.querySelectorAll("table.ov th .arrow").forEach(function(s) {{
    s.textContent = ""; }});
  th.querySelector(".arrow").textContent = sortState.asc ? "▲" : "▼";
}}

function showDetail(mailId) {{
  currentList = visibleRows().map(function(tr) {{ return tr.dataset.id; }});
  currentIdx = currentList.indexOf(mailId);
  document.querySelectorAll(".card").forEach(function(c) {{
    c.classList.remove("active"); }});
  var card = document.getElementById("card-" + mailId);
  if (card) card.classList.add("active");
  document.getElementById("view-list").style.display = "none";
  document.getElementById("view-detail").style.display = "block";
  updateNav();
  window.scrollTo(0, 0);
}}

function navDetail(step) {{
  var next = currentIdx + step;
  if (next < 0 || next >= currentList.length) return;
  currentIdx = next;
  document.querySelectorAll(".card").forEach(function(c) {{
    c.classList.remove("active"); }});
  document.getElementById("card-" + currentList[currentIdx])
          .classList.add("active");
  updateNav();
  window.scrollTo(0, 0);
}}

function updateNav() {{
  document.getElementById("nav-prev").disabled = currentIdx <= 0;
  document.getElementById("nav-next").disabled =
      currentIdx >= currentList.length - 1;
  document.getElementById("nav-pos").textContent =
      (currentIdx + 1) + " / " + currentList.length;
}}

function showList() {{
  document.getElementById("view-detail").style.display = "none";
  document.getElementById("view-list").style.display = "block";
}}

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

/* ── 우측 문답 사이드바 ── */
var chatHistory = [];
var chatServed = (location.protocol !== "file:");  // 서버로 열렸는가

function toggleChat() {{
  var opened = document.body.classList.toggle("chat-open");
  if (opened && !chatServed && !document.getElementById("chat-offline")) {{
    var b = chatAdd("bot",
      "이 리포트가 파일로 직접 열려 있어 챗봇 서버에 연결할 수 없습니다.\\n" +
      "run_dashboard.bat 실행 후 열리는 화면에서 [검토 대시보드]로 들어오면 " +
      "여기서 바로 문답할 수 있습니다.", true);
    b.id = "chat-offline";
    document.getElementById("chat-q").disabled = true;
    document.getElementById("chat-send").disabled = true;
  }}
  if (opened) document.getElementById("chat-q").focus();
}}

function chatAdd(role, text, err) {{
  var log = document.getElementById("chat-log");
  var div = document.createElement("div");
  div.className = "chat-msg " + role;
  var b = document.createElement("div");
  b.className = "chat-bubble" + (err ? " err" : "");
  b.textContent = text;
  div.appendChild(b);
  log.appendChild(div);
  log.scrollTop = log.scrollHeight;
  return b;
}}

function askChip(btn) {{
  document.getElementById("chat-q").value = btn.textContent;
  chatSend();
}}
function chatSubmit(ev) {{ ev.preventDefault(); chatSend(); return false; }}

function chatSend() {{
  if (!chatServed) return;
  var input = document.getElementById("chat-q");
  var text = input.value.trim();
  if (!text) return;
  input.value = "";
  document.getElementById("chat-send").disabled = true;
  chatAdd("user", text);
  var wait = chatAdd("bot", "답변 작성 중…");
  fetch("/chat/api", {{ method: "POST",
    headers: {{"Content-Type": "application/json"}},
    body: JSON.stringify({{ message: text, history: chatHistory }}) }})
  .then(function(res) {{ return res.json(); }})
  .then(function(data) {{
    if (data.ok) {{
      wait.textContent = data.reply;
      chatHistory.push({{ role: "user", content: text }});
      chatHistory.push({{ role: "assistant", content: data.reply }});
      if (chatHistory.length > 20) chatHistory = chatHistory.slice(-20);
    }} else {{
      wait.classList.add("err");
      wait.textContent = data.error || "오류가 발생했습니다.";
    }}
  }})
  .catch(function(e) {{
    wait.classList.add("err");
    wait.textContent = "서버 연결 오류: " + e;
  }})
  .then(function() {{
    document.getElementById("chat-send").disabled = false;
    input.focus();
  }});
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
