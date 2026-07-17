# -*- coding: utf-8 -*-
"""수집 → 분석 → 산출물 파이프라인 (CLI main.py와 웹 dashboard_server.py 공용).

실행 흐름 (작업지시서 §3):
  0. LLM 연결 확인 (실패 시 즉시 중단, T1)
  1. collector : Outlook 메일 수집 (원본 미변경, 기간 필터 지원)
  2. analyzer  : 사내 LLM 분석 (1건 실패 시 ERROR 표기 후 계속, T2)
  3. reporter  : HTML 대시보드 / tracker.xlsx / Outlook 초안(저장만)
  4. 처리 완료 메일 ID 기록 (work/processed_ids.json, T5 중복 방지)
"""
import json
import logging
import os

import llm_client
import analyzer
import reporter

log = logging.getLogger(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
PROCESSED_PATH = os.path.join(BASE_DIR, "work", "processed_ids.json")


def load_config():
    with open(CONFIG_PATH, encoding="utf-8") as f:
        return json.load(f)


def load_processed_ids():
    try:
        with open(PROCESSED_PATH, encoding="utf-8") as f:
            return set(json.load(f))
    except Exception:
        return set()


def save_processed_ids(ids):
    os.makedirs(os.path.dirname(PROCESSED_PATH), exist_ok=True)
    with open(PROCESSED_PATH, "w", encoding="utf-8") as f:
        json.dump(sorted(ids), f, ensure_ascii=False, indent=2)


def pending_inbox_ids(processed_ids):
    """work/inbox에 있으나 아직 처리 완료되지 않은 mail_id 목록."""
    inbox = analyzer.INBOX_DIR
    if not os.path.isdir(inbox):
        return []
    return [name[:-5] for name in sorted(os.listdir(inbox))
            if name.endswith(".json") and name[:-5] not in processed_ids]


def run(since=None, until=None, limit=None, skip_collect=False,
        save_drafts=True, open_browser=True, progress_cb=None):
    """파이프라인 전체를 실행하고 요약 dict를 반환한다.

    since/until : 'YYYY-MM-DD' 수집 기간 (경계 포함, 생략 시 전체)
    progress_cb(stage, done, total, detail) : 진행 상황 콜백(선택)
        stage ∈ {connect, collect, analyze, report}

    반환: {"ok": bool, "message": str, "report_path": str|None,
           "total": int, "errors": int}
    """
    def notify(stage, done=0, total=0, detail=""):
        if progress_cb:
            progress_cb(stage, done, total, detail)

    config = load_config()

    # 0. LLM 연결 확인 — 실패 시 분석을 진행하지 않고 즉시 중단 (§4.3, T1)
    notify("connect")
    ok, msg = llm_client.test_connection()
    log.info(msg)
    if not ok:
        log.error("LLM 연결에 실패하여 실행을 중단합니다. "
                  "llm_config.json(base_url/api_key/model)을 확인하세요.")
        return {"ok": False, "message": msg, "report_path": None,
                "total": 0, "errors": 0}

    processed_ids = load_processed_ids()

    # 1. 수집 (원본 메일 미변경 — C3)
    if skip_collect:
        log.info("수집 생략 — work/inbox의 기존 JSON을 분석합니다.")
    else:
        notify("collect", detail=f"{since or '처음'} ~ {until or '현재'}")
        try:
            import collector
            collector.collect(config, processed_ids, limit=limit,
                              since=since, until=until)
        except RuntimeError as e:
            log.error("메일 수집 실패: %s", e)
            return {"ok": False, "message": f"메일 수집 실패: {e}",
                    "report_path": None, "total": 0, "errors": 0}

    target_ids = pending_inbox_ids(processed_ids)
    if not target_ids:
        log.info("처리할 신규 메일이 없습니다.")
        return {"ok": True, "message": "처리할 신규 메일이 없습니다.",
                "report_path": None, "total": 0, "errors": 0}
    log.info("분석 대상: %d건", len(target_ids))

    # 2. 분석 (부분 실패 허용 — T2)
    def analyze_progress(done, total, subject):
        notify("analyze", done, total, subject)
    results = analyzer.analyze_all(target_ids, config,
                                   progress_cb=analyze_progress)
    error_count = sum(1 for r in results if r.get("status") == "ERROR")
    if error_count:
        log.warning("분석 실패 %d건 — 리포트에 ERROR로 표기됩니다.", error_count)

    # 3. 산출물
    notify("report")
    reporter.append_tracker(results)
    if save_drafts and config.get("outlook", {}).get(
            "save_drafts_to_outlook", True):
        reporter.save_outlook_drafts(results, config)
    report_path = reporter.write_html(results, open_browser=open_browser)

    # 4. 처리 완료 기록 (성공 건만 — 실패 건은 다음 실행에서 재시도)
    done_ids = {r["mail_id"] for r in results if r.get("status") != "ERROR"}
    save_processed_ids(processed_ids | done_ids)

    msg = f"완료: {len(results)}건 처리 (실패 {error_count}건)"
    log.info("=== %s ===", msg)
    return {"ok": True, "message": msg, "report_path": report_path,
            "total": len(results), "errors": error_count}
