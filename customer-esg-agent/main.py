# -*- coding: utf-8 -*-
"""고객 ESG 대응 Agent — 오케스트레이터.

실행 흐름 (작업지시서 §3):
  0. 로깅 설정 + LLM 연결 확인 (실패 시 즉시 중단, T1)
  1. collector : Outlook 메일 수집 (원본 미변경)
  2. analyzer  : 사내 LLM 분석 (1건 실패 시 ERROR 표기 후 계속, T2)
  3. reporter  : HTML 대시보드 / tracker.xlsx / Outlook 초안(저장만)
  4. 처리 완료 메일 ID 기록 (processed_ids.json, T5 중복 방지)

사용법:
  python main.py                 # 전체 실행
  python main.py --skip-collect  # 수집 생략, work/inbox의 기존 JSON 재분석
  python main.py --limit 10      # 최대 10건만 수집
  python main.py --no-browser    # 리포트 자동 오픈 생략
"""
import argparse
import datetime
import json
import logging
import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

import llm_client   # noqa: E402
import analyzer     # noqa: E402
import reporter     # noqa: E402

CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
PROCESSED_PATH = os.path.join(BASE_DIR, "work", "processed_ids.json")
LOG_DIR = os.path.join(BASE_DIR, "logs")

log = logging.getLogger("main")


def setup_logging():
    os.makedirs(LOG_DIR, exist_ok=True)
    log_path = os.path.join(LOG_DIR, f"{datetime.date.today():%Y%m%d}.log")
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        handlers=[logging.FileHandler(log_path, encoding="utf-8"),
                  logging.StreamHandler(sys.stdout)])


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
    ids = []
    for name in sorted(os.listdir(inbox)):
        if name.endswith(".json"):
            mail_id = name[:-5]
            if mail_id not in processed_ids:
                ids.append(mail_id)
    return ids


def main(argv=None):
    parser = argparse.ArgumentParser(description="고객 ESG 대응 Agent")
    parser.add_argument("--skip-collect", action="store_true",
                        help="Outlook 수집을 건너뛰고 work/inbox의 기존 JSON을 분석")
    parser.add_argument("--limit", type=int, default=None,
                        help="이번 실행에서 수집할 최대 메일 수")
    parser.add_argument("--no-browser", action="store_true",
                        help="완료 후 HTML 리포트 자동 오픈 생략")
    parser.add_argument("--no-drafts", action="store_true",
                        help="Outlook 임시보관함 초안 저장 생략")
    args = parser.parse_args(argv)

    setup_logging()
    log.info("=== 고객 ESG 대응 Agent 시작 ===")
    config = load_config()

    # 0. LLM 연결 확인 — 실패 시 분석을 진행하지 않고 즉시 중단 (§4.3, T1)
    ok, msg = llm_client.test_connection()
    log.info(msg)
    if not ok:
        log.error("LLM 연결에 실패하여 실행을 중단합니다. "
                  "llm_config.json(base_url/api_key/model)을 확인하세요.")
        return 1

    processed_ids = load_processed_ids()

    # 1. 수집 (원본 메일 미변경 — C3)
    if args.skip_collect:
        log.info("--skip-collect: Outlook 수집을 건너뜁니다.")
    else:
        try:
            import collector
            collector.collect(config, processed_ids, limit=args.limit)
        except RuntimeError as e:
            log.error("메일 수집 실패: %s", e)
            return 1

    target_ids = pending_inbox_ids(processed_ids)
    if not target_ids:
        log.info("처리할 신규 메일이 없습니다. 종료합니다.")
        return 0
    log.info("분석 대상: %d건", len(target_ids))

    # 2. 분석 (부분 실패 허용 — T2)
    results = analyzer.analyze_all(target_ids, config)
    error_count = sum(1 for r in results if r.get("status") == "ERROR")
    if error_count:
        log.warning("분석 실패 %d건 — 리포트에 ERROR로 표기됩니다.", error_count)

    # 3. 산출물
    reporter.append_tracker(results)
    if not args.no_drafts and config.get("outlook", {}).get(
            "save_drafts_to_outlook", True):
        reporter.save_outlook_drafts(results, config)
    reporter.write_html(results, open_browser=not args.no_browser)

    # 4. 처리 완료 기록 (성공 건만 — 실패 건은 다음 실행에서 재시도)
    done = {r["mail_id"] for r in results if r.get("status") != "ERROR"}
    save_processed_ids(processed_ids | done)

    log.info("=== 완료: 처리 %d건 (실패 %d건) ===", len(results), error_count)
    return 0


if __name__ == "__main__":
    sys.exit(main())
