# -*- coding: utf-8 -*-
"""고객 ESG 대응 Agent — CLI 진입점 (배치/스케줄러용).

사용법:
  python main.py                        # 전체 실행
  python main.py --since 2026-07-01 --until 2026-07-10   # 수신 기간 지정
  python main.py --skip-collect        # 수집 생략, work/inbox의 기존 JSON 재분석
  python main.py --limit 10            # 최대 10건만 수집
  python main.py --no-browser          # 리포트 자동 오픈 생략
  python main.py --no-drafts           # Outlook 초안 저장 생략

브라우저에서 기간을 지정해 실행하려면: python dashboard_server.py
"""
import argparse
import datetime
import logging
import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

import pipeline  # noqa: E402

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


def main(argv=None):
    parser = argparse.ArgumentParser(description="고객 ESG 대응 Agent")
    parser.add_argument("--since", help="수집 시작일 (YYYY-MM-DD, 경계 포함)")
    parser.add_argument("--until", help="수집 종료일 (YYYY-MM-DD, 경계 포함)")
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
    summary = pipeline.run(
        since=args.since, until=args.until, limit=args.limit,
        skip_collect=args.skip_collect, save_drafts=not args.no_drafts,
        open_browser=not args.no_browser)
    return 0 if summary["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
