# -*- coding: utf-8 -*-
"""사내 데이터플랫폼 브라우저 자동화 커넥터 (Phase C, opt-in).

설계 원칙:
- **로그인은 재사용, 클릭만 자동화.** persistent 브라우저 프로필로 실행해
  담당자가 평소 로그인해 둔 세션(쿠키)을 그대로 재사용한다. 자격증명을 저장·입력
  하지 않는다(C5). 세션이 만료돼 로그인 화면이 뜨면 일시정지하고 수동 로그인을
  기다린다(반자동 폴백).
- **의존성 격리.** 이 모듈만 Playwright에 의존한다(requirements-connector.txt).
  코어 agent는 표준 라이브러리만 쓴다. Playwright 미설치 시 명확히 안내하고,
  수동 드롭인(data/platform_drop/에 직접 넣기)으로 그대로 대체 가능하다.
- **클릭 경로는 설정.** platform_nav.json에 URL과 클릭 스텝을 기술 → 코드 수정 없이
  UI 변경에 대응.
- 접속 대상은 **사내 URL**이므로 C2(외부 통신 금지) 위반이 아니다.

사용:
  python platform_connector.py                 # platform_nav.json 사용, 다운로드→인덱싱
  python platform_connector.py --headless      # (로그인 세션 유효할 때만) 무인 실행
  python platform_connector.py --no-ingest     # 다운로드만, 인덱싱 생략
"""
import argparse
import json
import logging
import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
NAV_PATH = os.path.join(BASE_DIR, "platform_nav.json")
DROP_DIR = os.path.join(BASE_DIR, "data", "platform_drop")
DEFAULT_PROFILE = os.path.join(BASE_DIR, "browser_profile")

log = logging.getLogger("platform_connector")


def load_nav(path=NAV_PATH):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def _resolve_locator(page, step):
    """스텝 정의(selector 또는 text)를 Playwright locator로 변환."""
    if step.get("selector"):
        return page.locator(step["selector"])
    if step.get("text"):
        return page.get_by_text(step["text"], exact=step.get("exact", False))
    if step.get("role") and step.get("name"):
        return page.get_by_role(step["role"], name=step["name"])
    raise ValueError(f"스텝에 selector/text/role 중 하나가 필요합니다: {step}")


def _run_step(page, step, timeout):
    action = step.get("action", "click")
    if action == "wait":
        page.wait_for_timeout(step.get("ms", 1000))
        return
    if action == "goto":
        page.goto(step["url"], timeout=timeout)
        return
    loc = _resolve_locator(page, step)
    loc.wait_for(state="visible", timeout=timeout)
    if action == "click":
        loc.click(timeout=timeout)
    elif action == "fill":
        loc.fill(step.get("value", ""), timeout=timeout)
    elif action == "select":
        loc.select_option(step.get("value"), timeout=timeout)
    else:
        raise ValueError(f"알 수 없는 action: {action}")


def sync(nav=None, profile_dir=DEFAULT_PROFILE, drop_dir=DROP_DIR,
         headless=False, login_wait=None):
    """플랫폼에 접속해 클릭 경로를 따라 datasheet를 내려받아 drop_dir에 저장.

    login_wait: 로그인 화면 감지 시 호출되는 콜백(없으면 콘솔에서 Enter 대기).
    반환: 저장된 파일 경로.
    """
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        raise RuntimeError(
            "Playwright가 설치되어 있지 않습니다. 자동화를 쓰려면:\n"
            "  pip install -r requirements-connector.txt\n"
            "  python -m playwright install chromium\n"
            "설치가 어려우면 플랫폼에서 받은 xlsx를 data/platform_drop/에 직접 "
            "넣고 tools/platform_ingest.py를 실행하세요(수동 대체).")

    nav = nav or load_nav()
    os.makedirs(drop_dir, exist_ok=True)
    os.makedirs(profile_dir, exist_ok=True)
    timeout = nav.get("timeout_ms", 30000)
    # 폐쇄망 등에서 Chromium 경로를 직접 지정해야 할 때 (선택):
    #   set PLAYWRIGHT_CHROMIUM_PATH=C:\path\to\chrome.exe
    exe = os.environ.get("PLAYWRIGHT_CHROMIUM_PATH") or nav.get("executable_path")
    launch_kwargs = {"headless": headless, "accept_downloads": True}
    if exe:
        launch_kwargs["executable_path"] = exe

    with sync_playwright() as p:
        # persistent context = 실제 로그인 세션(쿠키) 재사용, 자격증명 저장 안 함
        ctx = p.chromium.launch_persistent_context(profile_dir, **launch_kwargs)
        try:
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            page.goto(nav["url"], timeout=timeout)

            # 로그인 화면 감지 → 반자동 폴백 (자격증명 자동입력 안 함)
            login_sel = nav.get("login_check")
            if login_sel and page.locator(login_sel).count() > 0:
                log.warning("로그인 세션이 만료된 것 같습니다. 브라우저에서 직접 "
                            "로그인해 주세요.")
                if login_wait:
                    login_wait(page)
                else:
                    input(">> 브라우저에서 로그인을 마친 뒤 이 창에서 Enter를 "
                          "누르세요... ")

            for step in nav.get("steps", []):
                _run_step(page, step, timeout)

            # 다운로드 트리거 → 파일 저장
            dl_step = nav["download"]
            with page.expect_download(timeout=timeout) as dl_info:
                _run_step(page, dl_step, timeout)
            download = dl_info.value
            fname = download.suggested_filename or "platform_export.xlsx"
            target = os.path.join(drop_dir, fname)
            download.save_as(target)
            log.info("다운로드 완료: %s", target)
            return target
        finally:
            ctx.close()


def main(argv=None):
    parser = argparse.ArgumentParser(description="사내 데이터플랫폼 자동 수집")
    parser.add_argument("--headless", action="store_true",
                        help="무인 실행 (로그인 세션이 유효할 때만 성공)")
    parser.add_argument("--no-ingest", action="store_true",
                        help="다운로드만 하고 인덱싱은 생략")
    parser.add_argument("--nav", default=NAV_PATH, help="platform_nav.json 경로")
    args = parser.parse_args(argv)

    logging.basicConfig(level=logging.INFO,
                        format="%(asctime)s [%(levelname)s] %(message)s")
    if not os.path.exists(args.nav):
        log.error("클릭 경로 설정 파일이 없습니다: %s\n"
                  "platform_nav.example.json을 복사해 만드세요.", args.nav)
        return 1
    try:
        path = sync(load_nav(args.nav))
    except Exception as e:
        log.error("자동 수집 실패: %s", e)
        return 1

    if not args.no_ingest:
        sys.path.insert(0, os.path.join(BASE_DIR, "tools"))
        import platform_ingest
        n, out = platform_ingest.convert(path)
        log.info("인덱싱 완료: %s (지표 %d개)", out, n)
    return 0


if __name__ == "__main__":
    sys.exit(main())
