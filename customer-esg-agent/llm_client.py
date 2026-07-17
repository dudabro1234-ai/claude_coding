# -*- coding: utf-8 -*-
"""사내 LLM 호출 클라이언트.

작업지시서 §4: ppa_dashboard/agent.py의 검증된 패턴을 이식한 구현.
- OpenAI 호환 /chat/completions 엔드포인트
- 표준 라이브러리(urllib)만 사용 (C2: 외부 인터넷 통신 금지, 사내 엔드포인트 전용)
- 설정 우선순위: 환경변수 > llm_config.json (C5: 키 하드코딩 금지)
"""
import json
import logging
import os
import time
import urllib.error
import urllib.request

log = logging.getLogger(__name__)

CONFIG_PATH = os.environ.get("ESG_LLM_CONFIG_PATH") or os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "llm_config.json")

MAX_RETRIES = 3
BACKOFF_BASE_SEC = 2  # 2s -> 4s -> 8s


def _file_config():
    try:
        with open(CONFIG_PATH, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def get_config():
    cfg = _file_config()
    return {
        "base_url": (os.environ.get("ESG_LLM_BASE_URL") or cfg.get("base_url") or "").strip(),
        "api_key":  (os.environ.get("ESG_LLM_API_KEY")  or cfg.get("api_key")  or "").strip(),
        "model":    (os.environ.get("ESG_LLM_MODEL")    or cfg.get("model")    or "").strip(),
    }


def _call_llm(messages, timeout=60):
    c = get_config()
    if not c["base_url"]:
        raise RuntimeError(
            "LLM base_url이 설정되지 않았습니다. "
            "llm_config.json 또는 ESG_LLM_BASE_URL 환경변수를 확인하세요.")
    body = json.dumps({"model": c["model"] or "default",
                       "messages": messages, "temperature": 0.2}).encode("utf-8")
    headers = {"Content-Type": "application/json"}
    if c["api_key"]:
        headers["Authorization"] = f"Bearer {c['api_key']}"
    req = urllib.request.Request(c["base_url"].rstrip("/") + "/chat/completions",
                                 data=body, headers=headers)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        data = json.loads(r.read().decode("utf-8"))
    return data["choices"][0]["message"]["content"]


def call_llm(messages, timeout=60, max_retries=MAX_RETRIES):
    """재시도 포함 LLM 호출 (§4.3: 최대 3회, 지수 백오프 2s→4s→8s)."""
    last_err = None
    for attempt in range(1, max_retries + 1):
        try:
            return _call_llm(messages, timeout=timeout)
        except Exception as e:
            last_err = e
            if attempt < max_retries:
                wait = BACKOFF_BASE_SEC * (2 ** (attempt - 1))
                log.warning("LLM 호출 실패 (%d/%d): %s — %ds 후 재시도",
                            attempt, max_retries, e, wait)
                time.sleep(wait)
    raise RuntimeError(f"LLM 호출이 {max_retries}회 모두 실패했습니다: {last_err}")


def call_llm_json(messages, timeout=90, max_retries=MAX_RETRIES):
    """LLM을 호출하고 응답에서 JSON 객체를 파싱해 반환한다.

    코드펜스(```json ... ```)로 감싸인 응답도 허용한다.
    """
    text = call_llm(messages, timeout=timeout, max_retries=max_retries)
    return extract_json(text)


def extract_json(text):
    """LLM 응답 텍스트에서 첫 번째 JSON 객체/배열을 파싱한다."""
    s = text.strip()
    if s.startswith("```"):
        # 코드펜스 제거
        first_nl = s.find("\n")
        if first_nl != -1:
            s = s[first_nl + 1:]
        if s.rstrip().endswith("```"):
            s = s.rstrip()[:-3]
        s = s.strip()
    # 본문 중간에 JSON이 있는 경우 대비: 첫 '{' 또는 '['부터 탐색
    for opener, closer in (("{", "}"), ("[", "]")):
        start = s.find(opener)
        if start == -1:
            continue
        depth = 0
        in_str = False
        esc = False
        for i in range(start, len(s)):
            ch = s[i]
            if in_str:
                if esc:
                    esc = False
                elif ch == "\\":
                    esc = True
                elif ch == '"':
                    in_str = False
            else:
                if ch == '"':
                    in_str = True
                elif ch == opener:
                    depth += 1
                elif ch == closer:
                    depth -= 1
                    if depth == 0:
                        return json.loads(s[start:i + 1])
        break
    raise ValueError(f"LLM 응답에서 JSON을 찾지 못했습니다: {text[:200]!r}")


def test_connection():
    """초소형 요청으로 연결 확인 (§4.3).

    실행 시작 시 1회 호출하고, 실패하면 분석을 진행하지 않고 즉시 중단해야 한다.
    반환: (성공여부: bool, 메시지: str)
    """
    try:
        reply = call_llm(
            [{"role": "user", "content": "OK라고만 답하세요"}],
            timeout=30, max_retries=1)
        return True, f"LLM 연결 확인: {reply.strip()[:50]}"
    except Exception as e:
        return False, f"LLM 연결 실패: {e}"
