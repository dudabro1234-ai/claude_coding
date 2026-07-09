"""
대화형 Agent 로직 — 결과 해석 + 시나리오 변경 제안

LLM 백엔드는 '교체 가능'하게 설계:
  - 환경변수 PPA_LLM_BASE_URL 이 설정되면 → OpenAI 호환 /chat/completions 호출 (사내 LLM)
  - 미설정이면 → 규칙 기반 Mock (개발/시연용, 전체 흐름 동일하게 동작)

서버(server.py)가 엔진을 돌려 결과(kpi)를 넘겨주고, 이 모듈은 자연어 이해/생성만 담당한다.
프로토콜:
  respond(message, history, params, kpi)
      → {"reply": str, "action": "none"|"propose", "params_patch": {...}, "label": str}
  summarize(patch, before_kpi, after_kpi) → str   (재실행 후 요약)
"""

import json
import os
import urllib.request

# ── LLM 접속 설정 ──────────────────────────────────────────────
# 우선순위: 환경변수 > llm_config.json (대시보드 ⚙설정 화면에서 저장)
# 사내 LLM(Qwen/GLM 등)은 보통 OpenAI 호환 /chat/completions 로 서빙됨.
CONFIG_PATH = os.environ.get("PPA_LLM_CONFIG_PATH") or os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "llm_config.json")


def _file_config():
    try:
        with open(CONFIG_PATH, encoding="utf-8") as fp:
            return json.load(fp)
    except Exception:
        return {}


def get_config():
    cfg = _file_config()
    return {
        "base_url": (os.environ.get("PPA_LLM_BASE_URL") or cfg.get("base_url") or "").strip(),
        "api_key": (os.environ.get("PPA_LLM_API_KEY") or cfg.get("api_key") or "").strip(),
        "model": (os.environ.get("PPA_LLM_MODEL") or cfg.get("model") or "").strip(),
    }


def save_config(base_url, api_key, model):
    cfg = {"base_url": (base_url or "").strip(),
           "api_key": (api_key or "").strip(),
           "model": (model or "").strip()}
    with open(CONFIG_PATH, "w", encoding="utf-8") as fp:
        json.dump(cfg, fp, ensure_ascii=False, indent=2)
    return cfg


def test_connection():
    """설정된 LLM에 초소형 요청을 보내 연결 상태를 확인한다."""
    c = get_config()
    if not c["base_url"]:
        return {"ok": False, "error": "Base URL이 비어 있습니다. ⚙설정에서 입력하세요."}
    try:
        out = _call_llm([{"role": "user", "content": "연결 테스트입니다. 'OK'라고만 답하세요."}], timeout=20)
        return {"ok": True, "reply": (out or "").strip()[:200], "model": c["model"]}
    except Exception as e:
        return {"ok": False, "error": f"{type(e).__name__}: {e}"}

# 엔진이 실제로 받는 파라미터(=Agent가 바꿀 수 있는 값)
PARAM_LABELS = {
    "re_goal_column": "RE 목표",
    "smr_mode": "SMR 포함",
    "flat_pv_mode": "균등태양광 포함",
    "flat_wt_mode": "균등풍력 포함",
    "sec_stop": "2043 SEC 가동중단",
}

SYSTEM_PROMPT = (
    "당신은 RE100 PPA 시뮬레이터의 분석 도우미입니다. 한국어로 간결하고 정확하게 답합니다.\n"
    "현재 시나리오 파라미터와 엔진 계산 결과(KPI)가 함께 제공됩니다. 결과 해석 질문에는 그 수치를 근거로 설명하세요.\n"
    "사용자가 시나리오 변경(SMR 추가/제거, RE 목표 변경, SEC 중단, 균등태양광/풍력 등)을 원하면, 바꿀 파라미터만 제안하고 사용자 확인을 요청하세요. 임의로 실행했다고 말하지 마세요.\n"
    "바꿀 수 있는 파라미터: re_goal_column(RE33_SEC|RE44_SEC|RE33|RE44), smr_mode(bool), flat_pv_mode(bool), flat_wt_mode(bool), sec_stop(bool).\n"
    "반드시 아래 JSON 형식만 출력하세요(다른 텍스트 금지):\n"
    '{"reply": "사용자에게 보일 한국어 답변", "action": "none" 또는 "propose", "params_patch": {바뀔 파라미터만, propose일 때만}}'
)


def llm_available():
    return bool(get_config()["base_url"])


def _call_llm(messages, timeout=60):
    c = get_config()
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


def _context(params, kpi):
    return (f"[현재 시나리오]\n{json.dumps(params, ensure_ascii=False)}\n"
            f"[현재 결과 KPI]\n25년 누적 총비용 {kpi.get('total25y_trillion')}조원, "
            f"2050 RE 달성률 {kpi.get('re_2050')}%, 2050 PPA 용량 {kpi.get('cap_2050')}MW, "
            f"EAC 평균 {kpi.get('eac_avg')}%")


def describe_patch(patch):
    parts = []
    for k, v in patch.items():
        label = PARAM_LABELS.get(k, k)
        if isinstance(v, bool):
            parts.append(f"{label} {'켜기' if v else '끄기'}")
        else:
            parts.append(f"{label}={v}")
    return ", ".join(parts)


# ──────────────────────────────────────────────────────────────────
# LLM 경로
# ──────────────────────────────────────────────────────────────────
def _respond_llm(message, history, params, kpi):
    messages = [{"role": "system", "content": SYSTEM_PROMPT}]
    for m in history[-8:]:
        if m.get("role") in ("user", "assistant"):
            messages.append({"role": m["role"], "content": m.get("content", "")})
    messages.append({"role": "user", "content": f"{_context(params, kpi)}\n\n[질문]\n{message}"})
    raw = _call_llm(messages)
    try:
        start, end = raw.find("{"), raw.rfind("}")
        obj = json.loads(raw[start:end + 1])
        reply = str(obj.get("reply", "")).strip() or "(빈 응답)"
        action = obj.get("action", "none")
        patch = obj.get("params_patch") or {}
        patch = {k: v for k, v in patch.items() if k in PARAM_LABELS}
        if action == "propose" and patch:
            return {"reply": reply, "action": "propose", "params_patch": patch, "label": describe_patch(patch)}
        return {"reply": reply, "action": "none"}
    except Exception:
        return {"reply": raw.strip()[:1500], "action": "none"}


def _summarize_llm(patch, before, after):
    prompt = (f"시나리오를 '{describe_patch(patch)}'로 변경해 다시 계산했습니다.\n"
              f"이전 KPI: {json.dumps(before, ensure_ascii=False)}\n"
              f"변경 후 KPI: {json.dumps(after, ensure_ascii=False)}\n"
              "변화의 핵심을 2~3문장으로 한국어로 요약해 주세요. JSON 없이 평문으로만.")
    messages = [{"role": "system", "content": "당신은 RE100 PPA 분석 도우미입니다. 한국어로 간결히."},
                {"role": "user", "content": prompt}]
    return _call_llm(messages).strip()


# ──────────────────────────────────────────────────────────────────
# Mock 경로 (규칙 기반)
# ──────────────────────────────────────────────────────────────────
NEG = ("빼", "제외", "끄", "off", "없애", "해제", "안 ", "안하", "않")


def _detect_patch(message):
    t = message.lower().replace(" ", "")
    raw = message
    patch = {}
    neg = any(n.replace(" ", "") in t for n in NEG)
    if "smr" in t:
        patch["smr_mode"] = not neg
    if "re44" in t or "44" in raw:
        patch["re_goal_column"] = "RE44_SEC"
    elif "re33" in t or "33" in raw:
        patch["re_goal_column"] = "RE33_SEC"
    if "sec" in t and any(k in raw for k in ("중단", "멈", "stop", "정지")):
        patch["sec_stop"] = True
    if "sec" in t and any(k in raw for k in ("유지", "계속", "가동")):
        patch["sec_stop"] = False
    if "균등풍력" in raw:
        patch["flat_wt_mode"] = not neg
    if "균등태양광" in raw and neg:
        patch["flat_pv_mode"] = False
    return patch


def _answer_mock(message, kpi):
    m = message
    total = kpi.get("total25y_trillion")
    re_2050 = kpi.get("re_2050")
    cap = kpi.get("cap_2050")
    eac = kpi.get("eac_avg")
    if any(k in m for k in ("비용", "원", "얼마", "cost")):
        return (f"현재 시나리오의 25년 누적 총비용은 약 **{total}조원**입니다(run_check 기준). "
                f"좌측 ‘연간 총비용 추이’ 차트에서 PPA·전력량요금·기본요금·EAC·망이용 구성을 볼 수 있어요. "
                f"시나리오를 바꿔보고 싶으면 예: ‘SMR 넣으면 어떻게 돼?’ 처럼 물어보세요.")
    if any(k in m for k in ("배분", "사업장", "IC", "CJ", "YI", "이천", "청주", "용인")):
        return ("사업장 배분은 좌측 ‘사업장 배분 실행’ 버튼으로 run_site를 돌리면 IC/CJ/YI별 MW가 나옵니다. "
                "일반적으로 기본요금 단가가 낮은 사업장에 더 많이 배분되어 전사 비용이 최소화됩니다. "
                "배분을 먼저 실행하신 뒤 다시 물어봐 주세요.")
    if any(k in m for k in ("달성", "re100", "재생", "비율", "%")):
        return (f"2050년 RE 달성률은 **{re_2050}%**, EAC 평균 비율은 **{eac}%**입니다. "
                f"EAC는 PPA로 못 채운 부분을 메우며 RE100 규칙상 최소 30%가 권장됩니다.")
    return (f"현재 시나리오 요약 — 25년 누적 총비용 **{total}조원**, 2050 RE 달성률 **{re_2050}%**, "
            f"PPA 총용량 **{cap}MW**, EAC 평균 **{eac}%**. "
            f"‘총비용 왜 이래?’, ‘사업장 배분 알려줘’, ‘SMR 넣으면?’ 처럼 물어보실 수 있어요.")


def _respond_mock(message, history, params, kpi):
    patch = _detect_patch(message)
    # 이미 같은 값이면 변경 의미 없음 → 질문으로 처리
    patch = {k: v for k, v in patch.items() if params.get(k) != v}
    if patch:
        desc = describe_patch(patch)
        return {"reply": f"**{desc}** 조건으로 다시 계산해 볼까요? 아래 ‘실행’을 누르면 시나리오를 다시 돌립니다.",
                "action": "propose", "params_patch": patch, "label": desc}
    if any(k in message for k in ("시나리오", "바꾸", "변경", "하면", "해보", "조건")) and not patch:
        return {"reply": "어떤 조건을 바꿀까요? 예: ‘SMR 포함’, ‘RE 목표 RE44로’, ‘SEC 가동중단’, ‘균등풍력 추가’.",
                "action": "none"}
    return {"reply": _answer_mock(message, kpi), "action": "none"}


# run_check 결과에는 영향이 없고 자동 최적화(run_greedy)에서만 의미 있는 파라미터
GREEDY_ONLY_KEYS = {"smr_mode", "flat_pv_mode", "flat_wt_mode"}


def _summarize_mock(patch, before, after):
    desc = describe_patch(patch)
    d = round(after["total25y_trillion"] - before["total25y_trillion"], 2)
    word = "증가" if d > 0 else ("감소" if d < 0 else "동일")
    sign = f"{abs(d):.2f}조 {word}" if d != 0 else "변화 없음"
    msg = (f"**{desc}** 적용해 다시 계산했습니다.\n"
           f"• 25년 누적 총비용: **{after['total25y_trillion']}조원** (이전 {before['total25y_trillion']}조 대비 {sign})\n"
           f"• 2050 RE 달성률 {after['re_2050']}% · EAC 평균 {after['eac_avg']}% · PPA 용량 {after['cap_2050']}MW\n"
           f"좌측 차트도 새 시나리오로 갱신했어요.")
    # SMR/균등 옵션은 기체결 평가(run_check)에는 영향 없음 → 자동 최적화 안내
    if d == 0 and set(patch) & GREEDY_ONLY_KEYS:
        msg += ("\n\nℹ️ 이 옵션은 **신규 계약을 자동으로 짜는 ‘자동 최적화(run_greedy)’에서만** 결과가 달라집니다. "
                "좌측 하단 **‘최적화 실행’ 버튼**을 눌러 보세요 — 방금 바꾼 옵션이 반영된 최적 포트폴리오를 새로 탐색합니다.")
    return msg


# ──────────────────────────────────────────────────────────────────
# 공개 API
# ──────────────────────────────────────────────────────────────────
def respond(message, history, params, kpi):
    if llm_available():
        try:
            return _respond_llm(message, history, params, kpi)
        except Exception as e:
            return {"reply": f"(사내 LLM 호출 실패 — Mock으로 답변)\n\n" + _respond_mock(message, history, params, kpi)["reply"],
                    "action": "none", "_llm_error": str(e)}
    return _respond_mock(message, history, params, kpi)


def summarize(patch, before, after):
    if llm_available():
        try:
            return _summarize_llm(patch, before, after)
        except Exception:
            return _summarize_mock(patch, before, after)
    return _summarize_mock(patch, before, after)
