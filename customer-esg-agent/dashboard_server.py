# -*- coding: utf-8 -*-
"""웹 대시보드 서버 — 브라우저에서 기간을 지정하고 [분석 시작]을 누르는 실행 방식.

- Python 표준 라이브러리(http.server)만 사용, 127.0.0.1 전용 바인딩.
  외부 네트워크로는 아무것도 열지 않는다 (C2 준수 — 사내 LLM 호출은 기존과 동일).
- 실행: python dashboard_server.py  (또는 run_dashboard.bat)
  → 브라우저가 자동으로 http://127.0.0.1:8765 를 연다.
"""
import datetime
import glob
import http.server
import json
import logging
import os
import sys
import threading
import webbrowser

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

import chatbot   # noqa: E402
import pipeline  # noqa: E402
import reporter  # noqa: E402

log = logging.getLogger("dashboard")

HOST = "127.0.0.1"
PORT = int(os.environ.get("ESG_DASHBOARD_PORT", "8765"))

# 실행 상태 (단일 실행만 허용)
STATE = {
    "running": False,
    "stage": None,      # connect | collect | analyze | report
    "done": 0,
    "total": 0,
    "detail": "",
    "result": None,     # pipeline.run() 반환값
}
STATE_LOCK = threading.Lock()

STAGE_LABEL = {
    "connect": "사내 LLM 연결 확인 중…",
    "collect": "Outlook 메일 수집 중…",
    "analyze": "LLM 분석 중…",
    "report": "리포트 생성 중…",
}


def _progress(stage, done=0, total=0, detail=""):
    with STATE_LOCK:
        STATE.update(stage=stage, done=done, total=total, detail=detail)


def _run_pipeline(params):
    try:
        result = pipeline.run(
            since=params.get("since") or None,
            until=params.get("until") or None,
            skip_collect=bool(params.get("skip_collect")),
            save_drafts=bool(params.get("save_drafts", True)),
            open_browser=False,          # 결과는 이 화면 안에서 보여준다
            progress_cb=_progress)
    except Exception as e:
        log.exception("파이프라인 실행 중 오류")
        result = {"ok": False, "message": f"실행 오류: {e}",
                  "report_path": None, "total": 0, "errors": 0}
    with STATE_LOCK:
        STATE.update(running=False, stage=None, result=result)


def _latest_report():
    files = sorted(glob.glob(os.path.join(reporter.OUTPUT_DIR, "report_*.html")))
    return files[-1] if files else None


CONTROL_PAGE = """<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<title>고객 ESG 대응 Agent</title>
<style>
  :root {
    --bg: #0f172a; --panel: #1e293b; --line: #334155;
    --text: #e2e8f0; --sub: #94a3b8; --accent: #34d399; --accent2: #22d3ee;
  }
  * { box-sizing: border-box; }
  body { margin: 0; font-family: 'Malgun Gothic', 'Segoe UI', sans-serif;
         background: var(--bg); color: var(--text); min-height: 100vh; }
  .hero { padding: 40px 20px 10px; text-align: center; }
  .hero h1 { margin: 0; font-size: 26px; letter-spacing: -0.5px;
    background: linear-gradient(90deg, var(--accent), var(--accent2));
    -webkit-background-clip: text; background-clip: text; color: transparent; }
  .hero p { color: var(--sub); font-size: 13px; }
  .panel { max-width: 560px; margin: 24px auto; background: var(--panel);
           border: 1px solid var(--line); border-radius: 16px; padding: 28px;
           box-shadow: 0 20px 50px rgba(0,0,0,.35); }
  label { display: block; font-size: 12px; color: var(--sub); margin: 14px 0 5px; }
  input[type=date] { width: 100%; padding: 10px 12px; border-radius: 10px;
    border: 1px solid var(--line); background: #0b1220; color: var(--text);
    font-size: 14px; color-scheme: dark; }
  .row { display: flex; gap: 14px; } .row > div { flex: 1; }
  .chk { display: flex; align-items: center; gap: 8px; margin-top: 14px;
         font-size: 13px; color: var(--sub); }
  button#go { width: 100%; margin-top: 22px; padding: 14px; border: 0;
    border-radius: 12px; font-size: 16px; font-weight: 700; cursor: pointer;
    background: linear-gradient(90deg, var(--accent), var(--accent2));
    color: #052e22; transition: transform .1s, opacity .2s; }
  button#go:hover { transform: translateY(-1px); }
  button#go:disabled { opacity: .45; cursor: not-allowed; transform: none; }
  .progress { display: none; margin-top: 22px; }
  .bar-wrap { height: 10px; background: #0b1220; border-radius: 6px;
              overflow: hidden; border: 1px solid var(--line); }
  .bar { height: 100%; width: 0%;
         background: linear-gradient(90deg, var(--accent), var(--accent2));
         transition: width .4s; }
  .stage { font-size: 13px; color: var(--sub); margin-top: 10px; min-height: 18px; }
  .result { display: none; margin-top: 20px; padding: 14px 16px;
            border-radius: 10px; font-size: 14px; }
  .result.ok { background: rgba(52,211,153,.12); border: 1px solid var(--accent); }
  .result.fail { background: rgba(248,113,113,.12); border: 1px solid #f87171; }
  .result a { color: var(--accent2); font-weight: 700; }
  .chat-link { display: block; text-align: center; margin-top: 18px;
    color: var(--accent2); font-size: 13px; font-weight: 700;
    text-decoration: none; }
  .chat-link:hover { text-decoration: underline; }
  .foot { text-align: center; color: #475569; font-size: 11px; margin: 30px 0; }
</style>
</head>
<body>
<div class="hero">
  <h1>고객 ESG 대응 Agent</h1>
  <p>분석할 메일의 수신 기간을 지정하고 [분석 시작]을 누르세요.<br>
     초안은 Outlook 임시보관함에 저장만 되며, 자동 발송되지 않습니다.</p>
</div>
<div class="panel">
  <div class="row">
    <div><label>수신 시작일</label><input type="date" id="since"></div>
    <div><label>수신 종료일</label><input type="date" id="until"></div>
  </div>
  <div class="chk"><input type="checkbox" id="drafts" checked>
    <span>Outlook 임시보관함에 초안 저장</span></div>
  <div class="chk"><input type="checkbox" id="skip">
    <span>수집 생략 (이전에 수집된 미처리 메일만 분석)</span></div>
  <button id="go" onclick="start()">분석 시작</button>
  <div class="progress" id="progress">
    <div class="bar-wrap"><div class="bar" id="bar"></div></div>
    <div class="stage" id="stage"></div>
  </div>
  <div class="result" id="result"></div>
  <a class="chat-link" href="/chat">💬 대응 현황 문답 (챗봇) →</a>
</div>
<div class="foot">사내 LLM 전용 · 원본 메일 미변경 · 자동 발송 없음</div>
<script>
const $ = id => document.getElementById(id);
const today = new Date(), week = new Date(today - 6*864e5);
$("until").value = today.toISOString().slice(0,10);
$("since").value = week.toISOString().slice(0,10);
let timer = null;

async function start() {
  $("go").disabled = true;
  $("result").style.display = "none";
  $("progress").style.display = "block";
  $("bar").style.width = "3%";
  $("stage").textContent = "시작 중…";
  await fetch("/run", { method: "POST",
    headers: {"Content-Type": "application/json"},
    body: JSON.stringify({ since: $("since").value, until: $("until").value,
      save_drafts: $("drafts").checked, skip_collect: $("skip").checked }) });
  timer = setInterval(poll, 1000);
}

async function poll() {
  const s = await (await fetch("/status")).json();
  if (s.running) {
    let pct = 5, label = s.stage_label || "";
    if (s.stage === "analyze" && s.total > 0) {
      pct = 10 + 80 * s.done / s.total;
      label = `LLM 분석 중… (${s.done}/${s.total}) ${s.detail || ""}`;
    } else if (s.stage === "report") { pct = 95; }
    else if (s.stage === "collect") { pct = 8; }
    $("bar").style.width = pct + "%";
    $("stage").textContent = label;
    return;
  }
  clearInterval(timer);
  $("go").disabled = false;
  $("bar").style.width = "100%";
  $("stage").textContent = "";
  const r = s.result || {};
  const el = $("result");
  el.className = "result " + (r.ok ? "ok" : "fail");
  el.style.display = "block";
  if (r.ok && r.report_path) {
    el.innerHTML = r.message +
      ' &nbsp;→&nbsp; <a href="/report" target="_blank">검토 대시보드 열기</a>';
    window.open("/report", "_blank");
  } else {
    el.textContent = r.message || "결과 없음";
  }
}
</script>
</body>
</html>"""


CHAT_PAGE = """<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<title>ESG 대응 현황 문답</title>
<style>
  :root {
    --bg: #0f172a; --panel: #1e293b; --line: #334155;
    --text: #e2e8f0; --sub: #94a3b8; --accent: #34d399; --accent2: #22d3ee;
  }
  * { box-sizing: border-box; }
  html, body { height: 100%; }
  body { margin: 0; font-family: 'Malgun Gothic', 'Segoe UI', sans-serif;
         background: var(--bg); color: var(--text);
         display: flex; flex-direction: column; }
  header { padding: 14px 22px; border-bottom: 1px solid var(--line);
           display: flex; align-items: center; gap: 14px; }
  header h1 { margin: 0; font-size: 16px;
    background: linear-gradient(90deg, var(--accent), var(--accent2));
    -webkit-background-clip: text; background-clip: text; color: transparent; }
  header a { color: var(--sub); font-size: 12px; text-decoration: none;
            margin-left: auto; }
  header a:hover { color: var(--text); }
  #log { flex: 1; overflow-y: auto; padding: 22px;
         max-width: 860px; width: 100%; margin: 0 auto; }
  .msg { display: flex; margin-bottom: 14px; }
  .msg.user { justify-content: flex-end; }
  .bubble { max-width: 78%; padding: 11px 16px; border-radius: 14px;
           font-size: 14px; line-height: 1.7; white-space: pre-wrap;
           word-break: break-word; }
  .user .bubble { background: linear-gradient(120deg, #065f46, #0e7490);
                 border-bottom-right-radius: 4px; }
  .bot .bubble { background: var(--panel); border: 1px solid var(--line);
                border-bottom-left-radius: 4px; }
  .bot .bubble.err { border-color: #f87171; color: #fca5a5; }
  .hint { color: var(--sub); font-size: 12px; text-align: center;
         margin: 8px 0 16px; }
  .chips { display: flex; gap: 8px; flex-wrap: wrap; justify-content: center;
          margin-bottom: 10px; }
  .chip { border: 1px solid var(--line); background: var(--panel);
         color: var(--sub); border-radius: 999px; padding: 7px 14px;
         font-size: 12px; cursor: pointer; font-family: inherit; }
  .chip:hover { color: var(--text); border-color: var(--accent); }
  form { display: flex; gap: 10px; padding: 16px 22px 22px;
        max-width: 860px; width: 100%; margin: 0 auto; }
  input { flex: 1; padding: 13px 16px; border-radius: 12px; font-size: 14px;
         border: 1px solid var(--line); background: #0b1220;
         color: var(--text); font-family: inherit; }
  input:focus { outline: none; border-color: var(--accent); }
  button { padding: 13px 24px; border: 0; border-radius: 12px; font-size: 14px;
          font-weight: 700; cursor: pointer; font-family: inherit;
          background: linear-gradient(90deg, var(--accent), var(--accent2));
          color: #052e22; }
  button:disabled { opacity: .45; cursor: not-allowed; }
  .typing { color: var(--sub); font-size: 13px; }
</style>
</head>
<body>
<header>
  <h1>💬 ESG 대응 현황 문답</h1>
  <a href="/">← 분석 실행 화면</a>
</header>
<div id="log">
  <div class="hint">분석 이력·관리대장(수기 처리상태 포함)·공개 Factsheet를
    근거로 답합니다. 데이터에 없는 내용은 답하지 않습니다.<br>
    조회 전용 도우미이며 메일 발송 등 어떤 동작도 수행하지 않습니다.</div>
  <div class="chips">
    <button class="chip" onclick="ask(this.textContent)">이번 주 마감 임박 요청은?</button>
    <button class="chip" onclick="ask(this.textContent)">아직 처리 안 된 요청 정리해줘</button>
    <button class="chip" onclick="ask(this.textContent)">고위험 리스크가 있는 건은 뭐야?</button>
    <button class="chip" onclick="ask(this.textContent)">Scope 1 배출량 요청 이력 보여줘</button>
  </div>
</div>
<form onsubmit="return submitMsg(event)">
  <input id="q" placeholder="예: OO 고객이 요청한 항목들 진행상황 알려줘"
         autocomplete="off" autofocus>
  <button id="send" type="submit">전송</button>
</form>
<script>
var chatHistory = [];

function add(role, text, err) {
  var log = document.getElementById("log");
  var div = document.createElement("div");
  div.className = "msg " + role;
  var b = document.createElement("div");
  b.className = "bubble" + (err ? " err" : "");
  b.textContent = text;
  div.appendChild(b);
  log.appendChild(div);
  log.scrollTop = log.scrollHeight;
  return b;
}

function ask(text) { document.getElementById("q").value = text; sendMsg(); }
function submitMsg(ev) { ev.preventDefault(); sendMsg(); return false; }

async function sendMsg() {
  var input = document.getElementById("q");
  var text = input.value.trim();
  if (!text) return;
  input.value = "";
  document.getElementById("send").disabled = true;
  add("user", text);
  var wait = add("bot", "답변 작성 중…");
  wait.classList.add("typing");
  try {
    var res = await fetch("/chat/api", { method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify({ message: text, history: chatHistory }) });
    var data = await res.json();
    wait.classList.remove("typing");
    if (data.ok) {
      wait.textContent = data.reply;
      chatHistory.push({ role: "user", content: text });
      chatHistory.push({ role: "assistant", content: data.reply });
      if (chatHistory.length > 20) chatHistory = chatHistory.slice(-20);
    } else {
      wait.classList.add("err");
      wait.textContent = data.error || "오류가 발생했습니다.";
    }
  } catch (e) {
    wait.classList.remove("typing");
    wait.classList.add("err");
    wait.textContent = "서버 연결 오류: " + e;
  }
  document.getElementById("send").disabled = false;
  input.focus();
}
</script>
</body>
</html>"""


class Handler(http.server.BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        log.debug(fmt, *args)

    def _send(self, code, body, ctype="text/html; charset=utf-8"):
        data = body.encode("utf-8") if isinstance(body, str) else body
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        if self.path == "/" or self.path.startswith("/index"):
            self._send(200, CONTROL_PAGE)
        elif self.path == "/chat":
            self._send(200, CHAT_PAGE)
        elif self.path == "/status":
            with STATE_LOCK:
                snap = dict(STATE)
            snap["stage_label"] = STAGE_LABEL.get(snap.get("stage"), "")
            self._send(200, json.dumps(snap, ensure_ascii=False),
                       "application/json; charset=utf-8")
        elif self.path == "/report":
            with STATE_LOCK:
                result = STATE.get("result") or {}
            path = result.get("report_path") or _latest_report()
            if path and os.path.exists(path):
                with open(path, "rb") as f:
                    self._send(200, f.read())
            else:
                self._send(404, "아직 생성된 리포트가 없습니다.")
        else:
            self._send(404, "not found")

    def do_POST(self):
        if self.path == "/chat/api":
            self._handle_chat()
            return
        if self.path != "/run":
            self._send(404, "not found")
            return
        length = int(self.headers.get("Content-Length", 0))
        try:
            params = json.loads(self.rfile.read(length).decode("utf-8") or "{}")
        except Exception:
            params = {}
        with STATE_LOCK:
            if STATE["running"]:
                self._send(409, json.dumps({"ok": False,
                           "message": "이미 실행 중입니다."}),
                           "application/json; charset=utf-8")
                return
            STATE.update(running=True, stage="connect", done=0, total=0,
                         detail="", result=None)
        threading.Thread(target=_run_pipeline, args=(params,),
                         daemon=True).start()
        self._send(200, json.dumps({"ok": True}),
                   "application/json; charset=utf-8")

    def _handle_chat(self):
        length = int(self.headers.get("Content-Length", 0))
        try:
            body = json.loads(self.rfile.read(length).decode("utf-8") or "{}")
            message = str(body.get("message", "")).strip()
            history = body.get("history") or []
            if not message:
                raise ValueError("질문이 비어 있습니다.")
            config = pipeline.load_config()
            reply = chatbot.chat(message, history=history, config=config)
            payload = {"ok": True, "reply": reply}
        except Exception as e:
            log.warning("챗봇 응답 실패: %s", e)
            payload = {"ok": False,
                       "error": f"답변 생성에 실패했습니다: {e}"}
        self._send(200, json.dumps(payload, ensure_ascii=False),
                   "application/json; charset=utf-8")


def main():
    os.makedirs(os.path.join(BASE_DIR, "logs"), exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        handlers=[logging.FileHandler(
                      os.path.join(BASE_DIR, "logs",
                                   f"{datetime.date.today():%Y%m%d}.log"),
                      encoding="utf-8"),
                  logging.StreamHandler(sys.stdout)])
    server = http.server.ThreadingHTTPServer((HOST, PORT), Handler)
    url = f"http://{HOST}:{PORT}"
    log.info("웹 대시보드 시작: %s (종료: Ctrl+C)", url)
    try:
        webbrowser.open(url)
    except Exception:
        pass
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        log.info("종료합니다.")


if __name__ == "__main__":
    main()
