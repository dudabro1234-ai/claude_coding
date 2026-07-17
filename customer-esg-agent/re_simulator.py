# -*- coding: utf-8 -*-
"""고객별 RE(재생에너지) 할당 시뮬레이션 — 목업 (상세 요구사항 확정 전).

배경: 고객 요청의 상당수가 "언제까지 RE를 할당해 달라"는 내용이라,
보유 RE 계약량을 입력하고 주요 고객(MS/AWS/Apple/Google/Meta/NVIDIA)의
요청량에 맞춰 할당을 시뮬레이션하는 화면을 대응 Agent 안에 메뉴로 둔다.

구현 형태 (사내 튜닝본에 이식하기 쉽도록):
- 이 파일 하나가 전부다. 화면은 100% 클라이언트(JS) 계산 — 서버·LLM·외부
  라이브러리 의존 없음. 입력값은 브라우저 localStorage에만 저장된다.
- 실행 방법 ①: python re_simulator.py  → output/re_simulator.html 생성·오픈
- 실행 방법 ②: dashboard_server에 /re-sim 라우트로 연결 (2줄 추가)

숫자 정확성: 시뮬레이션 결과는 입력값 기반 산술 계산일 뿐, 계약 이행·인증
가능성을 보증하지 않는다. 화면에 목업·검토용 배너를 상시 표시한다.
"""
import os
import sys
import webbrowser

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

PAGE = """<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<title>RE 할당 시뮬레이션 (목업)</title>
<style>
  :root {
    --bg: #f1f5f9; --card: #ffffff; --line: #e2e8f0;
    --text: #0f172a; --sub: #64748b;
    --accent: #10b981; --accent2: #0ea5e9;
    /* 차트 전용 (검증된 팔레트: PPA/REC/녹색프리미엄) */
    --s1: #2a78d6; --s2: #008300; --s3: #e87ba4;
    --ink: #0b0b0b; --ink2: #52514e; --muted: #898781;
    --grid: #e1e0d9; --baseline: #c3c2b7; --critical: #d03b3b;
  }
  * { box-sizing: border-box; }
  html { background: var(--bg); }
  body { margin: 0; font-family: 'Malgun Gothic', 'Segoe UI', system-ui,
         sans-serif; background: var(--bg); color: var(--text);
         min-height: 100vh; }
  header { background: linear-gradient(120deg, #0f172a, #1e3a5f 60%, #134e4a);
           color: #e2e8f0; padding: 22px 32px; }
  header h1 { margin: 0; font-size: 20px; letter-spacing: -0.3px; }
  header h1 small { font-weight: 400; color: #94a3b8; margin-left: 8px; }
  .mock-badge { display: inline-block; margin-left: 10px; font-size: 11px;
    font-weight: 800; color: #0f172a; background: #fbbf24;
    border-radius: 6px; padding: 2px 9px; vertical-align: middle; }
  header .warn { display: inline-block; margin-top: 9px; font-size: 12px;
    color: #fbbf24; background: rgba(251,191,36,.12);
    border: 1px solid rgba(251,191,36,.4); border-radius: 8px;
    padding: 5px 12px; }
  main { max-width: 1180px; margin: 0 auto; padding: 22px 32px 70px; }
  h2 { font-size: 15px; margin: 26px 0 10px; }
  h2 .hint { font-size: 12px; color: var(--sub); font-weight: 400;
             margin-left: 8px; }
  .panel { background: var(--card); border: 1px solid var(--line);
           border-radius: 16px; padding: 18px 22px;
           box-shadow: 0 2px 8px rgba(15,23,42,.06); }
  .grid2 { display: grid; grid-template-columns: 1.2fr 1fr; gap: 16px; }
  @media (max-width: 980px) { .grid2 { grid-template-columns: 1fr; } }
  table.edit { width: 100%; border-collapse: collapse; font-size: 13px; }
  table.edit th { background: #f8fafc; color: var(--sub); font-size: 11px;
    text-transform: uppercase; letter-spacing: .05em; text-align: left;
    padding: 8px 10px; border-bottom: 2px solid var(--line); }
  table.edit td { padding: 6px 8px; border-bottom: 1px solid var(--line); }
  table.edit input, table.edit select { width: 100%; padding: 7px 9px;
    border: 1px solid var(--line); border-radius: 8px; font-size: 13px;
    font-family: inherit; background: #fff; color: var(--text); }
  table.edit input[type=number] { text-align: right; }
  .row-del { border: 0; background: none; color: var(--sub); cursor: pointer;
             font-size: 15px; }
  .row-del:hover { color: var(--critical); }
  .add-btn { margin-top: 10px; border: 1px dashed var(--line);
    background: #f8fafc; border-radius: 8px; padding: 7px 14px;
    font-size: 12px; cursor: pointer; font-family: inherit; color: var(--sub); }
  .add-btn:hover { color: var(--text); border-color: var(--accent); }

  .controls { display: flex; gap: 12px; flex-wrap: wrap; align-items: flex-end;
              margin: 20px 0; }
  .ctl { display: flex; flex-direction: column; gap: 5px; }
  .ctl label { font-size: 11px; font-weight: 700; color: var(--sub);
               text-transform: uppercase; letter-spacing: .05em; }
  .ctl select { padding: 10px 12px; border: 1px solid var(--line);
    border-radius: 10px; font-size: 13px; font-family: inherit;
    background: #fff; }
  #run { padding: 12px 28px; border: 0; border-radius: 12px; font-size: 15px;
    font-weight: 800; cursor: pointer; font-family: inherit; color: #052e22;
    background: linear-gradient(90deg, var(--accent), var(--accent2)); }
  #run:hover { transform: translateY(-1px); }
  #reset { padding: 12px 16px; border: 1px solid var(--line);
    border-radius: 12px; font-size: 13px; cursor: pointer;
    font-family: inherit; background: #fff; color: var(--sub); }

  .kpi-bar { display: grid;
    grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
    gap: 12px; margin: 6px 0 18px; }
  .kpi { background: var(--card); border: 1px solid var(--line);
         border-radius: 14px; padding: 13px 18px; position: relative;
         overflow: hidden; }
  .kpi::before { content: ""; position: absolute; left: 0; top: 0; bottom: 0;
                 width: 4px; background: var(--kc, var(--accent2)); }
  .kpi-num { font-size: 25px; font-weight: 800; color: var(--kc, var(--accent2)); }
  .kpi-num small { font-size: 12px; font-weight: 600; color: var(--sub); }
  .kpi-label { font-size: 12px; color: var(--sub); margin-top: 2px; }

  /* 고객별 할당 bullet bars */
  .alloc-row { display: grid; grid-template-columns: 92px 1fr 250px;
    gap: 12px; align-items: center; padding: 7px 0; }
  .alloc-name { font-size: 13px; font-weight: 700; text-align: right; }
  .alloc-track { position: relative; height: 18px; background: #eef1f5;
    border-radius: 4px; overflow: hidden; }
  .alloc-fill { position: absolute; left: 0; top: 0; bottom: 0;
    background: var(--s1); border-radius: 4px 0 0 4px; }
  .alloc-fill.full { border-radius: 4px; }
  .alloc-val { font-size: 12px; color: var(--ink2);
    font-variant-numeric: tabular-nums; }
  .short { color: var(--critical); font-weight: 700; }
  .ok-txt { color: #006300; font-weight: 700; }

  /* 연도별 차트 */
  .chart-wrap { overflow-x: auto; }
  svg text { font-family: inherit; }
  .legend { display: flex; gap: 16px; flex-wrap: wrap; font-size: 12px;
            color: var(--ink2); margin: 4px 0 10px; }
  .legend .sw { display: inline-block; width: 11px; height: 11px;
    border-radius: 3px; margin-right: 5px; vertical-align: -1px; }
  .legend .ln { display: inline-block; width: 16px; height: 0;
    border-top: 2px dashed var(--ink2); margin-right: 5px; vertical-align: 3px; }
  #tooltip { position: fixed; display: none; background: #0f172a; color: #e2e8f0;
    font-size: 12px; padding: 6px 10px; border-radius: 8px; pointer-events: none;
    z-index: 99; white-space: pre; }
  table.result { width: 100%; border-collapse: collapse; font-size: 12px;
                 margin-top: 12px; }
  table.result th, table.result td { border: 1px solid var(--line);
    padding: 6px 9px; text-align: right;
    font-variant-numeric: tabular-nums; }
  table.result th { background: #f8fafc; color: var(--sub); }
  table.result td:first-child, table.result th:first-child { text-align: left; }
  .foot-note { font-size: 11px; color: var(--sub); margin-top: 14px;
               line-height: 1.7; }
  a.back { color: #7dd3fc; font-size: 12px; text-decoration: none;
           float: right; margin-top: 6px; }
</style>
</head>
<body>
<header>
  <a class="back" href="/" onclick="if(location.protocol==='file:'){alert('run_dashboard.bat 실행 시 메뉴로 연결됩니다.');return false;}">← 제어 화면</a>
  <h1>고객별 RE 할당 시뮬레이션 <small>Renewable Energy Allocation</small>
      <span class="mock-badge">목업 v0</span></h1>
  <div class="warn">⚠️ 검토용 산술 시뮬레이션입니다. 계약 이행·인증 가능성을 보증하지
      않으며, 결과를 고객에게 직접 회신하기 전 반드시 담당 조직 검토가 필요합니다.</div>
</header>
<main>
  <div class="grid2">
    <section>
      <h2>① 보유 RE 계약 <span class="hint">현재 확보한 재생에너지 계약을
          입력하세요 (연간 공급량 기준)</span></h2>
      <div class="panel">
        <table class="edit" id="contracts">
          <thead><tr><th>계약명</th><th style="width:130px">유형</th>
            <th style="width:110px">연간 공급량 (GWh)</th>
            <th style="width:80px">시작연도</th><th style="width:80px">종료연도</th>
            <th style="width:30px"></th></tr></thead>
          <tbody></tbody>
        </table>
        <button class="add-btn" onclick="addContract()">＋ 계약 추가</button>
      </div>
    </section>
    <section>
      <h2>② 고객 RE 요청량 <span class="hint">고객별 연간 요청량과 우선순위</span></h2>
      <div class="panel">
        <table class="edit" id="customers">
          <thead><tr><th>고객</th><th style="width:110px">요청량 (GWh/년)</th>
            <th style="width:90px">목표연도</th><th style="width:80px">우선순위</th>
            <th style="width:30px"></th></tr></thead>
          <tbody></tbody>
        </table>
        <button class="add-btn" onclick="addCustomer()">＋ 고객 추가</button>
      </div>
    </section>
  </div>

  <div class="controls">
    <div class="ctl"><label>기준연도</label>
      <select id="base-year"></select></div>
    <div class="ctl"><label>할당 방식</label>
      <select id="method">
        <option value="priority">우선순위 순차 할당</option>
        <option value="prorata">요청량 비례 배분</option>
      </select></div>
    <button id="run" onclick="simulate()">시뮬레이션 실행</button>
    <button id="reset" onclick="resetAll()">초기값으로</button>
  </div>

  <div id="results" style="display:none">
    <h2>③ 시뮬레이션 결과 <span class="hint" id="result-cond"></span></h2>
    <div class="kpi-bar" id="kpis"></div>

    <div class="grid2">
      <section class="panel">
        <h2 style="margin-top:0">고객별 할당 현황 <span class="hint">막대 = 요청량
            대비 할당 비율</span></h2>
        <div id="alloc-bars"></div>
      </section>
      <section class="panel">
        <h2 style="margin-top:0">연도별 공급 능력 vs 총 요청량</h2>
        <div class="legend">
          <span><span class="sw" style="background:var(--s1)"></span>PPA</span>
          <span><span class="sw" style="background:var(--s2)"></span>REC</span>
          <span><span class="sw" style="background:var(--s3)"></span>녹색프리미엄</span>
          <span><span class="ln"></span>총 요청량</span>
        </div>
        <div class="chart-wrap"><svg id="yearly" width="520" height="240"></svg></div>
      </section>
    </div>

    <section class="panel" style="margin-top:16px">
      <h2 style="margin-top:0">연도별 상세 (GWh)</h2>
      <table class="result" id="yearly-table"></table>
    </section>

    <div class="foot-note">
      · 공급 능력 = 해당 연도에 유효한(시작연도 ≤ 연도 ≤ 종료연도) 계약의 연간
      공급량 합계 · 우선순위 순차: 1순위부터 요청량을 채우고 남는 물량을 다음
      순위로 · 요청량 비례: 공급량을 요청량 비율대로 배분(요청량 초과분은 잔여
      고객에 재배분) · 입력값은 이 PC 브라우저에만 저장됩니다(localStorage).
    </div>
  </div>
</main>
<div id="tooltip"></div>

<script>
"use strict";
var TYPES = ["PPA", "REC", "녹색프리미엄"];
var TYPE_COLOR = { "PPA": "#2a78d6", "REC": "#008300", "녹색프리미엄": "#e87ba4" };
var YEARS_SPAN = 10;

var DEFAULTS = {
  contracts: [
    { name: "해상풍력 PPA-1", type: "PPA", gwh: 400, start: 2026, end: 2045 },
    { name: "태양광 PPA-2",   type: "PPA", gwh: 250, start: 2027, end: 2046 },
    { name: "REC 구매(연간)", type: "REC", gwh: 300, start: 2026, end: 2030 },
    { name: "녹색프리미엄",   type: "녹색프리미엄", gwh: 150, start: 2026, end: 2028 }
  ],
  customers: [
    { name: "MS",     demand: 220, target: 2027, prio: 1 },
    { name: "AWS",    demand: 180, target: 2028, prio: 2 },
    { name: "Apple",  demand: 260, target: 2027, prio: 3 },
    { name: "Google", demand: 150, target: 2029, prio: 4 },
    { name: "Meta",   demand: 120, target: 2028, prio: 5 },
    { name: "NVIDIA", demand: 100, target: 2030, prio: 6 }
  ]
};

function load(key) {
  try {
    var v = JSON.parse(localStorage.getItem("resim-" + key));
    if (Array.isArray(v) && v.length) return v;
  } catch (e) {}
  return JSON.parse(JSON.stringify(DEFAULTS[key]));
}
function save() {
  localStorage.setItem("resim-contracts", JSON.stringify(readContracts()));
  localStorage.setItem("resim-customers", JSON.stringify(readCustomers()));
}
function resetAll() {
  localStorage.removeItem("resim-contracts");
  localStorage.removeItem("resim-customers");
  renderInputs(DEFAULTS.contracts, DEFAULTS.customers);
  document.getElementById("results").style.display = "none";
}

function contractRow(c) {
  var opts = TYPES.map(function(t) {
    return "<option" + (t === c.type ? " selected" : "") + ">" + t + "</option>";
  }).join("");
  return "<tr>" +
    "<td><input class='c-name' value='" + c.name + "'></td>" +
    "<td><select class='c-type'>" + opts + "</select></td>" +
    "<td><input class='c-gwh' type='number' min='0' value='" + c.gwh + "'></td>" +
    "<td><input class='c-start' type='number' value='" + c.start + "'></td>" +
    "<td><input class='c-end' type='number' value='" + c.end + "'></td>" +
    "<td><button class='row-del' onclick='delRow(this)'>✕</button></td></tr>";
}
function customerRow(c) {
  return "<tr>" +
    "<td><input class='u-name' value='" + c.name + "'></td>" +
    "<td><input class='u-demand' type='number' min='0' value='" + c.demand + "'></td>" +
    "<td><input class='u-target' type='number' value='" + c.target + "'></td>" +
    "<td><input class='u-prio' type='number' min='1' value='" + c.prio + "'></td>" +
    "<td><button class='row-del' onclick='delRow(this)'>✕</button></td></tr>";
}
function renderInputs(contracts, customers) {
  document.querySelector("#contracts tbody").innerHTML =
      contracts.map(contractRow).join("");
  document.querySelector("#customers tbody").innerHTML =
      customers.map(customerRow).join("");
}
function delRow(btn) { btn.closest("tr").remove(); save(); }
function addContract() {
  var y = new Date().getFullYear();
  document.querySelector("#contracts tbody").insertAdjacentHTML("beforeend",
    contractRow({ name: "신규 계약", type: "PPA", gwh: 0, start: y, end: y + 9 }));
}
function addCustomer() {
  document.querySelector("#customers tbody").insertAdjacentHTML("beforeend",
    customerRow({ name: "신규 고객", demand: 0,
                  target: new Date().getFullYear() + 2, prio: 9 }));
}
function num(el) { var v = parseFloat(el.value); return isNaN(v) ? 0 : v; }
function readContracts() {
  return Array.prototype.map.call(
    document.querySelectorAll("#contracts tbody tr"), function(tr) {
      return { name: tr.querySelector(".c-name").value,
               type: tr.querySelector(".c-type").value,
               gwh: num(tr.querySelector(".c-gwh")),
               start: num(tr.querySelector(".c-start")),
               end: num(tr.querySelector(".c-end")) };
    });
}
function readCustomers() {
  return Array.prototype.map.call(
    document.querySelectorAll("#customers tbody tr"), function(tr) {
      return { name: tr.querySelector(".u-name").value,
               demand: num(tr.querySelector(".u-demand")),
               target: num(tr.querySelector(".u-target")),
               prio: num(tr.querySelector(".u-prio")) };
    });
}

function supplyByType(contracts, year) {
  var out = { "PPA": 0, "REC": 0, "녹색프리미엄": 0 };
  contracts.forEach(function(c) {
    if (c.start <= year && year <= c.end) out[c.type] += c.gwh;
  });
  return out;
}
function totalSupply(contracts, year) {
  var s = supplyByType(contracts, year);
  return s["PPA"] + s["REC"] + s["녹색프리미엄"];
}

function allocate(customers, supply, method) {
  var alloc = {};
  customers.forEach(function(c) { alloc[c.name] = 0; });
  if (method === "priority") {
    var remaining = supply;
    customers.slice().sort(function(a, b) { return a.prio - b.prio; })
      .forEach(function(c) {
        var a = Math.min(c.demand, remaining);
        alloc[c.name] = a;
        remaining -= a;
      });
  } else {
    var pool = supply;
    var open = customers.slice();
    // 비례 배분 + 초과분 재배분 (수렴까지 반복)
    for (var guard = 0; guard < 10 && pool > 1e-9 && open.length; guard++) {
      var tot = open.reduce(function(s, c) {
        return s + (c.demand - alloc[c.name]); }, 0);
      if (tot <= 1e-9) break;
      var used = 0;
      open.forEach(function(c) {
        var need = c.demand - alloc[c.name];
        var give = Math.min(need, pool * need / tot);
        alloc[c.name] += give;
        used += give;
      });
      pool -= used;
      open = open.filter(function(c) { return c.demand - alloc[c.name] > 1e-9; });
    }
  }
  return alloc;
}

function fmt(v) { return (Math.round(v * 10) / 10).toLocaleString("ko-KR"); }

function simulate() {
  save();
  var contracts = readContracts();
  var customers = readCustomers().filter(function(c) { return c.demand > 0; });
  var year = parseInt(document.getElementById("base-year").value, 10);
  var method = document.getElementById("method").value;
  var supply = totalSupply(contracts, year);
  var demand = customers.reduce(function(s, c) { return s + c.demand; }, 0);
  var alloc = allocate(customers, supply, method);
  var allocated = customers.reduce(function(s, c) { return s + alloc[c.name]; }, 0);
  var shortfall = Math.max(0, demand - allocated);

  document.getElementById("results").style.display = "block";
  document.getElementById("result-cond").textContent =
    year + "년 기준 · " + (method === "priority" ? "우선순위 순차" : "요청량 비례");

  var kpis = [
    ["공급 능력", fmt(supply), "#0ea5e9"],
    ["총 요청량", fmt(demand), "#64748b"],
    ["할당 완료", fmt(allocated), "#10b981"],
    ["미충족", fmt(shortfall), shortfall > 0 ? "#d03b3b" : "#10b981"]
  ];
  document.getElementById("kpis").innerHTML = kpis.map(function(k) {
    return "<div class='kpi' style='--kc:" + k[2] + "'><div class='kpi-num'>" +
      k[1] + " <small>GWh</small></div><div class='kpi-label'>" + k[0] +
      "</div></div>";
  }).join("");

  // 고객별 bullet bars
  var maxD = Math.max.apply(null, customers.map(function(c) { return c.demand; }));
  document.getElementById("alloc-bars").innerHTML =
    customers.slice().sort(function(a, b) { return a.prio - b.prio; })
    .map(function(c) {
      var a = alloc[c.name];
      var pct = c.demand ? a / c.demand * 100 : 0;
      var trackW = c.demand / maxD * 100;
      var lack = c.demand - a;
      var status = lack > 0.05
        ? "<span class='short'>▲ 부족 " + fmt(lack) + "</span>"
        : "<span class='ok-txt'>✓ 충족</span> <span>(목표 " + c.target + ")</span>";
      return "<div class='alloc-row'>" +
        "<div class='alloc-name'>" + c.name + "</div>" +
        "<div><div class='alloc-track' style='width:" + trackW + "%'>" +
        "<div class='alloc-fill" + (pct >= 99.95 ? " full" : "") +
        "' style='width:" + pct + "%'></div></div></div>" +
        "<div class='alloc-val'>" + fmt(a) + " / " + fmt(c.demand) +
        " GWh (" + Math.round(pct) + "%) " + status + "</div></div>";
    }).join("");

  drawYearly(contracts, demand, year);
  drawTable(contracts, customers, method, year);
}

function drawYearly(contracts, demand, baseYear) {
  var y0 = parseInt(document.getElementById("base-year").options[0].value, 10);
  var years = [];
  for (var i = 0; i < YEARS_SPAN; i++) years.push(y0 + i);
  var W = 520, H = 240, padL = 46, padB = 26, padT = 12, padR = 8;
  var plotW = W - padL - padR, plotH = H - padT - padB;
  var maxV = demand;
  years.forEach(function(y) { maxV = Math.max(maxV, totalSupply(contracts, y)); });
  maxV = maxV * 1.15 || 1;
  var bw = plotW / years.length;
  var svg = "";
  // 가로 그리드 + y라벨
  for (var g = 0; g <= 4; g++) {
    var gv = maxV / 4 * g;
    var gy = padT + plotH - plotH * g / 4;
    svg += "<line x1='" + padL + "' y1='" + gy + "' x2='" + (W - padR) +
      "' y2='" + gy + "' stroke='#e1e0d9' stroke-width='1'/>" +
      "<text x='" + (padL - 6) + "' y='" + (gy + 4) +
      "' font-size='10' fill='#898781' text-anchor='end'>" +
      Math.round(gv) + "</text>";
  }
  // 스택 바 (2px 갭)
  years.forEach(function(y, i) {
    var s = supplyByType(contracts, y);
    var x = padL + i * bw + bw * 0.18, w = bw * 0.64, cy = padT + plotH;
    TYPES.forEach(function(t) {
      var h = s[t] / maxV * plotH;
      if (h <= 0) return;
      cy -= h;
      var hh = Math.max(0, h - 2);  // 세그먼트 사이 2px 표면 갭
      svg += "<rect class='seg' x='" + x + "' y='" + (cy + 1) + "' width='" + w +
        "' height='" + hh + "' rx='2' fill='" + TYPE_COLOR[t] +
        "' data-tip='" + y + "년 " + t + ": " + fmt(s[t]) + " GWh'/>";
    });
    var hl = (y === baseYear);
    svg += "<text x='" + (x + w / 2) + "' y='" + (H - 8) +
      "' font-size='10' text-anchor='middle' fill='" +
      (hl ? "#0b0b0b" : "#898781") + "'" +
      (hl ? " font-weight='700'" : "") + ">" + y + "</text>";
  });
  // 총 요청량 기준선 (참조선 — 잉크 대시)
  var dy = padT + plotH - demand / maxV * plotH;
  svg += "<line x1='" + padL + "' y1='" + dy + "' x2='" + (W - padR) +
    "' y2='" + dy + "' stroke='#52514e' stroke-width='2' " +
    "stroke-dasharray='6 4'/>" +
    "<text x='" + (W - padR) + "' y='" + (dy - 5) +
    "' font-size='10' fill='#52514e' text-anchor='end'>총 요청량 " +
    fmt(demand) + "</text>";
  // 기준축
  svg += "<line x1='" + padL + "' y1='" + (padT + plotH) + "' x2='" +
    (W - padR) + "' y2='" + (padT + plotH) +
    "' stroke='#c3c2b7' stroke-width='1'/>";
  var el = document.getElementById("yearly");
  el.innerHTML = svg;
  // 호버 툴팁
  var tip = document.getElementById("tooltip");
  el.querySelectorAll(".seg").forEach(function(r) {
    r.addEventListener("mousemove", function(ev) {
      tip.style.display = "block";
      tip.textContent = r.getAttribute("data-tip");
      tip.style.left = (ev.clientX + 12) + "px";
      tip.style.top = (ev.clientY - 10) + "px";
    });
    r.addEventListener("mouseleave", function() {
      tip.style.display = "none"; });
  });
}

function drawTable(contracts, customers, method, baseYear) {
  var y0 = parseInt(document.getElementById("base-year").options[0].value, 10);
  var head = "<tr><th>구분</th>";
  var rows = { "PPA": "<td>PPA 공급</td>", "REC": "<td>REC 공급</td>",
               "녹색프리미엄": "<td>녹색프리미엄 공급</td>" };
  var totalRow = "<td><b>공급 합계</b></td>", allocRow = "<td>할당 합계</td>",
      shortRow = "<td>미충족</td>";
  for (var i = 0; i < YEARS_SPAN; i++) {
    var y = y0 + i;
    head += "<th" + (y === baseYear ? " style='background:#e0f2fe'" : "") +
            ">" + y + "</th>";
    var s = supplyByType(contracts, y);
    TYPES.forEach(function(t) { rows[t] += "<td>" + fmt(s[t]) + "</td>"; });
    var sup = s["PPA"] + s["REC"] + s["녹색프리미엄"];
    var alloc = allocate(customers, sup, method);
    var allocated = customers.reduce(function(sm, c) {
      return sm + alloc[c.name]; }, 0);
    var demand = customers.reduce(function(sm, c) { return sm + c.demand; }, 0);
    totalRow += "<td><b>" + fmt(sup) + "</b></td>";
    allocRow += "<td>" + fmt(allocated) + "</td>";
    var lack = demand - allocated;
    shortRow += "<td" + (lack > 0.05 ? " style='color:#d03b3b;font-weight:700'"
                                     : "") + ">" + fmt(Math.max(0, lack)) +
                "</td>";
  }
  document.getElementById("yearly-table").innerHTML =
    "<thead>" + head + "</tr></thead><tbody><tr>" + rows["PPA"] +
    "</tr><tr>" + rows["REC"] + "</tr><tr>" + rows["녹색프리미엄"] +
    "</tr><tr>" + totalRow + "</tr><tr>" + allocRow + "</tr><tr>" +
    shortRow + "</tr></tbody>";
}

// 초기화
(function init() {
  var y = new Date().getFullYear();
  var sel = document.getElementById("base-year");
  for (var i = 0; i < YEARS_SPAN; i++) {
    var o = document.createElement("option");
    o.value = y + i; o.textContent = (y + i) + "년";
    sel.appendChild(o);
  }
  renderInputs(load("contracts"), load("customers"));
  document.body.addEventListener("change", function(ev) {
    if (ev.target.closest("table.edit")) save();
  });
})();
</script>
</body>
</html>"""


def write_page(out_dir=OUTPUT_DIR):
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, "re_simulator.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(PAGE)
    return path


if __name__ == "__main__":
    p = write_page()
    print("RE 할당 시뮬레이터(목업) 생성:", p)
    if "--no-browser" not in sys.argv:
        webbrowser.open("file://" + os.path.abspath(p))
