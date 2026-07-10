# 고객 ESG 대응 Agent

고객사로부터 수신되는 ESG 요구사항 메일을 자동 수집·구조화하고,
사내 Factsheet 기반 **1차 답변 초안**과 **타부서 데이터 요청 초안**을 생성하는 도구입니다.

> ⚠️ **이 도구는 어떤 메일도 자동으로 발송하지 않습니다.**
> 모든 초안은 Outlook 임시보관함에 저장만 되며, 최종 검토와 발송은 사용자가 직접 수행합니다.

## 처리 흐름

```
Outlook 지정 폴더 → ① 수집(collector) → ② 사내 LLM 분석(analyzer) → ③ 산출물(reporter)
                                                                       ├─ output/report_YYYYMMDD.html (검토 대시보드)
                                                                       ├─ output/tracker.xlsx (누적 관리대장)
                                                                       └─ Outlook 임시보관함 초안 (저장만)
```

- 원본 메일은 **절대 변경하지 않습니다** (읽음처리 X, 이동 X, 삭제 X).
- LLM 호출은 **사내 엔드포인트로만** 이루어집니다 (표준 라이브러리 `urllib` 사용, 외부 API 없음).
- Factsheet의 `public_yn = N`(대외비) 행은 변환 단계에서 원천 제외되어 LLM에 전달되지 않습니다.

## 설치 (Windows 업무 PC)

1. Python 3.10+ 설치 확인: `python --version`
2. 패키지 설치:
   ```
   pip install pywin32 openpyxl python-dateutil
   ```
3. Outlook 데스크톱 앱이 설치·로그인되어 있어야 합니다.

## 초기 설정

### 1) LLM 접속 정보 (`llm_config.json`)
`llm_config.example.json`을 복사해 `llm_config.json`을 만들고 사내 LLM 정보를 입력합니다.
이 파일은 `.gitignore`에 등록되어 있어 커밋되지 않습니다.

```json
{ "base_url": "http://사내LLM주소/v1", "api_key": "발급받은 키", "model": "glm-5.2" }
```

환경변수(`ESG_LLM_BASE_URL`, `ESG_LLM_API_KEY`, `ESG_LLM_MODEL`)가 있으면 파일보다 우선합니다.

### 2) 업무 설정 (`config.json`)
- `outlook.folder_path` : 수집 대상 폴더 (예: `받은편지함/ESG요청`)
- `customers` : 고객사명과 발신 도메인 매핑 (발신 도메인으로 고객사를 자동 판별)
- `dept_contacts` : 부서별 데이터 요청 수신자 주소 (비우면 초안의 수신자가 빈 채로 저장됨)

### 3) Factsheet 준비
1. `data/factsheet.xlsx`를 아래 스키마로 작성합니다 (사람이 유지보수하는 원본):

   | 컬럼 | 설명 |
   |------|------|
   | item_code | 지표 고유코드 (예: `E-GHG-S1`) |
   | category | E/S/G 대분류 |
   | item_name | 지표명 |
   | year | 기준연도 |
   | site | 사이트 (`이천`/`전사` 등) |
   | value | 값 |
   | unit | 단위 |
   | source | 출처 |
   | **public_yn** | **공개가능여부 Y/N — N이면 LLM·산출물에서 원천 제외** |
   | note | 비고 |

2. agent용 사본 생성 (xlsx 수정 시마다 재실행):
   ```
   python tools\xlsx_to_md.py
   ```
   실행 결과에 "제외 N행" 수가 표시되므로 대외비 필터링 여부를 확인할 수 있습니다.

   형식 확인용 가짜 데이터가 필요하면 `python tools\make_sample_factsheet.py`로 샘플을 만들 수 있습니다.

## 실행

```
run_agent.bat                실행 (수집 → 분석 → 리포트, 완료 시 브라우저 자동 오픈)
run_agent.bat --no-browser   리포트 자동 오픈 없이 실행 (스케줄러용)
run_agent.bat --limit 10     이번 실행에서 최대 10건만 수집
run_agent.bat --skip-collect 수집 생략, work/inbox의 기존 JSON 재분석
run_agent.bat --no-drafts    Outlook 초안 저장 생략
```

시작 시 LLM 연결을 먼저 확인하며, **연결 실패 시 아무 분석도 하지 않고 즉시 종료**합니다.
메일 1건 분석이 실패해도 전체는 계속 진행되고, 실패 건은 리포트에 빨간 `ERROR`로 표시된 뒤
다음 실행에서 자동 재시도됩니다.

### 매일 자동 실행 (작업 스케줄러)
```
tools\register_task.bat      매일 08:30 자동 실행 등록
```

## 산출물 사용법

### `output/report_YYYYMMDD.html` — 검토 대시보드
- 상단 KPI: 신규 / 긴급 / 즉답가능 / 부서협조필요 / 분석실패 건수
- 고객사별 카드: 요구사항 표(상태 색상), 접이식 답변·부서요청 초안(**[복사] 버튼**),
  원본 메일 열기 링크
- 단일 파일이므로 사내망 오프라인에서 그대로 열립니다.

### `output/tracker.xlsx` — 누적 관리대장
- 실행할 때마다 **신규 행만 추가**됩니다 (기존 행은 절대 수정·삭제되지 않음).
- `처리상태(수기)`, `비고(수기)` 열은 사용자가 직접 기입하며 재실행 시에도 보존됩니다.
- 마지막 `mail_id(시스템)` 열은 중복 방지용 키이므로 수정하지 마세요.

### Outlook 임시보관함 초안
- 제목이 `[AI초안] `으로 시작하는 초안이 저장됩니다.
- 고객 답변 초안은 수신자가 비어 있고, 부서요청 초안은 `config.json`의 담당자 주소가 채워집니다.
- 모든 초안 본문 최상단에 `⚠️ [AI 생성 초안 — 발송 전 반드시 검토 필요]` 배너가 있습니다.
- **검토·수정 후 사용자가 직접 발송하세요.**

## 상태 파일

| 경로 | 용도 |
|------|------|
| `work/inbox/` | 수집된 메일 원문 JSON |
| `work/analyzed/` | 분석 결과 JSON |
| `work/processed_ids.json` | 처리 완료 메일 ID (중복 처리 방지). 특정 메일을 다시 처리하려면 해당 ID를 삭제 |
| `logs/YYYYMMDD.log` | 일자별 실행 로그 |

## 테스트

Outlook이 필요 없는 로직(추출 규칙, 대외비 필터, 트래커 보존, 부분 실패, 금지 코드 검사)은
로컬 목 LLM 서버로 자동 검증됩니다:
```
python tests\run_tests.py
```

## 설계 제약 (작업지시서 §0)

| # | 제약 |
|---|------|
| C1 | 메일 자동 발송 금지 — 초안은 임시보관함 저장까지만 |
| C2 | 외부 인터넷 통신 금지 — 사내 LLM 엔드포인트만 호출, `urllib` 표준 모듈만 사용 |
| C3 | 원본 메일·파일 삭제/이동/읽음처리 금지, tracker 덮어쓰기 금지 |
| C4 | `public_yn = N` 데이터는 어떤 출력물에도 미포함 (변환 단계에서 원천 제외) |
| C5 | API Key 하드코딩 금지 — `llm_config.json`(gitignore) 또는 환경변수로만 주입 |

## 운영 전 확인 필요 사항 (작업지시서 §9 — 미확정)

아래 항목은 담당자 확인 후 `config.json` 등에 반영해야 합니다. 현재는 플레이스홀더 상태입니다.

1. Outlook 수집 대상 폴더명 — 규칙 기반 자동분류(받은편지함 → ESG 폴더) 선행 여부
2. 고객사 목록 및 발신 도메인 매핑 (`config.json`의 `customers`)
3. 담당부서별 데이터 요청 수신자 주소 (`config.json`의 `dept_contacts`)
4. 사내 LLM의 컨텍스트 최대 길이·분당 호출 제한 (`config.json`의 `llm.max_body_chars` 조정)
5. `factsheet.xlsx` 현행 항목 수 및 PPT → Excel 전환 담당 주체
