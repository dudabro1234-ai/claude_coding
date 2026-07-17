이 폴더에 과거 고객대응이력(.xlsx)을 넣으세요.

  python tools\history_ingest.py

를 실행하면 가장 최근 파일을 읽어 data\history_index.json 을 생성합니다.

양식이 자유로워도 됩니다 — 아래 열만 있으면 자동 인식합니다(이름 별칭 허용):
  필수: 일자(날짜/대응일…) | 고객사(고객…) | 요청내용(요구사항/질문…) | 답변내용(회신내용…)
  선택: 요청유형 | 프레임워크 | 담당자 | 비고(결과…)

빈 양식이 필요하면: python tools\history_ingest.py --make-template
→ data\history_template.xlsx 생성

⚠️ 이 폴더의 파일과 history_index.json은 고객 커뮤니케이션 원문을 포함하므로
   git 커밋·외부 전송이 금지되어 있습니다(.gitignore 등록).
