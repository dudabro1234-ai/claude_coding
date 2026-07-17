이 폴더에 사내 데이터플랫폼에서 받은 datasheet(.xlsx)를 넣으세요.

  python tools\platform_ingest.py

를 실행하면 가장 최근 파일을 읽어 data\platform_index.json 을 생성합니다.

⚠️ 이 폴더의 파일과 생성된 platform_index.json은 대외비 원천값을 포함하므로
   git 커밋·외부 전송이 금지되어 있습니다(.gitignore 등록).
