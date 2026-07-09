# -*- coding: utf-8 -*-
"""
실데이터 인코딩·폴더 자동 인식 테스트

실행:  python tools/encoding_test.py

사내 실데이터에서 반복된 로드 실패(UTF-8 고정 읽기 → 엑셀 저장본 CP949에서 예외)를
재현·검증한다:

  [1] CP949(엑셀 'CSV' 저장)로 인코딩된 4종 CSV가 그대로 로드되고,
      계산 결과가 UTF-8 원본과 완전히 동일
  [2] UTF-8(BOM 포함) / BOM 없는 UTF-8 도 동일하게 동작
  [3] 컬럼명 앞뒤 공백이 있어도 매칭됨
  [4] 인식 불가 파일은 "엑셀에서 CSV UTF-8로 다시 저장" 안내와 함께 실패
  [5] data/real_data 폴더가 있으면 서버가 환경변수 없이 자동 선택
  [6] 엑셀 통합문서(.xlsx)로 저장한 입력 파일도 그대로 로드 (CSV와 혼용 가능)
  [7] xlsx 내용인데 이름만 *.csv 인 파일도 자동 감지해 로드 (엑셀 '다른 이름으로 저장' 실수 대응)
  [8] 파일 탐색 우선순위: 같은 이름의 .csv 가 없으면 .xlsx → .xls 순으로 찾음
"""

import os
import shutil
import subprocess
import sys
import tempfile

THIS_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.dirname(THIS_DIR)
sys.path.insert(0, BASE_DIR)

import pandas as pd

from core.config import SimulationParams
from core.engine import load_data, company_total_cost, read_csv_kr, find_input_file

SAMPLE_DIR = os.path.join(BASE_DIR, "data", "sample")
FILES = ["Hourly_Data.csv", "Annual_Usage.csv", "Annual_PPA.csv", "Annual_Rate.csv"]


def transcode_dir(dst, encoding):
    """샘플 CSV를 지정 인코딩으로 순수 텍스트 변환(수치 불변)해 dst에 만든다."""
    os.makedirs(dst, exist_ok=True)
    for name in FILES:
        with open(os.path.join(SAMPLE_DIR, name), encoding="utf-8-sig") as f:
            text = f.read()
        with open(os.path.join(dst, name), "w", encoding=encoding, newline="") as f:
            f.write(text)


def total(data_dir):
    params = SimulationParams()
    return company_total_cost(load_data(data_dir), params, params.fixed_ppas)[0]


def main():
    base_total = total(SAMPLE_DIR)
    print(f"[기준] data/sample(UTF-8 BOM) 총비용 {base_total:,.0f} 백만원")

    with tempfile.TemporaryDirectory() as tmp:
        # ── [1][2] 인코딩별 로드 + 결과 동일성 ──
        for enc, tag in [("cp949", "CP949(엑셀 저장본)"), ("utf-8", "UTF-8(BOM 없음)"),
                         ("utf-8-sig", "UTF-8(BOM)")]:
            d = os.path.join(tmp, enc)
            transcode_dir(d, enc)
            tc = total(d)
            assert abs(tc - base_total) < 1e-3, (enc, tc, base_total)
            print(f"[1·2] {tag} OK — 결과 동일")

        # ── [3] 컬럼명 앞뒤 공백 ──
        d = os.path.join(tmp, "spaced")
        transcode_dir(d, "cp949")
        p = os.path.join(d, "Annual_Rate.csv")
        with open(p, encoding="cp949") as f:
            lines = f.read().splitlines(True)
        lines[0] = ",".join(f" {c.strip()} " for c in lines[0].strip().split(",")) + "\n"
        with open(p, "w", encoding="cp949", newline="") as f:
            f.writelines(lines)
        assert abs(total(d) - base_total) < 1e-3
        print("[3] 컬럼명 공백 OK — strip 후 매칭")

        # ── [4] 인식 불가 인코딩 → 안내 메시지 ──
        bad = os.path.join(tmp, "bad.csv")
        with open(bad, "wb") as f:
            f.write("Year,값\n".encode("utf-16"))   # utf-16은 지원 목록에 없음
        try:
            read_csv_kr(bad)
            raise AssertionError("utf-16 파일이 오류 없이 읽힘")
        except UnicodeError as e:
            assert "CSV UTF-8" in str(e), str(e)
        print("[4] 오류 안내 OK — 재저장 방법 포함")

        # ── [6] 엑셀(.xlsx) 입력 — 연간 3종+worst 를 xlsx로, Hourly는 CSV (혼용) ──
        d = os.path.join(tmp, "excel")
        transcode_dir(d, "utf-8-sig")
        for stem in ["Annual_Usage", "Annual_PPA", "Annual_Rate"]:
            csv_path = os.path.join(d, stem + ".csv")
            pd.read_csv(csv_path).to_excel(os.path.join(d, stem + ".xlsx"), index=False)
            os.remove(csv_path)
        assert abs(total(d) - base_total) < 1e-3
        print("[6] 엑셀(.xlsx) OK — CSV와 혼용 로드, 결과 동일")

        # ── [7] xlsx 내용인데 이름만 *.csv ──
        d = os.path.join(tmp, "fakecsv")
        transcode_dir(d, "utf-8-sig")
        rate_csv = os.path.join(d, "Annual_Rate.csv")
        pd.read_csv(rate_csv).to_excel(rate_csv + ".tmp.xlsx", index=False)
        os.replace(rate_csv + ".tmp.xlsx", rate_csv)   # xlsx 바이트, 이름은 .csv
        assert abs(total(d) - base_total) < 1e-3
        print("[7] 이름만 .csv 인 엑셀 파일 OK — 매직 바이트로 감지")

        # ── [8] 파일 탐색 우선순위 (.csv 없으면 .xlsx) ──
        d = os.path.join(tmp, "find")
        os.makedirs(d)
        pd.DataFrame({"Year": [2026]}).to_excel(os.path.join(d, "Hourly_Data.xlsx"), index=False)
        assert find_input_file(d, "Hourly_Data").endswith(".xlsx")
        assert find_input_file(d, "Annual_PPA", required=False) is None
        print("[8] 파일 탐색 OK — .csv → .xlsx 순 폴백")

    # ── [5] data/real_data 자동 인식 (서버 DATA_DIR 선택 로직) ──
    real_dir = os.path.join(BASE_DIR, "data", "real_data")
    created = not os.path.isdir(real_dir)
    had_hourly = os.path.exists(os.path.join(real_dir, "Hourly_Data.csv"))
    if had_hourly:
        print("[5] data/real_data 에 실데이터가 이미 있어 생성 없이 확인만 합니다.")
    else:
        transcode_dir(real_dir, "cp949")   # 실데이터 상황 재현: CP949 파일 배치
    try:
        env = {**os.environ}
        env.pop("PPA_DATA_DIR", None)
        out = subprocess.run(
            [sys.executable, "-c", "import server; print(server.DATA_DIR)"],
            cwd=BASE_DIR, env=env, capture_output=True, text=True, check=True,
        ).stdout.strip()
        assert out.endswith(os.path.join("data", "real_data")), out
        print(f"[5] 자동 인식 OK — 환경변수 없이 DATA_DIR = {out}")
    finally:
        if not had_hourly:
            shutil.rmtree(real_dir) if created else [
                os.remove(os.path.join(real_dir, f)) for f in FILES
                if os.path.exists(os.path.join(real_dir, f))]

    print("\n✅ 인코딩·실데이터 폴더 테스트 통과")


if __name__ == "__main__":
    main()
