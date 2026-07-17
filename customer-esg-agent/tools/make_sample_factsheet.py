# -*- coding: utf-8 -*-
"""테스트용 샘플 factsheet.xlsx 생성기.

실데이터가 아닌 가상의 값으로 §5.1 스키마의 예시 파일을 만든다.
public_yn=N 행을 일부 포함시켜 C4 필터링(T4) 검증에 사용한다.

사용법: python tools/make_sample_factsheet.py [출력경로]
"""
import os
import sys

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(BASE_DIR, "data", "factsheet.xlsx")

HEADER = ["item_code", "category", "item_name", "year", "site",
          "value", "unit", "source", "public_yn", "note"]

SAMPLE_ROWS = [
    ["E-GHG-S1", "E", "Scope 1 배출량", 2025, "전사", 1234567, "tCO2eq",
     "2026 지속가능경영보고서 p.42", "Y", "제3자 검증 완료"],
    ["E-GHG-S2", "E", "Scope 2 배출량", 2025, "전사", 2345678, "tCO2eq",
     "2026 지속가능경영보고서 p.42", "Y", "시장기반"],
    ["E-GHG-S3", "E", "Scope 3 배출량", 2025, "전사", 3456789, "tCO2eq",
     "2026 지속가능경영보고서 p.43", "N", "내부 산정치 — 대외비"],
    ["E-ENE-RE", "E", "재생에너지 사용 비율", 2025, "전사", 31.5, "%",
     "2026 지속가능경영보고서 p.45", "Y", ""],
    ["E-WTR-USE", "E", "용수 사용량", 2025, "이천", 45678, "천톤",
     "2026 지속가능경영보고서 p.51", "Y", ""],
    ["S-LAB-RBA", "S", "RBA 인증 현황", 2025, "전사", "Full Member", "-",
     "RBA 회원 등록증", "Y", ""],
    ["S-SAF-LTIR", "S", "근로손실재해율(LTIR)", 2025, "전사", 0.08, "-",
     "내부 안전보건 통계", "N", "대외 미공개"],
    ["G-ETH-EDU", "G", "윤리교육 이수율", 2025, "전사", 99.2, "%",
     "2026 지속가능경영보고서 p.78", "Y", ""],
]


def make(out_path=DEFAULT_OUT):
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "factsheet"
    ws.append(HEADER)
    for row in SAMPLE_ROWS:
        ws.append(row)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    wb.save(out_path)
    return out_path


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_OUT
    print("샘플 factsheet 생성:", make(out))
    print("주의: 가상의 예시 데이터입니다. 실데이터로 교체하세요.")
