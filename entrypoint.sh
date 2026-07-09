#!/bin/sh
# 최초 배포 시 볼륨(/data/input)이 비어 있으면 샘플 데이터로 시딩
# → 실데이터 CSV로 교체하기 전까지 샘플로 동작 (교체 시 재시작 불필요)
DATA_DIR="${PPA_DATA_DIR:-/data/input}"
mkdir -p "$DATA_DIR"
if [ -z "$(ls "$DATA_DIR"/*.csv 2>/dev/null)" ]; then
  echo "[init] $DATA_DIR 이 비어 있어 샘플 데이터를 복사합니다."
  cp /app/data/sample/*.csv "$DATA_DIR"/
fi
exec python server.py
