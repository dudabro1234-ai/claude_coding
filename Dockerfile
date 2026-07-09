# RE100 PPA Simulator — 서버 배포용 이미지
#
# 빌드 (사내 레지스트리/미러 사용):
#   docker build -t docker-repo.skhynix.com/esg/ppa-dashboard:1.0.0 \
#     --build-arg PIP_INDEX_URL=http://<사내미러주소>/simple \
#     --build-arg PIP_TRUSTED_HOST=<사내미러주소> .
#
# 푸시:
#   docker push docker-repo.skhynix.com/esg/ppa-dashboard:1.0.0

ARG BASE_REGISTRY=dockerhub.skhynix.com
FROM ${BASE_REGISTRY}/python:3.13-slim

# 사내 pip 미러 (빌드 시 --build-arg로 지정, 미지정 시 기본 PyPI)
ARG PIP_INDEX_URL=https://pypi.org/simple
ARG PIP_TRUSTED_HOST=
ENV PIP_INDEX_URL=${PIP_INDEX_URL} \
    PIP_TRUSTED_HOST=${PIP_TRUSTED_HOST} \
    PYTHONUNBUFFERED=1 \
    PPA_PORT=8700 \
    PPA_DATA_DIR=/data/input \
    PPA_LLM_CONFIG_PATH=/data/llm_config.json

WORKDIR /app

COPY requirements-docker.txt .
RUN pip install --no-cache-dir -r requirements-docker.txt

# 앱 소스 (샘플 데이터 포함 — 볼륨 미마운트 시 폴백용)
COPY core/ core/
COPY web/ web/
COPY server.py agent.py entrypoint.sh ./
COPY data/sample/ data/sample/
RUN chmod +x entrypoint.sh

EXPOSE 8700

HEALTHCHECK --interval=60s --timeout=10s --start-period=30s \
  CMD python -c "import urllib.request,os;urllib.request.urlopen(f'http://127.0.0.1:{os.environ.get(\"PPA_PORT\",\"8700\")}/health',timeout=5)"

CMD ["./entrypoint.sh"]
