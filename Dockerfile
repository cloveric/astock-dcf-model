# astock-dcf-model 一体化镜像: Web 服务 + 建模引擎 + LibreOffice 验收工具链
FROM python:3.12-slim@sha256:2c941e860699f878900b0edc2403613c234d4b32eda3cc9fa7036991a2a63c4a

ENV PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    HOST=0.0.0.0 \
    PORT=8000 \
    WEB_MAX_JOBS=100 \
    WEB_MAX_ACTIVE_JOBS=8 \
    WEB_MAX_REQUEST_BYTES=65536 \
    WEB_ALLOW_LLM=0 \
    WEB_ALLOW_UNVERIFIED_DOWNLOAD=0

# curl: 数据层(东财F10/腾讯行情)只走系统 curl; libreoffice-calc: verify_model.py 重算验收(可选但建议)
RUN apt-get update \
    && apt-get install -y --no-install-recommends curl libreoffice-calc \
    && groupadd --system app \
    && useradd --system --gid app --home-dir /app --shell /usr/sbin/nologin app \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements-web.lock ./
RUN pip install --require-hashes -r requirements-web.lock
COPY . .
RUN mkdir -p /app/web/.data && chown -R app:app /app/web/.data

EXPOSE 8000

# /api 需要 WEB_API_TOKEN；静态首页可用于不携带凭证的存活探针。
HEALTHCHECK --interval=30s --timeout=5s --start-period=20s --retries=3 \
    CMD curl --fail --silent --show-error http://127.0.0.1:8000/ >/dev/null || exit 1

USER app

# 启动 Web 服务; 容器内也可直接执行:
#   python build_model.py --code 300476 && python verify_model.py --code 300476
# HOST=0.0.0.0 属于非回环绑定, 未通过 -e WEB_API_TOKEN=<随机长令牌> 配置令牌时会安全拒绝启动。
CMD ["python", "-m", "web.server"]
