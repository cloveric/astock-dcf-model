# astock-dcf-model 一体化镜像: Web 服务 + 建模引擎 + LibreOffice 验收工具链
FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    HOST=0.0.0.0 \
    PORT=8000

# curl: 数据层(东财F10/腾讯行情)只走系统 curl; libreoffice-calc: verify_model.py 重算验收(可选但建议)
RUN apt-get update \
    && apt-get install -y --no-install-recommends curl libreoffice-calc \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt ./
RUN pip install -r requirements.txt
COPY . .

EXPOSE 8000

# 启动 Web 服务; 容器内也可直接执行:
#   python build_model.py --code 300476 && python verify_model.py --code 300476
CMD ["python", "-m", "web.server"]
