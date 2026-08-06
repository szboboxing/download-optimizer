# ================================================================
# 下载流程优化工具 — Dockerfile
# 多阶段构建：builder（构建环境）+ runtime（运行环境）
# ================================================================

FROM python:3.13-slim AS builder

LABEL maintainer="szboboxing" \
      project="download-optimizer" \
      version="3.11" \
      description="下载流程优化工具 - 批量重命名、规约上传数据表准备与 AI 助手"

WORKDIR /build

ENV VIRTUAL_ENV=/opt/venv
RUN python -m venv "$VIRTUAL_ENV"
ENV PATH="$VIRTUAL_ENV/bin:$PATH"

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY main.py app.py .
COPY README.md LICENSE .

# 验证代码语法
RUN python -c "import py_compile; py_compile.compile('app.py', doraise=True); print('app.py syntax OK')"
RUN python -c "import py_compile; py_compile.compile('main.py', doraise=True); print('main.py syntax OK')"

# 验证依赖导入
RUN python -c "import openpyxl; import requests; print('Dependencies OK')"


# ================================================================
# 运行阶段
# ================================================================
FROM python:3.13-slim AS runtime

LABEL maintainer="szboboxing" \
      project="download-optimizer" \
      version="3.11"

WORKDIR /app

# 复制隔离的 Python 依赖
COPY --from=builder /opt/venv /opt/venv

# 从构建阶段复制产物
COPY --from=builder /build /app

# 设置环境变量
ENV PATH="/opt/venv/bin:$PATH" \
    PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    TZ=Asia/Shanghai \
    APP_HOME=/app

# 创建数据目录
RUN mkdir -p /app/data /app/input /app/output /app/logs

# 设置非 root 用户运行（安全最佳实践）
RUN groupadd -r appuser && useradd -r -g appuser -d /app appuser && \
    chown -R appuser:appuser /app

USER appuser

# 健康检查
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD python -c "import openpyxl, requests; print('Health OK')" || exit 1

# 默认入口：显示版本和环境信息
ENTRYPOINT ["python", "-c", "\nimport sys, openpyxl, requests\nprint('='*50)\nprint('下载流程优化工具 Docker 环境')\nprint('='*50)\nprint(f'Python: {sys.version}')\nprint(f'openpyxl: {openpyxl.__version__}')\nprint(f'requests: {requests.__version__}')\nprint('依赖检查通过 ✓')\nprint('应用需在 Windows 主机上运行，请下载 EXE:')\nprint('https://github.com/szboboxing/download-optimizer/releases')\n"]

CMD ["--help"]
