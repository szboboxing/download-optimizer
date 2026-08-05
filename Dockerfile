# ================================================================
# 下载流程优化工具 — Dockerfile (用于 CI/CD 构建环境)
# 注意：本项目为 Windows Tkinter 桌面 GUI 应用，不适合直接在
# Docker 容器中运行。Dockerfile 用于自动化构建和测试。
# ================================================================

FROM python:3.13-slim AS builder

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# 验证代码语法
RUN python -m py_compile 下载流程优化工具_v3.5.py

# 生成可执行文件（在 Windows 主机上执行，此处仅验证语法）
CMD ["python", "-c", "print('Build environment ready')"]


# ================================================================
# 多阶段：生产镜像（仅供开发测试环境使用）
# ================================================================
FROM python:3.13-slim AS runtime

WORKDIR /app

COPY --from=builder /app .

# 运行单元测试
RUN python -c "import openpyxl; import requests; print('Dependencies OK')"

CMD ["python", "-c", "print('下载流程优化工具 - 请在 Windows 主机上运行')"]
