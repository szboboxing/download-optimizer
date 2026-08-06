#!/bin/bash
# ================================================================
# 下载流程优化工具 — Docker 镜像构建脚本
# 使用方法：
#   chmod +x build-docker.sh
#   ./build-docker.sh              # 构建并本地运行
#   ./build-docker.sh --push       # 构建并推送到 Docker Hub
#   ./build-docker.sh --run        # 构建并运行容器
#   ./build-docker.sh --clean      # 清理构建产物
# ================================================================

set -e

IMAGE_NAME="download-optimizer"
IMAGE_TAG="v3.11"
DOCKER_HUB_USER="${DOCKER_HUB_USER:-szboboxing}"
FULL_IMAGE="${DOCKER_HUB_USER}/${IMAGE_NAME}:${IMAGE_TAG}"
LATEST_IMAGE="${DOCKER_HUB_USER}/${IMAGE_NAME}:latest"

COLOR_CYAN='\033[36m'
COLOR_GREEN='\033[32m'
COLOR_YELLOW='\033[33m'
COLOR_RED='\033[31m'
COLOR_RESET='\033[0m'

log_info()  { echo -e "${COLOR_CYAN}[INFO]${COLOR_RESET} $*"; }
log_ok()    { echo -e "${COLOR_GREEN}[OK]${COLOR_RESET} $*"; }
log_warn()  { echo -e "${COLOR_YELLOW}[WARN]${COLOR_RESET} $*"; }
log_error() { echo -e "${COLOR_RED}[ERROR]${COLOR_RESET} $*"; }

usage() {
    echo "下载流程优化工具 v3.11 — Docker 构建脚本"
    echo ""
    echo "用法: $0 [选项]"
    echo ""
    echo "选项:"
    echo "  (无参数)     构建镜像并本地验证"
    echo "  --run        构建并运行容器（交互模式）"
    echo "  --push       构建并推送到 Docker Hub"
    echo "  --clean      清理镜像和容器"
    echo "  --help       显示帮助"
    echo ""
    echo "环境变量:"
    echo "  DOCKER_HUB_USER  Docker Hub 用户名（默认: szboboxing）"
    echo ""
    echo "示例:"
    echo "  $0                       # 构建镜像"
    echo "  $0 --push                # 推送到 Docker Hub"
    echo "  DOCKER_HUB_USER=myuser $0 --push"
    exit 0
}

do_build() {
    log_info "开始构建 Docker 镜像..."
    log_info "镜像名称: ${IMAGE_NAME}:${IMAGE_TAG}"

    docker build \
        -t "${IMAGE_NAME}:${IMAGE_TAG}" \
        -t "${IMAGE_NAME}:latest" \
        --file Dockerfile \
        .

    log_ok "镜像构建完成"
    log_info "镜像信息:"
    docker images | grep "${IMAGE_NAME}"
}

do_run() {
    log_info "运行容器..."

    docker run --rm -it \
        --name "${IMAGE_NAME}-test" \
        -v "$(pwd)/data:/app/data" \
        -v "$(pwd)/input:/app/input" \
        -v "$(pwd)/output:/app/output" \
        -e TZ=Asia/Shanghai \
        "${IMAGE_NAME}:${IMAGE_TAG}"
}

do_push() {
    log_info "推送到 Docker Hub: ${FULL_IMAGE}"

    if ! docker info 2>/dev/null | grep -q "Username"; then
        log_warn "未检测到 Docker Hub 登录状态"
        log_info "请先执行: docker login"
        log_info "用户名: ${DOCKER_HUB_USER}"
    fi

    docker tag "${IMAGE_NAME}:${IMAGE_TAG}" "${FULL_IMAGE}"
    docker tag "${IMAGE_NAME}:latest" "${LATEST_IMAGE}"

    log_info "推送 ${FULL_IMAGE}..."
    docker push "${FULL_IMAGE}"

    log_info "推送 ${LATEST_IMAGE}..."
    docker push "${LATEST_IMAGE}"

    log_ok "推送完成！"
    log_info "Docker Hub: https://hub.docker.com/r/${DOCKER_HUB_USER}/${IMAGE_NAME}"
}

do_clean() {
    log_warn "清理镜像和容器..."

    docker rm -f "${IMAGE_NAME}-test" 2>/dev/null || true
    docker rmi "${IMAGE_NAME}:${IMAGE_TAG}" 2>/dev/null || true
    docker rmi "${IMAGE_NAME}:latest" 2>/dev/null || true
    docker rmi "${FULL_IMAGE}" 2>/dev/null || true
    docker rmi "${LATEST_IMAGE}" 2>/dev/null || true

    log_ok "清理完成"
}

# ---- 主流程 ----
cd "$(dirname "$0")"

case "${1:-}" in
    --help|-h)
        usage
        ;;
    --run)
        do_build
        do_run
        ;;
    --push)
        do_build
        do_push
        ;;
    --clean)
        do_clean
        ;;
    *)
        do_build
        echo ""
        log_ok "本地验证命令:"
        echo "  docker run --rm -it ${IMAGE_NAME}:${IMAGE_TAG}"
        echo ""
        log_info "飞牛 NAS / Docker Compose:"
        echo "  docker-compose up -d"
        ;;
esac
