<#
.SYNOPSIS
    下载流程优化工具 — Docker 镜像构建脚本 (PowerShell)

.DESCRIPTION
    构建、运行、推送 Docker 镜像到 Docker Hub。
    支持飞牛 NAS / Windows / Linux 环境。

.PARAMETER Action
    build   构建镜像并本地验证（默认）
    run     构建并运行容器（交互模式）
    push    构建并推送到 Docker Hub
    clean   清理镜像和容器
    help    显示帮助

.EXAMPLE
    .\build-docker.ps1                     # 构建镜像
    .\build-docker.ps1 -Action run         # 构建并运行
    .\build-docker.ps1 -Action push        # 推送到 Docker Hub
    $env:DOCKER_HUB_USER='myuser'; .\build-docker.ps1 -Action push
#>

param(
    [ValidateSet('build', 'run', 'push', 'clean', 'help')]
    [string]$Action = 'build'
)

$ErrorActionPreference = 'Stop'

$ImageName    = 'download-optimizer'
$ImageTag     = 'v3.5'
$DockerUser   = if ($env:DOCKER_HUB_USER) { $env:DOCKER_HUB_USER } else { 'szboboxing' }
$FullImage    = "${DockerUser}/${ImageName}:${ImageTag}"
$LatestImage  = "${DockerUser}/${ImageName}:latest"
$ProjectRoot  = $PSScriptRoot

function Write-Info    { param($msg) Write-Host "[INFO]    $msg" -ForegroundColor Cyan }
function Write-Ok      { param($msg) Write-Host "[OK]      $msg" -ForegroundColor Green }
function Write-Warn    { param($msg) Write-Host "[WARN]    $msg" -ForegroundColor Yellow }
function Write-Err     { param($msg) Write-Host "[ERROR]   $msg" -ForegroundColor Red }
function Write-Step    { param($msg) Write-Host "`n==== $msg ====" -ForegroundColor White }

function Test-Docker {
    try {
        $null = docker version 2>&1
        return $true
    } catch {
        Write-Err "Docker 未安装或未运行。请安装 Docker Desktop:"
        Write-Err "  Windows: https://docs.docker.com/desktop/install/windows-install/"
        Write-Err "  Linux:   curl -fsSL https://get.docker.com | sh"
        return $false
    }
}

function Invoke-Build {
    Write-Step "构建 Docker 镜像"
    Write-Info "镜像名称: ${ImageName}:${ImageTag}"
    Write-Info "项目目录: $ProjectRoot"

    docker build `
        -t "${ImageName}:${ImageTag}" `
        -t "${ImageName}:latest" `
        --file Dockerfile `
        $ProjectRoot

    if ($LASTEXITCODE -ne 0) {
        Write-Err "构建失败"
        exit 1
    }

    Write-Ok "镜像构建完成"
    Write-Host ""
    Write-Info "镜像列表:"
    docker images | Select-String $ImageName
}

function Invoke-Run {
    Write-Step "运行容器"

    $dataDir = Join-Path $ProjectRoot "data"
    $inputDir = Join-Path $ProjectRoot "input"
    $outputDir = Join-Path $ProjectRoot "output"

    foreach ($dir in @($dataDir, $inputDir, $outputDir)) {
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    }

    docker run --rm -it `
        --name "${ImageName}-test" `
        -v "${dataDir}:/app/data" `
        -v "${inputDir}:/app/input" `
        -v "${outputDir}:/app/output" `
        -e TZ=Asia/Shanghai `
        "${ImageName}:${ImageTag}"
}

function Invoke-Push {
    Write-Step "推送到 Docker Hub"

    $loginStatus = docker info 2>&1 | Select-String "Username"
    if (-not $loginStatus) {
        Write-Warn "未检测到 Docker Hub 登录状态"
        Write-Info "请先执行: docker login"
        Write-Info "用户名: $DockerUser"
    }

    Write-Info "标记镜像..."
    docker tag "${ImageName}:${ImageTag}" "${FullImage}"
    docker tag "${ImageName}:latest" "${LatestImage}"

    Write-Info "推送 ${FullImage}..."
    docker push "${FullImage}"
    if ($LASTEXITCODE -ne 0) {
        Write-Err "推送失败，请检查 Docker Hub 登录状态"
        exit 1
    }

    Write-Info "推送 ${LatestImage}..."
    docker push "${LatestImage}"

    Write-Ok "推送完成！"
    Write-Info "Docker Hub: https://hub.docker.com/r/${DockerUser}/${ImageName}"
}

function Invoke-Clean {
    Write-Step "清理镜像和容器"

    docker rm -f "${ImageName}-test" 2>$null | Out-Null
    docker rmi "${ImageName}:${ImageTag}" 2>$null | Out-Null
    docker rmi "${ImageName}:latest" 2>$null | Out-Null
    docker rmi "${FullImage}" 2>$null | Out-Null
    docker rmi "${LatestImage}" 2>$null | Out-Null

    Write-Ok "清理完成"
}

function Show-Help {
    Write-Host @"
下载流程优化工具 v3.5 — Docker 构建脚本

用法: .\build-docker.ps1 [-Action <操作>]

操作:
  build    构建镜像并本地验证 (默认)
  run      构建并运行容器（交互模式）
  push     构建并推送到 Docker Hub
  clean    清理镜像和容器
  help     显示帮助

环境变量:
  DOCKER_HUB_USER  Docker Hub 用户名 (默认: szboboxing)

示例:
  .\build-docker.ps1                     # 构建镜像
  .\build-docker.ps1 -Action run         # 构建并运行
  .\build-docker.ps1 -Action push        # 推送到 Docker Hub
  `$env:DOCKER_HUB_USER='myuser'; .\build-docker.ps1 -Action push

Docker Hub: https://hub.docker.com/repository/docker/${DockerUser}/${ImageName}
GitHub:     https://github.com/szboboxing/download-optimizer
"@
}

# ---- 主流程 ----
Set-Location $ProjectRoot

if (-not (Test-Docker)) {
    exit 1
}

switch ($Action) {
    'help'  { Show-Help }
    'run'   { Invoke-Build; Invoke-Run }
    'push'  { Invoke-Build; Invoke-Push }
    'clean' { Invoke-Clean }
    default {
        Invoke-Build
        Write-Host ""
        Write-Ok "本地验证命令:"
        Write-Host "  docker run --rm -it ${ImageName}:${ImageTag}"
        Write-Host ""
        Write-Info "飞牛 NAS / Docker Compose:"
        Write-Host "  docker-compose up -d"
        Write-Host ""
        Write-Info "推送到 Docker Hub:"
        Write-Host "  .\build-docker.ps1 -Action push"
    }
}
