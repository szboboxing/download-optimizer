# GitHub 开源操作完整指引

本教程将手把手教你如何将「下载流程优化工具」发布到 GitHub 开源。

---

## 目录

1. [准备工作](#准备工作)
2. [注册 GitHub 账号](#第一步注册-github-账号)
3. [安装 Git 客户端](#第二步安装-git-客户端)
4. [创建远程仓库](#第三步创建远程仓库)
5. [上传本地代码到 GitHub](#第四步上传本地代码到-github)
6. [发布 Release 版本](#第五步发布-release-版本)
7. [配置 Docker Compose 支持飞牛 NAS](#第六步配置-docker-compose-支持飞牛-nas)
8. [添加许可证](#第七步添加许可证)
9. [常见问题](#常见问题)

---

## 准备工作

在开始之前，请确保你已经：
- ✅ 安装了 **Git for Windows**（[下载地址](https://git-scm.com/download/win)）
- ✅ 有一个 GitHub 账号（没有的话先注册）
- ✅ 项目文件已整理好：

```
需要上传的文件清单：
├── 下载流程优化工具_v3.5.py   ← 主程序
├── requirements.txt           ← 依赖列表
├── Dockerfile                 ← Docker 配置
├── docker-compose.yml        ← 飞牛 NAS 配置
├── .gitignore                 ← 忽略规则
├── README.md                  ← 项目说明
├── LICENSE                    ← MIT 许可证
└── dist/下载流程优化工具_v3.5.exe  ← 可执行文件（Release 附件）
```

---

## 第一步：注册 GitHub 账号

1. 打开浏览器，访问 [https://github.com](https://github.com)
2. 点击右上角 **Sign up** 按钮
3. 依次填写：
   - **Username**：建议用英文，如 `zxcv2025`（这会成为你的 GitHub 用户名）
   - **Email**：填写你的邮箱（常用邮箱，用于接收验证邮件）
   - **Password**：设置密码（至少 8 位）
4. 点击 **Continue**
5. 完成人机验证（滑块拖动）
6. 选择个性化设置（可跳过，选择 `Skip personalization`）
7. 到邮箱中点击验证链接完成注册

---

## 第二步：安装 Git 客户端

### Windows 用户

1. 下载 Git for Windows：[https://git-scm.com/download/win](https://git-scm.com/download/win)
2. 运行安装程序，一路默认点击 **Next**
3. 安装完成后，在任意文件夹右键会出现：
   - `Git Bash Here`
   - `Git GUI Here`

### 验证安装

在任意文件夹右键 → **Git Bash Here**，输入：

```bash
git --version
```

显示类似 `git version 2.45.0.windows.1` 即安装成功。

### 配置 Git（必须做）

```bash
# 设置用户名（将 your-name 替换为你的 GitHub 用户名）
git config --global user.name "your-name"

# 设置邮箱（与 GitHub 注册邮箱一致）
git config --global user.email "your-email@example.com"
```

---

## 第三步：创建远程仓库

### 方法 A：网页创建（最简单）

1. 登录 GitHub → 点击右上角 **+** → **New repository**
2. 填写信息：

| 字段 | 填写内容 | 说明 |
|------|----------|------|
| Repository name | `download-optimizer` | 仓库名称，建议用英文 |
| Description | `下载流程优化工具 - 批量重命名与规约上传` | 可选，项目简介 |
| Visibility | **Public** | 公开仓库（免费用户只能选这个） |
| Initialize with README | **勾选** | 初始化一个 README |

3. 点击绿色按钮 **Create repository**

### 方法 B：命令行创建（需要已安装 GitHub CLI）

```bash
gh repo create download-optimizer --public --description "下载流程优化工具"
```

---

## 第四步：上传本地代码到 GitHub

### 准备上传的文件

确保你的项目目录包含以下文件：

```bash
# 在项目根目录打开 Git Bash Here
# 将你的文件整理好，只保留需要上传的
```

### 执行上传命令

```bash
# 1. 克隆空仓库到本地（把 your-name 替换为你的用户名）
git clone https://github.com/your-name/download-optimizer.git

# 2. 进入仓库目录
cd download-optimizer

# 3. 将所有项目文件复制到这个目录
# （手动复制 下载流程优化工具_v3.5.py, README.md, requirements.txt 等）

# 4. 添加所有文件到暂存区
git add .

# 5. 提交更改
git commit -m "feat: 初始化下载流程优化工具 v3.5

- 批量重命名文件夹功能
- 批量重命名文件功能
- 规约上传比对功能
- AI 智能助手模块
- Docker / docker-compose 配置
- MIT 许可证"

# 6. 推送到 GitHub
git push origin main
```

### 如果提示需要登录

当执行 `git push` 时，GitHub 会要求认证。GitHub 不再支持密码，需要使用 **Personal Access Token**：

1. GitHub → 右上角头像 → **Settings**
2. 左侧菜单 → **Developer settings** → **Personal access tokens** → **Tokens (classic)**
3. 点击 **Generate new token (classic)**
4. Note 填写：`GitHub Upload Token`
5. Expiration 选择：`90 days`（或自定义）
6. 勾选权限：勾选 `repo`（全部仓库权限）
7. 点击 **Generate token**
8. **立即复制 Token**（只显示一次！）
9. 回到 Git Bash，输入：
   - Username: 你的 GitHub 用户名
   - Password: 刚才复制的 Token

### 验证上传成功

刷新浏览器中的仓库页面，你应该能看到所有文件。

---

## 第五步：发布 Release 版本

Release 是 GitHub 上发布可下载版本的功能。

### 创建 Release

1. 打开仓库页面 → 点击 **Releases** 标签
2. 点击右侧的 **Draft a new release**
3. 填写：

| 字段 | 填写内容 |
|------|----------|
| Choose a tag | 输入 `v3.5` → 按回车创建 |
| Release title | `下载流程优化工具 v3.5` |
| Description | 使用下方模板 |

### Release 描述模板

```markdown
## 🚀 下载流程优化工具 v3.5

### 📋 更新内容

- ✨ **规约上传模块**：新增比对生成报网公司结算数据表格功能
- 🎨 **UI 优化**：主界面图标改为橙黄色盾牌+对勾设计
- 🐛 **Bug 修复**：
  - 修复合并单元格导致的读取错误
  - 修复条件格式化公式引用稳定性
- 📦 **版本**：升级至 v3.5

### 💻 系统要求

- Windows 10 / Windows 11
- 无需安装 Python

### 📥 安装方法

1. 点击下方 **Assets** 区域的 `下载流程优化工具_v3.5.exe`
2. 保存到本地
3. 双击运行即可

### 🐳 Docker 部署

```bash
# Docker Compose
docker-compose up -d
```

### 📖 飞牛 NAS

将 `docker-compose.yml` 内容粘贴到飞牛 Docker Compose 中启动即可。

### 🔗 完整文档

访问：[项目主页](https://github.com/your-name/download-optimizer)
```

### 上传 EXE 附件

1. 在 Release 编辑页面找到 **Attach binaries by dropping them here or selecting them**
2. 将 `dist/下载流程优化工具_v3.5.exe` 拖拽到上传区域
3. 等待上传完成（文件较大时耐心等待）
4. 点击 **Publish release** 按钮

### 验证 Release

在仓库 **Releases** 页面，应该能看到刚发布的版本。点击 `下载流程优化工具_v3.5.exe` 即可下载。

---

## 第六步：配置 Docker Compose 支持飞牛 NAS

### 飞牛 NAS 操作步骤

1. 登录飞牛 NAS 管理界面（通常是 `http://NAS-IP:5666`）
2. 进入 **Docker** → **Compose**（或 **应用中心** → **添加 Compose**）
3. 点击 **添加 Compose** 按钮
4. 填写信息：
   - **名称**：`下载流程优化工具`
   - **描述**：`批量重命名与规约上传工具`
   - **Compose 内容**：粘贴仓库中 `docker-compose.yml` 的内容
5. 点击 **立即启动**
6. 等待容器状态变为 **运行中**
7. 点击容器 → **日志**，查看环境检查结果

### 通用 Docker Compose 操作

```bash
# 克隆仓库
git clone https://github.com/your-name/download-optimizer.git
cd download-optimizer

# 启动所有服务
docker-compose up -d

# 查看状态
docker-compose ps

# 查看日志
docker-compose logs -f

# 停止服务
docker-compose down

# 停止并删除数据卷
docker-compose down -v
```

### 配置文件说明

`docker-compose.yml` 包含两个服务：

| 服务 | 说明 | 端口 |
|------|------|------|
| `app` | 主应用（环境验证） | 无端口（CLI 模式） |
| `api-proxy` | AI API 代理（可选） | 8080:80 |

---

## 第七步：添加许可证

1. 仓库页面 → **Add file** → **Create new file**
2. 文件名：`LICENSE`
3. 内容使用仓库中的 `LICENSE` 文件内容（MIT 许可证）
4. 点击 **Commit changes**

---

## 常见问题

### Q1: push 时报错 `remote Permission denied`？

**原因**：用户名或 Token 错误。

**解决**：
1. 重新生成 Token（Settings → Developer settings → Personal access tokens → Generate new token）
2. 勾选 `repo` 权限
3. 在 Git Bash 中重新推送，输入新 Token

### Q2: push 时报错 `repository not found`？

**原因**：仓库名拼写错误，或没有访问权限。

**解决**：
1. 确认仓库名正确：`download-optimizer`
2. 确认用户名正确
3. URL 应该是：`https://github.com/你的用户名/download-optimizer.git`

### Q3: git clone 下载慢？

**解决**：使用国内镜像：
```bash
# 使用 gitee 镜像（备选）
git clone https://gitee.com/your-name/download-optimizer.git
```

### Q4: 如何更新已发布的代码？

```bash
# 1. 修改本地文件
# 2. 添加
git add .
# 3. 提交（可以写更详细的提交信息）
git commit -m "fix: 修复规约上传比对逻辑"
# 4. 推送
git push origin main
# 5. GitHub 会自动更新仓库内容
```

### Q5: 如何发布新的 Release 版本？

重复第五步的操作：
1. 点击 Releases → Draft a new release
2. Tag version 填入新版本号（如 `v3.6`）
3. 上传新的 EXE
4. 发布

### Q6: 如何删除仓库？

1. 仓库页面 → **Settings**
2. 滚动到底部 → **Danger Zone**
3. 点击 **Delete this repository**
4. 输入仓库名确认删除

### Q7: 飞牛 NAS 上 Docker 启动失败怎么办？

1. 检查 NAS 是否开启了 Docker 功能
2. 检查 Compose 文件是否有格式错误
3. 查看容器日志：Docker → 容器 → 日志
4. 尝试精简 Compose 文件，只保留核心服务

---

## 进阶：添加 GitHub Actions 自动构建

在仓库中创建 `.github/workflows/build.yml` 文件：

```yaml
name: Build EXE

on:
  push:
    tags: ['v*']

jobs:
  build:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v4
      - name: Set up Python
        uses: actions/setup-python@v5
        with:
          python-version: '3.13'
      - run: pip install -r requirements.txt pyinstaller
      - run: pyinstaller --onefile --name "下载流程优化工具_${{ github.ref_name }}" --hidden-import win32com --hidden-import win32com.client --hidden-import pythoncom --hidden-import pywintypes --noconsole 下载流程优化工具_v3.5.py
      - uses: actions/upload-artifact@v4
        with:
          name: exe-artifact
          path: dist/*.exe
```

---

## 总结

完成以上 7 步后，你的项目已经开源成功：

| 步骤 | 状态 | 产出 |
|------|------|------|
| 注册账号 | ✅ | GitHub 账号 |
| 创建仓库 | ✅ | `download-optimizer` 仓库 |
| 上传代码 | ✅ | 源码在 GitHub 上 |
| 发布 Release | ✅ | 用户可下载 EXE |
| Docker 配置 | ✅ | 飞牛 NAS 可用 |
| 添加许可证 | ✅ | MIT License |

🎉 **恭喜！你已成为开源社区的一员！**
