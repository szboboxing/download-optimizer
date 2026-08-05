# 下载流程优化工具

> 一个用于批量重命名文件夹/文件、规约上传比对、AI 智能助手的 Windows 桌面 GUI 应用。

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.9+-blue.svg)](https://www.python.org/)
[![Version](https://img.shields.io/badge/version-v3.5-green.svg)](https://github.com/your-repo/download-optimizer/releases)

---

## 功能特点

| 模块 | 说明 |
|------|------|
| 📁 批量重命名文件夹 | 删除前 N 位字符、后加内容、前加内容，支持预览和回退 |
| 📄 批量重命名文件 | 添加日期、版本号、自定义字符，灵活配置 |
| 📋 规约上传 | WPS 表格与文件名比对，自动插入复选框、条件格式化 |
| 🤖 AI 智能助手 | 支持豆包 / DeepSeek / 元宝 / 硅基流动 等多平台 |

## 截图

主界面四大功能卡片 →

## 系统要求

- **操作系统**：Windows 10 / Windows 11
- **Python**：3.9 及以上（源码运行时需要）
- **依赖库**：openpyxl、requests、pywin32（源码运行时需要）

---

## 快速开始（三种安装方式）

### 方式一：下载 EXE 直接使用（推荐普通用户）

1. 前往 [GitHub Releases](https://github.com/your-repo/download-optimizer/releases)
2. 下载最新版本的 `下载流程优化工具_v3.5.exe`
3. 双击 EXE 即可运行，无需安装 Python

### 方式二：源码本地运行（开发者）

```bash
# 1. 克隆仓库
git clone https://github.com/your-repo/download-optimizer.git
cd download-optimizer

# 2. 创建虚拟环境
python -m venv venv
venv\Scripts\activate        # Windows
# source venv/bin/activate   # macOS/Linux

# 3. 安装依赖
pip install -r requirements.txt

# 4. 运行
python 下载流程优化工具_v3.5.py
```

### 方式三：Docker / 飞牛 NAS（环境验证）

```bash
# 克隆仓库
git clone https://github.com/your-repo/download-optimizer.git
cd download-optimizer

# Docker Compose 启动
docker-compose up -d

# 或直接使用 Docker
docker build -t download-optimizer .
docker run -it download-optimizer
```

> **⚠️ 说明**：本项目是 **Windows Tkinter GUI 桌面应用**，没有 Web 服务端口，因此 Docker 容器仅用于**环境验证和依赖检查**。实际应用请在 Windows 主机上运行 EXE 或源码。

---

## 飞牛 NAS Docker Compose 部署步骤

1. 登录飞牛 NAS 管理界面
2. 进入 **Docker** → **Compose** 或 **应用中心**
3. 点击 **添加 Compose** 或 **上传 Compose 文件**
4. 将仓库中的 `docker-compose.yml` 内容粘贴进去
5. 点击 **立即启动**，等待容器运行完成
6. 在日志中查看环境检查结果

---

## GitHub 开源操作完整指引

### 第一步：注册 GitHub 账号

1. 访问 [https://github.com](https://github.com)
2. 点击右上角 **Sign up**（注册）
3. 填写用户名、邮箱、密码（建议使用英文名，如 `your-name`）
4. 完成邮箱验证

### 第二步：创建仓库（Repository）

1. 登录后点击右上角 **+** → **New repository**
2. 填写：
   - **Repository name**：`download-optimizer`
   - **Description**：`下载流程优化工具 - 批量重命名与规约上传`
   - **Visibility**：选择 `Public`（公开）
   - **Initialize this repository with a README**：勾选
3. 点击 **Create repository**

### 第三步：上传代码到 GitHub

#### 方法 A：使用 Git 命令行（推荐开发者）

```bash
# 安装 Git（如果还没有）
# 下载地址：https://git-scm.com/download/win

# 克隆仓库到本地（将 your-username 替换为你的 GitHub 用户名）
git clone https://github.com/your-username/download-optimizer.git
cd download-optimizer

# 将项目文件复制到该目录
# 然后添加所有文件
git add .

# 提交
git commit -m "feat: 初始化下载流程优化工具 v3.5"

# 推送到 GitHub
git push origin main
```

#### 方法 B：网页直接上传（推荐新手）

1. 打开刚创建的仓库页面
2. 点击 **Add file** → **Upload files**
3. 将以下文件拖拽上传：
   - `下载流程优化工具_v3.5.py`（核心源码）
   - `requirements.txt`（依赖列表）
   - `README.md`（说明文档）
   - `Dockerfile`（Docker 配置）
   - `docker-compose.yml`（飞牛 NAS 配置）
   - `.gitignore`（忽略规则）
4. 在页面底部点击 **Commit changes**

### 第四步：发布 Release（发布可执行文件）

1. 打开仓库页面，点击 **Releases** 标签
2. 点击 **Draft a new release**
3. 填写：
   - **Tag version**：`v3.5`
   - **Release title**：`下载流程优化工具 v3.5`
   - **Description**：更新日志（见下方模板）
4. **上传 EXE 文件**：将 `dist/下载流程优化工具_v3.5.exe` 拖到附件区
5. 点击 **Publish release**

#### Release 描述模板

```markdown
## 🚀 下载流程优化工具 v3.5

### 更新内容
- ✨ 新增：规约上传模块，支持比对生成报网公司结算数据表格
- 🎨 优化：主界面图标改为橙黄色盾牌+对勾设计
- 📝 新增：规约卡片副标题「比对生成报网公司结算数据表格」
- 🐛 修复：合并单元格导致的读取错误
- 🐛 修复：条件格式化公式引用稳定性
- 📦 其他：版本号升级至 v3.5

### 系统要求
- Windows 10 / Windows 11
- 无需安装 Python，双击 EXE 即可使用

### 安装方法
1. 点击下方 Assets 下载 `下载流程优化工具_v3.5.exe`
2. 右键以管理员身份运行（可选）
3. 开始使用
```

### 第五步：配置 GitHub Pages（可选，托管文档）

1. 进入仓库 **Settings** → **Pages**
2. **Source** 选择 `main` 分支 / `root`
3. 保存后即可通过 `https://your-username.github.io/download-optimizer/` 访问

### 第六步：添加开源许可证

1. 仓库页面点击 **Add file** → **Create new file**
2. 文件名为 `LICENSE`
3. 内容使用 MIT 许可证（见下方模板）
4. 提交

#### MIT 许可证模板

```
MIT License

Copyright (c) 2025 下载流程优化工具

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

---

## 项目结构

```
download-optimizer/
├── 下载流程优化工具_v3.5.py    # 主程序源码
├── requirements.txt           # Python 依赖列表
├── Dockerfile                 # Docker 构建文件
├── docker-compose.yml         # 飞牛 NAS / Docker Compose 配置
├── .gitignore                 # Git 忽略规则
├── LICENSE                    # MIT 许可证
└── README.md                  # 项目说明文档
```

---

## 开发指南

### 本地开发

```bash
# 克隆项目
git clone https://github.com/your-username/download-optimizer.git

# 安装依赖
pip install -r requirements.txt

# 运行
python 下载流程优化工具_v3.5.py
```

### 打包为 EXE

```bash
# 安装 PyInstaller
pip install pyinstaller

# 打包（在项目根目录执行）
pyinstaller --onefile \
    --name "下载流程优化工具_v3.5" \
    --hidden-import win32com \
    --hidden-import win32com.client \
    --hidden-import pythoncom \
    --hidden-import pywintypes \
    --noconsole \
    下载流程优化工具_v3.5.py
```

打包完成后，EXE 文件位于 `dist/` 目录。

### 版本号规范

- 主版本号：重大架构变更
- 次版本号：功能新增
- 修订号：Bug 修复
- 示例：v3.5（第 3 代架构，第 5 次功能迭代）

---

## 常见问题 FAQ

### Q: 为什么我的 EXE 双击后闪退？
A: 请右键 → **以管理员身份运行**。如果仍然闪退，右键 → **发送到桌面快捷方式** → 在快捷方式上右键 → **属性** → **兼容性** → 勾选以管理员身份运行。

### Q: AI 助手模块如何配置 API Key？
A: 打开应用 → 点击 **AI 智能助手** → 在右侧输入框粘贴 API Key → 点击保存。支持豆包、DeepSeek、元宝、硅基流动等平台。

### Q: 规约上传的比对逻辑是什么？
A: 程序读取 WPS 表格，按文件名中的 `-` 分割：左边匹配表格 **S 列**，右边匹配表格 **A 列**。匹配成功的行置顶，并插入复选框和状态列。

### Q: 支持 Linux / macOS 吗？
A: 核心功能（文件操作、Excel 处理）支持跨平台，但 Tkinter GUI 和 `pywin32`（复选框插入）仅在 Windows 上可用。Linux/macOS 用户可通过 Docker 做环境验证。

---

## 技术栈

| 技术 | 用途 |
|------|------|
| Python 3.13 | 开发语言 |
| Tkinter | GUI 界面 |
| openpyxl | Excel/WPS 表格处理 |
| requests | HTTP 请求（AI API） |
| pywin32 | Windows COM 自动化（复选框） |
| PyInstaller | EXE 打包 |
| Docker | 环境验证 |

---

## 贡献指南

欢迎贡献代码！请遵循以下步骤：

1. Fork 本仓库
2. 创建特性分支：`git checkout -b feature/amazing-feature`
3. 提交变更：`git commit -m 'feat: add amazing feature'`
4. 推送到分支：`git push origin feature/amazing-feature`
5. 创建 Pull Request

---

## 更新日志（Changelog）

### v3.5 (2025-08-06)
- 新增：规约上传模块的完整功能
- 新增：比对生成报网公司结算数据表格
- 优化：UI 图标改为橙黄色盾牌+对勾
- 优化：应用标题改为「下载流程优化工具」

### v3.4 - v3.0
- 迭代修复合并单元格问题
- 迭代优化条件格式化规则
- 迭代改进文件/文件夹重命名功能

### v1.0 - v2.9
- 初始版本到功能完善的完整迭代

---

## 联系方式

如有问题或建议，请提交 [Issue](https://github.com/your-username/download-optimizer/issues)。

---

## 致谢

感谢所有为本项目做出贡献的开发者！

<p align="center">Made with ❤️ by 下载流程优化工具 Team</p>
