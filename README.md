# 下载流程优化工具

> 一个用于批量重命名文件夹/文件、规约上传数据表准备和 AI 智能助手的 Windows 桌面 GUI 应用。

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.9+-blue.svg)](https://www.python.org/)
[![Version](https://img.shields.io/badge/version-v3.11-green.svg)](https://github.com/szboboxing/download-optimizer/releases)

---

## 功能特点

| 模块 | 说明 |
|------|------|
| 📁 批量重命名文件夹 | 删除前 N 位字符 / 前加内容 / 后加内容，支持预览和一键回退 |
| 📄 批量重命名文件 | 添加日期、版本号、自定义字符，灵活配置 |
| 📋 规约上传数据表准备 | 自选比对列，复用现有 K/L 状态列并识别重复数据 |
| 🤖 AI 智能助手 | 支持豆包 / DeepSeek / 元宝 / 硅基流动 等多平台 |

## 核心功能详解

### 批量重命名文件夹
- **删除前 N 位字符**：可选 1-8 位，默认 3 位
- **前加内容**：在文件夹名前追加指定字符
- **后加内容**：在文件夹名后追加指定字符
- **预览功能**：实时显示将被重命名的文件夹列表
- **回退功能**：后悔时一键撤销上一次操作

### 批量重命名文件
- **添加当前日期**：支持多种日期格式
- **批量添加版本号**：自动递增版本号
- **添加自定义字符**：指定字符和数量

### 规约上传数据表准备
- 复制 WPS 表格并打开修改
- 文件名横杠左边默认比对表格 S 列，横杠右边默认比对表格 A 列
- 两个比对列均为可编辑选择框，可直接填写 W 等 A 至 XFD 范围内的列
- 左右任一侧内容匹配即视为匹配，原有生成规则保持不变
- 匹配行置顶到第 5 行开始
- 直接使用现有 K 列“复核”和 L 列“需上传”，不再插入新列
- 新匹配行在现有 K/L 中生成复选框和状态公式
- 已有状态的匹配行保留 K 列原值，并将 L 列改为“待上传（重复）”
- 本次未匹配的旧状态保持不变，行重排后自动同步复选框位置
- 生成完成后自动隐藏并折叠 K、L 列，状态数据和复选框仍保留
- 条件格式化：勾选后行背景 `#EBF1DE`、文字 `#00B050` 绿色加粗
- 术语与状态逻辑：[K/L 上传状态联动说明（Word）](./K-L上传状态联动说明.docx)

### AI 智能助手
- 支持主流 AI 平台：豆包、DeepSeek、元宝、硅基流动
- API Key 本地加密存储
- 多轮对话、上下文记忆

## 系统要求

- **操作系统**：Windows 10 / Windows 11
- **Python**：3.9+（源码运行时需要）
- **依赖库**：openpyxl、requests、pywin32

---

## 快速开始

### 方式一：下载 EXE（推荐普通用户）

前往 [GitHub Releases](https://github.com/szboboxing/download-optimizer/releases) 下载 `download-optimizer-v3.11.exe`，双击即可运行。

### 方式二：源码运行（开发者）

```bash
git clone https://github.com/szboboxing/download-optimizer.git
cd download-optimizer
python -m venv venv
venv\Scripts\activate              # Windows
pip install -r requirements.txt
python main.py
```

### 方式三：Docker / 飞牛 NAS

```bash
git clone https://github.com/szboboxing/download-optimizer.git
cd download-optimizer
docker-compose up -d
```

> **注意**：本项目为 Windows Tkinter GUI 桌面应用，Docker 仅用于环境验证。实际应用请在 Windows 主机上运行。

---

## 项目结构

```
download-optimizer/
├── main.py                 # 启动入口
├── app.py                  # 主程序（GUI + 核心逻辑）
├── requirements.txt        # Python 依赖
├── Dockerfile              # Docker 构建
├── docker-compose.yml      # 飞牛 NAS / Docker Compose
├── .gitignore              # Git 忽略规则
├── LICENSE                 # MIT 许可证
├── README.md               # 本文件
└── GITHUB_GUIDE.md         # GitHub 开源操作指引
```

## 开发指南

### 打包为 EXE

```bash
pip install pyinstaller
pyinstaller --onefile --name "下载流程优化工具_v3.11" ^
    --hidden-import win32com --hidden-import win32com.client ^
    --hidden-import pythoncom --hidden-import pywintypes ^
    --noconsole main.py
```

---

## 更新日志

### v3.11 (2026-08-07)
- 🧱 规约上传直接复用现有 K/L 状态列，不再重复增列
- ♻️ 保留已有 K 列状态及本次未匹配行的原有 L 列内容
- 🔁 已标注数据再次匹配时，将 L 列改为“待上传（重复）”
- ☑️ 行置顶后重建并同步全部 K 列复选框，保留原勾选状态
- 🛡️ 新增 K/L 表头安全校验，避免覆盖普通业务列
- 📄 随版本提供《K/L 上传状态联动说明》Word 文档
- 📦 构建 Windows 单文件 EXE

### v3.10 (2026-08-07)
- 🧹 移除已取消的双表状态同步功能及首页入口
- ✏️ 规约上传新增两个可编辑比对列选择框，默认保持横杠左边对应 S 列、右边对应 A 列
- 🔎 支持直接输入 W 等 A 至 XFD 范围内的列，左右任一侧匹配规则保持不变
- 🙈 生成完成后自动隐藏并折叠 K、L 列
- ✅ 新增 K/L 隐藏状态校验，并确认复选框、公式和前四行表头格式完整保留
- 📦 构建 Windows 单文件 EXE

### v3.7 - v3.9 (2026-08-06)
- 历史过渡版本；相关试验功能已在 v3.10 中移除

### v3.6 (2026-08-06)
- 📝 「规约上传」更名为「规约上传数据表准备」
- 🐛 修复生成文件前四行表头合并区域和样式错乱
- 🔧 比对与重排严格从第 5 行开始，完整保留表头行高和原列宽
- 📦 重新构建 Windows 单文件 EXE

### v3.5 (2025-08-06)
- 🎨 应用标题改为「下载流程优化工具」
- 🎨 规约上传图标改为橙黄色盾牌+对勾设计
- 📝 规约卡片新增副标题「比对生成报网公司结算数据表格」
- 🔧 规约上传：比对行从第 5 行开始置顶
- 🔧 规约上传：条件格式化颜色优化（行背景 #EBF1DE、文字 #00B050）
- 🐛 修复合并单元格导致的读取错误
- 🐛 修复启动时 icon_widget 未初始化错误
- 📦 开源发布：Docker、docker-compose、飞牛 NAS 全支持

### v3.0 - v3.4
- 规约上传模块完整实现（复选框、条件格式化、行重排序）
- UI 美化、统一 ttk 样式
- 文件夹/文件重命名功能完善

### v1.0 - v2.9
- 初始版本到功能完善的完整迭代（批量重命名核心功能、回退机制、AI 助手集成等）

---

## 技术栈

| 技术 | 用途 |
|------|------|
| Python 3.13 | 开发语言 |
| Tkinter | GUI 界面 |
| openpyxl | Excel/WPS 表格处理 |
| requests | HTTP 请求（AI API） |
| pywin32 | Windows COM 自动化 |
| PyInstaller | EXE 打包 |
| Docker | 环境验证 |

---

## 常见问题

**Q: EXE 双击闪退？** → 右键以管理员身份运行。

**Q: AI 助手如何配置？** → 点击 AI 智能助手 → 粘贴 API Key → 保存。

**Q: 支持 Linux/macOS 吗？** → 核心功能跨平台，GUI 和 COM 自动化仅 Windows 可用。

**Q: 飞牛 NAS 如何部署？** → 将 `docker-compose.yml` 粘贴到飞牛 Docker Compose 中启动。

---

## 贡献

欢迎贡献代码！Fork → Branch → PR。

## 许可证

[MIT License](LICENSE) © 2025 下载流程优化工具 Team
