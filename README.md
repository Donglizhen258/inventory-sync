# ERL 库存管理系统（联网版）

一个轻量级的实验室/仓库库存管理系统，采用「客户端 + 云端服务器」架构：

- **客户端**：Windows 桌面应用，负责库存的录入、查询、下载 Excel，以及多端同步。
- **云端服务器**：负责登录鉴权、表格版本存储、多端同步与历史版本回溯，可部署在任意一台常开电脑上，并支持内网穿透供远程访问。

数据以 Excel（`.xlsx`）为存储载体，无需数据库即可离线备份与查看；服务器端使用 SQLite 保存账号与版本元数据，资源占用极低（内存 < 100MB）。

## 功能特性

- 物料编码 `MAT-xxxxxx` 永久编号，自动生成
- 六种变动类型：入库 / 领用 / 寄出 / 报废 / 退回 / 整理，正负类型匹配、整数校验、库存不为负
- 项目分 Sheet 管理（每个项目一张工作表），物料编码全局唯一
- 权限体系：管理员（录入、回溯、用户管理、项目管理、替换表格）与普通用户（只读下载最新 Excel）
- 多端同步：先改先得冲突检测，提交前比对云端版本
- 历史版本回溯：每次提交自动留存版本，保留最近 14 天，可下载/恢复任意版本
- 数据备份与恢复：一键打包全部数据（含账号与配置），从备份 zip 整体迁移（岗位交接）
- 普通用户密码对称加密存储，管理员可查看明文
- 内网穿透：内置 cpolar / cloudflared 支持，远程同事可访问
- 断网自动检测与重连、下载自动保存、跨进程互斥锁、临时文件原子替换、自动备份保留 20 份

## 系统架构

```
┌─────────────────┐        HTTP / JSON         ┌──────────────────────┐
│   客户端 (client) │ ──────────────────────────▶ │  云端服务器 (server)   │
│  inventory_web   │  登录 / 同步 / 上传 / 回溯   │  FastAPI (main.py)     │
│  (本地HTTP服务)   │ ◀────────────────────────── │  SQLite + 文件版本存储  │
└─────────────────┘                             │  控制台 (server_console)│
                                                └──────────────────────┘
```

- 客户端：`inventory_web.py` 启动本地 HTTP 服务（127.0.0.1 随机端口），前端为原生 HTML/CSS/JS（`web/`），通过 8 个 JSON 接口与核心逻辑 `inventory_core.py` 交互。
- 云端服务器：`main.py`（FastAPI + SQLite + 文件存储，端口 18666）+ `server_console.py`（控制台，端口 18667，含穿透/备份/用户管理等）。
- 客户端同步：`sync_client.py`（标准库），通过 `sync.json` 保存 token/角色/基线版本。

## 目录结构

```
├── client/                      # 桌面客户端源码
│   ├── inventory_web.py         # 入口：本地 HTTP 服务与接口
│   ├── inventory_core.py        # 库存/编码/校验/迁移/备份核心逻辑
│   ├── sync_client.py           # 与云端同步的客户端
│   ├── gui_window.py / tray.py  # 窗口与托盘
│   ├── web/                     # 前端（index.html / app.js / styles.css）
│   └── test_*.py                # 自动化测试（与核心逻辑同目录）
├── server/                      # 云端服务器源码
│   ├── main.py                  # FastAPI 服务端（登录/版本/用户/同步）
│   ├── server_console.py        # 服务器控制台（穿透/备份/用户管理）
│   ├── console.html             # 控制台前端
│   ├── launcher.py              # 命令行启动/停止助手
│   └── gui_window.py / tray.py
├── requirements.txt             # 依赖
├── LICENSE                      # MIT 许可证
└── README.md
```

## 快速开始

### 环境要求

- Python 3.12+（64 位）
- Windows（客户端为 Windows 桌面应用；服务器端可运行于任意支持 Python 的系统）

### 安装依赖

```bash
pip install -r requirements.txt
```

### 运行云端服务器

```bash
cd server
python main.py          # 云端服务（18666）
python server_console.py  # 服务器控制台（18667，含穿透/备份/用户管理）
```

首次运行会自动生成 `server/data/` 与 `config.json`（含默认管理员密码 `admin123`，部署后请务必修改）。

### 运行客户端

```bash
cd client
python inventory_web.py
```

启动后按提示打开本机地址，登录即可使用（需先运行云端服务器）。

### 打包为 EXE（可选）

```bash
pip install pyinstaller
pyinstaller --noconfirm --onefile --windowed inventory_web.py
```

## 运行说明（重要）

- **首次运行自动初始化**：客户端首次启动会自动生成一个空的库存工作簿
  （含一个"未分类"项目表 + 隐藏的变更流水表），无需任何模板文件，数据存在
  系统用户目录 `%LOCALAPPDATA%\ERL库存管理\` 下。
- **独立窗口可选**：客户端/控制台使用 `pywebview` 弹出独立窗口（含系统托盘）。
  若未安装 `pywebview` 或缺少 WebView2 运行库，程序会自动降级为用系统默认
  浏览器打开界面，核心功能不受影响。
- **内网穿透需自行下载二进制**：服务器控制台的"一键穿透"依赖 `cpolar.exe` 和
  `cloudflared.exe`（第三方工具，不随源码分发）。如需使用，请分别放到
  `server/cpolar/cpolar.exe` 和 `server/cloudflared/cloudflared.exe`；
  不配置穿透也能在同一局域网内正常使用。

## 测试

```bash
cd client
python test_inventory_v3.py
python test_migrate_multi_sheet.py
```

> 说明：`test_inventory_v3.py` 与 `test_migrate_multi_sheet.py` 是深度回归测试，
> 需要用到 `assets/实验室库存管理模板.xlsx` 作为数据样本。该样本包含生产环境的
> 真实数据，出于隐私未随源码发布。若需运行完整测试，请自行准备一份结构一致的
> `.xlsx` 样本放到 `client/assets/` 目录下（软件首次运行会自动生成一份空工作簿，
> 可作为基础样本）。核心功能运行不依赖测试与样本。

## 技术栈

- 后端：Python 标准库（http.server / urllib）+ FastAPI + SQLite + openpyxl
- 前端：原生 HTML / CSS / JavaScript
- 打包：PyInstaller
- 内网穿透：cpolar / cloudflared

## 许可证

本项目采用 [MIT License](LICENSE)。

## 作者

董理臻 — 2026 年 8 月完成
联系方式：微信 18551780019 / QQ 1301535058
