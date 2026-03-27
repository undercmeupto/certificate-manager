# 证件管理系统 / Certificate Management System

基于 Flask + SQLite + 单页应用的证件管理系统，支持 Excel 导入导出、状态预警、会话锁定等企业级功能。

![License](https://img.shields.io/badge/license-MIT-blue)
![Python](https://img.shields.io/badge/python-3.7+-green)
![Flask](https://img.shields.io/badge/flask-2.0+-orange)
![SQLite](https://img.shields.io/badge/sqlite-3-blue)

## 功能特性 / Features

| 功能 | 说明 |
|------|------|
| **Excel 导入** | 支持简单格式（每行一证）和复杂格式（人员证件台帐） |
| **数据管理** | 完整的 CRUD 操作，单条添加、批量导入、编辑、删除 |
| **状态预警** | 自动计算剩余天数，四档状态分类 |
| **搜索过滤** | 按姓名/部门/证件名称搜索，按状态筛选 |
| **Excel 导出** | 全部导出 / 预警导出 / 按天数导出 / 原表更新导出 |
| **会话锁定** | 退出后清空数据并锁定系统，需重新上传文件 |
| **浅蓝色主题** | 清爽现代的 UI 设计，响应式布局 |
| **SQLite 数据库** | 使用 SQLite 数据库存储，支持索引和高效查询 |
| **打包支持** | 支持 PyInstaller 打包为独立 exe 程序 |

## 截图 / Screenshots

```
┌─────────────────────────────────────────────────────────────┐
│                    证件管理系统                              │
├─────────────────────────────────────────────────────────────┤
│ [上传] [添加] [删除] [导出全部] [导出预警] [按天数导出▼] [退出] │
├─────────────────────────────────────────────────────────────┤
│ 总计: 150  │ 已过期: 5  │ 紧急: 12  │ 预警: 28  │ 正常: 105 │
├─────────────────────────────────────────────────────────────┤
│ 搜索: [_______________] 状态: [全部▼] [搜索] [清除]          │
├─────────────────────────────────────────────────────────────┤
│ ☑ │ 姓名  │ 部门 │ 岗位 │ 证件名称 │ 到期日期 │ 剩余天数 │ 状态 │
├─────────────────────────────────────────────────────────────┤
│ ☑ │ 张三  │ HSE  │ 安全员│ 安全证   │ 2025-06-01│  45  │ ?预警 │
│ ☑ │ 李四  │ 钻井  │ 司钻 │ IADC    │ 2025-02-15│  -8  │ X过期 │
└─────────────────────────────────────────────────────────────┘
```

## 快速开始 / Quick Start

### 1. 安装依赖

```bash
# 克隆仓库
git clone https://github.com/undercmeupto/certificate-manager.git
cd certificate-manager/certificate-manager-web

# 安装依赖
pip install -r requirements.txt
```

### 2. 运行应用

```bash
python app.py
```

访问：http://localhost:5000

### 3. 使用流程

```
上传 Excel → 数据自动解析 → 状态自动计算 → 管理操作 → 导出/退出
```

## 项目结构 / Structure

```
certificate-manager-web/
├── app.py                         # Flask 主应用
├── config.py                      # 配置文件
├── models.py                      # SQLAlchemy ORM 模型
├── database.py                    # 数据库连接和会话管理
├── requirements.txt               # Python 依赖
├── utils/
│   ├── __init__.py
│   └── certificate_checker.py    # Excel 解析核心
├── migrations/
│   └── migrate_json_to_sqlite.py # JSON 到 SQLite 迁移脚本
├── static/
│   ├── favicon.png                # 网站图标
│   └── css/
│       └── industrial.css         # 浅蓝色主题样式
├── templates/
│   └── index.html                 # 单页应用
├── data/
│   └── certificates.json          # 旧版 JSON 存储（已弃用）
│   └── session_state.json         # 旧版会话状态（已弃用）
├── uploads/                       # 上传临时文件
├── exports/                       # 导出文件
├── certificates.db                # SQLite 数据库（自动生成）
└── README.md
```

## API 接口 / API Endpoints

| 端点 | 方法 | 描述 |
|------|------|------|
| `/` | GET | 主页 |
| `/api/data` | GET | 获取所有证件 |
| `/api/data/<id>` | GET | 获取单个证件 |
| `/api/data` | POST | 添加证件 |
| `/api/data/<id>` | PUT | 更新证件 |
| `/api/data/<id>` | DELETE | 删除证件 |
| `/api/upload` | POST | 上传 Excel（激活会话） |
| `/api/sheets` | POST | 获取 Excel 表格列表 |
| `/api/search` | POST | 搜索过滤证件 |
| `/api/export` | GET | 导出全部 Excel |
| `/api/export/warning` | GET | 导出预警 Excel |
| `/api/export/by-days` | GET | 按天数导出 Excel |
| `/api/export/original-updated` | GET | 导出原表更新 |
| `/api/statistics` | GET | 获取统计数据 |
| `/api/session/status` | GET | 获取会话状态 |
| `/api/exit` | POST | 退出会话（清空数据+锁定） |

## 状态说明 / Status

| 剩余天数 | 状态 | 图标 | 颜色 | 操作 |
|----------|------|------|------|------|
| < 0 | 已过期 | X | 🔴 红色 | 立即处理 |
| < 30 | 紧急 | ! | 🟠 橙色 | 尽快换证 |
| 30-90 | 预警 | ? | 🟡 黄色 | 提前准备 |
| ≥ 90 | 正常 | OK | 🟢 绿色 | 无需操作 |

## Excel 格式要求 / Excel Format

### 简单格式（每行一个证件）

| 姓名 | 部门 | 岗位 | 证件名称 | 证件号码 | 取证日期 | 到期日期 | 邮箱 | 手机号 |
|------|------|------|----------|----------|----------|----------|------|--------|
| 张三 | HSE | 安全员 | 安全证 | A001 | 2024-01-01 | 2025-01-01 | ... | ... |

### 复杂格式（人员证件台帐）

系统自动识别，支持以下证件类型：
- IADC/IWCF
- HSE证(H2S)
- 急救证
- 消防证
- 司索指挥证
- 防恐证
- 健康证
- 岗位证

## 会话锁定机制 / Session Lock

```
┌─────────────────────────────────────────────────────────┐
│  上传 Excel  │  激活会话  │  可执行所有操作              │
└─────────────────────────────────────────────────────────┘
                           │
                           ▼ 点击退出
┌─────────────────────────────────────────────────────────┐
│  清空数据    │  停用会话  │  系统锁定                    │
│  禁用添加/编辑/删除                                   │
│  只允许上传新文件解锁                                   │
└─────────────────────────────────────────────────────────┘
```

**退出后尝试操作会返回：**
> 会话已退出，请先上传 Excel 文件

## 配置 / Configuration

编辑 `config.py` 可调整：

```python
# 阈值设置
URGENT_DAYS = 30      # 紧急天数
WARNING_DAYS = 90     # 预警天数

# 服务器配置
HOST = '0.0.0.0'      # 监听地址
PORT = 5000           # 端口号
DEBUG = True          # 调试模式

# 文件上传
MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB

# 数据库配置
DATABASE_FILE = 'certificates.db'  # SQLite 数据库文件路径
```

## 数据库 / Database

系统使用 SQLite 数据存储数据，具有以下优势：

- **零配置** - 无需独立数据库服务器
- **高性能** - 支持索引，查询速度快
- **可移植** - 单文件存储，易于备份和迁移
- **可扩展** - 后期可轻松迁移到 PostgreSQL/MySQL

### 数据库结构

| 表名 | 说明 |
|------|------|
| `certificates` | 证件数据表 |
| `upload_metadata` | 上传元数据表 |
| `session_state` | 会话状态表 |

### 数据迁移

如果你有旧版本的 JSON 数据，可以使用迁移脚本：

```bash
python migrations/migrate_json_to_sqlite.py
```

迁移脚本会：
1. 自动备份原始 JSON 文件
2. 创建 SQLite 数据库
3. 迁移所有数据到新数据库
4. 验证数据完整性

## 打包为 EXE / Build Executable

使用 PyInstaller 将应用打包为独立可执行文件：

```bash
# 安装 PyInstaller
pip install pyinstaller

# 打包（单文件模式）
pyinstaller --onefile --windowed --add-data "templates;templates" --add-data "static;static" app.py

# 或使用 spec 文件（更灵活）
pyinstaller certificate-manager.spec
```

打包后：
- 可执行文件位于 `dist/` 目录
- 数据库文件存储在用户目录 `~/CertificateManager/certificates.db`
- 启动时自动打开浏览器访问 http://localhost:5000

## 技术栈 / Tech Stack

| 后端 | 前端 | 数据处理 | 数据库 |
|------|------|----------|--------|
| Flask 2.0+ | Vanilla JS | pandas | SQLite 3 |
| Flask-CORS | CSS3 | openpyxl | SQLAlchemy 2.0+ |
| Flask-SQLAlchemy | HTML5 | pdf2image | |
| Werkzeug | | | |

## 常见问题 / FAQ

**Q: 退出后数据还在吗？**
A: 退出会清空服务器数据，需重新上传 Excel 文件。

**Q: 支持多用户同时使用吗？**
A: 当前版本为单机版，数据存储在本地 SQLite 数据库中。

**Q: Excel 文件大小限制？**
A: 最大 16MB。

**Q: 如何修改状态阈值？**
A: 编辑 `config.py` 中的 `URGENT_DAYS` 和 `WARNING_DAYS`。

**Q: 按天数导出有什么用？**
A: 可以导出指定天数内到期的人员清单，例如"30天到期"、"60天到期"等，方便提前安排换证。

**Q: 导出的 Excel 文件格式？**
A: 导出文件包含标题行（含查询日期）、表头行、数据行和说明行，并根据状态自动添加颜色标记。

**Q: 如何备份数据？**
A: 直接复制 `certificates.db` 数据库文件即可。也可以使用迁移脚本导出为 JSON 格式。

**Q: 从旧版本 JSON 数据如何迁移？**
A: 运行 `python migrations/migrate_json_to_sqlite.py` 即可自动迁移，原始 JSON 文件会自动备份。

## 许可证 / License

MIT License

## 作者 / Author

[undercmeupto](https://github.com/undercmeupto)

---

⭐ 如果这个项目对您有帮助，请给个 Star！
