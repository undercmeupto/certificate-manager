# 证件管理系统 / Certificate Management System

工业风格Web版证件管理系统，支持Excel上传、CRUD操作、搜索过滤和Excel导出。

## 功能特性 / Features

- **Excel导入** - 支持简单格式和复杂格式（人员证件台帐）Excel文件
- **数据管理** - 完整的CRUD操作（创建、读取、更新、删除）
- **搜索过滤** - 按姓名、部门、证件名称搜索，按状态筛选
- **Excel导出** - 带状态颜色标记的Excel导出
- **状态预警** - 自动计算剩余天数，标记过期/紧急/预警/正常状态
- **工业风格UI** - 高对比度、清晰的工业界面设计

## 安装 / Installation

```bash
# 进入项目目录
cd C:\Users\HSE\certificate-manager-web

# 安装依赖
pip install -r requirements.txt
```

## 运行 / Run

```bash
python app.py
```

访问: http://localhost:5000

## 项目结构 / Structure

```
certificate-manager-web/
├── app.py                           # Flask主应用
├── config.py                        # 配置文件
├── utils/
│   └── certificate_checker.py      # 核心解析逻辑
├── static/
│   └── css/
│       └── industrial.css          # 工业风格样式
├── templates/
│   └── index.html                  # 单页应用
├── data/
│   └── certificates.json           # 数据存储
├── uploads/                        # 临时上传目录
├── exports/                        # 导出文件目录
├── requirements.txt
└── README.md
```

## API接口 / API Endpoints

| 端点 | 方法 | 描述 |
|------|------|------|
| `/` | GET | 主页 |
| `/api/data` | GET | 获取所有证件 |
| `/api/data` | POST | 添加证件 |
| `/api/data/<id>` | PUT | 更新证件 |
| `/api/data/<id>` | DELETE | 删除证件 |
| `/api/upload` | POST | 上传Excel |
| `/api/search` | POST | 搜索证件 |
| `/api/export` | GET | 导出Excel |
| `/api/statistics` | GET | 获取统计 |

## 状态说明 / Status

| 剩余天数 | 状态 | 图标 | 颜色 |
|----------|------|------|------|
| < 0 | 已过期 | X | 红色 |
| < 30 | 紧急 | ! | 红色 |
| 30-90 | 预警 | ? | 橙色 |
| > 90 | 正常 | OK | 绿色 |

## Excel格式要求 / Excel Format

### 简单格式（每行一个证件）

必需列: 姓名, 部门, 证件名称, 证件号码, 到期日期, 邮箱, 手机号

### 复杂格式（人员证件台帐）

自动识别，支持多证件列格式。
