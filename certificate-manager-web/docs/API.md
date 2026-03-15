# Certificate Manager API 文档

## 概述

证件管理系统 RESTful API 文档

**Base URL:** `http://localhost:5000`
**Content-Type:** `application/json`
**编码:** `UTF-8`

---

## 目录

- [会话管理](#会话管理)
- [数据操作](#数据操作)
- [Excel导入导出](#excel导入导出)
- [搜索筛选](#搜索筛选)
- [统计信息](#统计信息)
- [错误码](#错误码)

---

## 会话管理

### 获取会话状态

检查当前会话是否激活（是否已上传Excel文件）

**请求**
```http
GET /api/session/status
```

**响应**
```json
{
  "success": true,
  "active": true,
  "has_data": true
}
```

| 字段 | 类型 | 说明 |
|------|------|------|
| `active` | boolean | 会话是否激活 |
| `has_data` | boolean | 是否有数据 |

### 退出会话

清除所有数据并锁定系统，需重新上传文件

**请求**
```http
POST /api/exit
```

**响应**
```json
{
  "success": true,
  "message": "会话已退出，数据已清除，请上传新文件继续操作"
}
```

---

## 数据操作

### 获取所有证件

**请求**
```http
GET /api/data
```

**响应**
```json
{
  "success": true,
  "data": [
    {
      "id": "ac584e57-7945-4837-a5ee-24b2ac3ee1a8",
      "name": "张三",
      "department": "技术部",
      "position": "工程师",
      "certificate_name": "HSE证(H2S)",
      "certificate_number": "HSE2024001",
      "issue_date": "2024-01-15",
      "expiry_date": "2026-01-15",
      "email": "zhangsan@example.com",
      "phone": "13800138000",
      "days_remaining": 305,
      "status": "normal",
      "status_label": "正常",
      "status_icon": "OK",
      "status_color": "#27AE60",
      "created_at": "2026-03-15T13:11:38.228050"
    }
  ],
  "statistics": {
    "total": 172,
    "expired": 2,
    "urgent": 5,
    "warning": 15,
    "normal": 150
  }
}
```

### 获取单个证件

**请求**
```http
GET /api/data/{cert_id}
```

**路径参数**
| 参数 | 类型 | 说明 |
|------|------|------|
| `cert_id` | string (UUID) | 证件ID |

**响应**
```json
{
  "success": true,
  "data": {
    "id": "ac584e57-7945-4837-a5ee-24b2ac3ee1a8",
    "name": "张三",
    ...
  }
}
```

### 添加证件

⚠️ **需要会话激活**

**请求**
```http
POST /api/data
Content-Type: application/json

{
  "name": "张三",
  "department": "技术部",
  "position": "工程师",
  "certificate_name": "HSE证(H2S)",
  "certificate_number": "HSE2024001",
  "issue_date": "2024-01-15",
  "expiry_date": "2026-01-15",
  "email": "zhangsan@example.com",
  "phone": "13800138000"
}
```

**必填字段**
| 字段 | 类型 | 说明 |
|------|------|------|
| `name` | string | 姓名 |
| `department` | string | 部门 |
| `position` | string | 岗位 |
| `certificate_name` | string | 证件名称 |
| `issue_date` | string | 发证日期 (YYYY-MM-DD) |
| `expiry_date` | string | 到期日期 (YYYY-MM-DD) |

**可选字段**
| 字段 | 类型 | 说明 |
|------|------|------|
| `certificate_number` | string | 证件号码 |
| `email` | string | 邮箱 |
| `phone` | string | 手机号 |

**响应**
```json
{
  "success": true,
  "data": {
    "id": "ac584e57-7945-4837-a5ee-24b2ac3ee1a8",
    "name": "张三",
    ...
  }
}
```

### 更新证件

⚠️ **需要会话激活**

**请求**
```http
PUT /api/data/{cert_id}
Content-Type: application/json

{
  "name": "张三",
  "department": "技术部",
  "position": "高级工程师",
  "certificate_name": "HSE证(H2S)",
  "expiry_date": "2026-06-15"
}
```

**说明**
- 只需提供要更新的字段
- 系统会自动重新计算状态

**响应**
```json
{
  "success": true,
  "data": {
    "id": "ac584e57-7945-4837-a5ee-24b2ac3ee1a8",
    "position": "高级工程师",
    "expiry_date": "2026-06-15",
    "days_remaining": 457,
    "status": "normal"
  }
}
```

### 删除证件

⚠️ **需要会话激活**

**请求**
```http
DELETE /api/data/{cert_id}
```

**响应**
```json
{
  "success": true,
  "data": {
    "id": "ac584e57-7945-4837-a5ee-24b2ac3ee1a8",
    "name": "张三",
    "certificate_name": "HSE证(H2S)"
  }
}
```

---

## Excel导入导出

### 获取Sheet列表

获取Excel文件中的所有工作表名称

**请求**
```http
POST /api/sheets
Content-Type: multipart/form-data

file: <Excel文件>
```

**响应**
```json
{
  "success": true,
  "sheets": ["Sheet1", "人员证件台帐", "2024年数据"]
}
```

### 上传Excel

⚠️ **会清除现有数据**

**请求**
```http
POST /api/upload
Content-Type: multipart/form-data

file: <Excel文件>
sheet: Sheet1  // 可选，默认第一个
```

**文件限制**
- 格式：`.xlsx`, `.xls`
- 大小：最大 16MB

**响应**
```json
{
  "success": true,
  "message": "成功导入 172 条记录（已清除旧数据）",
  "imported": 172,
  "statistics": {
    "total": 172,
    "expired": 2,
    "urgent": 5,
    "warning": 15,
    "normal": 150
  },
  "has_original": true
}
```

### 导出全部数据

**请求**
```http
GET /api/export
```

**响应**
- Content-Type: `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
- 文件名: `certificates_export_20260315_143020.xlsx`

### 导出预警数据

仅导出已过期、紧急、预警状态的证件

**请求**
```http
GET /api/export/warning
```

**响应**
- Content-Type: `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
- 文件名: `certificates_warning_20260315_143020.xlsx`

### 导出原表更新

保留原始Excel格式，使用当前数据更新

**请求**
```http
GET /api/export/original-updated
```

**响应**
- Content-Type: `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
- 文件名: `updated_原表名_20260315_143020.xlsx`

---

## 搜索筛选

### 搜索证件

**请求**
```http
POST /api/search
Content-Type: application/json

{
  "search": "张三",
  "status": "expired"
}
```

**参数**
| 字段 | 类型 | 说明 |
|------|------|------|
| `search` | string | 搜索关键词（姓名/部门/岗位/证件名称/证件号码） |
| `status` | string | 状态筛选：`expired`/`urgent`/`warning`/`normal` |

**响应**
```json
{
  "success": true,
  "data": [
    {
      "id": "ac584e57-7945-4837-a5ee-24b2ac3ee1a8",
      "name": "张三",
      ...
    }
  ],
  "count": 3
}
```

---

## 统计信息

### 获取统计数据

**请求**
```http
GET /api/statistics
```

**响应**
```json
{
  "success": true,
  "data": {
    "total": 172,
    "expired": 2,
    "urgent": 5,
    "warning": 15,
    "normal": 150
  }
}
```

---

## 错误码

### HTTP状态码

| 状态码 | 说明 |
|--------|------|
| 200 | 请求成功 |
| 400 | 请求参数错误 |
| 403 | 会话未激活（需要先上传Excel） |
| 404 | 资源不存在 |
| 413 | 文件大小超过限制 |
| 500 | 服务器内部错误 |

### 错误响应格式

```json
{
  "success": false,
  "error": "错误描述信息"
}
```

### 常见错误信息

| 错误信息 | 说明 |
|----------|------|
| `会话已退出，请先上传Excel文件` | 需要先上传Excel激活会话 |
| `证件不存在` | 指定的证件ID不存在 |
| `缺少必填字段: {field}` | 请求缺少必填参数 |
| `不支持的文件格式` | 上传的文件格式不支持 |
| `文件大小不能超过16MB` | 文件超过大小限制 |
| `没有数据可导出` | 当前没有数据可导出 |

---

## 状态说明

### 证件状态计算规则

| 剩余天数 | 状态 | 标签 | 图标 | 颜色 |
|----------|------|------|------|------|
| < 0 | `expired` | 已过期 | X | #E74C3C (红色) |
| < 30 | `urgent` | 紧急 | ! | #F39C12 (橙色) |
| 30-90 | `warning` | 预警 | ? | #F1C40F (黄色) |
| > 90 | `normal` | 正常 | OK | #27AE60 (绿色) |

### 配置参数

| 参数 | 默认值 | 说明 |
|------|--------|------|
| `URGENT_DAYS` | 30 | 紧急天数阈值 |
| `WARNING_DAYS` | 90 | 预警天数阈值 |

---

## Excel格式说明

### 简单格式（每行一个证件）

| 列名 | 必填 | 说明 |
|------|------|------|
| 姓名 | ✅ | 员工姓名 |
| 部门 | ✅ | 所属部门 |
| 岗位 | ✅ | 职位岗位 |
| 证件名称 | ✅ | 证件类型 |
| 证件号码 | ❌ | 证件编号 |
| 到期日期 | ✅ | 格式：YYYY-MM-DD |
| 发证日期 | ❌ | 格式：YYYY-MM-DD |
| 邮箱 | ❌ | 电子邮箱 |
| 手机号 | ❌ | 联系电话 |

### 复杂格式（人员证件台帐）

自动检测格式，支持多证件列：
- IADC/IWCF
- HSE证(H2S)
- 急救证
- 消防证
- 司索指挥证
- 防恐证
- 健康证
- 岗位证

数据从第4行开始（前3行为标题行）

---

## 示例代码

### Python 示例

```python
import requests

BASE_URL = "http://localhost:5000"

# 获取所有证件
response = requests.get(f"{BASE_URL}/api/data")
data = response.json()

# 添加新证件
new_cert = {
    "name": "张三",
    "department": "技术部",
    "position": "工程师",
    "certificate_name": "HSE证(H2S)",
    "issue_date": "2024-01-15",
    "expiry_date": "2026-01-15"
}
response = requests.post(f"{BASE_URL}/api/data", json=new_cert)

# 上传Excel
files = {'file': open('data.xlsx', 'rb')}
data = {'sheet': 'Sheet1'}
response = requests.post(f"{BASE_URL}/api/upload", files=files, data=data)
```

### JavaScript 示例

```javascript
const BASE_URL = 'http://localhost:5000';

// 获取所有证件
async function getCertificates() {
  const response = await fetch(`${BASE_URL}/api/data`);
  return await response.json();
}

// 添加新证件
async function addCertificate(cert) {
  const response = await fetch(`${BASE_URL}/api/data`, {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify(cert)
  });
  return await response.json();
}

// 删除证件
async function deleteCertificate(id) {
  const response = await fetch(`${BASE_URL}/api/data/${id}`, {
    method: 'DELETE'
  });
  return await response.json();
}
```

---

## 更新日志

### v1.1.0 (2026-03-15)
- ✅ 修复删除功能竞态条件
- ✅ 添加批量删除功能
- ✅ 改进ID验证机制
- ✅ 优化代码质量

### v1.0.0 (2026-03-13)
- ✅ 基础CRUD功能
- ✅ Excel导入导出
- ✅ 状态自动计算
- ✅ 搜索筛选功能
