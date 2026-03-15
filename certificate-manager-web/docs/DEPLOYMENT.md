# 证件管理系统部署文档

## 目录

- [环境要求](#环境要求)
- [本地开发部署](#本地开发部署)
- [生产环境部署](#生产环境部署)
- [Docker部署](#docker部署)
- [云平台部署](#云平台部署)
- [常见问题](#常见问题)

---

## 环境要求

### 系统要求
- 操作系统：Windows 10+, Linux, macOS
- Python：3.8 或更高版本
- 内存：至少 512MB
- 磁盘：至少 100MB 可用空间

### Python 依赖

```
flask>=2.0.0
flask-cors>=3.0.0
pandas>=1.3.0
openpyxl>=3.0.0
python-dotenv>=0.19.0
werkzeug>=2.0.0
```

---

## 本地开发部署

### 1. 克隆项目

```bash
git clone https://github.com/undercmeupto/certificate-manager.git
cd certificate-manager/certificate-manager-web
```

### 2. 创建虚拟环境

**Windows (PowerShell)**
```powershell
python -m venv venv
.\\venv\\Scripts\\activate
```

**Windows (CMD)**
```cmd
python -m venv venv
venv\\Scripts\\activate.bat
```

**Linux/macOS**
```bash
python3 -m venv venv
source venv/bin/activate
```

### 3. 安装依赖

```bash
pip install -r requirements.txt
```

### 4. 配置环境变量（可选）

创建 `.env` 文件：

```env
# Flask 配置
FLASK_APP=app.py
FLASK_ENV=development
FLASK_DEBUG=True

# 服务器配置
HOST=0.0.0.0
PORT=5000

# 密钥（生产环境必须修改）
SECRET_KEY=your-secret-key-here-change-in-production

# 文件配置
MAX_CONTENT_LENGTH=16777216
```

### 5. 初始化目录

```bash
# 创建必要的目录
mkdir -p data uploads exports
```

### 6. 启动应用

```bash
python app.py
```

应用将在 `http://localhost:5000` 启动

---

## 生产环境部署

### 方案一：使用 Gunicorn (Linux/macOS)

#### 1. 安装 Gunicorn

```bash
pip install gunicorn
```

#### 2. 创建配置文件 `gunicorn.conf.py`

```python
import multiprocessing

# 服务器配置
bind = "0.0.0.0:5000"
workers = multiprocessing.cpu_count() * 2 + 1
threads = 2

# 超时配置
timeout = 120
keepalive = 5

# 日志配置
accesslog = "logs/access.log"
errorlog = "logs/error.log"
loglevel = "info"

# 进程命名
proc_name = "certificate-manager"

# 优雅重启
max_requests = 1000
max_requests_jitter = 100
```

#### 3. 启动服务

```bash
gunicorn -c gunicorn.conf.py app:app
```

#### 4. 使用 Systemd 管理服务

创建 `/etc/systemd/system/certificate-manager.service`:

```ini
[Unit]
Description=Certificate Manager Flask Application
After=network.target

[Service]
User=www-data
Group=www-data
WorkingDirectory=/var/www/certificate-manager-web
Environment="PATH=/var/www/certificate-manager-web/venv/bin"
ExecStart=/var/www/certificate-manager-web/venv/bin/gunicorn -c gunicorn.conf.py app:app
Restart=always

[Install]
WantedBy=multi-user.target
```

启动服务：

```bash
sudo systemctl daemon-reload
sudo systemctl start certificate-manager
sudo systemctl enable certificate-manager
```

### 方案二：使用 Waitress (Windows)

#### 1. 安装 Waitress

```bash
pip install waitress
```

#### 2. 创建启动脚本 `run_server.py`

```python
from waitress import serve
from app import app

serve(app,
    host='0.0.0.0',
    port=5000,
    threads=4,
    url_prefix='')

if __name__ == '__main__':
    print("Certificate Manager running on http://localhost:5000")
```

#### 3. 启动服务

```bash
python run_server.py
```

#### 4. 创建 Windows 服务

使用 NSSM 将应用注册为 Windows 服务：

```powershell
# 下载 NSSM: https://nssm.cc/download
# 安装服务
nssm install CertificateManager "C:\\path\\to\\venv\\Scripts\\python.exe" "C:\\path\\to\\run_server.py"
nssm set CertificateManager AppDirectory "C:\\path\\to\\certificate-manager-web"
nssm set CertificateManager DisplayName "Certificate Manager"
nssm set CertificateManager Description "Certificate Management Web Service"
nssm start CertificateManager
```

### 方案三：使用 Nginx 反向代理

#### Nginx 配置 `/etc/nginx/sites-available/certificate-manager`

```nginx
server {
    listen 80;
    server_name your-domain.com;

    client_max_body_size 16M;

    location / {
        proxy_pass http://127.0.0.1:5000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }

    location /static {
        alias /var/www/certificate-manager-web/static;
        expires 30d;
        add_header Cache-Control "public, immutable";
    }
}
```

启用配置：

```bash
sudo ln -s /etc/nginx/sites-available/certificate-manager /etc/nginx/sites-enabled/
sudo nginx -t
sudo systemctl reload nginx
```

---

## Docker部署

### 1. 创建 Dockerfile

```dockerfile
FROM python:3.9-slim

# 设置工作目录
WORKDIR /app

# 安装系统依赖
RUN apt-get update && apt-get install -y \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# 复制依赖文件
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 复制应用代码
COPY . .

# 创建必要的目录
RUN mkdir -p data uploads exports logs

# 暴露端口
EXPOSE 5000

# 设置环境变量
ENV FLASK_APP=app.py
ENV FLASK_ENV=production

# 启动命令
CMD ["gunicorn", "-c", "gunicorn.conf.py", "app:app"]
```

### 2. 创建 docker-compose.yml

```yaml
version: '3.8'

services:
  web:
    build: .
    ports:
      - "5000:5000"
    volumes:
      - ./data:/app/data
      - ./uploads:/app/uploads
      - ./exports:/app/exports
      - ./logs:/app/logs
    environment:
      - FLASK_ENV=production
      - SECRET_KEY=${SECRET_KEY}
    restart: unless-stopped
    healthcheck:
      test: ["CMD", "curl", "-f", "http://localhost:5000/api/session/status"]
      interval: 30s
      timeout: 10s
      retries: 3
```

### 3. 构建和运行

```bash
# 构建镜像
docker-compose build

# 启动服务
docker-compose up -d

# 查看日志
docker-compose logs -f

# 停止服务
docker-compose down
```

---

## 云平台部署

### Railway 部署

项目已包含 Railway 配置文件 `railway.json`：

```json
{
  "build": {
    "builder": "NIXPACKS"
  },
  "deploy": {
    "startCommand": "gunicorn -c gunicorn.conf.py app:app",
    "healthcheckPath": "/api/session/status"
  }
}
```

#### 部署步骤：

1. 访问 [railway.app](https://railway.app/)
2. 点击 "Deploy from GitHub repo"
3. 选择 `undercmeupto/certificate-manager` 仓库
4. 选择 `certificate-manager-web` 目录
5. 配置环境变量：
   - `SECRET_KEY`: 生成一个随机密钥
   - `FLASK_ENV`: `production`
6. 点击 "Deploy"

### Render 部署

#### 1. 创建 `render.yaml`

```yaml
services:
  - type: web
    name: certificate-manager
    env: python
    buildCommand: pip install -r requirements.txt
    startCommand: gunicorn app:app
    envVars:
      - key: FLASK_ENV
        value: production
      - key: SECRET_KEY
        generateValue: true
    disk:
      name: data
      mountPath: /app/data
      sizeGB: 1
```

#### 2. 部署步骤

1. 访问 [render.com](https://render.com/)
2. 连接 GitHub 账户
3. 选择 "New +"
4. 选择 "Web Service"
5. 连接仓库
6. 配置构建和启动命令
7. 点击 "Create Web Service"

---

## 数据备份

### 手动备份

```bash
# 备份数据文件
cp data/certificates.json data/certificates.backup.$(date +%Y%m%d).json

# 备份上传的文件
tar -czf uploads.backup.$(date +%Y%m%d).tar.gz uploads/

# 备份导出文件
tar -czf exports.backup.$(date +%Y%m%d).tar.gz exports/
```

### 自动备份脚本

创建 `scripts/backup.sh`:

```bash
#!/bin/bash
BACKUP_DIR="/backups/certificate-manager"
DATE=$(date +%Y%m%d_%H%M%S)

mkdir -p $BACKUP_DIR

# 备份数据
cp data/certificates.json $BACKUP_DIR/certificates_$DATE.json

# 压缩备份
tar -czf $BACKUP_DIR/backup_$DATE.tar.gz \
    data/ \
    uploads/ \
    exports/

# 删除30天前的备份
find $BACKUP_DIR -name "*.tar.gz" -mtime +30 -delete

echo "Backup completed: $DATE"
```

添加到 crontab:

```bash
# 每天凌晨2点备份
0 2 * * * /path/to/scripts/backup.sh
```

---

## 监控和日志

### 日志配置

创建 `logging_config.py`:

```python
import logging
from logging.handlers import RotatingFileHandler
import os

def setup_logging(app):
    log_dir = 'logs'
    os.makedirs(log_dir, exist_ok=True)

    # Flask日志
    handler = RotatingFileHandler(
        os.path.join(log_dir, 'app.log'),
        maxBytes=10485760,  # 10MB
        backupCount=10
    )
    handler.setFormatter(logging.Formatter(
        '[%(asctime)s] %(levelname)s in %(module)s: %(message)s'
    ))
    app.logger.addHandler(handler)
    app.logger.setLevel(logging.INFO)

    # Gunicorn日志
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s [%(process)d] [%(levelname)s] %(message)s',
        handlers=[
            logging.handlers.RotatingFileHandler(
                os.path.join(log_dir, 'gunicorn.log'),
                maxBytes=10485760,
                backupCount=10
            )
        ]
    )
```

### 健康检查

```python
@app.route('/health')
def health_check():
    """健康检查端点"""
    try:
        # 检查数据文件
        if os.path.exists(DATA_FILE):
            return jsonify({
                'status': 'healthy',
                'data_file': 'OK',
                'session_active': get_session_state()
            }), 200
        return jsonify({
            'status': 'warning',
            'message': 'No data file'
        }), 200
    except Exception as e:
        return jsonify({
            'status': 'unhealthy',
            'error': str(e)
        }), 503
```

---

## 安全加固

### 1. 生产环境检查清单

- [ ] 修改默认 SECRET_KEY
- [ ] 设置 DEBUG=False
- [ ] 配置 HTTPS/SSL
- [ ] 启用 CSRF 保护
- [ ] 配置防火墙规则
- [ ] 设置文件上传限制
- [ ] 定期备份数据
- [ ] 配置日志监控
- [ ] 限制 API 访问频率

### 2. SSL/HTTPS 配置

使用 Let's Encrypt 获取免费 SSL 证书：

```bash
# 安装 Certbot
sudo apt-get install certbot python3-certbot-nginx

# 获取证书
sudo certbot --nginx -d your-domain.com

# 自动续期
sudo certbot renew --dry-run
```

### 3. 防火墙配置

```bash
# UFW (Ubuntu)
sudo ufw allow 22/tcp    # SSH
sudo ufw allow 80/tcp    # HTTP
sudo ufw allow 443/tcp   # HTTPS
sudo ufw enable
```

---

## 性能优化

### 1. 启用缓存

```python
from flask_caching import Cache

cache = Cache(app, config={
    'CACHE_TYPE': 'filesystem',
    'CACHE_DIR': 'cache',
    'CACHE_THRESHOLD': 1000
})

@app.route('/api/statistics')
@cache.cached(timeout=60)  # 缓存60秒
def get_stats():
    # ...
```

### 2. 数据库迁移（可选）

考虑从 JSON 迁移到数据库：

```python
# SQLAlchemy 配置
from flask_sqlalchemy import SQLAlchemy

app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///certificates.db'
db = SQLAlchemy(app)

class Certificate(db.Model):
    id = db.Column(db.String(36), primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    # ... 其他字段
```

### 3. 分页查询

```python
@app.route('/api/data')
def get_data():
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 50, type=int)

    certificates = get_all_certificates()
    paginated = certificates[(page-1)*per_page : page*per_page]

    return jsonify({
        'data': paginated,
        'page': page,
        'per_page': per_page,
        'total': len(certificates)
    })
```

---

## 常见问题

### Q1: 端口被占用

**错误:** `Address already in use`

**解决:**
```bash
# Linux/macOS
lsof -ti:5000 | xargs kill -9

# Windows
netstat -ano | findstr :5000
taskkill /PID <PID> /F
```

### Q2: 权限错误

**错误:** `Permission denied`

**解决:**
```bash
# Linux/macOS
chmod +x app.py
chmod -R 755 data/ uploads/ exports/

# Windows（以管理员身份运行）
```

### Q3: 依赖安装失败

**错误:** `Failed building wheel for openpyxl`

**解决:**
```bash
# 升级 pip
pip install --upgrade pip

# 使用预编译包
pip install --only-binary :all: openpyxl
```

### Q4: 导出文件中文乱码

**解决:**
```python
# 确保 Excel 文件使用 UTF-8 编码
wb.save(filepath)
```

### Q5: 上传文件失败

**检查:**
1. 文件大小是否超过 16MB
2. 文件格式是否为 .xlsx 或 .xls
3. uploads 目录是否有写权限

---

## 联系支持

- GitHub: https://github.com/undercmeupto/certificate-manager
- Issues: https://github.com/undercmeupto/certificate-manager/issues

---

**文档版本:** v1.0
**最后更新:** 2026-03-15
