"""
证件管理系统配置
"""
import os

# ============ 路径配置 ============
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, 'data', 'certificates.json')
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
EXPORT_FOLDER = os.path.join(BASE_DIR, 'exports')

# ============ 阈值配置 ============
URGENT_DAYS = 30      # 紧急: < 30天
WARNING_DAYS = 90     # 预警: 30-90天

# ============ 文件配置 ============
ALLOWED_EXTENSIONS = {"xlsx", "xls"}
MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB

# ============ Flask 配置 ============
SECRET_KEY = os.environ.get('SECRET_KEY', 'dev-secret-key-change-in-production')
DEBUG = os.environ.get('DEBUG', 'True').lower() == 'true'
HOST = os.environ.get('HOST', '0.0.0.0')
PORT = int(os.environ.get('PORT', 5000))

# ============ 状态映射 ============
STATUS_MAP = {
    'expired': {'label': '已过期', 'icon': 'X', 'color': '#C0392B'},
    'urgent': {'label': '紧急', 'icon': '!', 'color': '#E74C3C'},
    'warning': {'label': '预警', 'icon': '?', 'color': '#F39C12'},
    'normal': {'label': '正常', 'icon': 'OK', 'color': '#27AE60'},
    'unknown': {'label': '日期无效', 'icon': '?', 'color': '#95A5A6'}
}

# ============ 会话状态配置 ============
SESSION_STATE_FILE = os.path.join(BASE_DIR, 'data', 'session_state.json')

# ============ 证件类型列定义（用于复杂格式Excel）============
CERTIFICATE_TYPES = [
    {'name': 'IADC/IWCF', 'num_col': 9, 'issue_col': 10, 'exp_col': 11},
    {'name': 'HSE证(H2S)', 'num_col': 12, 'issue_col': 13, 'exp_col': 14},
    {'name': '急救证', 'num_col': 15, 'issue_col': 16, 'exp_col': 17},
    {'name': '消防证', 'num_col': 18, 'issue_col': 19, 'exp_col': 20},
    {'name': '司索指挥证', 'num_col': 21, 'issue_col': 22, 'exp_col': 23},
    {'name': '防恐证', 'num_col': 24, 'issue_col': 25, 'exp_col': 26},
    {'name': '健康证', 'num_col': 30, 'issue_col': 31, 'exp_col': 32},
    {'name': '岗位证', 'num_col': 33, 'issue_col': 34, 'exp_col': 35},
]
