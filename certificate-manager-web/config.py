"""
证件管理系统配置
"""
import os
import sys

# ============ 路径配置 ============
def get_base_dir():
    """获取应用基础目录，支持PyInstaller打包"""
    if getattr(sys, 'frozen', False):
        # 如果是PyInstaller打包的exe
        return sys._MEIPASS
    else:
        # 开发环境
        return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = get_base_dir()

def get_user_data_dir():
    """获取用户数据目录，用于存放可写文件"""
    if getattr(sys, 'frozen', False):
        # 打包后使用用户目录
        home = os.path.expanduser('~')
        app_data = os.path.join(home, 'CertificateManager')
        os.makedirs(app_data, exist_ok=True)
        return app_data
    else:
        # 开发环境使用项目目录
        return BASE_DIR

USER_DATA_DIR = get_user_data_dir()

# ============ 数据库配置 ============
DATABASE_FILE = os.path.join(USER_DATA_DIR, 'certificates.db')

# ============ 文件路径配置（保留用于迁移参考）============
DATA_FILE = os.path.join(USER_DATA_DIR, 'certificates.json')  # Legacy JSON file
UPLOAD_FOLDER = os.path.join(USER_DATA_DIR, 'uploads')
EXPORT_FOLDER = os.path.join(USER_DATA_DIR, 'exports')
SESSION_STATE_FILE = os.path.join(USER_DATA_DIR, 'session_state.json')  # Legacy JSON file

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

# ============ 邮件通知配置 ============
# Email Notification Configuration

# SMTP默认配置
DEFAULT_SMTP_PORT = 587
DEFAULT_SMTP_USE_TLS = True
DEFAULT_AUTO_SEND_TIME = '09:00'
DEFAULT_AUTO_SEND_DAY = 1  # Monday

# 通知频率选项 / Notification schedule options
SCHEDULE_DAILY = 'daily'
SCHEDULE_WEEKLY = 'weekly'
SCHEDULE_MONTHLY = 'monthly'

NOTIFICATION_SCHEDULES = {
    SCHEDULE_DAILY: '每日发送',
    SCHEDULE_WEEKLY: '每周发送',
    SCHEDULE_MONTHLY: '每月发送'
}

# 发送状态常量 / Send status constants
SEND_STATUS_PENDING = 'pending'
SEND_STATUS_SENT = 'sent'
SEND_STATUS_FAILED = 'failed'

# 发送类型常量 / Send type constants
SEND_TYPE_MANUAL = 'manual'
SEND_TYPE_AUTO = 'auto'

# 邮件发送限制
EMAIL_BATCH_SIZE = 50  # 每批发送邮件数
EMAIL_DELAY_SECONDS = 1  # 批次间延迟（秒）

# 邮件模板（中文）
EMAIL_TEMPLATES = {
    'subject': '证件到期提醒 - {certificate_name}',
    'greeting': '尊敬的{name}：',
    'body_expired': '''您的证件【{certificate_name}】已过期！

证件信息：
- 证件名称：{certificate_name}
- 证件号码：{certificate_number}
- 到期日期：{expiry_date}
- 状态：已过期（过期 {days_overdue} 天）

请您尽快办理证件续期手续。

此邮件由证件管理系统自动发送，请勿回复。''',

    'body_urgent': '''您的证件【{certificate_name}】即将到期！

证件信息：
- 证件名称：{certificate_name}
- 证件号码：{certificate_number}
- 到期日期：{expiry_date}
- 剩余天数：{days_remaining} 天

请您尽快办理证件续期手续，以免影响正常工作。

此邮件由证件管理系统自动发送，请勿回复。''',

    'body_warning': '''温馨提醒：您的证件【{certificate_name}】将在{days_remaining}天后到期。

证件信息：
- 证件名称：{certificate_name}
- 证件号码：{certificate_number}
- 到期日期：{expiry_date}

请您提前安排证件续期事宜。

此邮件由证件管理系统自动发送，请勿回复。''',

    # 批量邮件模板（按人员汇总） / Batch Email Templates (Grouped by Person)
    'subject_batch': '证件到期提醒 - 您有 {count} 个证件即将到期',

    'body_batch': '''尊敬的{name}：

您好！以下是您的证件到期提醒：

{certificate_list}

请及时办理证件续期手续。

此邮件由证件管理系统自动发送，请勿回复。''',

    # 证件行（HTML表格格式） / Certificate Row (HTML Table Format)
    'certificate_row': '''
    <tr>
        <td style="padding: 12px; border-bottom: 1px solid #ddd;">{certificate_name}</td>
        <td style="padding: 12px; border-bottom: 1px solid #ddd; text-align: center;">{expiry_date}</td>
        <td style="padding: 12px; border-bottom: 1px solid #ddd; text-align: center;">
            <span style="color: {status_color}; font-weight: bold;">{status_label}</span>
            {status_note}
        </td>
    </tr>''',

    # 证件列表表头（HTML表格） / Certificate List Header (HTML Table)
    'certificate_list_header': '''
    <table style="width: 100%%; border-collapse: collapse; margin: 20px 0;">
        <thead>
            <tr style="background-color: #f5f5f5;">
                <th style="padding: 12px; text-align: left; border-bottom: 2px solid #ddd;">证件名称</th>
                <th style="padding: 12px; text-align: center; border-bottom: 2px solid #ddd;">到期日期</th>
                <th style="padding: 12px; text-align: center; border-bottom: 2px solid #ddd;">状态</th>
            </tr>
        </thead>
        <tbody>
{certificate_rows}
        </tbody>
    </table>'''
}

# 状态备注（用于批量邮件） / Status Notes (for Batch Emails)
STATUS_NOTES = {
    'expired': '（已过期 {days_overdue} 天）',
    'urgent': '（剩余 {days_remaining} 天）',
    'warning': '（剩余 {days_remaining} 天）'
}
