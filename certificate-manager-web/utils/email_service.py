"""
Email Service Module
邮件服务模块 - 处理SMTP连接和邮件发送
"""
import smtplib
import os
import base64
import logging
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
from email.header import Header
from typing import Dict, Tuple, List, Optional

try:
    from cryptography.fernet import Fernet
    FERNET_AVAILABLE = True
except ImportError:
    FERNET_AVAILABLE = False

logger = logging.getLogger(__name__)


class EmailService:
    """邮件服务类 / Email Service Class"""

    def __init__(self):
        self.smtp_config = None
        self.fernet_key = self._get_or_create_encryption_key() if FERNET_AVAILABLE else None

    @staticmethod
    def _get_or_create_encryption_key():
        """获取或创建加密密钥 / Get or create encryption key"""
        key_file = os.path.join(os.path.dirname(os.path.dirname(__file__)), '.email_key')
        if os.path.exists(key_file):
            try:
                with open(key_file, 'rb') as f:
                    return f.read()
            except Exception:
                pass

        # Generate new key
        if FERNET_AVAILABLE:
            key = Fernet.generate_key()
            try:
                with open(key_file, 'wb') as f:
                    f.write(key)
            except Exception:
                pass
            return key
        return None

    def encrypt_password(self, password: str) -> str:
        """加密密码 / Encrypt password"""
        if not FERNET_AVAILABLE or not self.fernet_key:
            return password  # Return as-is if encryption not available
        try:
            f = Fernet(self.fernet_key)
            encrypted = f.encrypt(password.encode('utf-8'))
            return base64.b64encode(encrypted).decode('utf-8')
        except Exception as e:
            logger.error(f"Password encryption failed: {e}")
            return password

    def decrypt_password(self, encrypted_password: str) -> str:
        """解密密码 / Decrypt password"""
        if not FERNET_AVAILABLE or not self.fernet_key:
            return encrypted_password  # Return as-is if encryption not available
        try:
            f = Fernet(self.fernet_key)
            decoded = base64.b64decode(encrypted_password.encode('utf-8'))
            decrypted = f.decrypt(decoded)
            return decrypted.decode('utf-8')
        except Exception as e:
            logger.error(f"Password decryption failed: {e}")
            return encrypted_password

    def configure_smtp(self, config: Dict) -> bool:
        """配置SMTP设置 / Configure SMTP settings"""
        self.smtp_config = {
            'server': config.get('smtp_server'),
            'port': config.get('smtp_port', 587),
            'use_tls': config.get('smtp_use_tls', True),
            'username': config.get('smtp_username'),
            'password': self.decrypt_password(config.get('smtp_password', '')),
            'from_email': config.get('smtp_from_email'),
            'from_name': config.get('smtp_from_name', '证件管理系统')
        }
        logger.info(f"SMTP configured: {self.smtp_config['server']}")
        return True

    def test_connection(self) -> Tuple[bool, str]:
        """测试SMTP连接 / Test SMTP connection"""
        if not self.smtp_config or not self.smtp_config.get('server'):
            return False, "SMTP未配置 / SMTP not configured"

        try:
            # Use SMTP_SSL for port 465, regular SMTP for other ports
            port = self.smtp_config['port']
            if port == 465:
                # SSL connection (direct SSL)
                server = smtplib.SMTP_SSL(self.smtp_config['server'], port, timeout=30)
                # Identify ourselves (required by some SMTP servers before login)
                try:
                    server.ehlo()
                except smtplib.SMTPException:
                    pass  # Some servers don't support EHLO in SSL mode
            else:
                # TLS connection (STARTTLS)
                server = smtplib.SMTP(self.smtp_config['server'], port, timeout=30)
                # Identify ourselves first
                server.ehlo()
                if self.smtp_config['use_tls']:
                    server.starttls()
                    # Re-identify after TLS
                    server.ehlo()

            if self.smtp_config['username'] and self.smtp_config['password']:
                server.login(self.smtp_config['username'], self.smtp_config['password'])
            server.quit()
            return True, "连接成功 / Connection successful"
        except smtplib.SMTPAuthenticationError as e:
            return False, f"SMTP认证失败: {str(e)} / SMTP authentication failed"
        except smtplib.SMTPException as e:
            return False, f"SMTP错误: {str(e)} / SMTP error"
        except Exception as e:
            return False, f"连接失败: {str(e)} / Connection failed"

    def send_email(self, to_email: str, subject: str, body: str,
                   is_html: bool = False) -> Tuple[bool, str]:
        """发送单封邮件 / Send single email"""
        if not self.smtp_config or not self.smtp_config.get('server'):
            return False, "SMTP未配置 / SMTP not configured"

        if not to_email:
            return False, "收件人地址为空 / Recipient email is empty"

        try:
            # Create message
            msg = MIMEMultipart('alternative')
            msg['Subject'] = Header(subject, 'utf-8')
            msg['From'] = formataddr((
                self.smtp_config['from_name'],
                self.smtp_config['from_email']
            ))
            msg['To'] = to_email

            # Add body
            mime_type = 'html' if is_html else 'plain'
            msg.attach(MIMEText(body, mime_type, 'utf-8'))

            # Send via SMTP
            port = self.smtp_config['port']
            if port == 465:
                # SSL connection (direct SSL)
                server = smtplib.SMTP_SSL(self.smtp_config['server'], port, timeout=30)
                # Identify ourselves (required by some SMTP servers before login)
                try:
                    server.ehlo()
                except smtplib.SMTPException:
                    pass  # Some servers don't support EHLO in SSL mode
            else:
                # TLS connection (STARTTLS)
                server = smtplib.SMTP(self.smtp_config['server'], port, timeout=30)
                # Identify ourselves first
                server.ehlo()
                if self.smtp_config['use_tls']:
                    server.starttls()
                    # Re-identify after TLS
                    server.ehlo()

            if self.smtp_config['username'] and self.smtp_config['password']:
                server.login(self.smtp_config['username'], self.smtp_config['password'])

            server.send_message(msg)
            server.quit()

            logger.info(f"Email sent successfully to {to_email}")
            return True, "发送成功 / Sent successfully"

        except smtplib.SMTPAuthenticationError:
            return False, "SMTP认证失败 / Authentication failed"
        except smtplib.SMTPRecipientsRefused:
            return False, "收件人地址被拒绝 / Recipient refused"
        except smtplib.SMTPException as e:
            return False, f"SMTP错误 / SMTP error: {str(e)}"
        except Exception as e:
            logger.error(f"Failed to send email to {to_email}: {e}")
            return False, f"发送失败 / Send failed: {str(e)}"

    def generate_certificate_email(self, certificate: Dict) -> Dict:
        """生成证件提醒邮件内容 / Generate certificate reminder email content"""
        from config import EMAIL_TEMPLATES

        # Select template based on status
        status = certificate.get('status', 'normal')
        if status == 'expired':
            template_key = 'body_expired'
        elif status == 'urgent':
            template_key = 'body_urgent'
        else:
            template_key = 'body_warning'

        # Get template
        template = EMAIL_TEMPLATES.get(template_key, EMAIL_TEMPLATES.get('body_warning', ''))

        # Prepare template variables
        days_remaining = certificate.get('days_remaining', 0)
        days_overdue = abs(days_remaining) if days_remaining < 0 else 0

        template_vars = {
            'name': certificate.get('name', ''),
            'certificate_name': certificate.get('certificate_name', ''),
            'certificate_number': certificate.get('certificate_number', ''),
            'expiry_date': certificate.get('expiry_date', ''),
            'days_remaining': days_remaining,
            'days_overdue': days_overdue
        }

        # Generate subject
        subject = EMAIL_TEMPLATES.get('subject', '证件到期提醒').format(**template_vars)

        # Generate body
        try:
            body = template.format(**template_vars)
        except KeyError as e:
            logger.error(f"Template formatting error: {e}")
            body = template

        return {
            'subject': subject,
            'body': body,
            'to_email': certificate.get('email', ''),
            'recipient_name': certificate.get('name', ''),
            'certificate_id': certificate.get('id', '')
        }

    def generate_batch_email(self, person_name: str, email: str,
                             certificates: List[Dict]) -> Dict:
        """生成批量证件提醒邮件（按人员汇总）/ Generate batch certificate reminder email (grouped by person)

        Args:
            person_name: 收件人姓名 / Recipient name
            email: 收件人邮箱 / Recipient email
            certificates: 证件列表 / List of certificates

        Returns:
            邮件信息字典 / Email information dict
        """
        from config import EMAIL_TEMPLATES, STATUS_NOTES, STATUS_MAP

        # 构建证件行 / Build certificate rows
        certificate_rows = []
        for cert in certificates:
            days_remaining = cert.get('days_remaining', 0)
            status = cert.get('status', 'unknown')

            # 获取状态备注 / Get status note
            status_note = ''
            if status == 'expired':
                days_overdue = abs(days_remaining)
                status_note = STATUS_NOTES['expired'].format(days_overdue=days_overdue)
            elif status == 'urgent':
                status_note = STATUS_NOTES['urgent'].format(days_remaining=days_remaining)
            elif status == 'warning':
                status_note = STATUS_NOTES['warning'].format(days_remaining=days_remaining)

            # 获取状态颜色和标签 / Get status color and label
            status_info = STATUS_MAP.get(status, {})
            status_color = status_info.get('color', '#666666')
            status_label = status_info.get('label', '未知')

            # 生成证件行 / Generate certificate row
            row = EMAIL_TEMPLATES['certificate_row'].format(
                certificate_name=cert.get('certificate_name', ''),
                expiry_date=cert.get('expiry_date', ''),
                status_label=status_label,
                status_color=status_color,
                status_note=status_note
            )
            certificate_rows.append(row)

        # 合并所有行 / Combine all rows
        all_rows = ''.join(certificate_rows)

        # 构建完整证件列表 / Build complete certificate list
        certificate_list = EMAIL_TEMPLATES['certificate_list_header'].format(
            certificate_rows=all_rows
        )

        # 生成主题 / Generate subject
        count = len(certificates)
        subject = EMAIL_TEMPLATES['subject_batch'].format(count=count)

        # 生成正文 / Generate body
        body = EMAIL_TEMPLATES['body_batch'].format(
            name=person_name,
            certificate_list=certificate_list
        )

        return {
            'subject': subject,
            'body': body,
            'to_email': email,
            'recipient_name': person_name,
            'certificate_count': count,
            'certificate_ids': [c.get('id') for c in certificates]
        }


# Global email service instance
email_service = EmailService()
