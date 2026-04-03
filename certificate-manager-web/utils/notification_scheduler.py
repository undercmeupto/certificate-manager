"""
Notification Scheduler Module
通知调度模块 - 自动定时发送邮件提醒
"""
import logging
import time
from datetime import datetime
from typing import List, Dict, Optional

try:
    import schedule
    SCHEDULE_AVAILABLE = True
except ImportError:
    SCHEDULE_AVAILABLE = False

logger = logging.getLogger(__name__)


class NotificationScheduler:
    """通知调度器 / Notification Scheduler"""

    def __init__(self):
        self.is_running = False
        self.last_run_time = None

    def start(self):
        """启动调度器 / Start scheduler"""
        if not SCHEDULE_AVAILABLE:
            logger.warning("Schedule library not available, auto-send disabled")
            return False

        if self.is_running:
            logger.warning("Scheduler already running")
            return True

        self.is_running = True
        logger.info("Notification scheduler started")
        return True

    def stop(self):
        """停止调度器 / Stop scheduler"""
        self.is_running = False
        if SCHEDULE_AVAILABLE:
            schedule.clear()
        logger.info("Notification scheduler stopped")

    def should_run_now(self) -> bool:
        """检查是否应该立即运行定时任务 / Check if scheduled task should run now"""
        from database import get_notification_settings

        settings = get_notification_settings()
        if not settings or not settings.get('auto_send_enabled'):
            return False

        now = datetime.now()
        schedule_time = settings.get('auto_send_time', '09:00')

        # Parse schedule time (HH:MM format)
        try:
            hour, minute = map(int, schedule_time.split(':'))
            scheduled_time = now.replace(hour=hour, minute=minute, second=0, microsecond=0)

            # Check if current time is within 1 minute of scheduled time
            time_diff = (now - scheduled_time).total_seconds()
            if 0 <= time_diff < 60:
                # Also check if we haven't run recently (avoid duplicate runs)
                if self.last_run_time:
                    last_run_diff = (now - self.last_run_time).total_seconds()
                    if last_run_diff < 3600:  # Don't run again within 1 hour
                        return False
                return True
        except Exception as e:
            logger.error(f"Error parsing schedule time: {e}")

        return False

    def tick(self):
        """调度器心跳检查 / Scheduler tick - call this periodically"""
        if self.is_running and SCHEDULE_AVAILABLE:
            if self.should_run_now():
                self._send_scheduled_notifications()
            schedule.run_pending()

    def _send_scheduled_notifications(self):
        """执行定时发送任务 / Execute scheduled notification send"""
        try:
            logger.info("Starting scheduled notification send...")
            self.last_run_time = datetime.now()

            from utils.email_service import email_service
            from database import (
                get_all_certificates,
                get_notification_settings_with_password,
                create_notification_record,
                get_notification_preference
            )

            # Get settings WITH PASSWORD
            settings = get_notification_settings_with_password()
            if not settings:
                logger.error("No notification settings found")
                return

            # Configure email service
            email_service.configure_smtp(settings)

            # Get all certificates
            certificates = get_all_certificates()

            # Filter certificates based on settings
            certificates_to_notify = self._filter_certificates(certificates, settings)

            logger.info(f"Found {len(certificates_to_notify)} certificates to notify")

            # Send notifications
            sent_count = 0
            failed_count = 0

            for cert in certificates_to_notify:
                try:
                    # Check user preference
                    pref = get_notification_preference(cert.get('email'))
                    if pref and not pref.get('email_enabled', True):
                        logger.info(f"Skipping {cert.get('email')} - notifications disabled")
                        continue

                    # Generate email
                    email_data = email_service.generate_certificate_email(cert)

                    # Send email
                    success, message = email_service.send_email(
                        email_data['to_email'],
                        email_data['subject'],
                        email_data['body']
                    )

                    # Parse expiry date for database
                    expiry_date = None
                    if cert.get('expiry_date'):
                        try:
                            expiry_date = datetime.strptime(cert.get('expiry_date'), '%Y-%m-%d').date()
                        except Exception:
                            pass

                    # Create record
                    record_data = {
                        'certificate_id': cert.get('id'),
                        'recipient_name': cert.get('name'),
                        'recipient_email': cert.get('email'),
                        'certificate_name': cert.get('certificate_name'),
                        'expiry_date': expiry_date,
                        'days_remaining': cert.get('days_remaining'),
                        'status': cert.get('status'),
                        'send_type': 'auto',
                        'send_status': 'sent' if success else 'failed',
                        'error_message': None if success else message,
                        'email_subject': email_data['subject'],
                        'email_body': email_data['body'],
                        'sent_at': datetime.now() if success else None
                    }
                    create_notification_record(record_data)

                    if success:
                        sent_count += 1
                    else:
                        failed_count += 1

                    # Rate limiting
                    time.sleep(0.5)

                except Exception as e:
                    logger.error(f"Failed to send notification to {cert.get('email')}: {e}")
                    failed_count += 1

            logger.info(f"Scheduled send completed: {sent_count} sent, {failed_count} failed")

        except Exception as e:
            logger.error(f"Scheduled notification task failed: {e}")

    def _filter_certificates(self, certificates: List[Dict], settings: Dict) -> List[Dict]:
        """根据设置过滤需要通知的证件 / Filter certificates based on settings"""
        filtered = []

        for cert in certificates:
            # Must have email
            if not cert.get('email'):
                continue

            status = cert.get('status')
            days_remaining = cert.get('days_remaining')

            # Check if should notify based on settings
            should_notify = False

            if settings.get('notify_expired') and status == 'expired':
                should_notify = True
            elif settings.get('notify_urgent') and status == 'urgent':
                should_notify = True
            elif settings.get('notify_warning') and status == 'warning':
                should_notify = True

            if should_notify:
                filtered.append(cert)

        return filtered

    def trigger_manual_send(self, certificate_ids: List[str] = None,
                           status_filter: str = None,
                           custom_recipient_email: str = None) -> Dict:
        """手动触发发送 / Trigger manual send

        Args:
            certificate_ids: 指定证件ID列表
            status_filter: 状态过滤 (expired, urgent, warning, all)
            custom_recipient_email: 自定义收件人邮箱（覆盖所有收件人）

        Returns:
            发送结果统计 / Send result statistics
        """
        import time

        from utils.email_service import email_service
        from database import (
            get_all_certificates,
            get_notification_settings_with_password,
            create_notification_record,
            get_certificate_by_id
        )

        # Get settings WITH PASSWORD
        settings = get_notification_settings_with_password()
        if not settings or not settings.get('smtp_server'):
            return {'success': False, 'error': '未配置SMTP设置 / SMTP not configured'}

        # Configure email service
        email_service.configure_smtp(settings)

        # Get certificates
        if certificate_ids:
            certificates = []
            for cid in certificate_ids:
                cert = get_certificate_by_id(cid)
                if cert:
                    certificates.append(cert)
        else:
            certificates = get_all_certificates()

        # Filter by status if specified
        if status_filter and status_filter != 'all':
            certificates = [c for c in certificates if c.get('status') == status_filter]
        else:
            # Filter to only include expired, urgent, warning
            certificates = [c for c in certificates if c.get('status') in ['expired', 'urgent', 'warning']]

        # Handle custom recipient email
        # When custom email is provided, include all certificates and use the custom email
        # When not using custom email, only include certificates with email
        if custom_recipient_email:
            logger.info(f"Using custom recipient email: {custom_recipient_email}")
        else:
            # Must have email
            certificates = [c for c in certificates if c.get('email')]
            logger.info(f"Filtered to {len(certificates)} certificates with email")

        logger.info(f"Manual send: {len(certificates)} certificates, custom_email: {bool(custom_recipient_email)}")

        sent_count = 0
        failed_count = 0
        results = []

        for cert in certificates:
            try:
                # Generate email
                email_data = email_service.generate_certificate_email(cert)

                # Use custom email if provided, otherwise use certificate email
                to_email = custom_recipient_email if custom_recipient_email else email_data['to_email']

                # Skip if no email (only when not using custom email)
                if not to_email:
                    logger.info(f"Skipping cert {cert.get('id')} - no email")
                    continue

                # Update email_data for send
                email_data['to_email'] = to_email

                # Send email
                success, message = email_service.send_email(
                    to_email,
                    email_data['subject'],
                    email_data['body']
                )

                # Parse expiry date for database
                expiry_date = None
                if cert.get('expiry_date'):
                    try:
                        expiry_date = datetime.strptime(cert.get('expiry_date'), '%Y-%m-%d').date()
                    except Exception:
                        pass

                # Create record
                record_data = {
                    'certificate_id': cert.get('id'),
                    'recipient_name': cert.get('name'),
                    'recipient_email': to_email,  # Use actual recipient email
                    'certificate_name': cert.get('certificate_name'),
                    'expiry_date': expiry_date,
                    'days_remaining': cert.get('days_remaining'),
                    'status': cert.get('status'),
                    'send_type': 'manual',
                    'send_status': 'sent' if success else 'failed',
                    'error_message': None if success else message,
                    'email_subject': email_data['subject'],
                    'email_body': email_data['body'],
                    'sent_at': datetime.now() if success else None
                }
                record_id = create_notification_record(record_data)

                results.append({
                    'certificate_id': cert.get('id'),
                    'email': to_email,  # Use actual recipient email
                    'original_email': cert.get('email'),  # Store original email for reference
                    'custom_email': custom_recipient_email,  # Store if custom email was used
                    'success': success,
                    'message': message,
                    'record_id': record_id
                })

                if success:
                    sent_count += 1
                else:
                    failed_count += 1

                time.sleep(0.5)  # Rate limiting

            except Exception as e:
                logger.error(f"Failed to send to {cert.get('email')}: {e}")
                failed_count += 1
                results.append({
                    'certificate_id': cert.get('id'),
                    'email': cert.get('email'),
                    'success': False,
                    'message': str(e)
                })

        return {
            'success': True,
            'total': len(certificates),
            'sent': sent_count,
            'failed': failed_count,
            'results': results
        }

    def _get_certificates_by_ids(self, certificate_ids: List[str]) -> List[Dict]:
        """根据ID获取证件列表 / Get certificates by IDs"""
        from database import get_certificate_by_id

        certificates = []
        for cid in certificate_ids:
            cert = get_certificate_by_id(cid)
            if cert:
                certificates.append(cert)
        return certificates

    def _filter_by_status(self, certificates: List[Dict], status_filter: str) -> List[Dict]:
        """根据状态过滤证件 / Filter certificates by status"""
        if not status_filter or status_filter == 'all':
            # Filter to only include expired, urgent, warning (exclude normal and unknown)
            return [c for c in certificates if c.get('status') in ['expired', 'urgent', 'warning']]
        return [c for c in certificates if c.get('status') == status_filter]

    def _filter_certificates_expiring(self, certificates: List[Dict]) -> List[Dict]:
        """过滤即将到期的证件（小于90天）/ Filter certificates expiring within 90 days"""
        return [c for c in certificates if c.get('days_remaining', float('inf')) < 90]

    def _group_certificates_by_person(self, certificates: List[Dict],
                                       default_email: str = None) -> List[Dict]:
        """按人员分组证件 / Group certificates by person

        Groups certificates by (email, name) tuple. Certificates without email
        will use the default_email if provided, or empty string for user to fill.

        Args:
            certificates: 证件列表 / List of certificates
            default_email: 默认邮箱（用于无邮箱的证件）/ Default email for certs without email

        Returns:
            分组后的证件列表 / Grouped certificates list
        """
        groups = {}

        for cert in certificates:
            # Get person name first
            name = cert.get('name', '未知')

            # Get email (use default if not available, otherwise use empty string)
            email = cert.get('email') or default_email or ''

            # Create grouping key - use name as primary when email is empty
            # This ensures that certificates without email are still grouped by person
            if email:
                key = (email, name)
            else:
                # For certificates without email, group by name only
                key = (f'no-email-{name}', name)

            if key not in groups:
                groups[key] = {
                    'person_name': name,
                    'email': email,
                    'certificates': []
                }

            groups[key]['certificates'].append(cert)

        # Convert to list
        return list(groups.values())

    def get_preview_data(self, certificate_ids: List[str] = None,
                         status_filter: str = None,
                         default_email: str = None) -> Dict:
        """获取批量发送预览数据 / Get preview data for batch send

        当用户选择特定证件时，会获取这些证件所属员工的所有即将过期证件（90天内）。
        When user selects specific certificates, this gets ALL expiring certificates
        for those employees (within 90 days).

        Args:
            certificate_ids: 指定证件ID列表 / Specific certificate IDs
            status_filter: 状态过滤 / Status filter
            default_email: 默认邮箱 / Default email

        Returns:
            预览数据 / Preview data
        """
        from database import get_all_certificates

        # Get initial certificates based on selection
        if certificate_ids:
            selected_certificates = self._get_certificates_by_ids(certificate_ids)
        else:
            # When no specific selection, get all certificates
            selected_certificates = get_all_certificates()

        # Filter by status if specified
        if status_filter:
            selected_certificates = self._filter_by_status(selected_certificates, status_filter)

        # NEW LOGIC: When certificates are selected, get ALL expiring certificates
        # for the employees who own those selected certificates
        if certificate_ids:
            # Extract unique employee NAMES from selected certificates (not just emails)
            # This handles employees who don't have email addresses
            employee_names = set()
            for cert in selected_certificates:
                name = cert.get('name')
                if name:
                    employee_names.add(name)

            logger.info(f"Found {len(employee_names)} unique employees from {len(selected_certificates)} selected certificates")

            # Get ALL certificates for these employees (by name matching)
            all_certificates = get_all_certificates()
            employee_certificates = []

            for cert in all_certificates:
                cert_name = cert.get('name')
                if cert_name in employee_names:
                    employee_certificates.append(cert)

            logger.info(f"Found {len(employee_certificates)} total certificates for these employees")

            # Filter to only include expiring certificates (< 90 days)
            certificates = self._filter_certificates_expiring(employee_certificates)

            logger.info(f"After filtering (<90 days): {len(certificates)} certificates")
        else:
            # When filtering by status (no specific selection)
            # Filter expiring only (< 90 days)
            certificates = self._filter_certificates_expiring(selected_certificates)

        # Group by person
        groups = self._group_certificates_by_person(certificates, default_email)

        logger.info(f"Grouped into {len(groups)} employees")

        # Format for response
        formatted_groups = []
        for group in groups:
            formatted_groups.append({
                'person_name': group['person_name'],
                'email': group['email'],
                'certificate_count': len(group['certificates']),
                'certificates': group['certificates']  # Include full cert data for detail view
            })

        return {
            'total_people': len(groups),
            'total_certificates': len(certificates),
            'groups': formatted_groups
        }

    def trigger_batch_send(self, recipients: List[Dict]) -> Dict:
        """触发批量邮件发送（按人员分组，支持修改收件人）/ Trigger batch email send

        Args:
            recipients: 修改后的收件人列表 / Modified recipients list
                Each recipient: {
                    'person_name': str,
                    'email': str,
                    'certificate_ids': List[str]
                }

        Returns:
            发送结果统计 / Send result statistics
        """
        from utils.email_service import email_service
        from database import (
            get_all_certificates,
            get_notification_settings_with_password,
            create_notification_record
        )

        # Get settings WITH PASSWORD
        settings = get_notification_settings_with_password()
        if not settings or not settings.get('smtp_server'):
            return {'success': False, 'error': '未配置SMTP设置 / SMTP not configured'}

        # Configure email service
        email_service.configure_smtp(settings)

        results = []
        sent_count = 0
        failed_count = 0

        for recipient in recipients:
            try:
                person_name = recipient['person_name']
                email = recipient['email']
                certificate_ids = recipient['certificate_ids']

                # Get certificates
                certificates = self._get_certificates_by_ids(certificate_ids)

                if not certificates:
                    logger.warning(f"No certificates found for recipient {person_name}")
                    continue

                # Generate batch email
                email_data = email_service.generate_batch_email(
                    person_name=person_name,
                    email=email,
                    certificates=certificates
                )

                # Send email
                success, message = email_service.send_email(
                    email,
                    email_data['subject'],
                    email_data['body'],
                    is_html=True  # Batch emails use HTML format
                )

                # Create notification records for each certificate
                for cert in certificates:
                    # Parse expiry date for database
                    expiry_date = None
                    if cert.get('expiry_date'):
                        try:
                            expiry_date = datetime.strptime(cert.get('expiry_date'), '%Y-%m-%d').date()
                        except Exception:
                            pass

                    record_data = {
                        'certificate_id': cert.get('id'),
                        'recipient_name': person_name,
                        'recipient_email': email,
                        'certificate_name': cert.get('certificate_name'),
                        'expiry_date': expiry_date,
                        'days_remaining': cert.get('days_remaining'),
                        'status': cert.get('status'),
                        'send_type': 'manual',
                        'send_status': 'sent' if success else 'failed',
                        'error_message': None if success else message,
                        'email_subject': email_data['subject'],
                        'email_body': email_data['body'],
                        'sent_at': datetime.now() if success else None
                    }
                    create_notification_record(record_data)

                results.append({
                    'person': person_name,
                    'email': email,
                    'success': success,
                    'message': message,
                    'certificate_count': len(certificates)
                })

                if success:
                    sent_count += 1
                else:
                    failed_count += 1

                # Rate limiting between emails
                time.sleep(1)

            except Exception as e:
                logger.error(f"Failed to send batch to {recipient.get('email')}: {e}")
                failed_count += 1
                results.append({
                    'person': recipient.get('person_name'),
                    'email': recipient.get('email'),
                    'success': False,
                    'message': str(e),
                    'certificate_count': 0
                })

        return {
            'success': True,
            'total_people': len(recipients),
            'total_certificates': sum(r.get('certificate_count', 0) for r in results),
            'sent': sent_count,
            'failed': failed_count,
            'results': results
        }


# Global scheduler instance
notification_scheduler = NotificationScheduler()
