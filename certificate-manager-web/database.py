"""
Database Connection and Session Management
数据库连接和会话管理
"""
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, scoped_session
from contextlib import contextmanager
from datetime import datetime
from models import (
    Base, Certificate, UploadMetadata, SessionState,
    NotificationSettings, NotificationRecord, NotificationPreference,
    parse_date
)
import config


# Global engine and session factory
_engine = None
_session_factory = None
Session = None


def init_db(db_path=None):
    """
    Initialize database connection and create tables

    Args:
        db_path: Database file path. If None, uses config.DATABASE_FILE
    """
    global _engine, _session_factory, Session

    if db_path is None:
        db_path = config.DATABASE_FILE

    # Create SQLite engine with proper settings
    _engine = create_engine(
        f'sqlite:///{db_path}',
        echo=config.DEBUG,  # Log SQL queries in debug mode
        connect_args={'check_same_thread': False},  # Allow multi-thread access
        pool_pre_ping=True,  # Verify connections before using
    )

    # Create session factory
    _session_factory = sessionmaker(bind=_engine)
    Session = scoped_session(_session_factory)

    # Create all tables
    Base.metadata.create_all(_engine)

    # Initialize session state if not exists
    _init_session_state()

    return _engine


def _init_session_state():
    """Initialize session state row if not exists"""
    session = get_session()
    try:
        state = session.query(SessionState).filter_by(id=1).first()
        if state is None:
            state = SessionState(id=1, active=False)
            session.add(state)
            session.commit()
    finally:
        close_session(session)


def get_session():
    """
    Get a new database session

    Returns:
        SQLAlchemy session object
    """
    if Session is None:
        init_db()
    return Session()


def close_session(session):
    """
    Close a database session

    Args:
        session: SQLAlchemy session to close
    """
    session.close()


@contextmanager
def session_scope():
    """
    Provide a transactional scope around a series of operations.

    Usage:
        with session_scope() as session:
            session.add(obj)
            # session will be committed or rolled back automatically
    """
    session = get_session()
    try:
        yield session
        session.commit()
    except Exception:
        session.rollback()
        raise
    finally:
        close_session(session)


def _update_model_fields(model, data_dict, exclude_fields=None):
    """
    Update model attributes from a dictionary.

    Args:
        model: SQLAlchemy model instance
        data_dict: Dictionary of field names to values
        exclude_fields: List of field names to exclude from updates

    Returns:
        True if any fields were updated, False otherwise
    """
    exclude = set(exclude_fields) if exclude_fields else {'id'}
    updated = False
    for key, value in data_dict.items():
        if key not in exclude and hasattr(model, key):
            setattr(model, key, value)
            updated = True
    return updated


def get_db_session():
    """
    Alias for get_session() for backward compatibility

    Returns:
        SQLAlchemy session object
    """
    return get_session()


# Certificate CRUD operations
def get_all_certificates(session=None):
    """Get all certificates"""
    if session is None:
        with session_scope() as sess:
            return [cert.to_dict() for cert in sess.query(Certificate).all()]
    return [cert.to_dict() for cert in session.query(Certificate).all()]


def get_certificate_by_id(cert_id, session=None):
    """Get certificate by ID"""
    if session is None:
        with session_scope() as sess:
            cert = sess.query(Certificate).filter_by(id=cert_id).first()
            return cert.to_dict() if cert else None
    cert = session.query(Certificate).filter_by(id=cert_id).first()
    return cert.to_dict() if cert else None


def add_certificate(cert_data, session=None):
    """Add a new certificate"""
    cert = Certificate.from_dict(cert_data)
    if session is None:
        with session_scope() as sess:
            sess.add(cert)
            sess.flush()
            return cert.id
    session.add(cert)
    session.flush()
    return cert.id


def update_certificate(cert_id, cert_data, session=None):
    """Update an existing certificate"""
    if session is None:
        with session_scope() as sess:
            cert = sess.query(Certificate).filter_by(id=cert_id).first()
            if cert is None:
                return False
            _update_cert_fields(cert, cert_data)
            return True
    cert = session.query(Certificate).filter_by(id=cert_id).first()
    if cert is None:
        return False
    _update_cert_fields(cert, cert_data)
    return True


def _update_cert_fields(cert, data):
    """Helper function to update certificate fields"""
    for field in ['name', 'department', 'position', 'certificate_name',
                  'certificate_number', 'email', 'phone', 'days_remaining',
                  'status', 'status_label', 'status_icon', 'status_color']:
        if field in data:
            setattr(cert, field, data[field])

    # Handle date fields separately
    if 'issue_date' in data and data['issue_date']:
        if isinstance(data['issue_date'], str):
            cert.issue_date = parse_date(data['issue_date'])
        else:
            cert.issue_date = data['issue_date']

    if 'expiry_date' in data and data['expiry_date']:
        if isinstance(data['expiry_date'], str):
            cert.expiry_date = parse_date(data['expiry_date'])
        else:
            cert.expiry_date = data['expiry_date']

    cert.updated_at = datetime.utcnow()


def delete_certificate(cert_id, session=None):
    """Delete a certificate by ID"""
    if session is None:
        with session_scope() as sess:
            cert = sess.query(Certificate).filter_by(id=cert_id).first()
            if cert is None:
                return False
            sess.delete(cert)
            return True
    cert = session.query(Certificate).filter_by(id=cert_id).first()
    if cert is None:
        return False
    session.delete(cert)
    return True


def save_certificates(certificates, metadata=None, session=None):
    """
    Save certificates to database (replace all existing data)

    Args:
        certificates: List of certificate dictionaries
        metadata: Upload metadata dictionary
        session: Optional database session

    Returns:
        True if successful
    """
    if session is None:
        with session_scope() as sess:
            _save_all(sess, certificates, metadata)
            return True
    _save_all(session, certificates, metadata)
    return True


def _save_all(session, certificates, metadata=None):
    """Helper to save all certificates within a session"""
    # Delete existing certificates
    session.query(Certificate).delete()

    # Add new certificates
    for cert_data in certificates:
        cert = Certificate.from_dict(cert_data)
        session.add(cert)

    # Update metadata if provided
    if metadata:
        session.query(UploadMetadata).delete()
        meta = UploadMetadata.from_dict(metadata)
        session.add(meta)


# Session state operations
def get_session_state(session=None):
    """Get current session state"""
    if session is None:
        with session_scope() as sess:
            state = sess.query(SessionState).filter_by(id=1).first()
            return state.to_dict() if state else {'active': False}
    state = session.query(SessionState).filter_by(id=1).first()
    return state.to_dict() if state else {'active': False}


def set_session_active(active=True, session=None):
    """Set session active state"""
    if session is None:
        with session_scope() as sess:
            state = sess.query(SessionState).filter_by(id=1).first()
            if state:
                state.active = active
                state.timestamp = datetime.utcnow()
            return
    state = session.query(SessionState).filter_by(id=1).first()
    if state:
        state.active = active
        state.timestamp = datetime.utcnow()


# Statistics
def get_statistics(session=None):
    """Get certificate statistics"""
    if session is None:
        with session_scope() as sess:
            return _calc_statistics(sess)
    return _calc_statistics(session)


def _calc_statistics(session):
    """Helper to calculate statistics"""
    total = session.query(Certificate).count()
    expired = session.query(Certificate).filter_by(status='expired').count()
    urgent = session.query(Certificate).filter_by(status='urgent').count()
    warning = session.query(Certificate).filter_by(status='warning').count()
    normal = session.query(Certificate).filter_by(status='normal').count()
    unknown = session.query(Certificate).filter_by(status='unknown').count()

    return {
        'total': total,
        'expired': expired,
        'urgent': urgent,
        'warning': warning,
        'normal': normal,
        'unknown': unknown,
    }


# Search/Filter operations
def search_certificates(search_term='', status_filter='all', session=None):
    """
    Search and filter certificates

    Args:
        search_term: Search string for name, department, or certificate name
        status_filter: Status filter (all, expired, urgent, warning, normal)
        session: Optional database session

    Returns:
        List of certificate dictionaries
    """
    if session is None:
        with session_scope() as sess:
            return _search(sess, search_term, status_filter)
    return _search(session, search_term, status_filter)


def _search(session, search_term, status_filter):
    """Helper to perform search query"""
    query = session.query(Certificate)

    # Apply search filter
    if search_term:
        term = f'%{search_term}%'
        query = query.filter(
            (Certificate.name.like(term)) |
            (Certificate.department.like(term)) |
            (Certificate.certificate_name.like(term)) |
            (Certificate.certificate_number.like(term))
        )

    # Apply status filter
    if status_filter and status_filter != 'all':
        query = query.filter(Certificate.status == status_filter)

    return [cert.to_dict() for cert in query.all()]


def get_certificates_by_days(days, session=None):
    """
    Get certificates due within specified days

    Args:
        days: Number of days
        session: Optional database session

    Returns:
        List of certificate dictionaries
    """
    if session is None:
        with session_scope() as sess:
            return _by_days(sess, days)
    return _by_days(session, days)


def _by_days(session, days):
    """Helper to get certificates by days"""
    certs = session.query(Certificate).filter(
        Certificate.days_remaining <= days,
        Certificate.days_remaining >= 0
    ).all()
    return [cert.to_dict() for cert in certs]


# ============ Notification Settings Operations ============

def get_notification_settings(session=None, include_password=False):
    """获取通知设置 / Get notification settings"""
    if session is None:
        with session_scope() as sess:
            settings = sess.query(NotificationSettings).first()
            return settings.to_dict(include_password=include_password) if settings else None
    settings = session.query(NotificationSettings).first()
    return settings.to_dict(include_password=include_password) if settings else None


def save_notification_settings(settings_data, session=None):
    """保存通知设置 / Save notification settings"""
    if session is None:
        with session_scope() as sess:
            return _save_settings(sess, settings_data)
    return _save_settings(session, settings_data)


def _save_settings(session, settings_data):
    """Helper to save settings"""
    settings = session.query(NotificationSettings).first()

    if settings:
        _update_model_fields(settings, settings_data, exclude_fields={'id'})
    else:
        # Create new
        settings = NotificationSettings(**{k: v for k, v in settings_data.items() if k != 'id'})
        session.add(settings)

    return True


# ============ Notification Record Operations ============

def create_notification_record(record_data, session=None):
    """创建通知记录 / Create notification record"""
    if session is None:
        with session_scope() as sess:
            record = NotificationRecord(**record_data)
            sess.add(record)
            sess.flush()
            return record.id
    record = NotificationRecord(**record_data)
    session.add(record)
    session.flush()
    return record.id


def get_notification_records(limit=100, offset=0, session=None):
    """获取通知记录列表 / Get notification records"""
    if session is None:
        with session_scope() as sess:
            records = sess.query(NotificationRecord)\
                .order_by(NotificationRecord.created_at.desc())\
                .limit(limit)\
                .offset(offset)\
                .all()
            return [r.to_dict() for r in records]

    records = session.query(NotificationRecord)\
        .order_by(NotificationRecord.created_at.desc())\
        .limit(limit)\
        .offset(offset)\
        .all()
    return [r.to_dict() for r in records]


def get_notification_records_by_status(status, session=None):
    """按状态获取通知记录 / Get notification records by status"""
    if session is None:
        with session_scope() as sess:
            records = sess.query(NotificationRecord)\
                .filter_by(send_status=status)\
                .order_by(NotificationRecord.created_at.desc())\
                .all()
            return [r.to_dict() for r in records]

    records = session.query(NotificationRecord)\
        .filter_by(send_status=status)\
        .order_by(NotificationRecord.created_at.desc())\
        .all()
    return [r.to_dict() for r in records]


def update_notification_record(record_id, update_data, session=None):
    """更新通知记录 / Update notification record"""
    if session is None:
        with session_scope() as sess:
            record = sess.query(NotificationRecord).filter_by(id=record_id).first()
            if record:
                _update_model_fields(record, update_data)
                return True
            return False

    record = session.query(NotificationRecord).filter_by(id=record_id).first()
    if record:
        _update_model_fields(record, update_data)
        return True
    return False


# ============ Notification Preferences Operations ============

def get_notification_preferences(session=None):
    """获取所有通知偏好 / Get all notification preferences"""
    if session is None:
        with session_scope() as sess:
            prefs = sess.query(NotificationPreference).all()
            return [p.to_dict() for p in prefs]
    prefs = session.query(NotificationPreference).all()
    return [p.to_dict() for p in prefs]


def get_notification_preference(email, session=None):
    """获取单个用户的通知偏好 / Get notification preference by email"""
    if session is None:
        with session_scope() as sess:
            pref = sess.query(NotificationPreference).filter_by(email=email).first()
            return pref.to_dict() if pref else None
    pref = session.query(NotificationPreference).filter_by(email=email).first()
    return pref.to_dict() if pref else None


def set_notification_preference(email, enabled=True, name=None, department=None, session=None):
    """设置用户通知偏好 / Set user notification preference"""
    if session is None:
        with session_scope() as sess:
            return _set_pref(sess, email, enabled, name, department)
    return _set_pref(session, email, enabled, name, department)


def _set_pref(session, email, enabled, name, department):
    """Helper to set preference"""
    pref = session.query(NotificationPreference).filter_by(email=email).first()
    if pref:
        pref.email_enabled = enabled
        if name:
            pref.name = name
        if department:
            pref.department = department
        pref.updated_at = datetime.utcnow()
    else:
        pref = NotificationPreference(
            email=email,
            email_enabled=enabled,
            name=name,
            department=department
        )
        session.add(pref)
    return True
