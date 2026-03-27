"""
SQLAlchemy ORM Models for Certificate Management System
证件管理系统数据库模型
"""
from datetime import datetime
from sqlalchemy import Column, String, Integer, Boolean, DateTime, Date, CheckConstraint
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.sql import func

Base = declarative_base()


class Certificate(Base):
    """证件表 / Certificates Table"""
    __tablename__ = 'certificates'

    # Primary key - use string ID to match existing UUID format
    id = Column(String(36), primary_key=True)

    # Basic information / 基本信息
    name = Column(String(100), nullable=False, index=True)  # 姓名
    department = Column(String(100), index=True)  # 部门
    position = Column(String(100))  # 岗位

    # Certificate information / 证件信息
    certificate_name = Column(String(100), nullable=False, index=True)  # 证件名称
    certificate_number = Column(String(100))  # 证件号码
    issue_date = Column(Date)  # 发证日期
    expiry_date = Column(Date, index=True)  # 到期日期

    # Contact information / 联系方式
    email = Column(String(200))  # 邮箱
    phone = Column(String(50))  # 手机号

    # Status fields / 状态字段
    days_remaining = Column(Integer)  # 剩余天数
    status = Column(String(20), index=True)  # status: expired, urgent, warning, normal, unknown
    status_label = Column(String(50))  # 状态标签
    status_icon = Column(String(10))  # 状态图标
    status_color = Column(String(20))  # 状态颜色

    # Timestamps / 时间戳
    created_at = Column(DateTime, default=datetime.utcnow)  # 创建时间
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)  # 更新时间

    def to_dict(self):
        """Convert model to dictionary / 转换为字典"""
        return {
            'id': self.id,
            'name': self.name,
            'department': self.department,
            'position': self.position,
            'certificate_name': self.certificate_name,
            'certificate_number': self.certificate_number,
            'issue_date': self.issue_date.isoformat() if self.issue_date else None,
            'expiry_date': self.expiry_date.isoformat() if self.expiry_date else None,
            'email': self.email,
            'phone': self.phone,
            'days_remaining': self.days_remaining,
            'status': self.status,
            'status_label': self.status_label,
            'status_icon': self.status_icon,
            'status_color': self.status_color,
            'created_at': self.created_at.isoformat() if self.created_at else None,
            'updated_at': self.updated_at.isoformat() if self.updated_at else None,
        }

    @classmethod
    def from_dict(cls, data):
        """Create model from dictionary / 从字典创建模型"""
        return cls(
            id=data.get('id'),
            name=data.get('name'),
            department=data.get('department'),
            position=data.get('position'),
            certificate_name=data.get('certificate_name'),
            certificate_number=data.get('certificate_number'),
            issue_date=datetime.fromisoformat(data['issue_date']) if data.get('issue_date') else None,
            expiry_date=datetime.fromisoformat(data['expiry_date']) if data.get('expiry_date') else None,
            email=data.get('email'),
            phone=data.get('phone'),
            days_remaining=data.get('days_remaining'),
            status=data.get('status'),
            status_label=data.get('status_label'),
            status_icon=data.get('status_icon'),
            status_color=data.get('status_color'),
            created_at=datetime.fromisoformat(data['created_at']) if data.get('created_at') else datetime.utcnow(),
            updated_at=datetime.fromisoformat(data['updated_at']) if data.get('updated_at') else datetime.utcnow(),
        )

    def __repr__(self):
        return f"<Certificate(id={self.id}, name={self.name}, certificate={self.certificate_name})>"


class UploadMetadata(Base):
    """上传元数据表 / Upload Metadata Table"""
    __tablename__ = 'upload_metadata'

    id = Column(Integer, primary_key=True, autoincrement=True)
    original_filename = Column(String(255), nullable=False)  # 原始文件名
    original_filepath = Column(String(500))  # 原始文件路径
    format_type = Column(String(20))  # 格式类型: simple, complex
    sheet_name = Column(String(100))  # 工作表名称
    upload_time = Column(DateTime, default=datetime.utcnow)  # 上传时间

    def to_dict(self):
        """Convert model to dictionary / 转换为字典"""
        return {
            'id': self.id,
            'original_filename': self.original_filename,
            'original_filepath': self.original_filepath,
            'format_type': self.format_type,
            'sheet_name': self.sheet_name,
            'upload_time': self.upload_time.isoformat() if self.upload_time else None,
        }

    @classmethod
    def from_dict(cls, data):
        """Create model from dictionary / 从字典创建模型"""
        return cls(
            id=data.get('id'),
            original_filename=data.get('original_filename'),
            original_filepath=data.get('original_filepath'),
            format_type=data.get('format_type'),
            sheet_name=data.get('sheet_name'),
            upload_time=datetime.fromisoformat(data['upload_time']) if data.get('upload_time') else datetime.utcnow(),
        )

    def __repr__(self):
        return f"<UploadMetadata(id={self.id}, filename={self.original_filename})>"


class SessionState(Base):
    """会话状态表 / Session State Table"""
    __tablename__ = 'session_state'

    # Single row table - always use id=1
    id = Column(Integer, primary_key=True)
    active = Column(Boolean, default=False)  # 会话是否激活
    timestamp = Column(DateTime, default=datetime.utcnow)  # 最后更新时间

    __table_args__ = (
        CheckConstraint('id = 1', name='check_single_row'),
    )

    def to_dict(self):
        """Convert model to dictionary / 转换为字典"""
        return {
            'active': self.active,
            'timestamp': self.timestamp.isoformat() if self.timestamp else None,
        }

    def __repr__(self):
        return f"<SessionState(id={self.id}, active={self.active})>"
