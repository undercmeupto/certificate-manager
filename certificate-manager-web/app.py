"""
证件管理系统 - Flask后端
Web-based certificate management with industrial UI
"""
import os
import uuid
import json
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file, session
from flask_cors import CORS
from werkzeug.utils import secure_filename

from config import (
    DATABASE_FILE, DATA_FILE, UPLOAD_FOLDER, EXPORT_FOLDER,
    ALLOWED_EXTENSIONS, MAX_CONTENT_LENGTH,
    SECRET_KEY, DEBUG, HOST, PORT, STATUS_MAP,
    SESSION_STATE_FILE
)
from utils.certificate_checker import (
    parse_excel_file,
    get_sheet_names,
    export_to_excel,
    calculate_days_remaining,
    get_status_indicator,
    detect_excel_format,
    export_updated_original
)
from database import (
    init_db,
    get_session,
    session_scope,
    get_all_certificates as db_get_all_certificates,
    get_certificate_by_id as db_get_certificate_by_id,
    add_certificate as db_add_certificate,
    update_certificate as db_update_certificate,
    delete_certificate as db_delete_certificate,
    save_certificates as db_save_certificates,
    get_session_state as db_get_session_state,
    set_session_active as db_set_session_active,
    get_statistics as db_get_statistics,
    search_certificates as db_search_certificates,
    get_certificates_by_days as db_get_certificates_by_days,
)

# ============ Flask App Setup ============
app = Flask(__name__)
app.config['SECRET_KEY'] = SECRET_KEY
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['EXPORT_FOLDER'] = EXPORT_FOLDER

CORS(app)

# 确保目录存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(EXPORT_FOLDER, exist_ok=True)
os.makedirs(os.path.dirname(DATA_FILE), exist_ok=True)

# Initialize database on startup
init_db()

# ============ Helper Functions ============


# ============ Session State Management ============

def get_session_state():
    """获取会话状态（从数据库）"""
    state = db_get_session_state()
    return state.get('active', False)


def set_session_state(active):
    """设置会话状态（保存到数据库）"""
    db_set_session_active(active)


def allowed_file(filename):
    """检查文件扩展名是否允许"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def get_all_certificates():
    """获取所有证件数据（从数据库）"""
    return db_get_all_certificates()


def get_metadata():
    """获取上传元数据（从数据库）"""
    from models import UploadMetadata
    with session_scope() as session:
        meta = session.query(UploadMetadata).first()
        if meta:
            return meta.to_dict()
        return None


def save_certificates(certificates, metadata=None):
    """保存证件数据到数据库"""
    return db_save_certificates(certificates, metadata=metadata)


def update_certificate_status(cert):
    """更新证件状态信息"""
    days = cert.get('days_remaining')
    if days is None:
        expiry_date = cert.get('expiry_date')
        if expiry_date:
            days = calculate_days_remaining(expiry_date)
            cert['days_remaining'] = days

    status_info = get_status_indicator(days)
    cert['status'] = status_info['status']
    cert['status_label'] = status_info['label']
    cert['status_icon'] = status_info['icon']
    cert['status_color'] = status_info['color']
    return cert


def get_statistics(certificates=None):
    """获取统计数据（从数据库或提供的列表）"""
    if certificates is None:
        # 从数据库获取统计数据
        stats = db_get_statistics()
        # Add 'unknown' count for compatibility
        if 'unknown' not in stats:
            stats['unknown'] = 0
        return stats
    else:
        # 从提供的列表计算统计数据
        stats = {
            'total': len(certificates),
            'expired': 0,
            'urgent': 0,
            'warning': 0,
            'normal': 0,
            'unknown': 0
        }
        for cert in certificates:
            status = cert.get('status', 'normal')
            if status in stats:
                stats[status] += 1
        return stats


# ============ Routes ============

@app.route('/')
def index():
    """Serve SPA"""
    return render_template('index.html')


@app.route('/api/data', methods=['GET'])
def get_data():
    """获取所有证件数据"""
    try:
        certificates = get_all_certificates()
        stats = get_statistics()  # 从数据库获取统计数据
        return jsonify({
            'success': True,
            'data': certificates,
            'statistics': stats
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/data/<cert_id>', methods=['GET'])
def get_certificate(cert_id):
    """获取单个证件"""
    try:
        cert = db_get_certificate_by_id(cert_id)
        if cert:
            return jsonify({'success': True, 'data': cert})
        return jsonify({'success': False, 'error': '证件不存在'}), 404
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/data', methods=['POST'])
def add_certificate():
    """添加新证件"""
    # 检查会话状态
    if not get_session_state():
        return jsonify({'success': False, 'error': '会话已退出，请先上传Excel文件'}), 403
    try:
        data = request.get_json()

        # 验证必填字段
        required_fields = ['name', 'department', 'position', 'certificate_name', 'issue_date', 'expiry_date']
        for field in required_fields:
            if not data.get(field):
                return jsonify({'success': False, 'error': f'缺少必填字段: {field}'}), 400

        # 创建新证件
        cert = {
            'id': str(uuid.uuid4()),
            'name': data['name'],
            'department': data['department'],
            'position': data['position'],
            'certificate_name': data['certificate_name'],
            'certificate_number': data.get('certificate_number', ''),
            'expiry_date': data['expiry_date'],
            'issue_date': data.get('issue_date', ''),
            'email': data.get('email', ''),
            'phone': data.get('phone', ''),
            'created_at': datetime.now().isoformat()
        }

        # 计算状态
        cert = update_certificate_status(cert)

        # 保存到数据库
        cert_id = db_add_certificate(cert)
        cert['id'] = cert_id

        return jsonify({'success': True, 'data': cert})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/data/<cert_id>', methods=['PUT'])
def update_certificate_endpoint(cert_id):
    """更新证件"""
    # 检查会话状态
    if not get_session_state():
        return jsonify({'success': False, 'error': '会话已退出，请先上传Excel文件'}), 403
    try:
        data = request.get_json()

        # 验证证件是否存在
        cert = db_get_certificate_by_id(cert_id)
        if not cert:
            return jsonify({'success': False, 'error': '证件不存在'}), 404

        # 更新字段
        for field in ['name', 'department', 'position', 'certificate_name', 'certificate_number',
                      'expiry_date', 'issue_date', 'email', 'phone']:
            if field in data:
                cert[field] = data[field]

        cert['updated_at'] = datetime.now().isoformat()

        # 重新计算状态
        cert = update_certificate_status(cert)

        # 保存到数据库
        db_update_certificate(cert_id, cert)

        return jsonify({'success': True, 'data': cert})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/data/<cert_id>', methods=['DELETE'])
def delete_certificate_endpoint(cert_id):
    """删除证件"""
    # 检查会话状态
    if not get_session_state():
        return jsonify({'success': False, 'error': '会话已退出，请先上传Excel文件'}), 403
    try:
        # 先获取要删除的证件
        cert = db_get_certificate_by_id(cert_id)
        if not cert:
            return jsonify({'success': False, 'error': '证件不存在'}), 404

        # 从数据库删除
        db_delete_certificate(cert_id)

        return jsonify({'success': True, 'data': cert})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/upload', methods=['POST'])
def upload_excel():
    """上传并解析Excel文件"""
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': '未选择文件'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'success': False, 'error': '未选择文件'}), 400

        if not allowed_file(file.filename):
            return jsonify({'success': False, 'error': '不支持的文件格式'}), 400

        # 保存文件（保留原始文件用于导出）
        filename = secure_filename(file.filename)
        original_filename = filename
        filepath = os.path.join(UPLOAD_FOLDER, f"original_{filename}")
        file.save(filepath)

        # 获取sheet参数
        sheet_name = request.form.get('sheet')

        # 解析Excel
        new_certificates, stats = parse_excel_file(filepath, sheet_name)

        # 检测Excel格式类型
        format_type = detect_excel_format(filepath)

        # 添加ID和原始文件信息
        for cert in new_certificates:
            cert['id'] = str(uuid.uuid4())
            cert['created_at'] = datetime.now().isoformat()

        # 保存原始文件路径信息到JSON
        metadata = {
            'original_filename': original_filename,
            'original_filepath': filepath,
            'format_type': format_type,
            'sheet_name': sheet_name,
            'upload_time': datetime.now().isoformat()
        }

        # 保存数据（包含元数据）
        save_certificates(new_certificates, metadata=metadata)

        # 激活会话
        set_session_state(True)

        return jsonify({
            'success': True,
            'message': f'成功导入 {len(new_certificates)} 条记录（已清除旧数据）',
            'imported': len(new_certificates),
            'statistics': get_statistics(new_certificates),
            'has_original': True
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/sheets', methods=['POST'])
def get_sheets():
    """获取Excel文件的sheet列表"""
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': '未选择文件'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'success': False, 'error': '未选择文件'}), 400

        # 保存临时文件
        filename = secure_filename(file.filename)
        filepath = os.path.join(UPLOAD_FOLDER, f"temp_{uuid.uuid4()}_{filename}")
        file.save(filepath)

        # 获取sheet列表
        sheets = get_sheet_names(filepath)

        # 删除临时文件
        try:
            os.remove(filepath)
        except:
            pass

        return jsonify({'success': True, 'sheets': sheets})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/search', methods=['POST'])
def search_data():
    """搜索证件"""
    try:
        data = request.get_json()
        search_term = data.get('search', '')
        status_filter = data.get('status', '')

        # 使用数据库搜索
        results = db_search_certificates(search_term, status_filter)

        return jsonify({
            'success': True,
            'data': results,
            'count': len(results)
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/export', methods=['GET'])
def export_data():
    """导出Excel"""
    try:
        certificates = get_all_certificates()

        if not certificates:
            return jsonify({'success': False, 'error': '没有数据可导出'}), 400

        # 生成文件名
        filename = f"certificates_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        filepath = os.path.join(EXPORT_FOLDER, filename)

        # 导出
        print(f"DEBUG: Calling export_to_excel with filepath={filepath}, export_type=all")
        result = export_to_excel(certificates, filepath, export_type='all')
        print(f"DEBUG: Export result={result}")
        if result:
            return send_file(filepath, as_attachment=True, download_name=filename)

        return jsonify({'success': False, 'error': '导出失败'}), 500
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/export/warning', methods=['GET'])
def export_warning_data():
    """导出预警Excel（仅包含已过期、紧急、预警状态的证件）"""
    try:
        certificates = get_all_certificates()

        if not certificates:
            return jsonify({'success': False, 'error': '没有数据可导出'}), 400

        # 过滤出需要预警的记录（已过期、紧急、预警）
        warning_certificates = [c for c in certificates if c.get('status') in ['expired', 'urgent', 'warning']]

        # 按优先级排序：紧急 > 预警 > 已过期
        status_priority = {'urgent': 1, 'warning': 2, 'expired': 3}
        warning_certificates.sort(key=lambda x: (status_priority.get(x.get('status', 'normal'), 999), x.get('days_remaining', 999)))

        if not warning_certificates:
            return jsonify({'success': False, 'error': '没有需要预警的记录'}), 400

        # 生成文件名
        filename = f"certificates_warning_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        filepath = os.path.join(EXPORT_FOLDER, filename)

        # 导出
        if export_to_excel(warning_certificates, filepath, export_type='warning'):
            return send_file(filepath, as_attachment=True, download_name=filename)

        return jsonify({'success': False, 'error': '导出失败'}), 500
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/export/by-days', methods=['GET'])
def export_by_days():
    """导出指定天数内到期的证件"""
    try:
        # 获取并验证天数参数
        days_str = request.args.get('days')
        if not days_str:
            return jsonify({'success': False, 'error': '缺少天数参数'}), 400

        try:
            days = int(days_str)
            if days <= 0:
                return jsonify({'success': False, 'error': '天数必须大于0'}), 400
        except ValueError:
            return jsonify({'success': False, 'error': '天数必须是有效的整数'}), 400

        # 使用数据库查询获取指定天数内到期的证件
        filtered_certificates = db_get_certificates_by_days(days)

        if not filtered_certificates:
            return jsonify({
                'success': False,
                'error': f'没有在 {days} 天内到期的证件'
            }), 400

        # 按剩余天数排序（升序，最紧急的在前）
        filtered_certificates.sort(key=lambda x: x.get('days_remaining', 999))

        # 生成文件名
        filename = f"certificates_{days}days_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        filepath = os.path.join(EXPORT_FOLDER, filename)

        # 导出（复用现有函数）
        if export_to_excel(filtered_certificates, filepath, export_type='by_days', days=days):
            return send_file(filepath, as_attachment=True, download_name=filename)

        return jsonify({'success': False, 'error': '导出失败'}), 500
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/export/original-updated', methods=['GET'])
def export_original_updated():
    """导出更新后的原表格（保留原始Excel格式，使用当前数据更新）"""
    try:
        metadata = get_metadata()
        if not metadata or not metadata.get('original_filepath'):
            return jsonify({'success': False, 'error': '没有原始上传文件，请先上传Excel'}), 400

        original_filepath = metadata['original_filepath']
        # 规范化路径（处理Windows反斜杠问题）
        original_filepath = os.path.normpath(original_filepath)
        # 如果是相对路径，转换为绝对路径
        if not os.path.isabs(original_filepath):
            base_dir = os.path.dirname(os.path.abspath(__file__))
            original_filepath = os.path.join(base_dir, original_filepath)
            original_filepath = os.path.normpath(original_filepath)

        if not os.path.exists(original_filepath):
            return jsonify({'success': False, 'error': f'原始文件不存在: {original_filepath}'}), 400

        certificates = get_all_certificates()
        if not certificates:
            return jsonify({'success': False, 'error': '没有数据可导出'}), 400

        # 生成文件名（带时间戳，避免缓存）
        original_name = metadata.get('original_filename', 'original.xlsx')
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        # 移除扩展名并添加时间戳（一次性解包）
        name_without_ext, ext = os.path.splitext(original_name)
        filename = f"updated_{name_without_ext}_{timestamp}{ext}"
        filepath = os.path.join(EXPORT_FOLDER, filename)

        # 使用原始格式更新导出
        if export_updated_original(original_filepath, filepath, certificates, metadata):
            return send_file(filepath, as_attachment=True, download_name=filename)

        return jsonify({'success': False, 'error': '导出失败'}), 500
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/exit', methods=['POST'])
def exit_session():
    """退出会话（清除数据并锁定，需重新上传文件）"""
    try:
        # 清空数据文件
        save_certificates([], metadata=None)

        # 停用会话
        set_session_state(False)

        return jsonify({
            'success': True,
            'message': '会话已退出，数据已清除，请上传新文件继续操作'
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/statistics', methods=['GET'])
def get_stats():
    """获取统计数据"""
    try:
        stats = get_statistics()  # 从数据库获取统计数据
        return jsonify({'success': True, 'data': stats})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/session/status', methods=['GET'])
def session_status():
    """获取会话状态"""
    from database import get_statistics as db_get_stats
    stats = db_get_stats()
    return jsonify({
        'success': True,
        'active': get_session_state(),
        'has_data': stats.get('total', 0) > 0
    })


# ============ Error Handlers ============

@app.errorhandler(413)
def request_entity_too_large(error):
    return jsonify({'success': False, 'error': '文件大小超过限制（最大16MB）'}), 413


@app.errorhandler(500)
def internal_server_error(error):
    return jsonify({'success': False, 'error': '服务器内部错误'}), 500


# ============ Main ============

def open_browser():
    """自动打开浏览器（仅在打包后的exe中执行）"""
    import sys
    import webbrowser
    from threading import Timer

    # 只在打包模式下自动打开浏览器
    if getattr(sys, 'frozen', False):
        Timer(1.5, lambda: webbrowser.open(f'http://localhost:{PORT}')).start()

if __name__ == '__main__':
    # 自动打开浏览器（exe模式）
    open_browser()

    print(f"""
    ╔═══════════════════════════════════════════════════════════╗
    ║           证件管理系统 - Certificate Management          ║
    ╠═══════════════════════════════════════════════════════════╣
    ║  启动地址: http://localhost:{PORT}                          ║
    ║  数据库: {DATABASE_FILE}     ║
    ╚═══════════════════════════════════════════════════════════╝
    """)
    app.run(host=HOST, port=PORT, debug=DEBUG)
