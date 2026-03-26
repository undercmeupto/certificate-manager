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
    DATA_FILE, UPLOAD_FOLDER, EXPORT_FOLDER,
    ALLOWED_EXTENSIONS, MAX_CONTENT_LENGTH,
    SECRET_KEY, DEBUG, HOST, PORT, STATUS_MAP,
    SESSION_STATE_FILE
)
from utils.certificate_checker import (
    parse_excel_file,
    get_sheet_names,
    save_to_json,
    load_from_json,
    export_to_excel,
    search_certificates,
    calculate_days_remaining,
    get_status_indicator,
    detect_excel_format
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

# ============ Helper Functions ============


# ============ Session State Management ============

SESSION_ACTIVE = False  # 全局会话状态


def get_session_state():
    """获取会话状态"""
    global SESSION_ACTIVE
    return SESSION_ACTIVE


def set_session_state(active):
    """设置会话状态"""
    global SESSION_ACTIVE
    SESSION_ACTIVE = active
    # 保存到文件
    try:
        os.makedirs(os.path.dirname(SESSION_STATE_FILE), exist_ok=True)
        with open(SESSION_STATE_FILE, 'w', encoding='utf-8') as f:
            json.dump({'active': active, 'timestamp': datetime.now().isoformat()}, f)
    except:
        pass


def load_session_state():
    """启动时加载会话状态"""
    global SESSION_ACTIVE
    try:
        if os.path.exists(SESSION_STATE_FILE):
            with open(SESSION_STATE_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                SESSION_ACTIVE = data.get('active', False)
    except:
        SESSION_ACTIVE = False


# 启动时加载会话状态
load_session_state()


def allowed_file(filename):
    """检查文件扩展名是否允许"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def get_all_certificates():
    """获取所有证件数据"""
    return load_from_json(DATA_FILE)


def get_metadata():
    """获取上传元数据"""
    try:
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            import json
            data = json.load(f)
            if isinstance(data, dict) and 'metadata' in data:
                return data['metadata']
        return None
    except:
        return None


def save_certificates(certificates, metadata=None):
    """保存证件数据，自动保留现有元数据"""
    if metadata is None:
        # 如果没有提供新元数据，尝试保留现有元数据
        existing_data = load_json_with_metadata(DATA_FILE)
        if isinstance(existing_data, dict) and 'metadata' in existing_data:
            metadata = existing_data['metadata']
    return save_to_json(certificates, DATA_FILE, metadata=metadata)


def load_json_with_metadata(filepath: str):
    """加载JSON文件（包含元数据）"""
    if not os.path.exists(filepath):
        return {}
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            import json
            return json.load(f)
    except:
        return {}


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


def get_statistics(certificates):
    """获取统计数据"""
    stats = {
        'total': len(certificates),
        'expired': 0,
        'urgent': 0,
        'warning': 0,
        'normal': 0
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
        stats = get_statistics(certificates)
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
        certificates = get_all_certificates()
        cert = next((c for c in certificates if c.get('id') == cert_id), None)
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

        # 保存
        certificates = get_all_certificates()
        certificates.append(cert)
        save_certificates(certificates)

        return jsonify({'success': True, 'data': cert})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/data/<cert_id>', methods=['PUT'])
def update_certificate(cert_id):
    """更新证件"""
    # 检查会话状态
    if not get_session_state():
        return jsonify({'success': False, 'error': '会话已退出，请先上传Excel文件'}), 403
    try:
        data = request.get_json()
        certificates = get_all_certificates()

        # 查找证件
        cert_index = next((i for i, c in enumerate(certificates) if c.get('id') == cert_id), None)
        if cert_index is None:
            return jsonify({'success': False, 'error': '证件不存在'}), 404

        # 更新字段
        cert = certificates[cert_index]
        for field in ['name', 'department', 'position', 'certificate_name', 'certificate_number',
                      'expiry_date', 'issue_date', 'email', 'phone']:
            if field in data:
                cert[field] = data[field]

        cert['updated_at'] = datetime.now().isoformat()

        # 重新计算状态
        cert = update_certificate_status(cert)
        certificates[cert_index] = cert

        # 保存
        save_certificates(certificates)

        return jsonify({'success': True, 'data': cert})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/data/<cert_id>', methods=['DELETE'])
def delete_certificate(cert_id):
    """删除证件"""
    # 检查会话状态
    if not get_session_state():
        return jsonify({'success': False, 'error': '会话已退出，请先上传Excel文件'}), 403
    try:
        certificates = get_all_certificates()
        cert_index = next((i for i, c in enumerate(certificates) if c.get('id') == cert_id), None)

        if cert_index is None:
            return jsonify({'success': False, 'error': '证件不存在'}), 404

        deleted_cert = certificates.pop(cert_index)
        save_certificates(certificates)

        return jsonify({'success': True, 'data': deleted_cert})
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

        certificates = get_all_certificates()
        results = search_certificates(certificates, search_term, status_filter)

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
        if export_to_excel(certificates, filepath):
            return send_file(filepath, as_attachment=True, download_name=filename)

        return jsonify({'success': False, 'error': '导出失败'}), 500
    except Exception as e:
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
        if export_to_excel(warning_certificates, filepath):
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

        # 获取所有证件数据
        certificates = get_all_certificates()

        if not certificates:
            return jsonify({'success': False, 'error': '没有数据可导出'}), 400

        # 过滤出指定天数内到期的证件（包括已过期）
        filtered_certificates = [
            c for c in certificates
            if c.get('days_remaining') is not None and c.get('days_remaining') <= days
        ]

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
        if export_to_excel(filtered_certificates, filepath):
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
        from utils.certificate_checker import export_updated_original
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
        certificates = get_all_certificates()
        stats = get_statistics(certificates)
        return jsonify({'success': True, 'data': stats})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/session/status', methods=['GET'])
def session_status():
    """获取会话状态"""
    return jsonify({
        'success': True,
        'active': get_session_state(),
        'has_data': len(get_all_certificates()) > 0
    })


# ============ Error Handlers ============

@app.errorhandler(413)
def request_entity_too_large(error):
    return jsonify({'success': False, 'error': '文件大小超过限制（最大16MB）'}), 413


@app.errorhandler(500)
def internal_server_error(error):
    return jsonify({'success': False, 'error': '服务器内部错误'}), 500


# ============ Main ============

if __name__ == '__main__':
    print(f"""
    ╔═══════════════════════════════════════════════════════════╗
    ║           证件管理系统 - Certificate Management          ║
    ╠═══════════════════════════════════════════════════════════╣
    ║  启动地址: http://localhost:{PORT}                          ║
    ║  数据文件: {DATA_FILE}     ║
    ╚═══════════════════════════════════════════════════════════╝
    """)
    app.run(host=HOST, port=PORT, debug=DEBUG)
