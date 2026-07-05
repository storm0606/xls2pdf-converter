import os, io, tempfile, time, threading, uuid, hashlib, json, re
from flask import Flask, request, jsonify, send_file, Response
from werkzeug.utils import secure_filename
from converter import xls_to_pdf

app = Flask(__name__)

# 限制上传文件大小为 50MB
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

_export_store = {}
_export_lock = threading.Lock()
_EXPORT_STORE_MAX = 200  # 最大缓存条目数

def _cleanup_export(token):
    with _export_lock:
        _export_store.pop(token, None)

def _sanitize_route_name(name):
    """消毒 route_name，防止 CRLF 注入和路径遍历"""
    if not name:
        return 'export'
    # 移除换行符和路径分隔符
    name = re.sub(r'[\r\n\\\/]', '', name)
    # 只保留安全字符
    name = re.sub(r'[^a-zA-Z0-9_\-\.\u4e00-\u9fff]', '_', name)
    # 限制长度
    return name[:100] if name else 'export'

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    f = request.files['file']
    route_name = _sanitize_route_name(request.form.get('route_name', 'export'))
    
    # 保存上传的XLS到临时文件
    fd, tmp_xls = tempfile.mkstemp(suffix='.xls')
    try:
        os.close(fd)
        f.save(tmp_xls)
        
        # 转换PDF
        fd2, tmp_pdf = tempfile.mkstemp(suffix='.pdf')
        os.close(fd2)
        try:
            result = xls_to_pdf(tmp_xls, tmp_pdf)
            
            with open(tmp_pdf, 'rb') as pf:
                pdf_bytes = pf.read()
            
            token = uuid.uuid4().hex
            filename = f'{route_name}.pdf'
            
            with _export_lock:
                # 限制缓存条目数
                if len(_export_store) >= _EXPORT_STORE_MAX:
                    return jsonify({'error': 'Server busy, try again later'}), 503
                _export_store[token] = (pdf_bytes, filename, 'application/pdf')
            
            # 5分钟后清理
            threading.Timer(300, _cleanup_export, args=[token]).start()
            
            return jsonify({'token': token, 'size': len(pdf_bytes)})
        finally:
            if os.path.exists(tmp_pdf):
                os.unlink(tmp_pdf)
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        if os.path.exists(tmp_xls):
            os.unlink(tmp_xls)

@app.route('/download/<token>.pdf')
def download(token):
    # 验证 token 格式，防止路径遍历
    if not re.match(r'^[a-f0-9]{32}$', token):
        return jsonify({'error': 'Invalid token format'}), 400
    
    with _export_lock:
        data = _export_store.get(token)
    if not data:
        return jsonify({'error': 'Token expired or invalid'}), 404
    
    pdf_bytes, filename, mime = data
    return send_file(io.BytesIO(pdf_bytes), mimetype=mime, as_attachment=True, download_name=filename)

@app.route('/convert/direct', methods=['POST'])
def convert_direct():
    """上传XLS，直接返回PDF二进制（不经过token中转）"""
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400

    f = request.files['file']
    route_name = _sanitize_route_name(request.form.get('route_name', 'export'))

    fd, tmp_xls = tempfile.mkstemp(suffix='.xls')
    try:
        os.close(fd)
        f.save(tmp_xls)

        fd2, tmp_pdf = tempfile.mkstemp(suffix='.pdf')
        os.close(fd2)
        try:
            result = xls_to_pdf(tmp_xls, tmp_pdf)

            with open(tmp_pdf, 'rb') as pf:
                pdf_bytes = pf.read()

            return Response(
                pdf_bytes,
                mimetype='application/pdf',
                headers={'Content-Disposition': f'attachment; filename="{route_name}.pdf"'}
            )
        finally:
            if os.path.exists(tmp_pdf):
                os.unlink(tmp_pdf)
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        if os.path.exists(tmp_xls):
            os.unlink(tmp_xls)


@app.route('/api/info')
def api_info():
    """返回服务信息"""
    # 列出可用字体
    font_dir = os.path.join(os.path.dirname(__file__), 'fonts')
    fonts_available = []
    if os.path.isdir(font_dir):
        for fname in os.listdir(font_dir):
            if fname.lower().endswith(('.ttf', '.ttc', '.otf')):
                fonts_available.append(fname)

    return jsonify({
        'version': '4.0',
        'supported_formats': ['xls', 'xlsx (via xlrd)'],
        'fonts_available': sorted(fonts_available),
        'endpoints': [
            {'method': 'POST', 'path': '/convert', 'desc': '上传XLS，返回JSON {token, size}'},
            {'method': 'GET',  'path': '/download/<token>.pdf', 'desc': '用token下载PDF'},
            {'method': 'POST', 'path': '/convert/direct', 'desc': '上传XLS，直接返回PDF二进制'},
            {'method': 'GET',  'path': '/api/info', 'desc': '返回服务信息'},
            {'method': 'GET',  'path': '/health', 'desc': '健康检查'},
        ]
    })


@app.route('/health')
def health():
    return jsonify({'status': 'ok'})

@app.errorhandler(413)
def request_entity_too_large(error):
    return jsonify({'error': 'File too large. Maximum size is 50MB.'}), 413

if __name__ == '__main__':
    os.makedirs(os.path.join(os.path.dirname(__file__), 'fonts'), exist_ok=True)
    app.run(host='0.0.0.0', port=28891, debug=False)
