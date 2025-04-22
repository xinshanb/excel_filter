from flask import Flask, request, send_file, jsonify
from werkzeug.utils import secure_filename
import pandas as pd
import os

app = Flask(__name__)

def init_upload_folder():
    upload_folder = '/tmp'
    if not os.path.exists(upload_folder):
        os.makedirs(upload_folder, exist_ok=True)
    return upload_folder

UPLOAD_FOLDER = init_upload_folder()
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET'])
def home():
    return '''
    <html>
        <head>
            <title>Excel文件过滤工具</title>
            <style>
                body { font-family: sans-serif; max-width: 800px; margin: 0 auto; padding: 20px; }
                .container { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
                h1 { color: #333; text-align: center; }
                .form-group { margin-bottom: 15px; }
                label { display: block; margin-bottom: 5px; }
                input, select { width: 100%; padding: 8px; margin-bottom: 10px; }
                button { background: #4CAF50; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; }
                button:hover { background: #45a049; }
            </style>
        </head>
        <body>
            <div class="container">
                <h1>Excel文件过滤工具</h1>
                <div class="form-group">
                    <label>选择Excel文件：</label>
                    <input type="file" id="file" accept=".xlsx,.xls">
                </div>
                <div id="columnSelect" style="display:none;">
                    <label>选择列：</label>
                    <select id="column"></select>
                </div>
                <div id="keywordInput" style="display:none;">
                    <label>输入关键词（用逗号分隔）：</label>
                    <input type="text" id="keywords">
                </div>
                <button onclick="processFile()" id="processBtn" disabled>处理文件</button>
                <div id="message"></div>
            </div>
            <script>
                let currentFilePath = '';
                
                document.getElementById('file').addEventListener('change', async (e) => {
                    const file = e.target.files[0];
                    if (!file) return;
                    
                    const formData = new FormData();
                    formData.append('file', file);
                    
                    try {
                        const response = await fetch('/api/upload', {
                            method: 'POST',
                            body: formData
                        });
                        
                        const data = await response.json();
                        if (response.ok) {
                            currentFilePath = data.filepath;
                            const select = document.getElementById('column');
                            select.innerHTML = data.columns.map(col => 
                                `<option value="${col}">${col}</option>`
                            ).join('');
                            document.getElementById('columnSelect').style.display = 'block';
                            document.getElementById('keywordInput').style.display = 'block';
                            document.getElementById('processBtn').disabled = false;
                        } else {
                            throw new Error(data.error);
                        }
                    } catch (error) {
                        document.getElementById('message').textContent = `错误：${error.message}`;
                    }
                });
                
                async function processFile() {
                    const column = document.getElementById('column').value;
                    const keywords = document.getElementById('keywords').value;
                    
                    if (!keywords.trim()) {
                        document.getElementById('message').textContent = '请输入关键词';
                        return;
                    }
                    
                    const formData = new FormData();
                    formData.append('filepath', currentFilePath);
                    formData.append('column', column);
                    formData.append('keywords', keywords);
                    
                    try {
                        const response = await fetch('/api/filter', {
                            method: 'POST',
                            body: formData
                        });
                        
                        if (response.ok) {
                            const blob = await response.blob();
                            const url = window.URL.createObjectURL(blob);
                            const a = document.createElement('a');
                            a.href = url;
                            a.download = 'filtered_excel.xlsx';
                            document.body.appendChild(a);
                            a.click();
                            window.URL.revokeObjectURL(url);
                            document.body.removeChild(a);
                            document.getElementById('message').textContent = '文件已成功过滤并下载！';
                        } else {
                            const data = await response.json();
                            throw new Error(data.error);
                        }
                    } catch (error) {
                        document.getElementById('message').textContent = `错误：${error.message}`;
                    }
                }
            </script>
        </body>
    </html>
    '''

@app.route('/api/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': '没有文件被上传'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '没有选择文件'}), 400
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        file.save(filepath)
        
        try:
            df = pd.read_excel(filepath)
            columns = df.columns.tolist()
            return jsonify({'columns': columns, 'filepath': filepath})
        except Exception as e:
            return jsonify({'error': str(e)}), 500
        finally:
            if os.path.exists(filepath):
                os.remove(filepath)
    
    return jsonify({'error': '不支持的文件类型'}), 400

@app.route('/api/filter', methods=['POST'])
def filter_excel():
    filepath = request.form.get('filepath')
    column = request.form.get('column')
    keywords = request.form.get('keywords')
    
    if not all([filepath, column, keywords]):
        return jsonify({'error': '缺少必要参数'}), 400
    
    try:
        df = pd.read_excel(filepath)
        keywords_list = [k.strip() for k in keywords.split(',')]
        
        mask = df[column].astype(str).apply(lambda x: any(k.lower() in x.lower() for k in keywords_list))
        filtered_df = df[mask]
        
        output_filename = 'filtered_' + os.path.basename(filepath)
        output_filepath = os.path.join(UPLOAD_FOLDER, output_filename)
        filtered_df.to_excel(output_filepath, index=False)
        
        try:
            return send_file(
                output_filepath,
                as_attachment=True,
                download_name=output_filename,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        finally:
            if os.path.exists(output_filepath):
                os.remove(output_filepath)
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500 