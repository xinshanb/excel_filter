from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)

# 设置上传文件夹路径
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': '没有文件被上传'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '没有选择文件'}), 400
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        try:
            df = pd.read_excel(filepath)
            columns = df.columns.tolist()
            return jsonify({'columns': columns, 'filepath': filepath})
        except Exception as e:
            return jsonify({'error': f'处理Excel文件时出错: {str(e)}'}), 500
    
    return jsonify({'error': '不支持的文件类型'}), 400

@app.route('/filter', methods=['POST'])
def filter_excel():
    filepath = request.form.get('filepath')
    column = request.form.get('column')
    keywords = request.form.get('keywords')
    
    if not all([filepath, column, keywords]):
        return jsonify({'error': '缺少必要参数'}), 400
    
    try:
        df = pd.read_excel(filepath)
        keywords_list = [k.strip() for k in keywords.split(',')]
        
        # 创建过滤条件
        mask = df[column].astype(str).apply(lambda x: any(k.lower() in x.lower() for k in keywords_list))
        filtered_df = df[mask]
        
        # 保存过滤后的文件
        output_filename = 'filtered_' + os.path.basename(filepath)
        output_filepath = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
        filtered_df.to_excel(output_filepath, index=False)
        
        return send_file(
            output_filepath,
            as_attachment=True,
            download_name=output_filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    
    except Exception as e:
        return jsonify({'error': f'处理文件时出错: {str(e)}'}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True) 