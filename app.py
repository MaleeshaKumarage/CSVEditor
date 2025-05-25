import os
from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import tempfile
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

ALLOWED_EXTENSIONS = {'csv', 'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_headers(file_path):
    if file_path.endswith('.csv'):
        df = pd.read_csv(file_path)
    else:
        df = pd.read_excel(file_path)
    return df.columns.tolist()

def process_filter(file_path, filters, max_records=None):
    if file_path.endswith('.csv'):
        df = pd.read_csv(file_path)
    else:
        df = pd.read_excel(file_path)
    
    for col, condition, value in filters:
        if not condition or not value:
            continue
        try:
            if condition == '==':
                df = df[df[col].astype(str) == value]
            elif condition == '>':
                df = df[df[col].astype(float) > float(value)]
            elif condition == '<':
                df = df[df[col].astype(float) < float(value)]
            elif condition == '>=':
                df = df[df[col].astype(float) >= float(value)]
            elif condition == '<=':
                df = df[df[col].astype(float) <= float(value)]
            elif condition == 'contains':
                df = df[df[col].astype(str).str.contains(value, case=False, na=False)]
            elif condition == 'not contains':
                df = df[~df[col].astype(str).str.contains(value, case=False, na=False)]
        except Exception:
            continue
    
    if max_records:
        df = df.head(max_records)
    
    return df

def process_update(file_path, updates, max_records=None):
    if file_path.endswith('.csv'):
        df = pd.read_csv(file_path)
    else:
        df = pd.read_excel(file_path)
    
    for col, condition, filter_value, new_value in updates:
        if not condition or not filter_value:
            continue
        try:
            mask = pd.Series(False, index=df.index)
            if condition == '==':
                mask = df[col].astype(str) == filter_value
            elif condition == '>':
                mask = df[col].astype(float) > float(filter_value)
            elif condition == '<':
                mask = df[col].astype(float) < float(filter_value)
            elif condition == '>=':
                mask = df[col].astype(float) >= float(filter_value)
            elif condition == '<=':
                mask = df[col].astype(float) <= float(filter_value)
            elif condition == 'contains':
                mask = df[col].astype(str).str.contains(filter_value, case=False, na=False)
            elif condition == 'not contains':
                mask = ~df[col].astype(str).str.contains(filter_value, case=False, na=False)
            
            df.loc[mask, col] = new_value
        except Exception:
            continue
    
    if max_records:
        df = df.head(max_records)
    
    return df

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    if not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file type'}), 400
    
    filename = secure_filename(file.filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(file_path)
    
    try:
        headers = get_headers(file_path)
        return jsonify({
            'success': True,
            'headers': headers,
            'filename': filename
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/filter', methods=['POST'])
def filter_file():
    data = request.json
    filename = data.get('filename')
    filters = data.get('filters', [])
    max_records = data.get('max_records')
    
    if not filename:
        return jsonify({'error': 'No filename provided'}), 400
    
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if not os.path.exists(file_path):
        return jsonify({'error': 'File not found'}), 404
    
    try:
        df = process_filter(file_path, filters, max_records)
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], f'filtered_{filename}')
        
        if filename.endswith('.csv'):
            df.to_csv(output_path, index=False)
        else:
            df.to_excel(output_path, index=False)
        
        return send_file(
            output_path,
            as_attachment=True,
            download_name=f'filtered_{filename}'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/update', methods=['POST'])
def update_file():
    data = request.json
    filename = data.get('filename')
    updates = data.get('updates', [])
    max_records = data.get('max_records')
    
    if not filename:
        return jsonify({'error': 'No filename provided'}), 400
    
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if not os.path.exists(file_path):
        return jsonify({'error': 'File not found'}), 404
    
    try:
        df = process_update(file_path, updates, max_records)
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], f'updated_{filename}')
        
        if filename.endswith('.csv'):
            df.to_csv(output_path, index=False)
        else:
            df.to_excel(output_path, index=False)
        
        return send_file(
            output_path,
            as_attachment=True,
            download_name=f'updated_{filename}'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True) 