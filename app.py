from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import os
from werkzeug.utils import secure_filename
import tempfile
from comparison_logic import compare_excel_files, get_column_suggestions
import shutil

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = tempfile.mkdtemp()
# Initialize file paths as None
app.config['PRELOAD_FILE'] = None
app.config['POSTLOAD_FILE'] = None

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/get_sheets/<file_type>', methods=['POST'])
def get_sheets(file_type):
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    try:
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], f"{file_type}_{filename}")
        
        # If file exists, remove it first
        if os.path.exists(filepath):
            os.remove(filepath)
            
        file.save(filepath)
        
        # Store the filepath in app config
        if file_type == 'preload':
            app.config['PRELOAD_FILE'] = filepath
        else:
            app.config['POSTLOAD_FILE'] = filepath
            
        # Read sheets using pandas
        sheets = pd.ExcelFile(filepath).sheet_names
        
        return jsonify(sheets)
    except Exception as e:
        print(f"Error in get_sheets: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/get_columns/<file_type>/<sheet_name>')
def get_columns(file_type, sheet_name):
    try:
        filepath = app.config.get(f'{file_type.upper()}_FILE')
        if not filepath:
            return jsonify({'error': f'No {file_type} file uploaded'}), 404
            
        df = pd.read_excel(filepath, sheet_name=sheet_name)
        columns = df.columns.tolist()
        
        return jsonify(columns)
    except Exception as e:
        print(f"Error in get_columns: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/compare', methods=['POST'])
def compare():
    try:
        data = request.get_json()
        if not app.config.get('PRELOAD_FILE') or not app.config.get('POSTLOAD_FILE'):
            return jsonify({'error': 'Files not uploaded'}), 400

        output_dir = tempfile.mkdtemp()
        sheet_mappings = data.get('sheetMappings', [])
        
        for mapping in sheet_mappings:
            if mapping.get('postloadSheet') != 'none':
                output_file = compare_excel_files(
                    app.config['PRELOAD_FILE'],
                    app.config['POSTLOAD_FILE'],
                    mapping['preloadSheet'],
                    mapping['postloadSheet'],
                    mapping['keyColumn'],
                    output_dir
                )
                app.config['COMPARISON_RESULT'] = output_file

        return jsonify({
            'message': 'Comparison completed successfully',
            'downloadReady': True
        })
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download_result')
def download_result():
    try:
        output_file = app.config.get('COMPARISON_RESULT')
        if not output_file or not os.path.exists(output_file):
            return jsonify({'error': 'No comparison result available'}), 404
        
        return send_file(
            output_file,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='comparison_result.xlsx'
        )
    except Exception as e:
        print(f"Error in download_result: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/clear', methods=['POST'])
def clear():
    try:
        # Clear stored files
        for key in ['PRELOAD_FILE', 'POSTLOAD_FILE', 'COMPARISON_RESULT']:
            filepath = app.config.get(key)
            if filepath and os.path.exists(filepath):
                os.remove(filepath)
            app.config[key] = None
            
        return jsonify({'message': 'Cleared successfully'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/get_column_suggestions', methods=['POST'])
def get_column_suggestions_route():
    try:
        data = request.get_json()
        preload_file = app.config['PRELOAD_FILE']
        postload_file = app.config['POSTLOAD_FILE']
        pre_sheet = data.get('preSheet')
        post_sheet = data.get('postSheet')
        
        if not all([preload_file, postload_file, pre_sheet, post_sheet]):
            return jsonify({'error': 'Missing required files or sheet names'}), 400
        
        pre_data = pd.read_excel(preload_file, sheet_name=pre_sheet)
        post_data = pd.read_excel(postload_file, sheet_name=post_sheet)
        
        suggestions = get_column_suggestions(pre_data.columns, post_data.columns)
        
        return jsonify({
            'success': True,
            'suggestions': suggestions
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        })

if __name__ == '__main__':
    # Create upload folder if it doesn't exist
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    app.run(debug=True) 