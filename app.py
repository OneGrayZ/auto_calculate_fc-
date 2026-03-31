from flask import Flask, render_template, request, jsonify, send_file, make_response
import pandas as pd
import numpy as np
import io
import os
import re
from datetime import datetime
from urllib.parse import quote

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size


def extract_fc_codes_from_addresses(series):
    """Extract FC-like codes (e.g. IND9, LAX9) from address text."""
    pattern = re.compile(r'[A-Z]{2,5}\d{1,3}')
    found = set()

    for value in series.dropna().astype(str):
        for match in pattern.findall(value.upper()):
            found.add(match)

    return sorted(found)

def calculation(fc, data_xlsx):
    """Calculate metrics for a single FC
    
    Filter logic explanation:
    - filter_1: All rows where '收件地址' contains the FC code (e.g., IND9)
    - filter_2: From filter_1, EXCLUDE rows where '跟进记录' contains '已转仓'
              This means if 跟进记录 has '已转仓LAX9', it WILL be filtered OUT (excluded)
              Only rows WITHOUT '已转仓' in 跟进记录 are kept
    - filter_3: From filter_2, only rows where '产品渠道' contains specific keywords
    """
    fc = fc.upper()  # Ensure FC is uppercase for matching
    filter_1 = data_xlsx[data_xlsx['收件地址'].str.contains(fc, na=False)]
    
    if filter_1.empty:
        print(f"Warning: No data found for FC: {fc}")  # Debug
        return {
            'FC': fc,
            '未出单': 0,
            '未转仓': 0,
            '未转仓核爆品': 0,
            '已转到该仓': 0,
            '总货量': 0
        }
    
    #locate the column name contains '体积' (case-insensitive) and use it for calculations
    volume_col = None
    for col in filter_1.columns:
        if '体积' in col:
            volume_col = col
            break

    if volume_col is None:
        raise ValueError("Excel file must contain a column with '体积' in the name")

    # 未出单 = sum of 预计总体积 for all rows matching this FC
    未出单 = filter_1[volume_col].sum()
    
    # 未转仓 = sum where 跟进记录 does NOT contain '已转仓'
    # YES, entries like '已转仓LAX9' WILL be filtered out (excluded)
    filter_2 = filter_1[~filter_1['跟进记录'].astype(str).str.contains('转仓', na=False)]
    未转仓 = filter_2[volume_col].sum() if volume_col in filter_2.columns else 0
    
    # 未转仓核爆品 = sum where also 产品渠道 contains specific keywords
    keywords = ['ONLY23', 'ONLY18', 'Z快线', '至尊达', '王者鲲运']
    pattern = '|'.join(keywords)
    filter_3 = filter_2[filter_2['产品渠道'].astype(str).str.contains(pattern, na=False)]
    未转仓核爆品 = filter_3[volume_col].sum() if not filter_3.empty else 0
    
    # 已转到该仓 = sum where 跟进记录 contains '转仓' + fc (e.g., '转仓IND9')
    change = '转仓' + str(fc)
    filter_4 = data_xlsx[data_xlsx['跟进记录'].astype(str).str.contains(change, na=False)]
    已转到该仓 = filter_4[volume_col].sum() if not filter_4.empty else 0

    # 总货量 = 未转仓 + 已转到该仓
    总货量 = float(未转仓) + float(已转到该仓)
    
    result = {
        'FC': str(fc),  # Ensure FC is a string
        '未出单': float(未出单),
        '未转仓': float(未转仓),
        '未转仓核爆品': float(未转仓核爆品),
        '已转到该仓': float(已转到该仓),
        '总货量': 总货量
    }
    print(f"Calculated result for {fc}: {result}")  # Debug
    return result

def process_data(fc_list, xlsx_file):
    """Process all FCs and return combined results"""
    # Read the Excel file
    df = pd.read_excel(xlsx_file)
    
    # Check if FC column exists
    if '收件地址' not in df.columns:
        raise ValueError("Excel file must contain a '收件地址' column")
    
    # Keep user input order when FCs are provided manually.
    # Only auto-detected FCs will be sorted by 总货量 at the end.
    has_manual_input = bool(fc_list)

    # If no FC list provided, auto-detect FC codes from address text
    if not has_manual_input:
        fc_list = extract_fc_codes_from_addresses(df['收件地址'])
        if not fc_list:
            raise ValueError("未能从'收件地址'自动识别FC代码，请在文本框中手动输入FC（每行一个）")
    
    # Calculate for each FC
    results = []
    for fc in fc_list:
        fc = fc.strip()
        if fc:  # Skip empty lines
            result = calculation(fc, df)
            results.append(result)
    
    # Create output DataFrame
    output_df = pd.DataFrame(results)

    # Sort only when FC list is auto-detected (no manual input).
    if not has_manual_input:
        output_df = output_df.sort_values(by='总货量', ascending=False).reset_index(drop=True)
    
    return output_df

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    try:
        # Get FC codes from text input
        fc_text = request.form.get('fc_codes', '')
        fc_list = [line.strip() for line in fc_text.split('\n') if line.strip()]
        
        # Get uploaded file
        if 'file' not in request.files:
            return jsonify({'error': '请上传Excel文件'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': '未选择文件'}), 400
        
        if not file.filename.endswith(('.xlsx', '.xls')):
            return jsonify({'error': '文件必须是Excel格式 (.xlsx 或 .xls)'}), 400
        
        # Process the data
        result_df = process_data(fc_list, file)
        
        # Convert to list of dictionaries for JSON response
        results = result_df.to_dict('records')
        
        return jsonify({
            'success': True,
            'results': results
        })
        
    except Exception as e:
        return jsonify({'error': f'处理错误: {str(e)}'}), 500

@app.route('/export', methods=['POST'])
def export():
    try:
        # Get FC codes from text input
        fc_text = request.form.get('fc_codes', '')
        fc_list = [line.strip() for line in fc_text.split('\n') if line.strip()]
        
        # Get uploaded file
        if 'file' not in request.files:
            return jsonify({'error': '请上传Excel文件'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': '未选择文件'}), 400
        
        if not file.filename.endswith(('.xlsx', '.xls')):
            return jsonify({'error': '文件必须是Excel格式 (.xlsx 或 .xls)'}), 400
        
        # Process the data
        result_df = process_data(fc_list, file)
        
        # Create output Excel file in memory
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False, sheet_name='计算结果')
        
        # Generate filename with current date and time
        now = datetime.now()
        filename = f"FC货量情况_{now.strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
        
        output.seek(0)
        response = make_response(output.getvalue())
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        response.headers['Content-Disposition'] = f"attachment; filename*=UTF-8''{quote(filename)}"
        return response
        
    except Exception as e:
        return jsonify({'error': f'导出错误: {str(e)}'}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
