from flask import Flask, render_template, request, jsonify
import pandas as pd
import numpy as np
import io
import os

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

def calculation(fc, data_xlsx):
    """Calculate metrics for a single FC
    
    Filter logic explanation:
    - filter_1: All rows where '收件地址' contains the FC code (e.g., IND9)
    - filter_2: From filter_1, EXCLUDE rows where '跟进记录' contains '已转仓'
              This means if 跟进记录 has '已转仓LAX9', it WILL be filtered OUT (excluded)
              Only rows WITHOUT '已转仓' in 跟进记录 are kept
    - filter_3: From filter_2, only rows where '产品渠道' contains specific keywords
    """
    filter_1 = data_xlsx[data_xlsx['收件地址'].str.contains(fc, na=False)]
    
    if filter_1.empty:
        print(f"Warning: No data found for FC: {fc}")  # Debug
        return {
            'FC': fc,
            '未出单': 0,
            '未转仓': 0,
            '未转仓核爆品': 0,
            '已转到该仓': 0
        }
    
    # 未出单 = sum of 预计总体积 for all rows matching this FC
    未出单 = filter_1['预计总体积'].sum()
    
    # 未转仓 = sum where 跟进记录 does NOT contain '已转仓'
    # YES, entries like '已转仓LAX9' WILL be filtered out (excluded)
    filter_2 = filter_1[~filter_1['跟进记录'].astype(str).str.contains('转仓', na=False)]
    未转仓 = filter_2['预计总体积'].sum() if '预计总体积' in filter_2.columns else 0
    
    # 未转仓核爆品 = sum where also 产品渠道 contains specific keywords
    keywords = ['ONLY23', 'ONLY18', 'Z快线', '至尊达', '王者鲲运']
    pattern = '|'.join(keywords)
    filter_3 = filter_2[filter_2['产品渠道'].astype(str).str.contains(pattern, na=False)]
    print(filter_3)  
    未转仓核爆品 = filter_3['预计总体积'].sum() if not filter_3.empty else 0
    
    # 已转到该仓 = sum where 跟进记录 contains '转仓' + fc (e.g., '转仓IND9')
    change = '转仓' + str(fc)
    filter_4 = data_xlsx[data_xlsx['跟进记录'].astype(str).str.contains(change, na=False)]
    已转到该仓 = filter_4['预计总体积'].sum() if not filter_4.empty else 0

    result = {
        'FC': str(fc),  # Ensure FC is a string
        '未出单': float(未出单),
        '未转仓': float(未转仓),
        '未转仓核爆品': float(未转仓核爆品),
        '已转到该仓': float(已转到该仓)
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
    
    # Calculate for each FC
    results = []
    for fc in fc_list:
        fc = fc.strip()
        if fc:  # Skip empty lines
            result = calculation(fc, df)
            results.append(result)
    
    # Create output DataFrame
    output_df = pd.DataFrame(results)
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
        
        if not fc_list:
            return jsonify({'error': '请输入FC代码'}), 400
        
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
        
        # Also prepare copy-friendly format (tab-separated for Excel paste)
        copy_data = []
        for row in results:
            copy_data.append({
                'fc': row['FC'],
                '未出单': row['未出单'],
                '未转仓': row['未转仓'],
                '未转仓核爆品': row['未转仓核爆品'],
                '已转到该仓': row['已转到该仓']
            })
        
        return jsonify({
            'success': True,
            'results': results,
            'copy_data': copy_data
        })
        
    except Exception as e:
        return jsonify({'error': f'处理错误: {str(e)}'}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
