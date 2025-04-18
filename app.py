from flask import Flask, request, send_file, render_template
import os
from src.doc_parser import DocParser
from src.doc_rebuilder import DocRebuilder
import tempfile
import logging

app = Flask(__name__)

# 配置日誌
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/convert', methods=['POST'])
def convert():
    try:
        # 獲取上傳的文件
        file = request.files['file']
        from_key = request.form['fromKey']
        to_key = request.form['toKey']
        
        # 創建臨時文件
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as temp_input:
            file.save(temp_input.name)
            
        # 創建臨時 JSON 文件
        temp_json = tempfile.NamedTemporaryFile(suffix='.json', delete=False)
        temp_output = tempfile.NamedTemporaryFile(suffix='.docx', delete=False)
        
        try:
            # 解析文檔
            parser = DocParser()
            data = parser.parse_docx(temp_input.name)
            parser.save_to_json(data, temp_json.name)
            
            # 重建文檔
            rebuilder = DocRebuilder()
            rebuilder.rebuild_docx(data, temp_output.name)
            
            # 返回處理後的文件
            return send_file(
                temp_output.name,
                as_attachment=True,
                download_name='converted.docx',
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
            
        finally:
            # 清理臨時文件
            os.unlink(temp_input.name)
            os.unlink(temp_json.name)
            os.unlink(temp_output.name)
            
    except Exception as e:
        logger.error(f"轉換失敗: {str(e)}")
        return {'error': str(e)}, 500

if __name__ == '__main__':
    app.run(debug=True) 