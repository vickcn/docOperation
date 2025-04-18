from flask import Flask, request, send_file
from flask_cors import CORS
import os
from src.doc_parser import DocParser
from src.doc_rebuilder import DocRebuilder
import tempfile

app = Flask(__name__)
CORS(app)  # 啟用CORS支持

@app.route('/api/convert', methods=['POST'])
def convert():
    try:
        # 獲取上傳的文件和參數
        file = request.files['file']
        from_key = request.form['fromKey']
        to_key = request.form['toKey']
        
        # 創建臨時文件來保存上傳的文件
        temp_input = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
        file.save(temp_input.name)
        
        # 創建臨時文件來保存輸出
        temp_output = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
        
        # 解析文檔
        parser = DocParser()
        doc_data = parser.parse_docx(temp_input.name)
        
        # 重建文檔
        rebuilder = DocRebuilder()
        rebuilder.rebuild_docx(doc_data, temp_output.name)
        
        # 返回處理後的文件
        return send_file(
            temp_output.name,
            as_attachment=True,
            download_name='converted.docx',
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        
    except Exception as e:
        return str(e), 500
        
    finally:
        # 清理臨時文件
        try:
            os.unlink(temp_input.name)
            os.unlink(temp_output.name)
        except:
            pass

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port) 