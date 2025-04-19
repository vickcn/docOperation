from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
from src.doc_parser import DocParser
from src.doc_rebuilder import DocRebuilder
import tempfile
import traceback

app = Flask(__name__)
CORS(app)  # 啟用CORS支持

@app.route('/')
def home():
    return jsonify({
        "status": "running",
        "message": "文檔轉換服務正在運行",
        "version": "vercel",
        "endpoints": {
            "/api/convert": "POST - 文檔轉換端點"
        }
    })

@app.route('/api/convert', methods=['POST'])
def convert():
    try:
        # 檢查請求中是否包含必要的文件和參數
        if 'file' not in request.files:
            return jsonify({
                "error": "未找到文件",
                "message": "請選擇一個.docx文件上傳"
            }), 400
            
        file = request.files['file']
        if not file.filename.endswith('.docx'):
            return jsonify({
                "error": "文件格式錯誤",
                "message": "只支持.docx格式的文件"
            }), 400
            
        if 'fromKey' not in request.form or 'toKey' not in request.form:
            return jsonify({
                "error": "參數缺失",
                "message": "請提供原調(fromKey)和目標調(toKey)"
            }), 400
            
        from_key = request.form['fromKey']
        to_key = request.form['toKey']
        
        print(f"開始處理文件: {file.filename}, 從 {from_key} 調到 {to_key}")
        
        # 創建臨時文件來保存上傳的文件
        temp_input = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
        file.save(temp_input.name)
        print(f"臨時輸入文件已創建: {temp_input.name}")
        
        # 創建臨時文件來保存輸出
        temp_output = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
        print(f"臨時輸出文件已創建: {temp_output.name}")
        
        try:
            # 解析文檔
            parser = DocParser()
            doc_data = parser.parse_docx(temp_input.name)
            print("文檔解析完成")
            
            # 重建文檔
            rebuilder = DocRebuilder()
            rebuilder.rebuild_docx(doc_data, temp_output.name)
            print("文檔重建完成")
            
            # 返回處理後的文件
            return send_file(
                temp_output.name,
                as_attachment=True,
                download_name='converted.docx',
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
            
        except Exception as e:
            print(f"文檔處理錯誤: {str(e)}")
            print(traceback.format_exc())
            return jsonify({
                "error": str(e),
                "message": "文檔處理過程中出錯",
                "details": traceback.format_exc()
            }), 500
            
        finally:
            # 清理臨時文件
            try:
                os.unlink(temp_input.name)
                os.unlink(temp_output.name)
                print("臨時文件已清理")
            except Exception as e:
                print(f"清理臨時文件時出錯: {str(e)}")
                
    except Exception as e:
        print(f"請求處理錯誤: {str(e)}")
        print(traceback.format_exc())
        return jsonify({
            "error": str(e),
            "message": "請求處理失敗",
            "details": traceback.format_exc()
        }), 400

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port) 