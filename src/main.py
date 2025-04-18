import argparse
import logging
import os
from pathlib import Path
from doc_parser import DocParser
from doc_rebuilder import DocRebuilder
from chord_transposer import ChordTransposer
import json
import sys

class LOGger:
    @staticmethod
    def myparser():
        parser = argparse.ArgumentParser(description='文档处理工具')
        parser.add_argument('--input', type=str, required=True, help='输入文件路径')
        parser.add_argument('--output', type=str, required=True, help='输出文件路径')
        parser.add_argument('--mode', type=str, choices=['parse', 'rebuild'], required=True,
                          help='操作模式：parse-解析文档，rebuild-重建文档')
        parser.add_argument('--from-key', type=str, help='原始调号（例如：C, D, E#, Bb等）')
        parser.add_argument('--to-key', type=str, help='目标调号（例如：C, D, E#, Bb等）')
        parser.add_argument('--preserve-spaces', type=bool, default=True,
                          help='是否保留原始空格（默认：True）')
        return parser

def setup_logging():
    """设置日志配置"""
    # 确保日志目录存在
    log_dir = Path('logs')
    log_dir.mkdir(exist_ok=True)
    
    # 设置日志文件路径
    log_file = log_dir / 'document_processor.log'
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger(__name__)

def transpose_data(data: dict, from_key: str, to_key: str, preserve_spaces: bool = True) -> dict:
    """转调文档数据"""
    if not from_key or not to_key or from_key == to_key:
        return data
        
    # 转调段落中的和弦
    for para in data.get("paragraphs", []):
        for run in para.get("runs", []):
            run["text"] = ChordTransposer.transpose_text(
                run["text"], 
                from_key, 
                to_key, 
                preserve_spaces=preserve_spaces
            )
            
    return data

def main():
    """主函數"""
    parser = LOGger.myparser()
    args = parser.parse_args()
    
    try:
        if args.mode == 'parse':
            # 解析文檔
            doc_parser = DocParser()
            data = doc_parser.parse_docx(args.input)
            doc_parser.save_to_json(data, args.output)
            print(f"文檔解析完成，結果保存到: {args.output}")
            
        elif args.mode == 'rebuild':
            # 讀取 JSON 數據
            with open(args.input, 'r', encoding='utf-8') as f:
                data = json.load(f)
                
            # 如果指定了調號，進行轉調
            if args.from_key and args.to_key:
                data = transpose_data(data, args.from_key, args.to_key, args.preserve_spaces)
                
            # 重建文檔
            rebuilder = DocRebuilder()
            rebuilder.rebuild_docx(data, args.output)
            print(f"文檔重建完成，保存到: {args.output}")
            
    except Exception as e:
        print(f"錯誤: {str(e)}")
        sys.exit(1)

if __name__ == '__main__':
    main() 