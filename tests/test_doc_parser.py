import unittest
import sys
import os

# 添加项目根目录到 Python 路径
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from src.doc_parser import DocParser

class TestDocParser(unittest.TestCase):
    def setUp(self):
        self.parser = DocParser()

    def test_merge_chord_runs(self):
        # 测试用例1：基本和弦合并
        test_case1 = [
            {'text': 'G', 'font_size': 12},
            {'text': '#', 'font_size': 12, 'superscript': True},
            {'text': '7', 'font_size': 12}
        ]
        expected1 = [
            {'text': 'G#', 'font_size': 12, 'accidentals': [{'text': '#', 'font_size': 12, 'superscript': True}]},
            {'text': '7', 'font_size': 12}
        ]
        result1 = self.parser.merge_chord_runs(test_case1)
        self.assertEqual(result1, expected1, "基本和弦合并测试失败")

        # 测试用例2：多个升降号
        test_case2 = [
            {'text': 'C', 'font_size': 12},
            {'text': '#', 'font_size': 12, 'superscript': True},
            {'text': '#', 'font_size': 12, 'superscript': True},
            {'text': 'm', 'font_size': 12}
        ]
        expected2 = [
            {'text': 'C##', 'font_size': 12, 'accidentals': [
                {'text': '#', 'font_size': 12, 'superscript': True},
                {'text': '#', 'font_size': 12, 'superscript': True}
            ]},
            {'text': 'm', 'font_size': 12}
        ]
        result2 = self.parser.merge_chord_runs(test_case2)
        self.assertEqual(result2, expected2, "多个升降号合并测试失败")

        # 测试用例3：不合并的情况
        test_case3 = [
            {'text': 'G', 'font_size': 12},
            {'text': 'm', 'font_size': 12},
            {'text': '7', 'font_size': 12}
        ]
        expected3 = [
            {'text': 'G', 'font_size': 12},
            {'text': 'm', 'font_size': 12},
            {'text': '7', 'font_size': 12}
        ]
        result3 = self.parser.merge_chord_runs(test_case3)
        self.assertEqual(result3, expected3, "不合并情况测试失败")

        # 测试用例4：空列表
        test_case4 = []
        expected4 = []
        result4 = self.parser.merge_chord_runs(test_case4)
        self.assertEqual(result4, expected4, "空列表测试失败")

        # 测试用例5：复杂和弦
        test_case5 = [
            {'text': 'A', 'font_size': 12},
            {'text': 'b', 'font_size': 12, 'superscript': True},
            {'text': 'm', 'font_size': 12},
            {'text': '7', 'font_size': 12},
            {'text': '/', 'font_size': 12},
            {'text': 'C', 'font_size': 12},
            {'text': '#', 'font_size': 12, 'superscript': True}
        ]
        expected5 = [
            {'text': 'Ab', 'font_size': 12, 'accidentals': [{'text': 'b', 'font_size': 12, 'superscript': True}]},
            {'text': 'm', 'font_size': 12},
            {'text': '7', 'font_size': 12},
            {'text': '/', 'font_size': 12},
            {'text': 'C#', 'font_size': 12, 'accidentals': [{'text': '#', 'font_size': 12, 'superscript': True}]}
        ]
        result5 = self.parser.merge_chord_runs(test_case5)
        self.assertEqual(result5, expected5, "复杂和弦测试失败")

if __name__ == '__main__':
    unittest.main() 