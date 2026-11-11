from copy import copy
import hashlib
from pathlib import Path
import tempfile
import unittest
import subprocess
import sys
import os
from unittest.mock import patch

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font
from filter_core import filter_workbook, req_years, row_pass, seg_for_degree

HEADERS = ['岗位', '学历', '专业', '招考对象', '基层工作最低年限', '政治面貌']
CFG = dict(degree='本科', fresh=False, years=0, face='群众', majors=['计算机'], allow_unlimited=True)


class RulesTest(unittest.TestCase):
    def test_degree_segment_never_falls_back_to_wrong_level(self):
        self.assertEqual(seg_for_degree('硕士：计算机', '本科'), '')
        self.assertEqual(seg_for_degree('本科：中文；硕士：计算机', '研究生'), '计算机')
        self.assertFalse(row_pass(['A', '本科', '本科：中文；硕士：计算机', '', '', ''], HEADERS, CFG))

    def test_dedicated_work_years_and_empty_years(self):
        self.assertFalse(row_pass(['A', '本科', '计算机', '不限', 2, '不限'], HEADERS, CFG))
        self.assertTrue(row_pass(['A', '本科', '计算机', '不限', '不限', '不限'], HEADERS, CFG))
        for text, result in [('需两年基层工作经历', 2), ('基层工作经历三年以上', 3), ('需十二年基层工作经历', 12)]:
            self.assertEqual(req_years(text), result)
        self.assertEqual(req_years('年龄35岁，2026年毕业'), 0)

    def test_fresh_graduate_and_political_requirements(self):
        self.assertFalse(row_pass(['A', '本科', '计算机', '仅限应届毕业生', 0, '不限'], HEADERS, CFG))
        self.assertTrue(row_pass(['A', '本科', '计算机', '应届毕业生和社会人员', 0, '不限'], HEADERS, CFG))
        self.assertFalse(row_pass(['A', '本科', '计算机', '', 0, '中共党员'], HEADERS, CFG))

    def test_work_years_header_variations_and_values(self):
        headers_with_suffix = ['岗位', '学历', '专业', '招考对象', '基层工作最低年限（年）', '政治面貌']
        cases = [
            # (name, raw_years, user_years, expected_pass, reason)
            ('整数0', 0, 0, True, '要求0年，用户具备0年经验，符合条件保留'),
            ('字符串"0"', '0', 0, True, '字符串0年，符合0年年限要求保留'),
            ('整数2', 2, 0, False, '要求2年，用户仅具备0年经验，应排除'),
            ('字符串"2"', '2', 0, False, '字符串2年，超过配置的0年年限应排除'),
            ('带“年”文本"2年"', '2年', 0, False, '带“年”文本2年，超过0年应排除'),
            ('带中文数字“两年”', '两年', 0, False, '中文数字两年，超过0年应排除'),
            ('文本"不限"', '不限', 0, True, '不限基层工作年限，应保留'),
            ('空值None', None, 0, True, '未填写年限默认不作限制，应保留'),
            ('空格文本', '   ', 0, True, '空白文本默认不作限制，应保留'),
            ('用户有2年经验要求2年', '2年', 2, True, '要求2年且用户具备2年经验，应保留'),
            ('用户有2年经验要求3年', '3年', 2, False, '要求3年超过用户具备的2年经验，应排除'),
        ]
        for name, raw_years, user_years, expected_pass, reason in cases:
            with self.subTest(case=name, reason=reason):
                cfg = {**CFG, 'years': user_years}
                row = ['岗位A', '本科', '计算机', '不限', raw_years, '不限']
                self.assertEqual(row_pass(row, headers_with_suffix, cfg), expected_pass, f"{name}: {reason}")

    def test_degree_segmented_majors_table_driven(self):
        headers = ['岗位', '学历', '专业', '招考对象', '基层工作最低年限', '政治面貌']
        cases = [
            # (name, major_text, target_degree, search_majors, expected_pass, reason)
            (
                '本科段命中计算机且配置本科',
                '本科：计算机科学与技术；硕士：工商管理',
                '本科',
                ['计算机'],
                True,
                '当前学历为本科，匹配本科段内的计算机专业，应保留',
            ),
            (
                '硕士段为其他专业且配置研究生',
                '本科：计算机科学与技术；硕士：工商管理',
                '研究生',
                ['计算机'],
                False,
                '当前学历为研究生，仅匹配硕士段（工商管理），不包含计算机，应排除',
            ),
            (
                '本科段为其他专业且配置本科',
                '本科：汉语言文学；硕士：软件工程',
                '本科',
                ['软件工程'],
                False,
                '当前学历为本科，本科段为汉语言文学，不包含软件工程，应排除',
            ),
            (
                '硕士段命中软件工程且配置研究生',
                '本科：汉语言文学；硕士：软件工程',
                '研究生',
                ['软件工程'],
                True,
                '当前学历为研究生，硕士段为软件工程，应保留',
            ),
            (
                '英文冒号和分号格式下本科段匹配',
                '本科:计算机类;硕士:电子信息',
                '本科',
                ['计算机'],
                True,
                '英文冒号和分号分隔下正确提取本科段计算机专业，应保留',
            ),
            (
                '多学历段落中博士段不干扰硕士匹配',
                '本科：数学；硕士：统计学；博士：计算机',
                '硕士',
                ['统计学'],
                True,
                '当前配置硕士，仅提取硕士段统计学，应保留',
            ),
        ]
        for name, major_text, target_degree, search_majors, expected_pass, reason in cases:
            with self.subTest(case=name, reason=reason):
                cfg = {**CFG, 'degree': target_degree, 'majors': search_majors, 'allow_unlimited': False}
                row = ['岗位A', target_degree, major_text, '不限', 0, '不限']
                self.assertEqual(row_pass(row, headers, cfg), expected_pass, f"{name}: {reason}")


class WorkbookTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.src, self.dst = self.root/'input.xlsx', self.root/'output.xlsx'
        wb = Workbook()
        ws = wb.active
        ws.title = '职位表'
        ws.append(['演示数据，不是真实招聘信息'])
        ws.merge_cells('A1:F1')
        ws.append(HEADERS)
        ws.append(['不匹配', '研究生', '计算机', '不限', 0, '不限'])
        ws.append(['保留', '本科', '计算机', '不限', 0, '不限'])
        ws.append([None]*6)
        ws['A4'].fill = PatternFill('solid', fgColor='34D399')
        ws['A4'].font = Font(bold=True)
        ws.row_dimensions[4].height = 39
        ws.column_dimensions['A'].width = 24
        ws.auto_filter.ref = 'A2:F5'
        note = wb.create_sheet('使用说明'); note.append(['原样保留的说明页'])
        wb.save(self.src); wb.close()
        self.original = self.src.read_bytes()

    def tearDown(self):
        self.temp.cleanup()

    def test_result_styles_dimensions_and_original_are_preserved(self):
        self.assertEqual(filter_workbook(self.src, self.dst, CFG, log=lambda _: None), 1)
        self.assertEqual(self.src.read_bytes(), self.original)
        wb = load_workbook(self.dst); ws = wb['职位表']
        self.assertEqual(ws['A3'].value, '保留')
        self.assertEqual(ws['A3'].fill.fgColor.rgb, '0034D399')
        self.assertTrue(ws['A3'].font.bold)
        self.assertEqual(ws.row_dimensions[3].height, 39)
        self.assertEqual(ws.column_dimensions['A'].width, 24)
        self.assertIn('A1:F1', ws.merged_cells)
        self.assertEqual(ws.auto_filter.ref, 'A2:F3')
        self.assertEqual(wb['使用说明']['A1'].value, '原样保留的说明页')
        wb.close()

    def test_never_overwrite_source_even_with_overwrite_enabled(self):
        with self.assertRaises(ValueError):
            filter_workbook(self.src, self.src, CFG, overwrite=True)
        self.assertEqual(self.src.read_bytes(), self.original)

    def test_existing_output_requires_explicit_overwrite(self):
        self.dst.write_bytes(b'existing result')
        with self.assertRaises(FileExistsError):
            filter_workbook(self.src, self.dst, CFG)
        self.assertEqual(self.dst.read_bytes(), b'existing result')

    def test_failed_save_leaves_existing_output_untouched_and_no_temp_file(self):
        self.dst.write_bytes(b'existing result')
        with patch('openpyxl.workbook.workbook.Workbook.save', side_effect=OSError('disk unavailable')):
            with self.assertRaises(OSError):
                filter_workbook(self.src, self.dst, CFG, overwrite=True, log=lambda _: None)
        self.assertEqual(self.dst.read_bytes(), b'existing result')
        self.assertEqual(list(self.root.glob('.job-filter-*')), [])

    def test_formula_and_data_merges_fail_before_saving(self):
        for mode in ['formula', 'merge']:
            with self.subTest(mode=mode):
                self.src.write_bytes(self.original)
                wb = load_workbook(self.src)
                if mode == 'formula': wb['使用说明']['B1'] = "='职位表'!A4"
                else: wb['职位表'].merge_cells('A3:A4')
                wb.save(self.src); wb.close()
                with self.assertRaises(ValueError): filter_workbook(self.src, self.dst, CFG)
                self.assertFalse(self.dst.exists())

    def test_no_header_is_not_reported_as_success(self):
        wb = Workbook(); wb.active.append(['not a job list']); wb.save(self.src); wb.close()
        with self.assertRaisesRegex(ValueError, '表头'): filter_workbook(self.src, self.dst, CFG)
        self.assertFalse(self.dst.exists())

    def test_no_matches_is_a_valid_empty_result(self):
        cfg = {**CFG, 'majors': ['不存在的专业'], 'allow_unlimited': False}
        self.assertEqual(filter_workbook(self.src, self.dst, cfg, log=lambda _: None), 0)
        wb = load_workbook(self.dst); self.assertEqual(wb['职位表'].max_row, 2); wb.close()

    def test_negative_years_and_changed_extension_are_rejected(self):
        with self.assertRaises(ValueError): filter_workbook(self.src, self.dst, {**CFG, 'years': -1})
        with self.assertRaises(ValueError): filter_workbook(self.src, self.root/'out.xlsm', CFG)

    def test_cli_works_with_non_chinese_output_encoding(self):
        cli = Path(__file__).resolve().parents[1] / 'filter_cli.py'
        result = subprocess.run([sys.executable, str(cli), str(self.src), str(self.dst), '--majors', '计算机'],
                                capture_output=True, env={**os.environ, 'PYTHONIOENCODING': 'ascii'})
        self.assertEqual(result.returncode, 0, result.stderr.decode('utf-8'))
        self.assertIn('完成', result.stdout.decode('utf-8'))
        self.assertTrue(self.dst.exists())

    def test_cli_supports_chinese_comma_separated_majors(self):
        cli = Path(__file__).resolve().parents[1] / 'filter_cli.py'
        result = subprocess.run(
            [sys.executable, str(cli), str(self.src), str(self.dst), '--majors', '软件工程，计算机类，大数据', '--overwrite'],
            capture_output=True,
            env={**os.environ, 'PYTHONIOENCODING': 'utf-8'}
        )
        self.assertEqual(result.returncode, 0, result.stderr.decode('utf-8'))
        self.assertIn('完成', result.stdout.decode('utf-8'))
        self.assertTrue(self.dst.exists())


if __name__ == '__main__':
    unittest.main()
