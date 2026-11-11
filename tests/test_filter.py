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


if __name__ == '__main__':
    unittest.main()
