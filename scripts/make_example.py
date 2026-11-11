"""Generate a fully synthetic, reproducible first-run example."""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from filter_core import filter_workbook

root = Path(__file__).resolve().parents[1] / 'examples'
root.mkdir(exist_ok=True)
wb = Workbook(); ws = wb.active; ws.title = '演示职位表'
ws.append(['完全虚构的演示数据 / Synthetic demonstration data'])
ws.merge_cells('A1:F1')
ws.append(['岗位', '学历', '专业', '招考对象', '基层工作最低年限', '政治面貌'])
for row in [
    ['A · 软件支持', '本科及以上', '计算机类、软件工程', '不限', 0, '不限'],
    ['B · 数据研究', '研究生', '计算机', '不限', 0, '不限'],
    ['C · 校园招聘', '本科', '计算机', '仅限应届毕业生', 0, '不限'],
    ['D · 综合服务', '本科', '专业不限', '不限', 0, '不限'],
    ['E · 技术管理', '本科', '计算机', '不限', 3, '不限'],
    ['F · 文字编辑', '本科', '汉语言文学', '不限', 0, '不限'],
]: ws.append(row)
for row in ws:
    for cell in row:
        cell.alignment = Alignment(vertical='center', wrap_text=True)
        cell.font = Font(name='Arial', size=11, color='172D3A')
        if cell.row <= 2:
            cell.fill = PatternFill('solid', fgColor='153E46'); cell.font = Font(color='FFFFFF', bold=True)
        elif cell.row % 2:
            cell.fill = PatternFill('solid', fgColor='EFF8F5')
    ws.row_dimensions[row[0].row].height = 32
for col in 'ABCDEF': ws.column_dimensions[col].width = 25
ws.freeze_panes = 'A3'; ws.auto_filter.ref = 'A2:F8'
source = root/'demo.xlsx'; output = root/'demo-filtered.xlsx'
wb.save(source); wb.close()
cfg = dict(degree='本科', fresh=False, years=0, face='群众', majors=['计算机'], allow_unlimited=True)
filter_workbook(source, output, cfg, overwrite=True)
