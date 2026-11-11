"""Offline, conservative filtering for plain Excel job lists."""
from copy import copy
import os
from pathlib import Path
import re
import tempfile

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

DEGREE_RANK = {"大专": 1, "专科": 1, "本科": 2, "研究生": 3, "硕士": 3, "博士": 4}
OBJ_KEYS = ["考录对象", "招录对象", "招考对象", "招聘对象", "考录范围", "招考范围", "报考范围", "是否限应届", "其他条件", "其他要求", "备注"]
CN_NUM = {"一": 1, "两": 2, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}


def find_header(ws):
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=min(15, ws.max_row), values_only=True), 1):
        text = " ".join(str(x) for x in row if x is not None)
        if "专业" in text and "学历" in text:
            return i, list(row)
    return None, None


def locate(headers, keys):
    return [i for i, h in enumerate(headers) if h and any(k in str(h) for k in keys)]


def seg_for_degree(major, degree):
    labels = r"本科|研究生|硕士|博士|专科|大专"
    sections = list(re.finditer(rf"({labels})[：:]", major))
    if not sections:
        return major
    accepted = {"大专": {"专科", "大专"}, "专科": {"专科", "大专"}, "硕士": {"研究生", "硕士"}, "研究生": {"研究生", "硕士"}}.get(degree, {degree})
    return " ".join(major[m.end():sections[i + 1].start() if i + 1 < len(sections) else len(major)] for i, m in enumerate(sections) if m.group(1) in accepted)


def req_degree_rank(text):
    ranks = [rank for label, rank in DEGREE_RANK.items() if label in str(text)]
    return min(ranks) if ranks else 0


def req_years(text, dedicated=False):
    text = str(text).strip()
    if dedicated and text.isdigit():
        return int(text)
    if not dedicated:
        # A general age or date in the remarks is not a work-experience requirement.
        match = re.search(r"(?:基层工作|基层经历).*?([0-9一两二三四五六七八九十]+)\s*年", text)
        if not match:
            match = re.search(r"([0-9一两二三四五六七八九十]+)\s*年[^，。；;]*基层工作", text)
    else:
        match = re.search(r"([0-9一两二三四五六七八九十]+)\s*年", text)
    if not match:
        return 0
    value = match.group(1)
    if value.isdigit():
        return int(value)
    if "十" in value:
        left, right = value.split("十", 1)
        return (CN_NUM.get(left, 1) * 10) + CN_NUM.get(right, 0)
    return CN_NUM.get(value, 0)


def validate_config(cfg):
    if cfg.get('degree') not in DEGREE_RANK:
        raise ValueError('请选择有效学历。')
    if type(cfg.get('years')) is not int or cfg['years'] < 0:
        raise ValueError('基层工作年限必须是非负整数。')
    if cfg.get('face') not in {'群众', '共青团员', '中共党员'}:
        raise ValueError('请选择有效政治面貌。')
    if not isinstance(cfg.get('majors'), list) or any(not isinstance(k, str) or not k.strip() for k in cfg['majors']):
        raise ValueError('专业关键词必须是非空文本。')
    if not cfg['majors'] and not cfg.get('allow_unlimited'):
        raise ValueError('请填写专业关键词，或勾选专业不限。')


def row_pass(cells, headers, cfg):
    def col_text(keys):
        return " ".join(str(cells[i]) for i in locate(headers, keys) if i < len(cells) and cells[i] is not None)
    major = col_text(["专业"])
    if not major.strip():
        return False
    if req_degree_rank(col_text(["学历"])) > DEGREE_RANK[cfg['degree']]:
        return False
    segment = seg_for_degree(major, cfg['degree'])
    words = cfg['majors'] + (["不限"] if cfg['allow_unlimited'] else [])
    if not any(word in segment for word in words):
        return False
    obj = col_text(OBJ_KEYS)
    if not cfg['fresh'] and '应届' in obj:
        if not re.search(r"应届[^，。；;]*(?:社会|在职)|(?:社会|在职)[^，。；;]*应届|不限应届|应届[^，。；;]*不限", obj):
            return False
    years = max(req_years(obj), req_years(col_text(['基层工作', '基层经历']), dedicated=True))
    if years > cfg['years']:
        return False
    face = col_text(['政治面貌'])
    if cfg['face'] != '中共党员' and ('党员' in face and '不限' not in face or re.search(r'(?:须|限|要求|仅限)[^，。；;]*党员', obj)):
        return False
    return True


def _check_structure(wb, eligible):
    for ws in wb.worksheets:
        if any(c.data_type == 'f' for row in ws for c in row):
            raise ValueError(f'[{ws.title}] 含有公式。请在副本中粘贴为值后重试，以免删除行后引用失效。')
        if ws._charts or ws._images or ws.tables or len(ws.conditional_formatting) or ws.data_validations.count:
            raise ValueError(f'[{ws.title}] 含图表、图片、表格对象、条件格式或数据验证；请先另存为普通职位表。')
        if ws.title in eligible and any(r.max_row > eligible[ws.title][0] for r in ws.merged_cells.ranges):
            raise ValueError(f'[{ws.title}] 数据区域含合并单元格，请在副本中取消合并并补齐内容后重试。')
    if wb.defined_names:
        raise ValueError('工作簿含命名范围，请先另存为不含命名范围的普通职位表。')


def filter_workbook(src, dst, cfg, log=print, overwrite=False):
    validate_config(cfg)
    src, dst = Path(src).expanduser().resolve(), Path(dst).expanduser().resolve()
    if not src.is_file() or src.suffix.lower() not in {'.xlsx', '.xlsm'}:
        raise ValueError('请选择存在的 .xlsx 或 .xlsm 文件；不支持旧版 .xls。')
    if src == dst or dst.exists() and os.path.samefile(src, dst):
        raise ValueError('输出文件不能与原文件相同，请选择新的文件名。')
    if dst.suffix.lower() != src.suffix.lower():
        raise ValueError('输出文件扩展名必须与输入相同；宏工作簿请保留 .xlsm。')
    if dst.exists() and not overwrite:
        raise FileExistsError('输出文件已存在，请选择新文件名或明确允许覆盖。')
    if not dst.parent.is_dir():
        raise ValueError('输出文件夹不存在，请先创建。')
    wb = load_workbook(src, keep_vba=src.suffix.lower() == '.xlsm', keep_links=True)
    temp_path = None
    try:
        eligible = {}
        for ws in wb.worksheets:
            header, headers = find_header(ws)
            if header is not None:
                eligible[ws.title] = (header, headers)
        if not eligible:
            raise ValueError('前 15 行内未找到同时含“专业”和“学历”的表头，未生成输出。')
        _check_structure(wb, eligible)
        total = 0
        for ws in wb.worksheets:
            if ws.title not in eligible:
                log(f'[{ws.title}] 未识别为职位表，保留原内容。')
                continue
            header, headers = eligible[ws.title]
            remove, kept, blank = [], 0, 0
            old_max = ws.max_row
            dimensions = {r: copy(d) for r, d in ws.row_dimensions.items()}
            for r in range(header + 1, old_max + 1):
                cells = [c.value for c in ws[r]]
                if not any(v is not None and str(v).strip() for v in cells):
                    remove.append(r)
                    blank += 1
                elif row_pass(cells, headers, cfg):
                    kept += 1
                else:
                    remove.append(r)
            # Delete contiguous runs from bottom to top; cell styles move with cells.
            groups = []
            for r in remove:
                if groups and groups[-1][0] + groups[-1][1] == r:
                    groups[-1][1] += 1
                else:
                    groups.append([r, 1])
            for start, count in reversed(groups):
                ws.delete_rows(start, count)
            deleted = set(remove)
            ws.row_dimensions.clear()
            new_r = 0
            for old_r in range(1, old_max + 1):
                if old_r not in deleted:
                    new_r += 1
                    if old_r in dimensions:
                        dimensions[old_r].index = new_r
                        ws.row_dimensions[new_r] = dimensions[old_r]
            # A plain filter range can be rebuilt; row-specific sort state cannot.
            if ws.auto_filter.ref:
                ws.auto_filter.ref = f'A{header}:{get_column_letter(ws.max_column)}{max(header, ws.max_row)}'
                ws.auto_filter.sortState = None
                ws.auto_filter.filterColumn = []
            total += kept
            log(f'[{ws.title}] 保留 {kept} 个职位，排除 {len(remove)-blank} 行，清理 {blank} 个空行。')
        fd, temp = tempfile.mkstemp(prefix='.job-filter-', suffix=dst.suffix, dir=dst.parent)
        os.close(fd)
        temp_path = Path(temp)
        wb.save(temp_path)
        # Recheck after processing; GUI may have taken time to filter a large workbook.
        if dst.exists() and not overwrite:
            raise FileExistsError('处理期间输出文件已出现，请换一个文件名。')
        os.replace(temp_path, dst)
        temp_path = None
        return total
    finally:
        wb.close()
        if temp_path is not None:
            temp_path.unlink(missing_ok=True)
