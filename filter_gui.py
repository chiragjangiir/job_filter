#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
公考职位表筛选器 (GUI版)
- 双击运行 -> 选择职位表Excel -> 勾选个人条件 -> 选择保存位置 -> 输出筛选结果
- 输出方式: 在原Excel副本上删除不符合条件的行(保留原表所有格式/列宽/表头), 等效于Excel筛选后另存
打包命令(在对应系统上执行):
    pip install openpyxl pyinstaller
    Windows: pyinstaller -F -w -n 职位筛选器 filter_gui.py
    macOS:   pyinstaller -F -w -n 职位筛选器 filter_gui.py
"""
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from openpyxl import load_workbook

DEGREE_RANK = {"大专": 1, "专科": 1, "本科": 2, "研究生": 3, "硕士": 3, "博士": 4}
OBJ_KEYS = ["考录对象", "招录对象", "招考对象", "招聘对象", "考录范围", "招考范围", "报考范围",
            "是否限应届", "基层工作最低年限", "服务基层项目", "其他条件", "其他要求", "备注"]
CN_NUM = {"一": 1, "两": 2, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8}


# ---------------- 核心筛选逻辑 ----------------
def find_header(ws):
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=15, values_only=True), 1):
        s = " ".join(str(x) for x in row if x)
        if "专业" in s and "学历" in s:
            return i, list(row)
    return None, None


def locate(headers, keys):
    return [i for i, h in enumerate(headers) for k in keys if h and k in str(h)]


def seg_for_degree(major, degree):
    """按本人学历取专业栏对应段落, 无分段则取全文"""
    key = "本科" if degree == "本科" else ("研究生" if degree in ("研究生", "硕士") else "专科|大专")
    m = re.search(rf"(?:{key})[：:]\s*(.*?)(?=(本科|研究生|硕士|博士|专科|大专)[：:]|$)", major, re.S)
    return m.group(1) if m else major


def req_degree_rank(text):
    t = str(text)
    if re.search(r"(专科|大专)", t):
        return 1
    if "本科" in t:
        return 2
    if "研究生" in t or "硕士" in t:
        return 3
    if "博士" in t:
        return 4
    return 0  # 未写学历, 不据此剔除


def req_years(text):
    """提取基层工作经历年限要求, 无要求返回0"""
    t = str(text)
    if "基层工作" not in t:
        return 0
    m = re.search(r"([0-9一两二三四五六七八]+)\s*年", t)
    if not m:
        return 0
    v = m.group(1)
    return int(v) if v.isdigit() else CN_NUM.get(v, 0)


def row_pass(cells, headers, cfg):
    def col_text(keys):
        return " ".join(str(cells[i]) for i in locate(headers, keys) if i < len(cells) and cells[i] is not None)

    major = col_text(["专业"])
    if not major.strip():
        return False  # 非数据行
    degree_req = col_text(["学历"])
    obj = col_text(OBJ_KEYS)
    face = col_text(["政治面貌"])

    # 1 学历: 岗位要求学历 > 本人学历 -> 剔除
    if req_degree_rank(degree_req) > DEGREE_RANK.get(cfg["degree"], 2):
        return False
    # 2 专业: 本人学历对应段落须命中关键词
    seg = seg_for_degree(major, cfg["degree"])
    kws = cfg["majors"] + (["不限"] if cfg["allow_unlimited"] else [])
    if not any(k in seg for k in kws):
        return False
    # 3 应届: 本人非应届 且 岗位仅限应届 -> 剔除
    if not cfg["fresh"] and re.search(r"应届", obj) and not re.search(r"(社会|在职|不限)", obj):
        return False
    # 4 基层工作经历: 要求年限 > 本人年限 -> 剔除
    if req_years(obj) > cfg["years"]:
        return False
    # 5 政治面貌: 岗位要求党员而本人非党员 -> 剔除
    if cfg["face"] != "中共党员" and re.search(r"党员", face + obj) and "不限" not in face:
        return False
    return True


def filter_workbook(src, dst, cfg, log=print):
    wb = load_workbook(src)  # 保留全部格式
    total_kept = 0
    for ws in wb.worksheets:
        h, headers = find_header(ws)
        if h is None:
            log(f"[{ws.title}] 未找到表头, 跳过")
            continue
        del_rows = []
        for r in range(h + 1, ws.max_row + 1):
            cells = [c.value for c in ws[r]]
            if any(v is not None and str(v).strip() for v in cells):
                if not row_pass(cells, headers, cfg):
                    del_rows.append(r)
        # 自底向上按连续区段删除(保留原格式, 速度快)
        del_rows.sort(reverse=True)
        i = 0
        while i < len(del_rows):
            j = i
            while j + 1 < len(del_rows) and del_rows[j + 1] == del_rows[j] - 1:
                j += 1
            ws.delete_rows(del_rows[j], del_rows[i] - del_rows[j] + 1)
            i = j + 1
        kept = ws.max_row - h
        total_kept += max(kept, 0)
        log(f"[{ws.title}] 剔除 {len(del_rows)} 行, 保留 {max(kept,0)} 个职位")
    wb.save(dst)
    return total_kept


# ---------------- GUI ----------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("公考职位表筛选器")
        self.geometry("560x430")
        self.resizable(False, False)
        self.src = tk.StringVar()
        pad = dict(padx=10, pady=6, sticky="w")

        f = ttk.Frame(self); f.pack(fill="both", expand=True, padx=10, pady=8)
        ttk.Label(f, text="① 职位表Excel:").grid(row=0, column=0, **pad)
        ttk.Entry(f, textvariable=self.src, width=44).grid(row=0, column=1, **pad)
        ttk.Button(f, text="选择文件…", command=self.pick).grid(row=0, column=2, **pad)

        ttk.Label(f, text="② 个人条件").grid(row=1, column=0, **pad)
        self.degree = tk.StringVar(value="本科")
        ttk.Label(f, text="学历:").grid(row=2, column=0, **pad)
        ttk.Combobox(f, textvariable=self.degree, values=["大专", "本科", "研究生"],
                     state="readonly", width=10).grid(row=2, column=1, **pad)
        self.fresh = tk.BooleanVar(value=False)
        ttk.Checkbutton(f, text="我是应届毕业生", variable=self.fresh).grid(row=2, column=2, **pad)

        ttk.Label(f, text="基层工作年限:").grid(row=3, column=0, **pad)
        self.years = tk.Spinbox(f, from_=0, to=20, width=8); self.years.delete(0, "end"); self.years.insert(0, "2")
        self.years.grid(row=3, column=1, **pad)
        self.face = tk.StringVar(value="群众")
        ttk.Label(f, text="政治面貌:").grid(row=3, column=2, **pad)
        ttk.Combobox(f, textvariable=self.face, values=["群众", "共青团员", "中共党员"],
                     state="readonly", width=10).grid(row=3, column=3, **pad)

        ttk.Label(f, text="专业关键词(逗号分隔):").grid(row=4, column=0, **pad)
        self.majors = ttk.Entry(f, width=44)
        self.majors.insert(0, "软件工程,计算机类,计算机科学与技术")
        self.majors.grid(row=4, column=1, columnspan=2, **pad)
        self.unlimited = tk.BooleanVar(value=True)
        ttk.Checkbutton(f, text='包含"专业不限"岗位', variable=self.unlimited).grid(row=5, column=1, **pad)

        ttk.Button(f, text="③ 开始筛选并另存为…", command=self.run).grid(row=6, column=1, pady=12)
        self.logbox = tk.Text(f, height=8, width=70, state="disabled")
        self.logbox.grid(row=7, column=0, columnspan=4, padx=10)

    def pick(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm")])
        if p:
            self.src.set(p)

    def log(self, msg):
        self.logbox.configure(state="normal")
        self.logbox.insert("end", msg + "\n")
        self.logbox.see("end")
        self.logbox.configure(state="disabled")
        self.update()

    def run(self):
        src = Path(self.src.get())
        if not src.exists():
            messagebox.showwarning("提示", "请先选择职位表Excel文件"); return
        dst = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                           initialfile=src.stem + "_筛选结果.xlsx",
                                           filetypes=[("Excel", "*.xlsx")])
        if not dst:
            return
        cfg = dict(degree=self.degree.get(), fresh=self.fresh.get(),
                   years=int(self.years.get() or 0), face=self.face.get(),
                   majors=[k.strip() for k in self.majors.get().split(",") if k.strip()],
                   allow_unlimited=self.unlimited.get())
        try:
            n = filter_workbook(src, dst, cfg, log=self.log)
            self.log(f"完成! 共保留 {n} 个职位 -> {dst}")
            messagebox.showinfo("完成", f"筛选完成, 共保留 {n} 个职位。\n已保存到:\n{dst}")
        except Exception as e:
            messagebox.showerror("出错", str(e))


if __name__ == "__main__":
    App().mainloop()
