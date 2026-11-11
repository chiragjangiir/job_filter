#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Offline desktop interface for plain Excel job lists.

Filtering rules and file protection live in filter_core.py.
See docs/USAGE.md for supported workbook structures.
"""
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

from filter_core import filter_workbook, validate_config
import queue
import threading


# ---------------- GUI ----------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("公考职位表筛选器")
        self.geometry("820x620")
        self.minsize(760, 580)
        self.events = queue.Queue()
        self.running = False
        self.after(100, self.poll_events)
        self.src = tk.StringVar()
        pad = dict(padx=10, pady=6, sticky="w")

        f = ttk.Frame(self, padding=20); f.pack(fill="both", expand=True)
        f.columnconfigure(1, weight=1)
        f.rowconfigure(7, weight=1)
        ttk.Label(f, text="① 职位表Excel:").grid(row=0, column=0, **pad)
        ttk.Entry(f, textvariable=self.src, width=36).grid(row=0, column=1, columnspan=2, padx=10, pady=6, sticky="ew")
        ttk.Button(f, text="选择文件…", command=self.pick).grid(row=0, column=3, **pad)

        ttk.Label(f, text="② 个人条件").grid(row=1, column=0, **pad)
        self.degree = tk.StringVar(value="本科")
        ttk.Label(f, text="学历:").grid(row=2, column=0, **pad)
        ttk.Combobox(f, textvariable=self.degree, values=["大专", "本科", "研究生", "博士"],
                     state="readonly", width=10).grid(row=2, column=1, **pad)
        self.fresh = tk.BooleanVar(value=False)
        ttk.Checkbutton(f, text="我是应届毕业生", variable=self.fresh).grid(row=2, column=2, **pad)

        ttk.Label(f, text="基层工作年限:").grid(row=3, column=0, **pad)
        self.years = tk.Spinbox(f, from_=0, to=20, width=8); self.years.delete(0, "end"); self.years.insert(0, "0")
        self.years.grid(row=3, column=1, **pad)
        self.face = tk.StringVar(value="群众")
        ttk.Label(f, text="政治面貌:").grid(row=3, column=2, **pad)
        ttk.Combobox(f, textvariable=self.face, values=["群众", "共青团员", "中共党员"],
                     state="readonly", width=10).grid(row=3, column=3, **pad)

        ttk.Label(f, text="专业关键词(逗号分隔):").grid(row=4, column=0, **pad)
        self.majors = ttk.Entry(f, width=44)
        self.majors.insert(0, "软件工程,计算机类,计算机科学与技术")
        self.majors.grid(row=4, column=1, columnspan=3, padx=10, pady=6, sticky="ew")
        self.unlimited = tk.BooleanVar(value=True)
        ttk.Checkbutton(f, text='包含"专业不限"岗位', variable=self.unlimited).grid(row=5, column=1, **pad)

        self.run_button = ttk.Button(f, text="③ 筛选并另存为…", command=self.run)
        self.run_button.grid(row=6, column=1, pady=16, sticky="w")
        self.progress = ttk.Progressbar(f, mode="indeterminate")
        self.progress.grid(row=6, column=2, columnspan=2, sticky="ew", padx=10)
        self.logbox = tk.Text(f, height=8, width=70, state="disabled")
        self.logbox.grid(row=7, column=0, columnspan=4, padx=10, sticky="nsew")
        ttk.Label(f, text="本地离线处理 · 原文件保留 · 结果仅供初筛，报名前请人工复核。", wraplength=700).grid(row=8, column=0, columnspan=4, padx=10, pady=(16, 0), sticky="w")
        self.log("选择职位表，设置条件，再选择新的输出文件名。")

    def pick(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm")])
        if p:
            self.src.set(p)

    def log(self, msg):
        self.logbox.configure(state="normal")
        self.logbox.insert("end", msg + "\n")
        self.logbox.see("end")
        self.logbox.configure(state="disabled")

    def poll_events(self):
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == 'log':
                    self.log(payload)
                else:
                    self.running = False
                    self.progress.stop()
                    self.run_button.configure(state='normal')
                    if kind == 'done':
                        count, dst = payload
                        self.log(f'完成：{count} 个职位 → {dst}')
                        messagebox.showinfo('筛选完成', f'共保留 {count} 个职位。\n原文件未修改，结果已另存到：\n{dst}')
                    else:
                        self.log(f'未完成：{payload}')
                        messagebox.showerror('请检查后重试', payload)
        except queue.Empty:
            pass
        self.after(100, self.poll_events)

    def run(self):
        if self.running:
            return
        src = Path(self.src.get())
        if not src.is_file():
            messagebox.showwarning('请选择文件', '请先选择一个存在的 Excel 职位表。')
            return
        try:
            cfg = dict(degree=self.degree.get(), fresh=self.fresh.get(),
                       years=int(self.years.get()), face=self.face.get(),
                       majors=[k.strip() for k in re.split('[,，]', self.majors.get()) if k.strip()],
                       allow_unlimited=self.unlimited.get())
            validate_config(cfg)
        except ValueError as error:
            messagebox.showwarning('请检查个人条件', str(error))
            return
        suffix = src.suffix.lower()
        dst = filedialog.asksaveasfilename(defaultextension=suffix,
            initialfile=src.stem + '_筛选结果' + suffix, filetypes=[('Excel', '*' + suffix)])
        if not dst:
            return
        # The native save dialog confirms replacing an existing output file.
        overwrite = Path(dst).exists()
        self.running = True
        self.run_button.configure(state='disabled')
        self.progress.start(12)
        self.log('正在处理，请稍候…')
        def work():
            try:
                count = filter_workbook(src, dst, cfg,
                    log=lambda message: self.events.put(('log', message)), overwrite=overwrite)
                self.events.put(('done', (count, dst)))
            except Exception as error:
                self.events.put(('error', str(error)))
        threading.Thread(target=work, daemon=True).start()


if __name__ == "__main__":
    import sys
    application = App()
    if '--self-test' in sys.argv:
        application.update_idletasks()
        assert application.winfo_width() >= application.winfo_reqwidth()
        assert application.winfo_height() >= application.winfo_reqheight()
        application.after(500, application.destroy)
    application.mainloop()
