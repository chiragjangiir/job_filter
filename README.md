# 公考职位表筛选器 · Job Filter

**把一整份 Excel 职位表，缩小成值得逐条核对的候选清单。**

[English](README.en.md) · [试用样例](examples/demo.xlsx) · [预期结果](examples/demo-filtered.xlsx) · [使用指南](docs/USAGE.md) · [版本记录](CHANGELOG.md)

[![Checks](https://github.com/shanghe000/job_filter/actions/workflows/quality.yml/badge.svg)](https://github.com/shanghe000/job_filter/actions/workflows/quality.yml)
[![Build](https://github.com/shanghe000/job_filter/actions/workflows/build.yml/badge.svg)](https://github.com/shanghe000/job_filter/actions/workflows/build.yml)

一个本地运行的 Python 桌面工具：选择 Excel，设置学历、专业、应届身份等条件，将候选岗位另存为新文件。也可通过命令行处理。

- **离线处理**：不上传职位表或个人条件，不需要账号。
- **保留原件**：拒绝覆盖输入文件，保存失败时保留已有输出。
- **可核对**：提供完全虚构的输入与输出样例，明确哪些条件属于启发式匹配。
- **保留基础格式**：对支持的普通表格保留单元格样式、列宽、行高与标题区合并单元格。

> 本工具用于候选初筛，不判断报考资格。正式报名时请逐条核对原始公告和岗位要求。

![桌面应用实际界面](docs/images/desktop.png)

## 一分钟试用

需要 Python 3.10+。图形界面需要 Tcl/Tk 8.6+；macOS 建议使用 python.org 的安装包，避免系统旧版 Tk。

```sh
git clone https://github.com/shanghe000/job_filter.git
cd job_filter
python -m venv .venv
# macOS / Linux
source .venv/bin/activate
# Windows PowerShell 改为：.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python filter_gui.py
```

选择 `examples/demo.xlsx`，保持默认条件，另存为新文件。应保留 **A · 软件支持** 和 **D · 综合服务** 两个岗位。
可与 `examples/demo-filtered.xlsx` 比较。样例中的岗位、条件均为虚构。

### 不打开界面

```sh
python filter_cli.py examples/demo.xlsx result.xlsx --degree 本科 --majors 计算机 --years 0
```

重复运行需使用新的输出名，或明确加上 `--overwrite`。输入文件始终不能作为输出目标。

## 下载桌面包

[Actions → Build desktop apps](https://github.com/shanghe000/job_filter/actions/workflows/build.yml) 提供 Windows x64 与 macOS Apple Silicon 的打包产物，下载 Actions 产物需登录 GitHub。
发布流程会先生成包含压缩包和校验值的 **Release 草稿**；只有维护者核对后公开的版本，才会出现在 [Releases](https://github.com/shanghe000/job_filter/releases)。
桌面包尚未进行商业代码签名或 Apple 公证。Intel Mac 请从源码运行。

## 支持范围

| 项目 | 当前行为 |
| --- | --- |
| 格式 | `.xlsx`；`.xlsm` 保留 VBA 存储，不执行宏；不支持 `.xls` |
| 表头 | 前 15 行中同时出现“专业”和“学历”的一行 |
| 多工作表 | 处理符合条件的工作表，未识别的页原样保留并写入日志 |
| 学历、专业 | 关键词初筛；按学历对应的专业段落匹配 |
| 空行 | 从已识别工作表的数据区移除，不计作职位 |
| 复杂工作簿 | 含公式、图表、图片、表格对象、条件格式、验证或命名范围时拒绝处理 |
| 合并单元格 | 仅支持表头及其上方；数据区合并需先在副本中展开 |

不承诺完整保留 Excel 的所有高级功能。详见 [使用指南与限制](docs/USAGE.md)。

## 参与开发

```sh
python -m unittest discover -s tests -v
python scripts/make_example.py
```

核心逻辑在 `filter_core.py`，桌面界面与命令行分别位于 `filter_gui.py` 和 `filter_cli.py`。
请使用合成数据复现问题，避免在 Issue 中上传个人信息或未授权表格。

[贡献说明](CONTRIBUTING.md) · [安全政策](SECURITY.md) · [路线图](docs/ROADMAP.md)

## 许可证

[MIT License](LICENSE) © shanghe000。欢迎使用、修改和贡献；分发时请保留许可声明。
