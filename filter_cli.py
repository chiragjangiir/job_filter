"""Command-line entry point; never requires a display server."""
import argparse
import re
import sys
from filter_core import filter_workbook


def main():
    # Redirected output on Windows can default to a non-Chinese code page.
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            stream.reconfigure(encoding='utf-8', errors='backslashreplace')
    parser = argparse.ArgumentParser(description='离线筛选 Excel 公考职位表。结果仅供初筛，报名前请人工复核。')
    parser.add_argument('source')
    parser.add_argument('destination')
    parser.add_argument('--degree', choices=['大专', '本科', '研究生', '博士'], default='本科')
    parser.add_argument('--majors', default='软件工程,计算机类,计算机科学与技术')
    parser.add_argument('--years', type=int, default=0)
    parser.add_argument('--fresh', action='store_true')
    parser.add_argument('--face', choices=['群众', '共青团员', '中共党员'], default='群众')
    parser.add_argument('--exclude-unlimited', action='store_true')
    parser.add_argument('--overwrite', action='store_true', help='允许替换已存在的输出文件；永不覆盖输入文件')
    args = parser.parse_args()
    cfg = dict(degree=args.degree, majors=[x.strip() for x in re.split('[,，]', args.majors) if x.strip()], years=args.years, fresh=args.fresh, face=args.face, allow_unlimited=not args.exclude_unlimited)
    try:
        count = filter_workbook(args.source, args.destination, cfg, overwrite=args.overwrite)
    except (ValueError, OSError) as error:
        print(f'未完成：{error}', file=sys.stderr)
        return 1
    print(f'完成：{count} 个职位 → {args.destination}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
