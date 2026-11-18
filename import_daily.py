# -*- coding: utf-8 -*-
"""
import_daily.py —— 运营日耗看板 · 数据导入脚本
================================================
把「运营日耗」文件夹里的每日 xlsx（宽表：一列一天）合并进看板主数据 data.js。
重复导入是安全的：同一「日期 + 账户 + 媒体 + 运营」的记录按最新值覆盖，不会重复计算。

用法（在 my-project 目录下运行）：
    python import_daily.py                     # 扫描默认日耗文件夹，导入所有新数据
    python import_daily.py --file "路径.xlsx"   # 只导入指定文件
    python import_daily.py --folder "某文件夹"  # 指定扫描文件夹
    python import_daily.py --dry-run           # 只预览，不写 data.js

默认日耗文件夹：C:\\Users\\Administrator\\Desktop\\运营日耗
默认数据文件：  本目录下 data.js
"""
import json
import os
import re
import shutil
import sys
from datetime import datetime, date
from collections import Counter

try:
    import openpyxl
except ImportError:
    print('缺少依赖，正在安装 openpyxl ...')
    import subprocess
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'openpyxl', '-q'])
    import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_JS = os.path.join(HERE, 'data.js')
DEFAULT_FOLDER = r'C:\Users\Administrator\Desktop\运营日耗'

# 日表组长名 → 主数据运营组（保持与历史数据同一口径）
GROUP_MAP = {'夏玲欢': '钱世豪', '余念情': '刘静'}
# 离职人员：从该日期起（含）不再计入（历史数据保留，仅剔除离职后的记录）
RETIRED = {'任洁': '2026-11-01', '敖丽': '2026-11-07'}
FILE_PAT = re.compile(r'^运营日耗.*\.xlsx$', re.IGNORECASE)

HEADERS = {
    'op': ['账户操作人', '运营', '操作人', '优化', '优化师', '运营人员'],
    'group': ['盟效运营组', '运营组组长', '组长', '运营组'],
    'client': ['客户', '客户名称', '广告主'],
    'media': ['媒体', '渠道平台', '平台'],
    'channel': ['渠道', '渠道类型'],
    'belong': ['运营归属'],
    'port': ['端口', '账户端口'],
    'accid': ['账户id', '账户ID', '账号id'],
    'sale': ['销售'],
    'cost': ['消耗', '消耗金额', '日耗', '消费', '金额'],
    'date': ['日期'],
}


def find_header(headers, keys):
    norm = lambda h: str(h or '').replace(' ', '').lower()
    for k in keys:
        for i, h in enumerate(headers):
            if norm(h) == norm(k):
                return i
    for k in keys:
        for i, h in enumerate(headers):
            if norm(k) in norm(h):
                return i
    return -1


def norm_date(v, ref_year=None):
    if v is None or v == '':
        return ''
    if isinstance(v, datetime):
        return v.strftime('%Y-%m-%d')
    if isinstance(v, date):
        return v.strftime('%Y-%m-%d')
    if isinstance(v, (int, float)):
        # Excel 日期序列号必须是整数（消耗金额如 77292.87 不能误判成日期）
        if isinstance(v, float) and not v.is_integer():
            return ''
        if 20000 < v < 80000:
            base = datetime(1899, 12, 30)
            return (base + __import__('datetime').timedelta(days=int(v))).strftime('%Y-%m-%d')
        return ''
    s = str(v).strip()
    if not s:
        return ''
    s = s.replace('年', '-').replace('月', '-').replace('日', '').replace('.', '-').replace('/', '-')
    m = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})', s)
    if m:
        return '%s-%02d-%02d' % (m.group(1), int(m.group(2)), int(m.group(3)))
    m = re.match(r'^(\d{1,2})-(\d{1,2})$', s)
    if m:
        y = ref_year or datetime.now().year
        return '%s-%02d-%02d' % (y, int(m.group(1)), int(m.group(2)))
    if re.match(r'^\d{5}$', s):
        base = datetime(1899, 12, 30)
        return (base + __import__('datetime').timedelta(days=int(s))).strftime('%Y-%m-%d')
    return ''


def to_num(v):
    if v is None or v == '':
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(',', '').replace('¥', '').replace(' ', '')
    try:
        return float(s)
    except ValueError:
        return None


def parse_xlsx(path):
    """返回记录列表（与 data.js 记录结构一致）"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()

    hi = -1
    cand = -1
    for i in range(min(len(rows), 8)):
        cells = rows[i]
        if find_header(cells, HEADERS['op']) >= 0 or find_header(cells, HEADERS['group']) >= 0:
            if cand < 0:
                cand = i
            if any(norm_date(c) for c in cells) or find_header(cells, HEADERS['date']) >= 0:
                hi = i
                break
    if hi < 0:
        # 有表头但没有日期列：多半是「月度汇总表」（如只有「8月消耗」一列）
        if cand >= 0:
            mcols = [str(c).strip() for c in rows[cand]
                     if re.match(r'^\d{1,2}\s*月\s*(消耗|消费|消耗金额|金额)?$',
                                 str(c or '').strip())]
            if mcols:
                raise ValueError('月度汇总表（只有「%s」这类列，无日期列），不能按日导入' % mcols[0])
            raise ValueError('表头里找不到日期列（宽表需「2026-08-01」式日期表头，明细表需「日期」列）')
        raise ValueError('没有识别到表头行（前 8 行未见「运营 / 运营组组长」列）')

    headers = rows[hi]
    idx = {k: find_header(headers, v) for k, v in HEADERS.items()}
    if idx['op'] < 0:
        raise ValueError('没有识别到「运营/账户操作人」列')

    date_cols = [(i, norm_date(h)) for i, h in enumerate(headers) if norm_date(h)]
    is_wide = bool(date_cols) and idx['cost'] < 0
    if not is_wide and idx['date'] < 0:
        raise ValueError('明细表需要「日期」列，或使用宽表（一列一天）')
    if not is_wide and idx['cost'] < 0:
        raise ValueError('没有识别到「消耗」列')

    recs = []
    for row in rows[hi + 1:]:
        op = str(row[idx['op']] or '').strip()
        if not op or op == '/':
            continue
        grp_raw = str(row[idx['group']] or '').strip() if idx['group'] >= 0 else ''
        grp = GROUP_MAP.get(grp_raw, grp_raw or '未分组')

        def g(k, dflt='/'):
            v = row[idx[k]] if idx[k] >= 0 else None
            return str(v) if v not in (None, '') else dflt

        if is_wide:
            for ci, dt in date_cols:
                if ci >= len(row):
                    continue
                v = to_num(row[ci])
                if v is None or v <= 0:
                    continue
                recs.append({
                    '日期': dt, '销售': g('sale'), '客户': g('client'), '媒体': g('media'),
                    '渠道': g('channel'), '运营归属': g('belong', '盟效运营'), '端口': g('port'),
                    '账户id': g('accid', ''), '盟效运营组': grp, '优化': op,
                    '账户负责人': op, '账户操作人': op, '消耗': round(v, 2),
                })
        else:
            dt = norm_date(row[idx['date']])
            if not dt:
                continue
            v = to_num(row[idx['cost']])
            if v is None or v <= 0:
                continue
            recs.append({
                '日期': dt, '销售': g('sale'), '客户': g('client'), '媒体': g('media'),
                '渠道': g('channel'), '运营归属': g('belong', '盟效运营'), '端口': g('port'),
                '账户id': g('accid', ''), '盟效运营组': grp, '优化': op,
                '账户负责人': op, '账户操作人': op, '消耗': round(v, 2),
            })
    # 离职人员：离职日期（含）之后的数据一律不计入
    if RETIRED:
        recs = [r for r in recs
                if not (r['账户操作人'] in RETIRED and r['日期'] >= RETIRED[r['账户操作人']])]
    return recs


def rec_key(r):
    acc = r['账户id'] if r['账户id'] not in ('', '/') else r['客户']
    # 客户名要进键：同 id 可能对应「（盟效）xxx」「（盟耀）xxx」两个主体
    return '|'.join([r['日期'], acc, r['客户'], r['端口'], r['媒体'], r['账户操作人']])


def load_data_js():
    if not os.path.exists(DATA_JS):
        return []
    with open(DATA_JS, encoding='utf-8') as f:
        s = f.read()
    i, j = s.find('['), s.rindex(']')
    return json.loads(s[i:j + 1])


def save_data_js(records, dry=False):
    records.sort(key=lambda r: (r['日期'], r['盟效运营组'], str(r['消耗'])))
    out = 'const RAW_DATA = ' + json.dumps(records, ensure_ascii=False, separators=(', ', ': '))
    if dry:
        return
    bak = DATA_JS + '.bak-' + datetime.now().strftime('%Y%m%d-%H%M%S')
    shutil.copy2(DATA_JS, bak)
    with open(DATA_JS, 'w', encoding='utf-8') as f:
        f.write(out)
    # 只保留最近 5 个备份
    baks = sorted(f for f in os.listdir(HERE) if f.startswith('data.js.bak-'))
    for old in baks[:-5]:
        os.remove(os.path.join(HERE, old))
    print('已备份原数据 ->', os.path.basename(bak))


def collect_files(args):
    files = []
    if args.file:
        files.append(os.path.abspath(args.file))
    else:
        folder = os.path.abspath(args.folder) if args.folder else DEFAULT_FOLDER
        if not os.path.isdir(folder):
            print('日耗文件夹不存在：', folder)
            print('可用 --folder 指定其他目录，或 --file 指定单个文件')
            return files
        for name in sorted(os.listdir(folder)):
            if FILE_PAT.match(name) and not name.startswith('~$'):
                files.append(os.path.join(folder, name))
    return [f for f in files if os.path.exists(f)]


def main():
    import argparse
    ap = argparse.ArgumentParser(description='把每日运营日耗 xlsx 合并进看板 data.js')
    ap.add_argument('--file', help='只导入这一个 xlsx')
    ap.add_argument('--folder', help='扫描此文件夹（默认: %s)' % DEFAULT_FOLDER)
    ap.add_argument('--dry-run', action='store_true', help='只预览，不写入')
    args = ap.parse_args()

    print('=' * 56)
    print(' 运营日耗看板 · 数据导入')
    print('=' * 56)

    current = load_data_js()
    keymap = {rec_key(r): r for r in current}
    print('现有记录: %d 条, 日期 %s ~ %s' % (
        len(current),
        min((r['日期'] for r in current), default='-'),
        max((r['日期'] for r in current), default='-')))

    files = collect_files(args)
    if not files:
        print('没有找到可导入的 xlsx 文件')
        return

    added = replaced = 0
    date_counter = Counter()
    parsed = []
    for f in files:
        name = os.path.basename(f)
        try:
            recs = parse_xlsx(f)
        except Exception as e:
            print('  [跳过] %-28s %s' % (name, e))
            continue
        if not recs:
            print('  [跳过] %-28s 没有解析出有效记录' % name)
            continue
        parsed.append((min(r['日期'] for r in recs), name, recs))

    # 按文件内最早日期升序处理：越新的文件越后合并，其数值覆盖旧文件的
    parsed.sort(key=lambda x: x[0])
    for min_date, name, recs in parsed:
        a = b = 0
        for r in recs:
            k = rec_key(r)
            if k in keymap:
                old = keymap[k]
                if abs(float(old['消耗']) - float(r['消耗'])) > 0.01:
                    old.update(r)
                    b += 1
            else:
                keymap[k] = r
                current.append(r)
                a += 1
        dates = sorted({r['日期'] for r in recs})
        print('  [%s] %-30s 新增 %4d 条, 更新 %3d 条 (%s ~ %s)' % (
            '预览' if args.dry_run else '导入', name, a, b,
            dates[0] if dates else '-', dates[-1] if dates else '-'))
        added += a
        replaced += b
        for r in recs:
            date_counter[r['日期']] += 1

    if added == 0 and replaced == 0:
        print('\n数据已是最新，无需更新。')
        return

    print('\n合计: 新增 %d 条, 更新 %d 条' % (added, replaced))
    print('覆盖日期: %s ~ %s' % (min(date_counter), max(date_counter)))
    if args.dry_run:
        print('（dry-run 预览模式，未写入 data.js）')
        return
    save_data_js(current)
    print('完成！data.js 已更新，刷新浏览器看板即可看到新数据。')


if __name__ == '__main__':
    main()
