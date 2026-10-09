#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
证监会《境内企业境外发行证券和上市备案情况表（首次公开发行及全流通）》自动采集脚本

功能链路：
  1. 用官方栏目接口解析“境外证券发行”栏目的 channelId（不写死，栏目调整也不影响）
  2. 调用官方列表接口，找到最新一期备案情况表及其 xlsx 附件直链
  3. 下载附件并计算 SHA-256，与上一期比对：没变化直接退出，不重复解析
  4. 有变化则解析 Excel → 与上期做变更比对（新增 / 状态流转 / 移除）
  5. 输出 data/filings.json（前端直接 fetch）、data/history.json（期次留痕）、
     data/raw/<期次>.xlsx（原始附件存档）

依赖：requests, openpyxl      （pip install requests openpyxl）

用法：
  python scripts/fetch_filings.py           # 检查更新；无变化则跳过（退出码 0）
  python scripts/fetch_filings.py --force   # 忽略哈希比对，强制重新抓取解析
  python scripts/fetch_filings.py --no-archive   # 不保存原始 xlsx 存档

数据来源均为证监会官网公开信息，仅用于学术演示。
"""

import argparse
import hashlib
import json
import os
import re
import sys
import time
from datetime import datetime, timezone, timedelta

import requests
from openpyxl import load_workbook

# ---------------------------------------------------------------- 常量配置
BASE = 'https://www.csrc.gov.cn'
# 栏目编码 c101935 = 政务信息公开 / 境外证券发行（如栏目调整，改这里即可）
CHANNEL_CODE = 'c101935'
LIST_API = BASE + '/searchList/{channel_id}?_isAgg=true&_isJson=true&_pageSize=30&_template=index&_rangeTimeGte=&_channelName=&page={page}'
CHANNEL_API = BASE + '/getChannelList?channelCode=' + CHANNEL_CODE

# 目标数据表的标题特征
TARGET_TITLE_KEY = '境内企业境外发行证券和上市备案情况表'

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 '
                  '(KHTML, like Gecko) Chrome/124.0 Safari/537.36',
    'Referer': BASE + '/csrc/%s/zfxxgk_zdgk.shtml' % CHANNEL_CODE,
    'Accept-Language': 'zh-CN,zh;q=0.9',
}

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(ROOT, 'data')
RAW_DIR = os.path.join(DATA_DIR, 'raw')
FILINGS_JSON = os.path.join(DATA_DIR, 'filings.json')
HISTORY_JSON = os.path.join(DATA_DIR, 'history.json')

CST = timezone(timedelta(hours=8))   # 北京时间


# ---------------------------------------------------------------- 网络请求
def http_get(url, binary=False, retries=3, timeout=40):
    """带重试的 GET。官网偶发 5xx/超时，重试 3 次并退避。"""
    last_err = None
    for attempt in range(1, retries + 1):
        try:
            resp = requests.get(url, headers=HEADERS, timeout=timeout)
            resp.raise_for_status()
            return resp.content if binary else resp.text
        except Exception as err:              # noqa: BLE001
            last_err = err
            print('  ! 第 %d 次请求失败：%s' % (attempt, err))
            if attempt < retries:
                time.sleep(2 * attempt)
    raise RuntimeError('请求失败：%s（%s）' % (url, last_err))


# ---------------------------------------------------------------- 步骤 1：栏目 ID
def resolve_channel_id():
    """通过 channelCode 换取 channelId，避免把 ID 写死在代码里。"""
    data = json.loads(http_get(CHANNEL_API))
    channel_id = (data.get('results') or {}).get('channelId')
    if not channel_id:
        raise RuntimeError('未能解析栏目 channelId，请检查 CHANNEL_CODE：%s' % CHANNEL_CODE)
    return channel_id


# ---------------------------------------------------------------- 步骤 2：找最新一期
def find_latest_table(channel_id, max_pages=3):
    """在栏目列表中找到最新一期备案情况表，返回期次、详情页、附件直链等信息。"""
    for page in range(1, max_pages + 1):
        data = json.loads(http_get(LIST_API.format(channel_id=channel_id, page=page)))
        results = (data.get('data') or {}).get('results') or []
        for item in results:
            meta = (item.get('domainMetaList') or [{}])[0] if isinstance(item.get('domainMetaList'), list) else (item.get('domainMetaList') or {})
            title = (item.get('title') or meta.get('title') or '').strip()
            if TARGET_TITLE_KEY not in title:
                continue
            # 在附件列表里挑 xlsx
            xlsx = next((f for f in (item.get('resList') or [])
                         if str(f.get('fileName', '')).lower().endswith('.xlsx')), None)
            if not xlsx:
                continue
            period = extract_period(title)
            return {
                'title': title,
                'period': period,                                  # 例：截至2026年9月30日
                'periodDate': period_to_iso(period),               # 例：2026-09-30
                'pageUrl': normalize_url(item.get('url') or meta.get('url') or ''),
                'fileUrl': BASE + (xlsx.get('filePath') or ''),
                'fileName': xlsx.get('fileName'),
                'publishedAt': meta.get('publishedTimeStr') or '',
            }
        if not results:
            break
    raise RuntimeError('列表中未找到《%s》，请确认官网栏目结构是否变化' % TARGET_TITLE_KEY)


def extract_period(title):
    """从标题中取“（截至YYYY年M月D日）”里的期次描述。"""
    m = re.search(r'（截至([^）]+)）', title)
    return ('截至' + m.group(1)) if m else ''


def period_to_iso(period):
    """'截至2026年9月30日' -> '2026-09-30'"""
    m = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日', period or '')
    if not m:
        return ''
    return '%s-%02d-%02d' % (m.group(1), int(m.group(2)), int(m.group(3)))


def normalize_url(url):
    if not url:
        return ''
    if url.startswith('//'):
        return 'https:' + url
    if url.startswith('/'):
        return BASE + url
    return url


# ---------------------------------------------------------------- 步骤 3：下载与哈希
def download_attachment(file_url):
    content = http_get(file_url, binary=True)
    return content, hashlib.sha256(content).hexdigest()


# ---------------------------------------------------------------- 步骤 4：解析 Excel
def parse_workbook(content):
    """解析官方备案情况表：第 1 行大标题、第 3-4 行双表头、之后为数据行。

    列序：A 序号 | B 企业名称 | C 申报类型 | D 申报主体 | E 拟上市证券交易所
         | F 保荐人/主承销商 | G 境内律师 | H 接收日期 | I 备案状态 | J 备注
    """
    import io
    wb = load_workbook(io.BytesIO(content), data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]

    rows_all = list(ws.iter_rows(values_only=True))
    header_idx = None
    for i, row in enumerate(rows_all):
        if row and str(row[0] or '').strip() == '序号':
            header_idx = i
            break
    if header_idx is None:
        raise RuntimeError('未找到表头（序号列），Excel 结构可能已调整')

    rows = []
    for row in rows_all[header_idx + 1:]:
        first = str(row[0] or '').strip()
        if not re.match(r'^\d+$', first):        # 数据行结束
            continue
        rows.append({
            'no': int(first),
            'name': clean(row[1]),
            'type': clean(row[2]),
            'subject': clean(row[3]),
            'venue': clean(row[4]),
            'sponsor': clean(row[5]),
            'lawyer': clean(row[6]),
            'date': clean_date(row[7]),
            'status': clean(row[8]),
            'remark': clean(row[9]),
        })
    wb.close()
    return rows


def clean(v):
    """去掉单元格内换行、首尾空白；空值统一为 ''。"""
    if v is None:
        return ''
    return re.sub(r'\s+', ' ', str(v)).strip()


def clean_date(v):
    """Excel 里接收日期是日期类型，统一成 '2026年9月30日' 与官网展示口径一致。"""
    if v is None:
        return ''
    if isinstance(v, datetime):
        return '%d年%d月%d日' % (v.year, v.month, v.day)
    s = str(v).strip()
    m = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})', s)
    if m:
        return '%d年%d月%d日' % (int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return s


# ---------------------------------------------------------------- 步骤 5：变更比对
def diff_rows(old_rows, new_rows):
    """对比两期数据：新增企业 / 备案状态流转 / 从名单移除。

    以企业名称为键（同一企业可能存在多条记录，用状态集合比较）。
    """
    def index(rows):
        m = {}
        for r in rows:
            m.setdefault(r['name'], []).append(r['status'])
        return m

    old_map, new_map = index(old_rows), index(new_rows)
    added = [n for n in new_map if n not in old_map]
    removed = [n for n in old_map if n not in new_map]
    changed = []
    for name, statuses in new_map.items():
        if name in old_map and sorted(statuses) != sorted(old_map[name]):
            changed.append({
                'name': name,
                'from': '、'.join(sorted(set(old_map[name]))),
                'to': '、'.join(sorted(set(statuses))),
            })
    return {
        'added': sorted(added),
        'removed': sorted(removed),
        'statusChanged': changed,
        'addedCount': len(added),
        'statusChangedCount': len(changed),
        'removedCount': len(removed),
    }


def count_by(rows, key):
    out = {}
    for r in rows:
        out[r[key]] = out.get(r[key], 0) + 1
    return dict(sorted(out.items(), key=lambda kv: -kv[1]))


# ---------------------------------------------------------------- 读写 JSON
def load_json(path, default):
    if os.path.exists(path):
        try:
            with open(path, encoding='utf-8') as f:
                return json.load(f)
        except Exception:                     # noqa: BLE001
            return default
    return default


def dump_json(path, meta, rows):
    """meta 美化输出、rows 每行一条，便于 git diff 逐行比对。"""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    meta_txt = json.dumps(meta, ensure_ascii=False, indent=2).replace('\n', '\n  ')
    with open(path, 'w', encoding='utf-8') as f:
        f.write('{\n  "meta": ' + meta_txt + ',\n  "rows": [\n')
        for i, r in enumerate(rows):
            f.write('    ' + json.dumps(r, ensure_ascii=False) + (',' if i < len(rows) - 1 else '') + '\n')
        f.write('  ]\n}\n')


# ---------------------------------------------------------------- 主流程
def main():
    ap = argparse.ArgumentParser(description='证监会备案情况表自动采集')
    ap.add_argument('--force', action='store_true', help='忽略哈希比对，强制重新解析')
    ap.add_argument('--no-archive', action='store_true', help='不保存原始 xlsx 存档')
    args = ap.parse_args()

    os.makedirs(DATA_DIR, exist_ok=True)
    prev = load_json(FILINGS_JSON, {})
    prev_meta = prev.get('meta') or {}
    prev_rows = prev.get('rows') or []

    # 1) 栏目
    channel_id = resolve_channel_id()
    print('✔ 栏目 channelId：%s' % channel_id)

    # 2) 最新一期
    latest = find_latest_table(channel_id)
    print('✔ 最新一期：%s（%s）' % (latest['period'], latest['periodDate']))
    print('  详情页：%s' % latest['pageUrl'])

    # 3) 下载 + 哈希比对
    content, sha256 = download_attachment(latest['fileUrl'])
    print('✔ 附件已下载：%s（%.0f KB，sha256=%s…）' % (latest['fileName'], len(content) / 1024, sha256[:12]))

    if not args.force and prev_meta.get('sha256') == sha256:
        print('· 文件哈希与上期一致，数据无更新，跳过解析。')
        return 0

    # 4) 解析
    rows = parse_workbook(content)
    if not rows:
        raise RuntimeError('解析结果为空，已中止以免覆盖既有数据')
    print('✔ 解析完成：%d 条记录' % len(rows))

    # 5) 变更比对
    changes = diff_rows(prev_rows, rows) if prev_rows else {
        'added': [], 'removed': [], 'statusChanged': [],
        'addedCount': 0, 'statusChangedCount': 0, 'removedCount': 0,
    }
    print('✔ 较上期：新增 %d 家 / 状态变更 %d 家 / 移除 %d 家'
          % (changes['addedCount'], changes['statusChangedCount'], changes['removedCount']))

    # 6) 输出
    now = datetime.now(CST)
    meta = {
        'title': latest['title'],
        'period': latest['period'],
        'periodDate': latest['periodDate'],
        'publishedAt': latest['publishedAt'],
        'pageUrl': latest['pageUrl'],
        'fileUrl': latest['fileUrl'],
        'fileName': latest['fileName'],
        'sha256': sha256,
        'fetchedAt': now.strftime('%Y-%m-%d %H:%M:%S'),
        'total': len(rows),
        'statusCount': count_by(rows, 'status'),
        'typeCount': count_by(rows, 'type'),
        'source': '中国证监会官网公开信息（自动采集）',
        'changes': changes,
    }
    dump_json(FILINGS_JSON, meta, rows)
    print('✔ 已写出：%s' % FILINGS_JSON)

    # 期次留痕
    history = load_json(HISTORY_JSON, [])
    history = [h for h in history if h.get('periodDate') != meta['periodDate']]
    history.append({
        'periodDate': meta['periodDate'],
        'period': meta['period'],
        'total': meta['total'],
        'statusCount': meta['statusCount'],
        'sha256': sha256,
        'addedCount': changes['addedCount'],
        'statusChangedCount': changes['statusChangedCount'],
        'fetchedAt': meta['fetchedAt'],
        'fileName': meta['fileName'],
    })
    history.sort(key=lambda h: h.get('periodDate') or '')
    with open(HISTORY_JSON, 'w', encoding='utf-8') as f:
        json.dump(history, f, ensure_ascii=False, indent=2)
    print('✔ 已更新期次留痕：%s（累计 %d 期）' % (HISTORY_JSON, len(history)))

    # 原始附件存档
    if not args.no_archive and meta['periodDate']:
        os.makedirs(RAW_DIR, exist_ok=True)
        raw_path = os.path.join(RAW_DIR, meta['periodDate'] + '.xlsx')
        if not os.path.exists(raw_path):
            with open(raw_path, 'wb') as f:
                f.write(content)
            print('✔ 原始附件已存档：%s' % raw_path)

    print('UPDATED')     # 供 CI 判断是否产生变更
    return 0


if __name__ == '__main__':
    try:
        sys.exit(main())
    except Exception as exc:                  # noqa: BLE001
        print('✘ 采集失败：%s' % exc, file=sys.stderr)
        sys.exit(1)
