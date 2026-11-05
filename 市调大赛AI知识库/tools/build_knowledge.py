#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成 20_知识 下的章节知识卡片，并自动追加本章题目索引。

卡片正文在 cards_ug_1.py / cards_ug_2.py / cards_zk.py 中维护。
运行：python3 tools/build_knowledge.py
"""
import json
import os
import re
from collections import Counter

from cards_ug_1 import CARDS as CARDS_UG_1
from cards_ug_2 import CARDS as CARDS_UG_2
from cards_zk import CARDS as CARDS_ZK

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
QDIR = os.path.join(ROOT, '10_题库')
OUT_UG = os.path.join(ROOT, '20_知识', '本科')
OUT_ZK = os.path.join(ROOT, '20_知识', '专科')

META = {
    'U01': ('本科', '01_市场调查概述.md'),
    'U02': ('本科', '02_数据收集.md'),
    'U03': ('本科', '03_调查设计.md'),
    'U04': ('本科', '04_问卷与量表.md'),
    'U05': ('本科', '05_数据处理与报告.md'),
    'U06': ('本科', '06_单变量分析.md'),
    'U07': ('本科', '07_抽样估计.md'),
    'U08': ('本科', '08_双变量关联分析.md'),
    'U09': ('本科', '09_多变量分析（一）.md'),
    'U10': ('本科', '10_多变量分析（二）.md'),
    'Z01': ('专科', 'Z01_市场调查概述.md'),
    'Z02': ('专科', 'Z02_调查方案设计.md'),
    'Z03': ('专科', 'Z03_问卷及量表设计.md'),
    'Z04': ('专科', 'Z04_资料采集方法.md'),
    'Z05': ('专科', 'Z05_调查实施与报告.md'),
    'Z06': ('专科', 'Z06_描述统计基础.md'),
    'Z07': ('专科', 'Z07_抽样调查与估计.md'),
    'Z08': ('专科', 'Z08_数据分析方法.md'),
}


def clean(s):
    s = re.sub(r'\s+', ' ', s or '').strip()
    return s.replace('|', '\\|')


def short(s, n=72):
    s = clean(s)
    return s if len(s) <= n else s[:n] + '…'


def related_section(cid, qs):
    members = [q for q in qs if q['chapter_id'] == cid]
    members.sort(key=lambda q: (q['year'], 0 if q['group'] == '本科' else 1, q['num']))
    hf = [q for q in members if q['high_freq']]
    ep = [q for q in members if q['error_prone']]
    reviewed = [q for q in members if q.get('explanation')]
    topics = Counter(q['topic'] for q in members)
    priority = sorted({q['id']: q for q in hf + ep}.values(),
                      key=lambda q: (not q['error_prone'], not q['high_freq'], q['year'], q['num']))
    lines = ['<!-- AUTO-RELATED:START -->', '## 本章题目索引', '',
             f'- 本章共 **{len(members)}** 题；跨届高频 **{len(hf)}** 题；已写解析的易错题 **{len(ep)}** 题；已有解析 **{len(reviewed)}** 题。',
             f'- 完整题面：`../10_题库/按章节/{cid}_{META[cid][1].split("_", 1)[1].replace(".md", "")}.md`',
             '- 主题分布：' + '、'.join(f'{k} {v}' for k, v in topics.most_common())]
    if priority:
        lines += ['', '### 优先看（高频 / 易错）', '',
                  '| ID | 题型 | 答案 | 题干 | 标记 |', '|---|---|---|---|---|']
        for q in priority:
            mark = ('★' if q['high_freq'] else '') + ('⚠' if q['error_prone'] else '')
            lines.append(f"| `{q['id']}` | {q['type']} | **{q['answer']}** | {short(q['stem'])} | {mark} |")
    lines += ['', '### 全部题目 ID', '']
    by_year = {}
    for q in members:
        by_year.setdefault((q['year'], q['group']), []).append(q['id'])
    for (year, group), ids in sorted(by_year.items()):
        lines.append(f"- 第{year}届{group}（{len(ids)}）：" + '、'.join(f'`{i}`' for i in ids))
    lines += ['', '<!-- AUTO-RELATED:END -->', '']
    return '\n'.join(lines)


def main():
    qs = [json.loads(line) for line in open(os.path.join(QDIR, 'questions.jsonl'), encoding='utf8') if line.strip()]
    cards = {}
    cards.update(CARDS_UG_1)
    cards.update(CARDS_UG_2)
    cards.update(CARDS_ZK)
    missing = [cid for cid in META if cid not in cards]
    if missing:
        raise SystemExit(f'缺少章节正文：{missing}')
    for cid, (group, filename) in META.items():
        outdir = OUT_UG if group == '本科' else OUT_ZK
        os.makedirs(outdir, exist_ok=True)
        body = cards[cid].strip() + '\n\n' + related_section(cid, qs)
        open(os.path.join(outdir, filename), 'w', encoding='utf8').write(body)
    # 知识库总索引
    lines = ['# 知识卡片索引', '',
             '按第17届网考提纲章节组织。每章包含：核心概念、分类对比、公式与条件、易混点、答题套路、本章题目索引。',
             '', '## 本科（10 章）', '', '| 章节 | 知识卡片 | 题量 | 高频 | 易错 |', '|---|---|---:|---:|---:|']
    for cid, (group, filename) in META.items():
        if group != '本科':
            continue
        mem = [q for q in qs if q['chapter_id'] == cid]
        lines.append(f"| {mem[0]['chapter'] if mem else cid} | [{filename}](本科/{filename}) | {len(mem)} | "
                     f"{sum(1 for q in mem if q['high_freq'])} | {sum(1 for q in mem if q['error_prone'])} |")
    lines += ['', '## 专科（8 章）', '', '| 章节 | 知识卡片 | 题量 | 高频 | 易错 |', '|---|---|---:|---:|---:|']
    for cid, (group, filename) in META.items():
        if group != '专科':
            continue
        mem = [q for q in qs if q['chapter_id'] == cid]
        lines.append(f"| {mem[0]['chapter'] if mem else cid} | [{filename}](专科/{filename}) | {len(mem)} | "
                     f"{sum(1 for q in mem if q['high_freq'])} | {sum(1 for q in mem if q['error_prone'])} |")
    lines += ['', '## 怎么用', '',
              '1. 先按题干关键词在 `30_速查/关键词索引.md` 或 `tools/search.py` 找题。',
              '2. 找到题目后，回到对应章节卡片核对概念边界和公式条件。',
              '3. 卡片中的“待核对”表示教材口径可能有差异，考试时以本校教材为准。', '']
    open(os.path.join(ROOT, '20_知识', 'README.md'), 'w', encoding='utf8').write('\n'.join(lines))
    print('知识卡片生成完成：', len(META), '章')


if __name__ == '__main__':
    main()
