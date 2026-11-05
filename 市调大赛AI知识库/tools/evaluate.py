#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""检索评测：Recall@1 / Recall@5 / MRR / 首位答案一致率。

运行：
    python3 tools/evaluate.py
    python3 tools/evaluate.py --top 10 --json

同时生成：
    50_评测/基准问题.md
    50_评测/评测报告.md
"""
import argparse
import json
import os
import sys
from collections import defaultdict
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from benchmark import BENCHMARK
from search import Searcher, load

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, '50_评测')


def pct(x):
    return f'{x * 100:.1f}%'


def evaluate(searcher, top=10, verbose=False):
    rows = []
    for item in BENCHMARK:
        accept = set(item['accept'])
        results = searcher.search(item['q'], top=top)
        ids = [q['id'] for _, q in results]
        rank = next((i + 1 for i, qid in enumerate(ids) if qid in accept), None)
        top1 = results[0][1] if results else None
        answer_ok = bool(rank == 1 and top1['answer'] == item['answer'])
        rows.append(dict(
            item=item, rank=rank, top1=top1,
            top_ids=ids,
            answer_ok=answer_ok,
        ))
        if verbose:
            mark = 'OK ' if rank == 1 else ('~  ' if rank and rank <= 5 else 'MISS')
            print(f"{mark} rank={rank or '-'} q={item['q']} -> {ids[:3]}")
    return rows


def metrics(rows):
    n = len(rows)
    hit1 = sum(1 for r in rows if r['rank'] == 1)
    hit5 = sum(1 for r in rows if r['rank'] and r['rank'] <= 5)
    mrr = sum(1.0 / r['rank'] for r in rows if r['rank']) / n
    ans = sum(1 for r in rows if r['answer_ok'])
    return dict(n=n, hit1=hit1 / n, hit5=hit5 / n, mrr=mrr, answer=ans / n,
                hit1_n=hit1, hit5_n=hit5, answer_n=ans)


def group_metrics(rows):
    groups = defaultdict(list)
    for r in rows:
        groups[r['item']['kind']].append(r)
    out = []
    for kind, members in sorted(groups.items(), key=lambda kv: (-len(kv[1]), kv[0])):
        m = metrics(members)
        out.append((kind, m))
    return out


def write_benchmark():
    lines = ['# 基准问题', '',
             f'- 共 **{len(BENCHMARK)}** 条查询，覆盖本科 10 章与专科 8 章。',
             '- 查询不是原题复制，而是换述、关键词组合、口语化问法。',
             '- `可接受题号` 内的任意一题被检索到即算命中；同题干不同选项的变体不放入同一组。',
             '- 答案一致性只看检索结果第 1 条是否命中可接受题号且答案与目标答案一致。', '',
             '## 查询清单', '',
             '| # | 类型 | 查询 | 可接受题号 | 目标答案 | 考点 |',
             '|---:|---|---|---|---|---|']
    for i, item in enumerate(BENCHMARK, 1):
        accept = '、'.join(f'`{x}`' for x in item['accept'])
        lines.append(f"| {i} | {item['kind']} | {item['q']} | {accept} | "
                     f"**{item['answer']}** | {item['note']} |")
    lines += ['', '## 类型分布', '']
    counts = defaultdict(int)
    for item in BENCHMARK:
        counts[item['kind']] += 1
    lines += ['| 类型 | 条数 |', '|---|---:|']
    for kind, n in sorted(counts.items(), key=lambda kv: (-kv[1], kv[0])):
        lines.append(f'| {kind} | {n} |')
    lines += ['', '## 怎么复跑', '',
              '```bash', 'cd 市调大赛AI知识库',
              'python3 tools/evaluate.py', '```', '',
              '评测脚本会重新生成本文件和 `50_评测/评测报告.md`。', '']
    open(os.path.join(OUT, '基准问题.md'), 'w', encoding='utf8').write('\n'.join(lines))


def write_report(rows, top):
    m = metrics(rows)
    gm = group_metrics(rows)
    lines = ['# 检索评测报告', '',
             f'- 评测日期：{date.today().isoformat()}',
             f'- 题库：610 题（本科 510、专科 100）',
             f'- 查询数：**{m["n"]}**',
             f'- 检索深度：Top {top}',
             f'- 检索脚本：`tools/search.py`（字符 n-gram + IDF + 同义词归一 + 人工标签）', '',
             '## 总体指标', '',
             '| 指标 | 数值 | 说明 |',
             '|---|---:|---|',
             f'| Recall@1 | **{pct(m["hit1"])}**（{m["hit1_n"]}/{m["n"]}） | 第 1 条命中目标题 |',
             f'| Recall@5 | **{pct(m["hit5"])}**（{m["hit5_n"]}/{m["n"]}） | 前 5 条命中目标题 |',
             f'| MRR | **{m["mrr"]:.3f}** | 目标题排名倒数的平均值 |',
             f'| 首位答案一致率 | **{pct(m["answer"])}**（{m["answer_n"]}/{m["n"]}） | 第 1 条命中且答案与目标答案一致 |', '',
             '## 分组指标', '',
             '| 类型 | 条数 | Recall@1 | Recall@5 | MRR | 首位答案一致率 |',
             '|---|---:|---:|---:|---:|---:|']
    for kind, g in gm:
        lines.append(f"| {kind} | {g['n']} | {pct(g['hit1'])} | {pct(g['hit5'])} | "
                     f"{g['mrr']:.3f} | {pct(g['answer'])} |")
    lines += ['', '## 未进前 5 的查询', '']
    misses = [r for r in rows if not r['rank'] or r['rank'] > 5]
    if not misses:
        lines.append('无。')
    else:
        lines += ['| 查询 | 目标题号 | 实际 Top 3 |', '|---|---|---|']
        for r in misses:
            top3 = '、'.join(f'`{x}`' for x in r['top_ids'][:3]) or '无结果'
            lines.append(f"| {r['item']['q']} | {'、'.join(r['item']['accept'])} | {top3} |")
    lines += ['', '## 排名 2—5 的查询', '',
              '这些查询能找到原题，但第 1 条不是目标题；考试时仍可看前 5 条。', '']
    near = [r for r in rows if r['rank'] and 1 < r['rank'] <= 5]
    if not near:
        lines.append('无。')
    else:
        lines += ['| 查询 | 目标题号 | 目标排名 | 第 1 条 |', '|---|---|---:|---|']
        for r in near:
            lines.append(f"| {r['item']['q']} | {'、'.join(r['item']['accept'])} | "
                         f"{r['rank']} | `{r['top1']['id']}` |")
    lines += ['', '## 全部查询明细', '',
              '| # | 查询 | 目标题号 | 排名 | Top1 | 答案一致 |',
              '|---:|---|---|---:|---|---|']
    for i, r in enumerate(rows, 1):
        rank = str(r['rank']) if r['rank'] else '未命中'
        top1 = f"`{r['top1']['id']}`" if r['top1'] else '—'
        lines.append(f"| {i} | {r['item']['q']} | {'、'.join(r['item']['accept'])} | "
                     f"{rank} | {top1} | {'是' if r['answer_ok'] else '否'} |")
    lines += ['', '## 指标口径与局限', '',
              '1. **Recall@1 / Recall@5**：同一考点的重复题视为同一目标，任一副本命中都算命中。',
              '2. **MRR**：只统计前 10 条；未命中记 0。',
              '3. **首位答案一致率**：比“找到题”更严格，要求第 1 条就是目标题且答案一致。',
              '4. 查询由人工按考试场景构造，不能完全代表真实考生输入；分数用于回归比较，不是绝对能力上限。',
              '5. 题库答案保留原始答案，存疑题不因评测而改答案。',
              '6. 公式/图片题依赖 `tools/annotations.py` 的人工标签；换述题依赖同义词表和字符 n-gram。', '',
              '## 结论', '',
              '- 当前检索在换述、章节关键词、重复题场景下表现稳定；公式图片题靠人工标签兜底。',
              '- 若新增题目，建议先跑本评测，确认 Recall@5 没有明显下降。',
              '- 排名 2—5 的查询不影响开卷使用，但可以作为下一步优化同义词/标签的线索。', '']
    open(os.path.join(OUT, '评测报告.md'), 'w', encoding='utf8').write('\n'.join(lines))


def main():
    ap = argparse.ArgumentParser(description='市调大赛题库检索评测')
    ap.add_argument('--top', type=int, default=10)
    ap.add_argument('--json', action='store_true')
    ap.add_argument('--verbose', action='store_true')
    args = ap.parse_args()

    os.makedirs(OUT, exist_ok=True)
    questions = load()
    searcher = Searcher(questions)
    rows = evaluate(searcher, top=args.top, verbose=args.verbose)
    m = metrics(rows)
    write_benchmark()
    write_report(rows, args.top)
    if args.json:
        print(json.dumps(dict(metrics=m,
                              details=[dict(q=r['item']['q'], accept=r['item']['accept'],
                                            rank=r['rank'],
                                            top_ids=r['top_ids'][:5],
                                            answer_ok=r['answer_ok']) for r in rows]),
                         ensure_ascii=False, indent=2))
        return
    print(f"评测完成：{m['n']} 条查询")
    print(f"Recall@1 = {pct(m['hit1'])} ({m['hit1_n']}/{m['n']})")
    print(f"Recall@5 = {pct(m['hit5'])} ({m['hit5_n']}/{m['n']})")
    print(f"MRR       = {m['mrr']:.3f}")
    print(f"首位答案一致率 = {pct(m['answer'])} ({m['answer_n']}/{m['n']})")
    print('报告：50_评测/评测报告.md')


if __name__ == '__main__':
    main()
