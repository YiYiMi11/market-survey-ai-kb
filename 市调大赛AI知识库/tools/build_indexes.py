# -*- coding: utf-8 -*-
"""从 questions.jsonl 生成速查索引、统计与解析文件。"""
import os, json, re
from collections import Counter, defaultdict

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
QDIR = os.path.join(ROOT, '10_题库')
OUT = os.path.join(ROOT, '30_速查')
QUAL = os.path.join(ROOT, '40_质量')
os.makedirs(OUT, exist_ok=True)
os.makedirs(QUAL, exist_ok=True)

QS = [json.loads(l) for l in open(os.path.join(QDIR, 'questions.jsonl'), encoding='utf8')]
QS.sort(key=lambda q: (q['year'], 0 if q['group'] == '本科' else 1, q['num']))
BY_ID = {q['id']: q for q in QS}

def clean(s):
    s = re.sub(r'!\[[^\]]*\]\([^)]*\)', '[图]', s)
    s = re.sub(r'\[\[[^\]]*\]\]', '[图]', s)
    return re.sub(r'\s+', ' ', s).strip()

def short(s, n=80):
    s = clean(s)
    return s if len(s) <= n else s[:n] + '…'

# ---------------------------------------------------------------- 题干索引
lines = ['# 题干索引', '',
         '一行一题，按届次和题号排序。用 `grep` 或 `tools/search.py` 检索。',
         '', '| ID | 题型 | 答案 | 章节 | 主题 | 高频 | 易错 | 题干 |',
         '|---|---|---|---|---|---|---|---|']
for q in QS:
    lines.append(f"| `{q['id']}` | {q['type']} | **{q['answer']}** | {q['chapter']} | {q['topic']} | "
                 f"{'★' if q['high_freq'] else ''} | {'⚠' if q['error_prone'] else ''} | {short(q['stem'], 110)} |")
open(os.path.join(OUT, '题干索引.md'), 'w', encoding='utf8').write('\n'.join(lines) + '\n')

# ---------------------------------------------------------------- 关键词索引
kw_map = defaultdict(set)
for q in QS:
    kws = set(q.get('keywords') or [])
    if not kws:
        kws.add(q['topic'])
    for k in kws:
        if k:
            kw_map[k].add(q['id'])
    kw_map[q['chapter']].add(q['id'])

lines = ['# 关键词索引', '',
         '按关键词反查题目 ID。先查“核心概念别名”，再查“关键词”。', '',
         '## 核心概念别名', '',
         '| 概念 | 建议搜索词 |', '|---|---|',
         '| 调查方式 | 普查、重点调查、典型调查、抽样调查、统计报表 |',
         '| 测量尺度 | 定类、定序、定距、定比、等比 |',
         '| 信度效度 | 信度、效度、再测信度、结构效度 |',
         '| 误差类型 | 抽样误差、非抽样误差、计量误差、抽样框误差、无回答误差 |',
         '| 抽样设计 | 简单随机、分层、系统、整群、多阶段、二重抽样、配额、自愿 |',
         '| 估计与检验 | 参数估计、区间估计、置信度、假设检验、两类错误 |',
         '| 变量关系 | 列联分析、相关分析、一元回归、多元回归、Logistic |',
         '| 多元方法 | 方差分析、因子分析、聚类分析、对应分析、多维标度、联合分析、综合评价 |',
         '', '## 关键词', '']
for kw in sorted(kw_map, key=lambda k: (-len(kw_map[k]), k)):
    ids = sorted(kw_map[kw], key=lambda x: (BY_ID[x]['year'], BY_ID[x]['num']))
    shown = '、'.join(f'`{i}`' for i in ids[:25])
    more = f' 等 {len(ids)} 题' if len(ids) > 25 else f'（{len(ids)} 题）'
    lines.append(f'- **{kw}**：{shown}{more}')
open(os.path.join(OUT, '关键词索引.md'), 'w', encoding='utf8').write('\n'.join(lines) + '\n')

# ---------------------------------------------------------------- 高频考点
topic = defaultdict(list)
for q in QS:
    topic[(q['chapter_id'], q['chapter'], q['topic'])].append(q)
rows = []
for key, members in topic.items():
    rows.append((sum(1 for x in members if x['high_freq']),
                 sum(1 for x in members if x['error_prone']),
                 len(members), key, members))
rows.sort(key=lambda r: (-r[0], -r[1], -r[2]))
lines = ['# 高频考点', '',
         '按“章节 / 主题”聚合。★ 为跨届重复题，⚠ 为已写解析的易错题。',
         '', '| 章节 | 主题 | 题量 | ★高频 | ⚠易错 | 代表题目 |', '|---|---|---:|---:|---:|---|']
for hf, ep, n, key, members in rows:
    rep = '、'.join(f"`{x['id']}`" for x in sorted(members, key=lambda x: (-x['high_freq'], -x['error_prone'], x['year']))[:6])
    lines.append(f'| {key[1]} | {key[2]} | {n} | {hf} | {ep} | {rep} |')
open(os.path.join(OUT, '高频考点.md'), 'w', encoding='utf8').write('\n'.join(lines) + '\n')

# ---------------------------------------------------------------- 高频题解析
def render_full(q, include_options=True):
    out = [f"### {q['id']} · {q['type']} · 答案 {q['answer']}", '',
           f"**题目**：{clean(q['stem'])}", '']
    if include_options and q['options']:
        for k in sorted(q['options']):
            out.append(f'- {k}. {clean(q["options"][k])}')
        out.append('')
    if q.get('explanation'):
        out += [f"**解析**：{q['explanation']}", '']
    if q.get('trap'):
        out += [f"**易错点**：{q['trap']}", '']
    return out

groups = defaultdict(list)
for q in QS:
    if q.get('duplicate_group'):
        groups[q['duplicate_group']].append(q)
lines = ['# 高频题解析', '',
         '同一道题在不同届次重复出现时，归为一个题组。下面给出题组、答案、解析和易错点。',
         '答案以原始题库为准；如遇教材表述差异，以本校教材为准。', '']
for gid in sorted(groups):
    members = sorted(groups[gid], key=lambda x: (x['year'], x['num']))
    canon = next((x for x in members if x.get('explanation')), members[-1])
    lines += [f"## {gid} · {canon['chapter']} / {canon['topic']}", '',
              f"**题组**：{'、'.join(f'`{x['id']}`' for x in members)}", '',
              f"**答案**：**{canon['answer']}**", '']
    lines += render_full(canon)
open(os.path.join(OUT, '高频题解析.md'), 'w', encoding='utf8').write('\n'.join(lines) + '\n')

# ---------------------------------------------------------------- 易错题解析
lines = ['# 易错题解析', '',
         '按章节整理已写解析的易错题。建议先看“易错点”，再回题干核对。', '']
by_ch = defaultdict(list)
for q in QS:
    if q['error_prone']:
        by_ch[q['chapter']].append(q)
for ch in sorted(by_ch):
    lines += [f'## {ch}', '']
    for q in sorted(by_ch[ch], key=lambda x: (x['year'], x['num'])):
        lines += render_full(q)
open(os.path.join(OUT, '易错题解析.md'), 'w', encoding='utf8').write('\n'.join(lines) + '\n')

# ---------------------------------------------------------------- 判断题陷阱
lines = ['# 判断题陷阱', '',
         '判断题最常见的陷阱：绝对化表述、因果倒置、概念偷换、把“相关”说成“因果”、把“样本”说成“总体”。',
         '', '| ID | 答案 | 题干 | 提醒 |', '|---|---|---|---|']
for q in QS:
    if q['type'] == '判断题':
        note = q.get('trap') or ('按原始题库答案' if not q.get('explanation') else '')
        lines.append(f"| `{q['id']}` | **{q['answer']}** | {short(q['stem'], 100)} | {short(note, 60)} |")
open(os.path.join(OUT, '判断题陷阱.md'), 'w', encoding='utf8').write('\n'.join(lines) + '\n')

# ---------------------------------------------------------------- 多选题清单
lines = ['# 多选题清单', '', '多选题答案可能为多个字母，注意“全选”与“漏选”都是错。', '']
by_ch = defaultdict(list)
for q in QS:
    if q['type'] == '多选题':
        by_ch[q['chapter']].append(q)
for ch in sorted(by_ch):
    lines += [f'## {ch}', '', '| ID | 答案 | 题干 |', '|---|---|---|']
    for q in sorted(by_ch[ch], key=lambda x: (x['year'], x['num'])):
        lines.append(f"| `{q['id']}` | **{q['answer']}** | {short(q['stem'], 110)} |")
    lines.append('')
open(os.path.join(OUT, '多选题清单.md'), 'w', encoding='utf8').write('\n'.join(lines) + '\n')

# ---------------------------------------------------------------- 统计
lines = ['# 题库统计', '',
         f'- 总题数：**{len(QS)}**',
         f"- 本科：**{sum(1 for q in QS if q['group']=='本科')}**",
         f"- 专科：**{sum(1 for q in QS if q['group']=='专科')}**",
         f"- 高频题（重复题组成员）：**{sum(1 for q in QS if q['high_freq'])}**",
         f"- 易错题（已写解析）：**{sum(1 for q in QS if q['error_prone'])}**",
         f"- 已写解析：**{sum(1 for q in QS if q['explanation'])}**", '',
         '## 按届次', '', '| 届次 | 组别 | 题量 |', '|---|---|---:|']
for (y, g), n in sorted(Counter((q['year'], q['group']) for q in QS).items()):
    lines.append(f'| 第{y}届 | {g} | {n} |')
lines += ['', '## 按章节', '', '| 章节 | 题量 | 高频 | 易错 |', '|---|---:|---:|---:|']
for ch in sorted(set(q['chapter'] for q in QS)):
    mem = [q for q in QS if q['chapter'] == ch]
    lines.append(f"| {ch} | {len(mem)} | {sum(1 for x in mem if x['high_freq'])} | {sum(1 for x in mem if x['error_prone'])} |")
lines += ['', '## 按题型', '', '| 题型 | 题量 |', '|---|---:|']
for t, n in Counter(q['type'] for q in QS).most_common():
    lines.append(f'| {t} | {n} |')
open(os.path.join(QUAL, '题库统计.md'), 'w', encoding='utf8').write('\n'.join(lines) + '\n')

# ---------------------------------------------------------------- 存疑
flagged = [q for q in QS if q.get('answer_flag')]
lines = ['# 答案与图片存疑', '',
         '以下题目存在表述歧义、答案口径差异或源文件图片缺失。考试时请以本校教材为准。', '']
for q in flagged:
    lines += [f"## {q['id']} · {q['type']} · 答案 {q['answer']}", '',
              f"**题目**：{clean(q['stem'])}", '',
              f"**说明**：{q['answer_flag']}", '']
open(os.path.join(QUAL, '答案存疑.md'), 'w', encoding='utf8').write('\n'.join(lines) + '\n')

print('索引生成完成：', len(QS), '题')
print('高频解析', sum(1 for q in QS if q['high_freq']), '题；易错解析', sum(1 for q in QS if q['error_prone']), '题')
