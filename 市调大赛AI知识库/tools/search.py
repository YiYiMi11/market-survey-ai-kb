#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""市调大赛题库本地检索。

用法：
    python3 tools/search.py "抽样误差是指什么"
    python3 tools/search.py "右偏分布 均值 中位数 众数" --chapter U06 --top 3
    python3 tools/search.py --id Q16-BK-056
    python3 tools/search.py "调整后的判定系数" --json
    echo "问卷 量表 绝对零点" | python3 tools/search.py

设计目标：题干换述、同义术语、只给关键词时，仍能找回原题。
不依赖第三方库，610 题规模下毫秒级返回。
"""
import argparse
import json
import math
import os
import re
import sys
import unicodedata
from collections import Counter, defaultdict

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
QUESTIONS = os.path.join(ROOT, '10_题库', 'questions.jsonl')

# 同义术语归一：把常见换述折叠到题库用词。长词优先替换。
SYNONYMS = {
    '算术平均数': '均值', '平均数': '均值', '平均值': '均值',
    '均方差': '标准差', '离散系数': '变异系数',
    '决定系数': '判定系数', '可决系数': '判定系数', 'r方': '判定系数', 'r²': '判定系数',
    '置信范围': '置信区间', '置信水平': '置信度', '可靠程度': '置信度',
    '弃真': '第一类错误', '取伪': '第二类错误',
    '名义尺度': '定类尺度', '名义变量': '定类变量', '分类尺度': '定类尺度',
    '顺序尺度': '定序尺度', '等级尺度': '定序尺度',
    '等距尺度': '定距尺度', '区间尺度': '定距尺度',
    '比率尺度': '定比尺度', '等比尺度': '定比尺度', '比例尺度': '定比尺度',
    '双重抽样': '二重抽样', '多阶抽样': '多阶段抽样', '多级抽样': '多阶段抽样',
    '小组座谈': '焦点小组', '小组座谈会': '焦点小组', '焦点小组座谈': '焦点小组',
    '深度访谈': '深层访谈', '深访': '深层访谈',
    '桌面研究': '文案调查', '案头调查': '文案调查', '文献调查': '文案调查',
    '未回答': '无回答', '无应答': '无回答',
    '测量误差': '计量误差', '记录误差': '计量误差',
    '共线性': '多重共线性', '多重共线': '多重共线性',
    'lambda': 'λ', 'wilkslambda': 'wilksλ', 'wilksλ': 'wilksλ',
    'anova': '方差分析', 'logit模型': 'logistic回归', '逻辑回归': 'logistic回归',
    '集群分析': '聚类分析', '分层聚类': '系统聚类',
    '多维尺度': '多维标度', '多维尺度分析': '多维标度分析',
    '联合分析': '联合分析', 'conjoint': '联合分析',
    '综合评价': '综合评价', '多指标评价': '综合评价',
    '无量纲化': '无量纲化', '标准化处理': '无量纲化',
    '全面调查': '普查', '一次性全面调查': '普查',
    '典型调查': '典型调查', '重点调查': '重点调查',
    '抽样框': '抽样框', '名录框': '抽样框', '名单框': '抽样框',
    '总体参数': '参数', '样本统计量': '统计量',
    '显著性检验': '假设检验', '假设检验': '假设检验',
    '独立样本t检验': 't检验', '配对样本t检验': 't检验',
    '回归直线': '回归方程', '拟合直线': '回归方程',
    '误差平方和': '残差平方和', '剩余平方和': '残差平方和',
    '可解释变异': '判定系数', '解释方差': '判定系数',
    '调查表': '问卷', '访问问卷': '问卷',
    '态度量表': '量表', '测量量表': '量表',
    '信度': '信度', '可靠性': '信度', '效度': '效度', '有效性': '效度',
    '系统抽样': '系统抽样', '等距抽样': '系统抽样', '机械抽样': '系统抽样',
    '整群抽样': '整群抽样', '聚类抽样': '整群抽样', '分群抽样': '整群抽样',
    '分层抽样': '分层抽样', '类型抽样': '分层抽样', '分类抽样': '分层抽样',
    '简单随机抽样': '简单随机抽样', '纯随机抽样': '简单随机抽样',
    '任意抽样': '便利抽样', '方便抽样': '便利抽样', '偶遇抽样': '便利抽样',
    '滚雪球抽样': '滚雪球抽样', '雪球抽样': '滚雪球抽样',
    '判断抽样': '判断抽样', '目的抽样': '判断抽样', '立意抽样': '判断抽样',
    '配额抽样': '配额抽样', '定额抽样': '配额抽样',
}

# 单字术语不参与归一，避免误伤。
PUNCT = re.compile(r'[\s，。、；：（）()《》“”"\'‘’？?！!·\-—_.<>/\\|,;:%]+')

# 口语化提问中的填充词。先删长词，再删短词；不删“为什么”“哪些”等带语义的词。
FILLERS = (('的完整定义', '定义'), ('的完整含义', '含义'), ('的定义', '定义'),
           ('的含义', '含义'), ('的概念', '概念'),
           ('是什么', ''), ('什么是', ''), ('什么', ''), ('完整', ''),
           ('请问', ''), ('一下', ''))


def normalize(text):
    text = unicodedata.normalize('NFKC', text or '').lower()
    text = PUNCT.sub('', text)
    for src, dst in FILLERS:
        if src in text:
            text = text.replace(src, dst)
    for src in sorted(SYNONYMS, key=len, reverse=True):
        if src in text:
            text = text.replace(src, SYNONYMS[src])
    return text


def grams(text, with_unigram=True):
    """字符 n-gram + 英文/数字整词。返回 {gram: 权重}。"""
    s = normalize(text)
    out = {}
    for m in re.findall(r'[a-z0-9]+', s):
        out['w:' + m] = max(out.get('w:' + m, 0), 1.6)
    for run in re.findall(r'[\u0370-\u03ff]+', s):
        for ch in run:
            out[ch] = max(out.get(ch, 0), 1.5)
    for run in re.findall(r'[\u4e00-\u9fff]+', s):
        n = len(run)
        if with_unigram:
            for ch in run:
                out[ch] = max(out.get(ch, 0), 0.25)
        for size, weight in ((2, 1.0), (3, 1.5)):
            for i in range(n - size + 1):
                g = run[i:i + size]
                out[g] = max(out.get(g, 0), weight)
    return out


def load():
    qs = [json.loads(line) for line in open(QUESTIONS, encoding='utf8') if line.strip()]
    qs.sort(key=lambda q: (q['year'], 0 if q['group'] == '本科' else 1, q['num']))
    return qs


class Searcher:
    def __init__(self, questions):
        self.qs = questions
        self.N = len(questions)
        self.docs = []
        self.df = Counter()
        for q in questions:
            d = Counter()
            for g, w in grams(q['stem']).items():
                d[g] += w
            for g, w in grams(' '.join(q['options'].values())).items():
                d[g] += 0.35 * w
            extra = ' '.join((q.get('keywords') or []) + (q.get('tags') or [])) + ' ' + q.get('topic', '') + ' ' + q.get('chapter', '')
            for g, w in grams(extra).items():
                d[g] += 0.45 * w
            self.docs.append(d)
            for g in d:
                self.df[g] += 1
        self.idf = {g: math.log((self.N + 1) / (df + 0.5)) + 0.5 for g, df in self.df.items()}
        self.default_idf = math.log(self.N + 1) + 0.5
        self.norms = []
        for d in self.docs:
            self.norms.append(math.sqrt(sum((w * self.idf.get(g, self.default_idf)) ** 2 for g, w in d.items())) or 1.0)

    def search(self, query, chapter=None, group=None, year=None, qtype=None, top=5):
        qn = normalize(query)
        qg = grams(query)
        if not qn or not qg:
            return []
        q_weight = {g: w * self.idf.get(g, self.default_idf) for g, w in qg.items()}
        q_norm = math.sqrt(sum(v * v for v in q_weight.values())) or 1.0
        q_total = sum(q_weight.values()) or 1.0
        results = []
        for i, q in enumerate(self.qs):
            if chapter and not (q['chapter_id'].lower() == chapter.lower() or chapter.lower() in q['chapter'].lower()):
                continue
            if group and q['group'] != group:
                continue
            if year and q['year'] != int(year):
                continue
            if qtype and qtype not in q['type']:
                continue
            d = self.docs[i]
            dot = 0.0
            matched = 0.0
            for g, wq in q_weight.items():
                wd = d.get(g)
                if wd:
                    dot += wq * wd * self.idf.get(g, self.default_idf)
                    matched += wq
            cos = dot / (q_norm * self.norms[i])
            coverage = matched / q_total
            s = 0.62 * cos + 0.38 * coverage
            stem_n = normalize(q['stem'])
            if qn == stem_n:
                s = 1.0
            elif len(qn) >= 6 and qn in stem_n:
                s = max(s, 0.94)
            elif len(stem_n) >= 6 and stem_n in qn:
                s = max(s, 0.90)
            # 命中题干开头的疑问核心，轻微加权
            if len(qn) >= 4 and stem_n.startswith(qn[:min(len(qn), 12)]):
                s += 0.02
            results.append((min(s, 1.0), q))
        results.sort(key=lambda x: (-x[0], -int(x[1]['high_freq']), -x[1]['year'], x[1]['num']))
        return results[:top]


def fmt(q, score, full=True, brief=False):
    lines = [f"[{score:.3f}] {q['id']} · {q['type']} · 答案 {q['answer']}"]
    if brief:
        lines.append(q['stem'])
        return '\n'.join(lines)
    lines.append(f"章节：{q['chapter']} / {q['topic']}" + ('  ★高频' if q['high_freq'] else '') + ('  ⚠易错' if q['error_prone'] else ''))
    lines.append(f"题目：{q['stem']}")
    if full and q['options']:
        for k in sorted(q['options']):
            mark = ' ✓' if q['answer'] and k in q['answer'] else ''
            lines.append(f"  {k}. {q['options'][k]}{mark}")
    if q.get('explanation'):
        lines.append(f"解析：{q['explanation']}")
    if q.get('trap'):
        lines.append(f"易错点：{q['trap']}")
    if q.get('answer_flag'):
        lines.append(f"待核对：{q['answer_flag']}")
    lines.append(f"来源：{q['source']}")
    return '\n'.join(lines)


def main():
    ap = argparse.ArgumentParser(description='市调大赛题库检索')
    ap.add_argument('query', nargs='*', help='题干、关键词或自然语言描述')
    ap.add_argument('--id', help='按题目 ID 精确查询')
    ap.add_argument('--chapter', help='章节过滤，如 U07 / Z06 / 抽样')
    ap.add_argument('--group', choices=['本科', '专科'])
    ap.add_argument('--year', type=int)
    ap.add_argument('--type', dest='qtype', help='题型过滤：单选/多选/判断')
    ap.add_argument('--top', type=int, default=5)
    ap.add_argument('--brief', action='store_true', help='只输出 ID、答案、题干')
    ap.add_argument('--json', action='store_true', help='输出 JSON')
    ap.add_argument('--list-chapters', action='store_true')
    args = ap.parse_args()

    qs = load()
    by_id = {q['id']: q for q in qs}

    if args.list_chapters:
        seen = {}
        for q in qs:
            seen.setdefault(q['chapter_id'], [q['chapter'], 0])
            seen[q['chapter_id']][1] += 1
        for cid in sorted(seen):
            print(f"{cid}\t{seen[cid][0]}\t{seen[cid][1]} 题")
        return

    if args.id:
        q = by_id.get(args.id.upper())
        if not q:
            print(f'未找到题目：{args.id}', file=sys.stderr)
            sys.exit(1)
        print(fmt(q, 1.0, brief=args.brief))
        return

    query = ' '.join(args.query).strip()
    if not query and not sys.stdin.isatty():
        query = sys.stdin.read().strip()
    if not query:
        ap.print_help()
        return

    searcher = Searcher(qs)
    results = searcher.search(query, chapter=args.chapter, group=args.group,
                              year=args.year, qtype=args.qtype, top=args.top)
    if args.json:
        out = []
        for score, q in results:
            item = {k: q[k] for k in ('id', 'type', 'chapter', 'chapter_id', 'topic', 'stem',
                                      'options', 'answer', 'high_freq', 'error_prone',
                                      'explanation', 'trap', 'answer_flag', 'tags', 'duplicate_group',
                                      'duplicate_ids', 'source')}
            item['score'] = round(score, 4)
            out.append(item)
        print(json.dumps(out, ensure_ascii=False, indent=2))
        return
    if not results:
        print('未检索到匹配题目。建议：减少限定词、改用核心术语，或先看 30_速查/关键词索引.md。')
        return
    print(f'检索：{query}  （返回 {len(results)} 条）\n')
    for n, (score, q) in enumerate(results, 1):
        print(f'--- 结果 {n} ' + '-' * 46)
        print(fmt(q, score, brief=args.brief))
        print()


if __name__ == '__main__':
    main()
