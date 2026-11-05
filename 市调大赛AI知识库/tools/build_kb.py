# -*- coding: utf-8 -*-
"""从原始 docx/xlsx 构建市调大赛 AI 知识库的结构化题库。

用法：python3 tools/build_kb.py
"""
import os, re, json, csv, html, zipfile, shutil
from collections import Counter, defaultdict
from xml.etree import ElementTree as ET
import openpyxl
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from explanations import HIGH_FREQ, ERROR_PRONE, FLAGGED
from annotations import TAGS

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_ROOT = os.path.dirname(ROOT)
QDIR = os.path.join(ROOT, '10_题库')
ASSET_DIR = os.path.join(QDIR, 'assets')
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

# ---------------------------------------------------------------- 原始文件
DOCX_JOBS = [
    (13, '本科', '第12-16届网考相关资料/第13届市调大赛本科组模拟题.docx'),
    (14, '本科', '第12-16届网考相关资料/第14届市调大赛本科组模拟题(9).docx'),
    (15, '本科', '第12-16届网考相关资料/第15届市调大赛本科组模拟题1.docx'),
    (16, '本科', '第12-16届网考相关资料/第16届市调大赛本科组模拟题一(1).docx'),
    (17, '本科', '2026年第17届市调大赛网考说明及网考模拟题/17届市调大赛本科组模拟题.docx'),
    (17, '专科', '2026年第17届市调大赛网考说明及网考模拟题/17届市调大赛专科组模拟题.docx'),
]
XLSX_JOBS = [
    (12, '本科', '第12-16届网考相关资料/第12届本科模拟题（10.19）.xlsx'),
]

def para_content(p):
    """按文档顺序收集段落文字、公式文字和内嵌图片占位符。"""
    out = []
    for e in p.iter():
        tag = e.tag.split('}')[-1]
        if tag == 't':
            out.append(e.text or '')
        elif tag == 'blip':
            rid = e.get(R + 'embed')
            if rid:
                out.append(f'[[IMG:{rid}]]')
        elif tag == 'tab':
            out.append('\t')
        elif tag == 'br':
            out.append('\n')
    return ''.join(out)

def parse_docx(path, year, group):
    path = os.path.join(SRC_ROOT, path)
    z = zipfile.ZipFile(path)
    root = ET.fromstring(z.read('word/document.xml'))
    rels = {}
    try:
        rr = ET.fromstring(z.read('word/_rels/document.xml.rels'))
        for rel in rr:
            rels[rel.get('Id')] = rel.get('Target')
    except KeyError:
        pass
    qid_prefix = f'Q{year}-{"BK" if group == "本科" else "ZK"}'
    paras = [para_content(p).replace('\n', '').strip() for p in root.iter(W + 'p')]
    paras = [p for p in paras if p]
    questions, cur, mode = [], None, None
    for ln in paras:
        m = re.match(r'^(\d+)\s*[-－]\s*题目\s*[:：]\s*(.*)$', ln)
        if m:
            if cur:
                questions.append(cur)
            cur = {'num': int(m.group(1)), 'body': m.group(2), 'opts': []}
            mode = 'body'
            continue
        m = re.match(r'^(\d+)\s*[-－]\s*选项\s*[:：]?\s*(.*)$', ln)
        if m and cur:
            mode = 'opts'
            if m.group(2).strip():
                cur['opts'].append(m.group(2).strip())
            continue
        if cur is None:
            continue
        if mode == 'opts':
            if re.match(r'^[（(]\s*[A-Ha-h]\s*[）)]', ln):
                cur['opts'].append(ln)
            elif cur['opts']:
                cur['opts'][-1] += ln
            else:
                cur['body'] += ln
        else:
            cur['body'] += ln
    if cur:
        questions.append(cur)

    out = []
    for q in questions:
        m = re.search(r'\[题型[:：](.*?)\]\[答案[:：](.*?)\]', q['body'])
        if not m:
            continue
        typ, ans = m.group(1).strip(), m.group(2).strip()
        stem = q['body'][:m.start()].strip()
        qid = f'{qid_prefix}-{q["num"]:03d}'
        # 保存该题用到的图片；同一 rId 只存一份
        img_map = {}
        for rid in re.findall(r'\[\[IMG:([^\]]+)\]\]', stem + '\n' + '\n'.join(q['opts'])):
            if rid in img_map:
                continue
            target = rels.get(rid, '')
            target = target.lstrip('/')
            if not target.startswith('word/'):
                target = 'word/' + target
            try:
                data = z.read(target)
            except KeyError:
                continue
            ext = os.path.splitext(target)[1].lower() or '.png'
            fname = f'{qid}-{rid}{ext}'
            os.makedirs(ASSET_DIR, exist_ok=True)
            with open(os.path.join(ASSET_DIR, fname), 'wb') as f:
                f.write(data)
            img_map[rid] = f'10_题库/assets/{fname}'
        def repl(s):
            return re.sub(r'\[\[IMG:([^\]]+)\]\]',
                          lambda mm: f'[[IMG:{img_map.get(mm.group(1), "")}]]', s)
        stem = repl(stem)
        opts = []
        for o in q['opts']:
            mo = re.match(r'^[（(]\s*([A-Ha-h])\s*[）)]\s*(.*)$', o)
            if mo:
                opts.append((mo.group(1).upper(), repl(mo.group(2).strip())))
        out.append({
            'id': qid, 'num': q['num'], 'year': year, 'group': group,
            'type': typ, 'stem': stem,
            'options': dict(opts), 'answer': ans,
            'images': [v for k, v in sorted(img_map.items())],
            'source': os.path.relpath(path, SRC_ROOT),
        })
    return out

def parse_xlsx(path, year, group):
    path = os.path.join(SRC_ROOT, path)
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    header = [str(c).strip() if c is not None else '' for c in rows[0]]
    idx = {h: i for i, h in enumerate(header)}
    out = []
    for r in rows[1:]:
        if not r or r[idx['题目']] is None:
            continue
        def clean(v):
            s = html.unescape(str(v))
            s = re.sub(r'<img[^>]*src="([^"]+)"[^>]*>',
                       r'[[原题图片:\1（源文件未嵌入）]]', s)
            s = re.sub(r'<[^>]+>', '', s)
            return re.sub(r'\s+', ' ', s).strip()
        opts = {}
        for j in range(1, 6):
            key = f'选项{j}'
            if key in idx and r[idx[key]] is not None and clean(r[idx[key]]):
                opts[chr(64 + j)] = clean(r[idx[key]])
        out.append({
            'id': f'Q{year}-{"BK" if group == "本科" else "ZK"}-{int(float(r[idx["序号"]])):03d}',
            'num': int(float(r[idx['序号']])), 'year': year, 'group': group,
            'type': str(r[idx['类型']]).strip(), 'stem': clean(r[idx['题目']]),
            'options': opts, 'answer': str(r[idx['答案']]).strip(),
            'images': [], 'source': os.path.relpath(path, SRC_ROOT),
        })
    return out

# ---------------------------------------------------------------- 章节规则
CHAPTERS = {
    'U01': ('本科', '01 市场调查概述', '市场调查基本知识、基本程序、应用范围'),
    'U02': ('本科', '02 数据收集', '定性与定量数据收集方法、组织与实施'),
    'U03': ('本科', '03 调查设计', '调查策划与调查方案设计（不含抽样设计和问卷设计）'),
    'U04': ('本科', '04 问卷与量表', '问卷及量表的设计与应用'),
    'U05': ('本科', '05 数据处理与报告', '数据处理汇总、调查报告撰写与演示'),
    'U06': ('本科', '06 单变量分析', '规模结构比较、集中度、差异性、参数估计、显著性检验'),
    'U07': ('本科', '07 抽样估计', '抽样设计、抽样估计'),
    'U08': ('本科', '08 双变量关联分析', '列联分析、相关分析、回归分析'),
    'U09': ('本科', '09 多变量分析（一）', '多变量方差分析、多元回归、Logistic 模型'),
    'U10': ('本科', '10 多变量分析（二）', '因子、聚类、对应、多维标度、联合分析、综合评价'),
    'Z01': ('专科', 'Z01 市场调查概述', '基本知识与基本程序、应用范围'),
    'Z02': ('专科', 'Z02 调查方案设计', '调查策划与方案设计'),
    'Z03': ('专科', 'Z03 问卷及量表设计', '问卷及量表的设计与应用'),
    'Z04': ('专科', 'Z04 资料采集方法', '定性、定量、一手和二手数据收集'),
    'Z05': ('专科', 'Z05 调查实施与报告', '组织实施、数据整理、报告撰写'),
    'Z06': ('专科', 'Z06 描述统计基础', '统计学基础、规模结构比较、集中度、差异性'),
    'Z07': ('专科', 'Z07 抽样调查与估计', '抽样设计与抽样估计'),
    'Z08': ('专科', 'Z08 数据分析方法', '时间序列、双变量相关与回归'),
    'U00': ('本科', '00 综合/待分类', '暂未归入具体章节'),
    'Z00': ('专科', 'Z00 综合/待分类', '暂未归入具体章节'),
}

# 关键词 -> (章节, 主题, 权重)
RULES = [
    # U01
    ('调查基本程序', 'U01', '调查程序', 5), ('基本程序', 'U01', '调查程序', 5),
    ('调查流程', 'U01', '调查程序', 5), ('调查分类', 'U01', '调查分类', 5),
    ('分类标志', 'U01', '调查分类', 5), ('市场调查特点', 'U01', '调查特点', 5),
    ('市场调查原则', 'U01', '调查原则', 5), ('市场调查作用', 'U01', '调查作用', 5),
    ('市场营销', 'U01', '调查作用', 4), ('广告研究', 'U01', '应用范围', 5),
    ('顾客满意', 'U01', '应用范围', 4), ('市场环境', 'U01', '应用范围', 4),
    ('竞争情报', 'U01', '应用范围', 4), ('探索性调查', 'U01', '研究类型', 4),
    ('描述性调查', 'U01', '研究类型', 4), ('预测性调查', 'U01', '研究类型', 4),
    ('因果性调查', 'U01', '研究类型', 4), ('市场调查', 'U01', '综合', 2),
    # U02
    ('焦点小组', 'U02', '焦点小组', 6), ('小组座谈', 'U02', '焦点小组', 6),
    ('深层访谈', 'U02', '深层访谈', 6), ('深度访谈', 'U02', '深层访谈', 6),
    ('神秘顾客', 'U02', '观察法', 6), ('观察法', 'U02', '观察法', 5),
    ('实验法', 'U02', '实验法', 5), ('实验设计', 'U02', '实验法', 5),
    ('文案调查', 'U02', '文案调查', 6), ('二手资料', 'U02', '二手数据', 5),
    ('二手数据', 'U02', '二手数据', 5), ('文献', 'U02', '二手数据', 4),
    ('年鉴', 'U02', '二手数据', 4), ('德尔菲', 'U02', '德尔菲法', 6),
    ('头脑风暴', 'U02', '头脑风暴', 5), ('投影技法', 'U02', '投影技法', 5),
    ('深度访谈', 'U02', '深层访谈', 5), ('访谈', 'U02', '访谈调查', 4),
    ('面访', 'U02', '面访调查', 4), ('电话调查', 'U02', '电话调查', 5),
    ('邮寄调查', 'U02', '邮寄调查', 5), ('网络调查', 'U02', '网络调查', 5),
    ('网上调查', 'U02', '网络调查', 5), ('在线调查', 'U02', '网络调查', 5),
    ('观察', 'U02', '观察法', 3), ('定性调查', 'U02', '定性调查', 5),
    ('定量调查', 'U02', '定量调查', 5), ('调查员', 'U02', '调查员管理', 5),
    ('访问员', 'U02', '调查员管理', 5), ('追问', 'U02', '访问技巧', 5),
    ('提问方式', 'U02', '访问技巧', 4), ('固定样本', 'U02', '固定样本', 5),
    ('计算机辅助', 'U02', '计算机辅助调查', 5), ('遥感', 'U02', '遥感调查', 5),
    ('统计报表', 'U02', '统计报表', 5), ('普查', 'U02', '普查', 4),
    ('数据收集', 'U02', '数据收集', 4), ('资料采集', 'U02', '数据收集', 4),
    ('信息内部来源', 'U02', '数据来源', 5), ('信息外部来源', 'U02', '数据来源', 5),
    ('询问法', 'U02', '询问法', 4), ('访谈法', 'U02', '访谈调查', 4),
    # U03
    ('调查方案', 'U03', '调查方案', 6), ('方案设计', 'U03', '调查方案', 6),
    ('调查计划', 'U03', '调查方案', 5), ('调查时间', 'U03', '调查时间', 5),
    ('调查预算', 'U03', '调查预算', 5), ('操作定义', 'U03', '操作定义', 5),
    ('调查问题框架', 'U03', '调查设计', 5), ('调查目标', 'U03', '调查目标', 4),
    ('调查对象', 'U03', '调查对象', 3), ('调查范围', 'U03', '调查对象', 4),
    ('预调查', 'U03', '预调查', 5), ('可行性', 'U03', '可行性研究', 4),
    ('调查设计', 'U03', '调查设计', 4), ('调查工作', 'U03', '调查时间', 3),
    # U04
    ('问卷', 'U04', '问卷设计', 4), ('量表', 'U04', '量表', 5),
    ('问句', 'U04', '问题设计', 5), ('答案设计', 'U04', '答案设计', 5),
    ('问题排列', 'U04', '问题排列', 5), ('问卷编码', 'U04', '问卷编码', 5),
    ('测量尺度', 'U04', '测量尺度', 5), ('定类', 'U04', '测量尺度', 5),
    ('定序', 'U04', '测量尺度', 5), ('定距', 'U04', '测量尺度', 5),
    ('定比', 'U04', '测量尺度', 5), ('信度', 'U04', '信度', 5),
    ('效度', 'U04', '效度', 5), ('语义差别', 'U04', '语义差别量表', 5),
    ('李克特', 'U04', '李克特量表', 5), ('评比量表', 'U04', '评比量表', 5),
    ('配对比较', 'U04', '配对比较量表', 5), ('排序', 'U04', '排序量表', 3),
    ('开放式', 'U04', '问题类型', 4), ('封闭式', 'U04', '问题类型', 4),
    ('甄别', 'U04', '问卷结构', 4), ('说明词', 'U04', '问卷结构', 4),
    ('问题表述', 'U04', '问题设计', 5), ('测量', 'U04', '测量', 2),
    # U05
    ('数据整理', 'U05', '数据整理', 5), ('数据处理', 'U05', '数据整理', 5),
    ('数据审核', 'U05', '数据审核', 5), ('数据净化', 'U05', '数据净化', 5),
    ('净化', 'U05', '数据净化', 4), ('录入', 'U05', '数据录入', 4),
    ('插补', 'U05', '缺失值插补', 5), ('缺失', 'U05', '缺失值插补', 4),
    ('频数分布', 'U05', '频数分布', 5), ('组距', 'U05', '统计分组', 5),
    ('统计分组', 'U05', '统计分组', 5), ('统计表', 'U05', '统计表', 4),
    ('统计图', 'U05', '统计图', 4), ('直方图', 'U05', '统计图', 5),
    ('条形图', 'U05', '统计图', 5), ('茎叶图', 'U05', '统计图', 5),
    ('箱线图', 'U05', '统计图', 5), ('调查报告', 'U05', '调查报告', 6),
    ('书面报告', 'U05', '调查报告', 6), ('演示报告', 'U05', '报告演示', 5),
    ('附录', 'U05', '调查报告', 4), ('报告撰写', 'U05', '调查报告', 5),
    ('问卷的接受', 'U05', '问卷检查', 6), ('不能接收的问卷', 'U05', '问卷检查', 6),
    # U06
    ('集中趋势', 'U06', '集中趋势', 5), ('平均数', 'U06', '集中趋势', 4),
    ('均值', 'U06', '集中趋势', 4), ('中位数', 'U06', '集中趋势', 5),
    ('众数', 'U06', '集中趋势', 5), ('离散程度', 'U06', '离散程度', 5),
    ('标准差', 'U06', '离散程度', 5), ('方差', 'U06', '离散程度', 4),
    ('变异系数', 'U06', '离散程度', 5), ('极差', 'U06', '离散程度', 4),
    ('分位数', 'U06', '分位数', 4), ('偏态', 'U06', '偏态与峰度', 5),
    ('峰度', 'U06', '偏态与峰度', 5), ('参数估计', 'U06', '参数估计', 5),
    ('点估计', 'U06', '参数估计', 5), ('区间估计', 'U06', '区间估计', 5),
    ('置信区间', 'U06', '区间估计', 5), ('置信度', 'U06', '区间估计', 4),
    ('假设检验', 'U06', '假设检验', 5), ('显著性', 'U06', '假设检验', 4),
    ('第一类错误', 'U06', '两类错误', 5), ('第二类错误', 'U06', '两类错误', 5),
    ('t检验', 'U06', 't检验', 5), ('检验统计量', 'U06', '假设检验', 5),
    ('正态分布', 'U06', '正态分布', 4), ('标准分数', 'U06', '标准分数', 5),
    ('经验法则', 'U06', '正态分布', 4), ('季节指数', 'U06', '时间序列', 5),
    ('时间序列', 'U06', '时间序列', 5), ('增长速度', 'U06', '时间序列', 4),
    ('指数', 'U06', '统计指数', 3), ('描述统计', 'U06', '描述统计', 4),
    ('参数', 'U06', '统计基础', 2), ('统计量', 'U06', '统计基础', 3),
    ('概率分布', 'U06', '概率分布', 3), ('随机变量', 'U06', '概率分布', 3),
    # U07
    ('抽样框', 'U07', '抽样框', 6), ('概率抽样', 'U07', '概率抽样', 6),
    ('非概率抽样', 'U07', '非概率抽样', 6), ('简单随机', 'U07', '简单随机抽样', 6),
    ('分层抽样', 'U07', '分层抽样', 6), ('系统抽样', 'U07', '系统抽样', 6),
    ('整群抽样', 'U07', '整群抽样', 6), ('多阶段', 'U07', '多阶段抽样', 6),
    ('多阶抽样', 'U07', '多阶段抽样', 6), ('二阶抽样', 'U07', '多阶段抽样', 5),
    ('二重抽样', 'U07', '二重抽样', 6), ('配额抽样', 'U07', '配额抽样', 6),
    ('滚雪球', 'U07', '滚雪球抽样', 6), ('便利抽样', 'U07', '便利抽样', 5),
    ('判断抽样', 'U07', '判断抽样', 5), ('自愿抽样', 'U07', '自愿抽样', 5),
    ('设计效应', 'U07', '设计效应', 6), ('样本量', 'U07', '样本量', 6),
    ('样本容量', 'U07', '样本量', 6), ('抽样误差', 'U07', '抽样误差', 6),
    ('非抽样误差', 'U07', '非抽样误差', 6), ('无回答', 'U07', '无回答', 5),
    ('权数', 'U07', '权数', 5), ('标准误', 'U07', '标准误', 4),
    ('抽样分布', 'U07', '抽样分布', 5), ('估计量', 'U07', '估计量', 4),
    ('估计值', 'U07', '估计量', 4), ('比率估计', 'U07', '比率估计', 5),
    ('回归估计', 'U07', '回归估计', 5), ('辅助变量', 'U07', '辅助变量', 6),
    ('抽样设计', 'U07', '抽样设计', 5), ('抽样调查', 'U07', '抽样调查', 4),
    ('抽样方法', 'U07', '抽样方法', 5), ('抽样', 'U07', '抽样', 2),
    ('总体', 'U07', '抽样基础', 2), ('样本', 'U07', '抽样基础', 2),
    # U08
    ('列联', 'U08', '列联分析', 6), ('卡方', 'U08', '列联分析', 5),
    ('独立性检验', 'U08', '列联分析', 5), ('一致性检验', 'U08', '列联分析', 5),
    ('拟合优度', 'U08', '列联分析', 5), ('相关分析', 'U08', '相关分析', 6),
    ('相关系数', 'U08', '相关分析', 6), ('相关关系', 'U08', '相关分析', 5),
    ('散点图', 'U08', '相关分析', 5), ('一元线性回归', 'U08', '一元回归', 6),
    ('回归方程', 'U08', '回归分析', 5), ('回归系数', 'U08', '回归分析', 5),
    ('最小二乘', 'U08', '回归分析', 5), ('判定系数', 'U08', '回归分析', 5),
    ('残差', 'U08', '回归诊断', 5), ('DW', 'U08', '回归诊断', 5),
    ('回归', 'U08', '回归分析', 3), ('双变量', 'U08', '双变量分析', 5),
    # U09
    ('方差分析', 'U09', '方差分析', 6), ('MSA', 'U09', '方差分析', 6),
    ('MSE', 'U09', '方差分析', 6), ('组间', 'U09', '方差分析', 5),
    ('组内', 'U09', '方差分析', 5), ('多变量方差', 'U09', '多变量方差分析', 6),
    ('MANOVA', 'U09', '多变量方差分析', 6), ('Wilks', 'U09', '多变量方差分析', 6),
    ('多元回归', 'U09', '多元回归', 6), ('多重判定系数', 'U09', '多元回归', 6),
    ('调整的判定系数', 'U09', '多元回归', 6), ('多重共线性', 'U09', '多元回归', 6),
    ('VIF', 'U09', '多元回归', 6), ('Logistic', 'U09', 'Logistic回归', 6),
    ('优势比', 'U09', 'Logistic回归', 6), ('判别分析', 'U09', '判别分析', 6),
    ('Fisher判别', 'U09', '判别分析', 6), ('结构方程', 'U09', '结构方程', 5),
    ('非参数', 'U09', '非参数检验', 4), ('Wilcoxon', 'U09', '非参数检验', 6),
    ('Mood', 'U09', '非参数检验', 6),
    # U10
    ('因子分析', 'U10', '因子分析', 6), ('公共因子', 'U10', '因子分析', 6),
    ('因子载荷', 'U10', '因子分析', 6), ('因子得分', 'U10', '因子分析', 6),
    ('聚类', 'U10', '聚类分析', 6), ('K均值', 'U10', '聚类分析', 6),
    ('系统聚类', 'U10', '聚类分析', 6), ('快速聚类', 'U10', '聚类分析', 6),
    ('对应分析', 'U10', '对应分析', 6), ('多维标度', 'U10', '多维标度', 6),
    ('联合分析', 'U10', '联合分析', 6), ('综合评价', 'U10', '综合评价', 6),
    ('评价指标', 'U10', '综合评价', 5), ('权重', 'U10', '综合评价', 4),
    ('无量纲', 'U10', '综合评价', 5), ('主成分', 'U10', '主成分分析', 5),
    ('正交设计', 'U10', '正交设计', 5),

    # 补充规则（二轮）
    ('市场调研', 'U01', '综合', 3), ('市场研究', 'U01', '应用范围', 3),
    ('调查提供者', 'U01', '调查行业', 5), ('调查使用者', 'U01', '调查行业', 5),
    ('调查机构', 'U01', '调查行业', 4), ('调查部门', 'U01', '调查行业', 4),
    ('调查组织模式', 'U01', '调查行业', 4), ('调查内容', 'U01', '调查内容', 3),
    ('市场信息', 'U01', '调查作用', 3), ('调查的作用', 'U01', '调查作用', 5),
    ('调查作用', 'U01', '调查作用', 4), ('调查特点', 'U01', '调查特点', 4),
    ('调查原则', 'U01', '调查原则', 4), ('调查目的', 'U01', '调查目标', 3),
    ('随机拨号', 'U02', '电话调查', 5), ('行风评议', 'U02', '电话调查', 5),
    ('联网直报', 'U02', '统计报表', 5), ('非随机化实验', 'U02', '实验法', 6),
    ('实验单位', 'U02', '实验法', 5), ('交互效应', 'U02', '实验效度', 6),
    ('内部有效性', 'U02', '实验效度', 6), ('外部有效性', 'U02', '实验效度', 6),
    ('访问', 'U02', '访问技巧', 3), ('询问', 'U02', '询问法', 3),
    ('调查策划', 'U03', '调查方案', 5), ('计划书', 'U03', '调查方案', 4),
    ('调查组织计划', 'U03', '调查方案', 4),
    ('问卷结构', 'U04', '问卷结构', 5), ('问题类型', 'U04', '问题类型', 4),
    ('态度', 'U04', '量表', 3), ('等级法', 'U04', '量表', 5),
    ('多项选择', 'U04', '问题类型', 5), ('问卷测试', 'U04', '问卷设计', 4),
    ('数据分类', 'U05', '数据整理', 6), ('预处理', 'U05', '数据整理', 6),
    ('数据检查', 'U05', '数据审核', 5), ('逻辑分析', 'U05', '数据净化', 5),
    ('数据分布', 'U05', '统计图', 3), ('统计整理', 'U05', '数据整理', 5),
    ('MLE', 'U06', '参数估计', 6), ('最大似然', 'U06', '参数估计', 6),
    ('无偏', 'U06', '估计量性质', 5), ('有效性', 'U06', '估计量性质', 3),
    ('充分性', 'U06', '估计量性质', 5), ('随机抽样', 'U07', '概率抽样', 5),
    ('重点调查', 'U07', '重点调查', 6), ('典型调查', 'U07', '典型调查', 6),
    ('全面调查', 'U07', '普查', 5), ('非全面调查', 'U07', '非全面调查', 5),
    ('抽样精度', 'U07', '抽样误差', 5), ('调查误差', 'U07', '抽样误差', 4),
    ('估计精度', 'U07', '抽样误差', 4), ('SSE', 'U08', '回归分析', 6),
    ('估计标准误差', 'U08', '回归分析', 6), ('实验数据', 'U09', '方差分析', 6),
    ('协方差分析', 'U09', '方差分析', 5), ('多元统计', 'U09', '多变量概述', 5),
]
# 专科规则：复用本科关键词，但章节重映射
ZMAP = {
    'U01': 'Z01', 'U03': 'Z02', 'U04': 'Z03', 'U02': 'Z04', 'U05': 'Z05',
    'U06': 'Z06', 'U07': 'Z07', 'U08': 'Z08', 'U09': 'Z08', 'U10': 'Z08',
}

def norm_text(s):
    return re.sub(r'\s+', '', s or '')

def classify(q):
    text = norm_text(q['stem'])
    scores = defaultdict(int)
    topics = defaultdict(Counter)
    kws = set()
    for kw, ch, topic, w in RULES:
        if kw and kw in text:
            scores[ch] += w
            topics[ch][topic] += w
            if w >= 4:
                kws.add(kw)
    if not scores:
        ch = 'U00' if q['group'] == '本科' else 'Z00'
        topic = '综合'
        conf = 'low'
    else:
        ch = max(scores, key=lambda k: (scores[k], -int(k[1:])))
        topic = topics[ch].most_common(1)[0][0]
        # 最高分与次高分差距过小则标记为待复核
        ranked = sorted(scores.values(), reverse=True)
        conf = 'high' if len(ranked) == 1 or ranked[0] - ranked[1] >= 2 else 'medium'
    if q['group'] == '专科':
        ch = ZMAP.get(ch, 'Z01')
    return ch, topic, sorted(kws), conf

# ---------------------------------------------------------------- 重复题聚类
def ngram(s):
    s = re.sub(r'[\s，。、；：（）()《》“”"\'‘’？?！!·\-—_.<>/]', '', s or '')
    return set(s[i:i + 2] for i in range(len(s) - 1))

def jaccard(a, b):
    return len(a & b) / len(a | b) if a and b else 0.0

def cluster_questions(questions):
    qs = [q for q in questions if q['group'] == '本科']
    for q in qs:
        q['_bg'] = ngram(q['stem'])
        q['_obg'] = ngram(''.join(q['options'].values()))
    par = list(range(len(qs)))
    def find(x):
        while par[x] != x:
            par[x] = par[par[x]]
            x = par[x]
        return x
    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            par[rb] = ra
    same_stem_variants = []
    for i in range(len(qs)):
        for j in range(i + 1, len(qs)):
            a, b = qs[i], qs[j]
            if a['year'] == b['year']:
                continue
            sj = jaccard(a['_bg'], b['_bg'])
            oj = jaccard(a['_obg'], b['_obg'])
            if a['stem'] == b['stem'] and oj >= 0.5:
                union(i, j)
            elif sj >= 0.85 and oj >= 0.5:
                union(i, j)
            elif a['stem'] == b['stem']:
                same_stem_variants.append((a['id'], b['id'], round(oj, 3)))
    groups = defaultdict(list)
    for i in range(len(qs)):
        groups[find(i)].append(qs[i])
    clusters = {k: v for k, v in groups.items() if len(v) > 1}
    for cid, members in enumerate(sorted(clusters.values(), key=lambda v: (-len(v), min(x['year'] for x in v))), 1):
        gid = f'G{cid:02d}'
        for q in members:
            q['duplicate_group'] = gid
            q['duplicate_ids'] = [x['id'] for x in members if x['id'] != q['id']]
    for q in qs:
        q.setdefault('duplicate_group', '')
        q.setdefault('duplicate_ids', [])
    return same_stem_variants

def main():
    all_q = []
    for year, group, rel in XLSX_JOBS:
        all_q += parse_xlsx(rel, year, group)
    for year, group, rel in DOCX_JOBS:
        all_q += parse_docx(rel, year, group)
    all_q.sort(key=lambda q: (q['year'], 0 if q['group'] == '本科' else 1, q['num']))
    for q in all_q:
        ch, topic, kws, conf = classify(q)
        q['chapter_id'] = ch
        q['chapter'] = CHAPTERS[ch][1]
        q['topic'] = topic
        q['keywords'] = kws
        q['tags'] = sorted(set(TAGS.get(q['id'], [])))
        q['classify_confidence'] = conf
        q['answer_letters'] = q['answer']
        q['verified'] = False
        q['explanation'] = ''
        q['trap'] = ''
        q['answer_flag'] = ''
        q['reviewed'] = False
        q['error_prone'] = False
        q['high_freq'] = False

    # 人工校准：低置信度或规则易错的题目
    OVERRIDES = {
        # 12届
        'Q12-BK-012': ('U04', '问题类型'), 'Q12-BK-014': ('U07', '非抽样误差'),
        'Q12-BK-016': ('U04', '测量尺度'), 'Q12-BK-023': ('U04', '问卷编码'),
        'Q12-BK-024': ('U05', '数据整理'), 'Q12-BK-025': ('U05', '数据审核'),
        'Q12-BK-027': ('U05', '统计图'), 'Q12-BK-029': ('U06', '集中趋势'),
        'Q12-BK-032': ('U06', '集中趋势'), 'Q12-BK-051': ('U08', '相关分析'),
        'Q12-BK-056': ('U09', '方差分析'), 'Q12-BK-058': ('U09', '判别分析'),
        'Q12-BK-062': ('U02', '数据收集'), 'Q12-BK-063': ('U05', '数据整理'),
        'Q12-BK-065': ('U04', '问题设计'), 'Q12-BK-068': ('U06', '描述统计'),
        'Q12-BK-074': ('U10', '因子分析'),
        # 13届
        'Q13-BK-001': ('U01', '研究类型'), 'Q13-BK-007': ('U02', '投影技法'),
        'Q13-BK-009': ('U01', '应用范围'), 'Q13-BK-015': ('U02', '调查方法比较'),
        'Q13-BK-020': ('U03', '调查设计'), 'Q13-BK-027': ('U05', '数据整理'),
        'Q13-BK-043': ('U07', '非抽样误差'), 'Q13-BK-061': ('U02', '二手数据'),
        'Q13-BK-062': ('U07', '调查分类'), 'Q13-BK-068': ('U06', '概率基础'),
        'Q13-BK-069': ('U07', '抽样估计'), 'Q13-BK-071': ('U08', '回归分析'),
        'Q13-BK-075': ('U01', '调查项目'), 'Q13-BK-084': ('U09', '多变量概述'),
        # 14届
        'Q14-BK-004': ('U07', '重点调查'), 'Q14-BK-005': ('U02', '访问技巧'),
        'Q14-BK-008': ('U02', '焦点小组'), 'Q14-BK-012': ('U02', '实验效度'),
        'Q14-BK-025': ('U03', '操作定义'), 'Q14-BK-026': ('U05', '数据整理'),
        'Q14-BK-031': ('U06', '集中趋势'), 'Q14-BK-036': ('U06', '统计指数'),
        'Q14-BK-042': ('U07', '概率抽样'), 'Q14-BK-051': ('U08', '回归分析'),
        'Q14-BK-054': ('U08', '回归分析'), 'Q14-BK-058': ('U09', '判别分析'),
        'Q14-BK-065': ('U06', '假设检验'), 'Q14-BK-068': ('U06', '统计推断'),
        'Q14-BK-070': ('U07', '抽样基础'), 'Q14-BK-084': ('U09', '多元回归'),
        # 15届
        'Q15-BK-004': ('U02', '实验法'), 'Q15-BK-005': ('U01', '应用范围'),
        'Q15-BK-007': ('U07', '重点调查'), 'Q15-BK-012': ('U02', '现场控制'),
        'Q15-BK-018': ('U04', '问题类型'), 'Q15-BK-027': ('U05', '编码'),
        'Q15-BK-032': ('U06', '集中趋势'), 'Q15-BK-033': ('U06', '集中趋势'),
        'Q15-BK-034': ('U06', '描述统计'), 'Q15-BK-039': ('U07', '系统抽样'),
        'Q15-BK-044': ('U07', '概率抽样'), 'Q15-BK-053': ('U08', '回归分析'),
        'Q15-BK-055': ('U08', '回归分析'), 'Q15-BK-057': ('U09', 'Logistic回归'),
        'Q15-BK-067': ('U02', '调查实施'), 'Q15-BK-068': ('U06', '时间序列'),
        'Q15-BK-073': ('U08', '回归分析'), 'Q15-BK-080': ('U06', '时间序列'),
        # 16届
        'Q16-BK-003': ('U07', '抽样调查'), 'Q16-BK-014': ('U02', '数据收集'),
        'Q16-BK-015': ('U02', '数据收集'), 'Q16-BK-017': ('U02', '数据收集'),
        'Q16-BK-018': ('U04', '问题类型'), 'Q16-BK-025': ('U04', '测量尺度'),
        'Q16-BK-026': ('U05', '缺失值插补'), 'Q16-BK-029': ('U05', '缺失值插补'),
        'Q16-BK-034': ('U06', '假设检验'), 'Q16-BK-038': ('U06', '区间估计'),
        'Q16-BK-052': ('U08', '双变量分析'), 'Q16-BK-053': ('U08', '列联分析'),
        'Q16-BK-056': ('U09', '多元回归'), 'Q16-BK-058': ('U09', '方差分析'),
        'Q16-BK-060': ('U06', '统计基础'), 'Q16-BK-061': ('U02', '文案调查'),
        'Q16-BK-068': ('U06', '概率分布'), 'Q16-BK-070': ('U06', '估计量性质'),
        'Q16-BK-072': ('U09', '非参数检验'), 'Q16-BK-078': ('U03', '操作定义'),
        # 17届本科
        'Q17-BK-001': ('U01', '调查作用'), 'Q17-BK-006': ('U04', '测量尺度'),
        'Q17-BK-009': ('U01', '应用范围'), 'Q17-BK-011': ('U02', '实验法'),
        'Q17-BK-017': ('U09', '方差分析'), 'Q17-BK-029': ('U04', '问卷编码'),
        'Q17-BK-040': ('U07', '系统抽样'), 'Q17-BK-044': ('U06', '区间估计'),
        'Q17-BK-056': ('U09', '多元回归'), 'Q17-BK-057': ('U09', '多元回归'),
        'Q17-BK-061': ('U05', '数据整理'), 'Q17-BK-068': ('U07', '抽样估计'),
        'Q17-BK-074': ('U10', '因子分析'), 'Q17-BK-080': ('U06', '描述统计'),
        # 17届专科
        'Q17-ZK-001': ('Z01', '应用范围'), 'Q17-ZK-004': ('Z01', '调查程序'),
        'Q17-ZK-005': ('Z01', '应用范围'), 'Q17-ZK-007': ('Z07', '重点调查'),
        'Q17-ZK-009': ('Z07', '抽样调查'), 'Q17-ZK-022': ('Z03', '问题类型'),
        'Q17-ZK-027': ('Z04', '调查方法'), 'Q17-ZK-029': ('Z04', '调查方法'),
        'Q17-ZK-035': ('Z05', '调查实施'), 'Q17-ZK-036': ('Z03', '问题设计'),
        'Q17-ZK-038': ('Z05', '统计图'), 'Q17-ZK-041': ('Z05', '调查报告'),
        'Q17-ZK-044': ('Z05', '调查实施'), 'Q17-ZK-047': ('Z06', '描述统计'),
        'Q17-ZK-050': ('Z06', '集中趋势'), 'Q17-ZK-051': ('Z06', '统计指数'),
        'Q17-ZK-062': ('Z07', '抽样'), 'Q17-ZK-070': ('Z08', '时间序列'),
        'Q17-ZK-082': ('Z03', '问题设计'), 'Q17-ZK-089': ('Z03', '问卷编码'),
    }
    for _q in all_q:
        if _q['id'] in OVERRIDES:
            _ch, _topic = OVERRIDES[_q['id']]
            _q['chapter_id'] = _ch
            _q['chapter'] = CHAPTERS[_ch][1]
            _q['topic'] = _topic
            _q['classify_confidence'] = 'manual'
    variants = cluster_questions(all_q)
    for q in all_q:
        q.setdefault('duplicate_group', '')
        q.setdefault('duplicate_ids', [])
    # 高频题解析：同一个重复题组共享同一份解析
    groups_tmp = defaultdict(list)
    for q in all_q:
        if q.get('duplicate_group'):
            groups_tmp[q['duplicate_group']].append(q)
    for gid, members in groups_tmp.items():
        canon = next((x for x in members if x['id'] in HIGH_FREQ), None)
        if not canon:
            continue
        item = HIGH_FREQ[canon['id']]
        for q in members:
            q['high_freq'] = True
            q['explanation'] = item['why']
            q['trap'] = item['trap']
            q['reviewed'] = True
    # 易错题解析
    for q in all_q:
        if q['id'] in ERROR_PRONE:
            q['error_prone'] = True
            if not q['explanation']:
                q['explanation'] = ERROR_PRONE[q['id']]['why']
                q['trap'] = ERROR_PRONE[q['id']]['trap']
            q['reviewed'] = True
    # 答案/图片存疑标记
    for q in all_q:
        if q['id'] in FLAGGED:
            q['answer_flag'] = FLAGGED[q['id']]
    # 清理内部字段并输出
    for q in all_q:
        for k in ['_bg', '_obg']:
            q.pop(k, None)
    os.makedirs(QDIR, exist_ok=True)
    with open(os.path.join(QDIR, 'questions.jsonl'), 'w', encoding='utf8') as f:
        for q in all_q:
            f.write(json.dumps(q, ensure_ascii=False) + '\n')
    with open(os.path.join(QDIR, 'questions.csv'), 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(['id', 'year', 'group', 'type', 'chapter_id', 'chapter', 'topic',
                    'stem', 'A', 'B', 'C', 'D', 'E', 'answer', 'high_freq',
                    'error_prone', 'duplicate_group', 'verified', 'reviewed',
                    'explanation', 'trap', 'answer_flag', 'source'])
        for q in all_q:
            o = q['options']
            w.writerow([q['id'], q['year'], q['group'], q['type'], q['chapter_id'],
                        q['chapter'], q['topic'], q['stem'], o.get('A', ''), o.get('B', ''),
                        o.get('C', ''), o.get('D', ''), o.get('E', ''), q['answer'],
                        q['high_freq'], q['error_prone'], q['duplicate_group'],
                        q['verified'], q['reviewed'], q['explanation'], q['trap'],
                        q['answer_flag'], q['source']])
    # 按章节
    chap_dir = os.path.join(QDIR, '按章节')
    os.makedirs(chap_dir, exist_ok=True)
    for ch in CHAPTERS:
        members = [q for q in all_q if q['chapter_id'] == ch]
        if not members:
            continue
        lines = [f"# {CHAPTERS[ch][1]}", '',
                 f"> {CHAPTERS[ch][2]} ｜ 共 {len(members)} 题", '']
        for q in sorted(members, key=lambda x: (x['year'], x['num'])):
            lines += render_question(q, image_prefix='../assets')
        open(os.path.join(chap_dir, f'{ch}_{CHAPTERS[ch][1].split(" ",1)[-1]}.md'),
             'w', encoding='utf8').write('\n'.join(lines) + '\n')
    # 按届
    year_dir = os.path.join(QDIR, '按届')
    os.makedirs(year_dir, exist_ok=True)
    for year in sorted(set(q['year'] for q in all_q)):
        for group in ['本科', '专科']:
            members = [q for q in all_q if q['year'] == year and q['group'] == group]
            if not members:
                continue
            title = f'第{year}届市调大赛{group}模拟题'
            lines = [f'# {title}', '', f'> 共 {len(members)} 题', '']
            for q in sorted(members, key=lambda x: x['num']):
                lines += render_question(q, image_prefix='../assets')
            fn = f'第{year}届_{group}_模拟题.md'
            open(os.path.join(year_dir, fn), 'w', encoding='utf8').write('\n'.join(lines) + '\n')
    # 质量报告
    qual_dir = os.path.join(ROOT, '40_质量')
    os.makedirs(qual_dir, exist_ok=True)
    lines = ['# 重复题聚类', '',
             '同一题在不同届次重复出现时归入同一组。保留全部原题，不删除任何版本。', '']
    groups = defaultdict(list)
    for q in all_q:
        if q.get('duplicate_group'):
            groups[q['duplicate_group']].append(q)
    for gid in sorted(groups):
        members = sorted(groups[gid], key=lambda x: (x['year'], x['num']))
        lines += [f'## {gid}（{len(members)} 题）', '']
        for q in members:
            lines.append(f"- `{q['id']}` {q['type']} 答案 {q['answer']}：{q['stem'][:80]}")
        lines += ['', '**共同考点**：' + members[0]['chapter'] + ' / ' + members[0]['topic'], '']
    open(os.path.join(qual_dir, '重复题聚类.md'), 'w', encoding='utf8').write('\n'.join(lines) + '\n')
    lines = ['# 答案冲突检查', '',
             '仅当题干相同且选项高度一致时，才视为同一道题；否则视为同题干变体，不判为冲突。', '']
    if variants:
        lines += ['## 同题干但选项不同（不是冲突）', '']
        for a, b, oj in variants:
            lines.append(f'- `{a}` 与 `{b}`：选项重合度 {oj}，属于不同版本的“下列说法中错误的是”，需分别按各自选项作答。')
    else:
        lines.append('未发现同题干变体。')
    lines += ['', '## 同题答案冲突', '']
    conflict = 0
    for gid, members in groups.items():
        by_opts = defaultdict(set)
        for q in members:
            key = tuple(sorted(q['options'].items()))
            by_opts[key].add(q['answer'])
        for key, answers in by_opts.items():
            if len(answers) > 1:
                conflict += 1
                lines.append(f'- {gid}: {answers}')
    if not conflict:
        lines.append('未发现同一题、同一组选项但答案不同的情况。')
    open(os.path.join(qual_dir, '答案冲突.md'), 'w', encoding='utf8').write('\n'.join(lines) + '\n')
    # 校验记录
    issues = []
    for q in all_q:
        if not q['stem']:
            issues.append((q['id'], '题干为空'))
        if len(q['options']) < 2:
            issues.append((q['id'], f'选项不足：{len(q["options"])}'))
        if q['type'] == '判断题' and set(q['options']) != {'A', 'B'}:
            issues.append((q['id'], '判断题选项不是 A/B'))
        if any(c not in q['options'] for c in q['answer'] if c.isalpha()):
            issues.append((q['id'], f'答案字母不在选项中：{q["answer"]}'))
        if q['classify_confidence'] == 'low':
            issues.append((q['id'], '章节归类低置信度'))
    lines = ['# 校验记录', '', f'- 总题数：{len(all_q)}',
             f'- 本科：{sum(1 for q in all_q if q["group"]=="本科")}',
             f'- 专科：{sum(1 for q in all_q if q["group"]=="专科")}',
             f'- 发现待处理项：{len(issues)}', '', '## 待处理项', '']
    for qid, msg in issues:
        lines.append(f'- `{qid}`：{msg}')
    if not issues:
        lines.append('无。')
    open(os.path.join(qual_dir, '校验记录.md'), 'w', encoding='utf8').write('\n'.join(lines) + '\n')
    print('题目总数', len(all_q), '重复组', len(groups), '待处理', len(issues))
    print('章节分布：')
    for ch in CHAPTERS:
        n = sum(1 for q in all_q if q['chapter_id'] == ch)
        print(' ', ch, CHAPTERS[ch][1], n)
    print('低置信度：')
    for qid, msg in issues:
        if '低置信度' in msg:
            q = next(x for x in all_q if x['id'] == qid)
            print(' ', qid, q['stem'][:60])

def render_question(q, image_prefix='assets'):
    def img(s):
        return re.sub(r'\[\[IMG:([^\]]+)\]\]',
                      lambda m: f'![图]({image_prefix}/{os.path.basename(m.group(1))})' if m.group(1) else '![图缺失]',
                      s)
    lines = [f"## 第{q['num']}题 · {q['type']}", '',
             f"> `{q['id']}` ｜ {q['chapter']} / {q['topic']} ｜ 答案 **{q['answer']}**", '',
             f"**题目**：{img(q['stem'])}", '']
    if q['options']:
        lines += ['**选项**：', '']
        for lab in sorted(q['options']):
            lines.append(f"- {lab}. {img(q['options'][lab])}")
        lines.append('')
    if q.get('explanation'):
        lines += [f"**解析**：{q['explanation']}", '']
    if q.get('trap'):
        lines += [f"**易错点**：{q['trap']}", '']
    if q.get('answer_flag'):
        lines += [f"**待核对**：{q['answer_flag']}", '']
    return lines

if __name__ == '__main__':
    main()
