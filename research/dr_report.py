# -*- coding: utf-8 -*-
"""深度研究档案 (/dr 档案.md) 解析: 章节切分 + 量化结论提取 + Assumptions依据回填"""
import re

# 形如 "## 3. 业务量价拆解" / "### 4.2 毛利率" / "## 近况" 的标题
_SEC_RE = re.compile(r'^(#{1,4})\s*(?:(\d+(?:\.\d+)*)[\.、\s]+)?(.*)$')
# 量化特征: 数字+单位/财务关键词
_QUANT_RE = re.compile(r'\d+(?:\.\d+)?\s*(?:%|％|亿|百万|万㎡|万平米|万平方米|万片|万平|元|x|倍|pct|天)')
_DOMAIN_KW = ['HDI', '高多层', 'MLP', 'MLPCB', 'FPC', '单双层', '毛利率', '产能', 'Capex',
              '资本开支', '销量', 'ASP', '均价', 'WACC', '税率', '费用率', 'DSO', 'DPO',
              'DIO', '应收', '存货', '应付', '股利', '分红', 'β', 'beta', '增速', '收入',
              '营收', '归母', '净利', '折现', '永续']
# 假设行label → 档案关键词映射 (命中即回填依据; 前面的关键词权重高)
_LABEL_KW = {
    'HDI': ['HDI'], '高多层': ['高多层', 'MLP', 'MLPCB'], 'FPC': ['FPC'], '单双层': ['单双层'],
    '毛利率': ['毛利率'], 'ASP': ['ASP', '均价'], '销量': ['销量', '产能'],
    'Capex': ['Capex', '资本开支'], '资本开支': ['Capex', '资本开支'],
    '税率': ['税率'], 'DSO': ['DSO', '应收'], 'DIO': ['DIO', '存货'], 'DPO': ['DPO', '应付'],
    '股利': ['股利', '分红'], 'WACC': ['WACC'], '永续': ['永续'], '增速': ['增速', '收入', '营收'],
}


def parse(path):
    """解析md档案 → {path, text, sections, findings}
    sections: [{no, title, lines}]; findings: [{section, title, text, keywords}]"""
    with open(path, encoding='utf-8') as f:
        text = f.read()
    sections = []
    cur = {'no': '0', 'title': '全文', 'lines': []}
    auto_no = 0
    for line in text.splitlines():
        m = _SEC_RE.match(line.strip())
        if m and line.strip().startswith('#'):
            if cur['lines']:
                sections.append(cur)
            auto_no += 1
            cur = {'no': m.group(2) or str(auto_no), 'title': (m.group(3) or '').strip(), 'lines': []}
        elif line.strip():
            cur['lines'].append(line.strip())
    if cur['lines']:
        sections.append(cur)

    findings = []
    for s in sections:
        for ln in s['lines']:
            if len(ln) < 8 or ln.startswith('|') and ln.count('|') < 3:
                continue
            if _QUANT_RE.search(ln) and any(k in ln for k in _DOMAIN_KW):
                findings.append({'section': s['no'], 'title': s['title'], 'text': ln,
                                 'keywords': [k for k in _DOMAIN_KW if k in ln]})
    return {'path': path, 'text': text, 'sections': sections, 'findings': findings}


def annotate(findings, label, max_hits=2):
    """对Assumptions行label匹配档案量化结论 → ['dr档案§N: ...', ...] (最多max_hits条)
    按命中关键词数打分, 优先返回与label最相关的结论 (如同为HDI行, 优先取含'HDI'的段落)"""
    kws = []
    for lk, ks in _LABEL_KW.items():
        if lk in label:
            kws.extend(ks)
    if not kws:
        return []
    scored = []
    for idx, f in enumerate(findings):
        score = sum(1 for k in set(kws) if k in f['text'])
        if score:
            scored.append((-score, idx, f))
    scored.sort()
    hits = []
    for _, _, f in scored[:max_hits]:
        t = f['text']
        if len(t) > 60:
            t = t[:57] + '...'
        hits.append(f"dr档案§{f['section']}: {t}")
    return hits
