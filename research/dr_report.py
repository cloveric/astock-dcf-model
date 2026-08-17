# -*- coding: utf-8 -*-
"""深度研究档案 (/dr 档案.md) 解析: 章节切分 + 量化结论提取 + Assumptions依据回填"""
import re

# 形如 "## 3. 业务量价拆解" / "### 4.2 毛利率" / "## 近况" 的标题
_SEC_RE = re.compile(r'^(#{1,4})\s*(?:(\d+(?:\.\d+)*)[\.、\s]+)?(.*)$')
# 量化特征: 数字+单位/财务关键词
_QUANT_RE = re.compile(r'\d+(?:\.\d+)?\s*(?:%|％|亿|百万|万㎡|万平米|万平方米|万片|万平|元|x|倍|pct|天)')
# 通用财务/经营词 (公司无关) — DOMAIN 关键词基础; 产品词不再硬编码,
# 由 set_segments() 从 cfg segments 运行时派生 (见下)
_DOMAIN_BASE = ['毛利率', '净利率', '产能', 'Capex', '资本开支', '销量', 'ASP', '均价',
                'WACC', '税率', '费用率', 'DSO', 'DPO', 'DIO', '应收', '存货', '应付',
                '股利', '分红', 'β', 'beta', '增速', '收入', '营收', '归母', '净利',
                '折现', '永续', '占比', '份额', '渗透率', '利用率', '出货', '涨价', '目标价']
# 假设行label → 档案关键词映射 (通用财务词部分; 命中即回填依据)
_LABEL_KW_BASE = {
    '毛利率': ['毛利率'], 'ASP': ['ASP', '均价'], '均价': ['ASP', '均价'],
    '销量': ['销量', '产能'], 'Capex': ['Capex', '资本开支'], '资本开支': ['Capex', '资本开支'],
    '税率': ['税率'], 'DSO': ['DSO', '应收'], 'DIO': ['DIO', '存货'], 'DPO': ['DPO', '应付'],
    '股利': ['股利', '分红'], 'WACC': ['WACC'], '永续': ['永续'], '增速': ['增速', '收入', '营收'],
}
# 分词时丢弃的过泛词元
_STOP = {'其他', '其它', '合计', '业务', '产品', '路径', '假设', '采用值'}
# set_segments() 注入的动态状态: 业务/产品关键词 与 label词→同业务别名组
_seg_domain = []
_seg_label = {}

_BRK_RE = re.compile(r'[（(]([^（）()]*)[)）]')


def _split_tokens(text):
    """拆词: 括号外主名 + 括号内按 / 、 , 空白拆分的别名; 过滤过短/纯数字/停用词"""
    out = []
    main = _BRK_RE.sub(' ', str(text))
    for part in _BRK_RE.findall(str(text)):
        main += ' ' + part.replace('/', ' ')
    for p in re.split(r'[\s/、,，:：|]+', main):
        p = p.strip()
        if len(p) >= 2 and not p.isdigit() and p not in _STOP and p not in out:
            out.append(p)
    return out


def set_segments(segments):
    """注入 cfg segments (list of dict, 取 name/short 字段), 运行时派生业务关键词

    派生规则: 每段取 全名(去括号)/short/括号内别名(按/拆分) 为词元;
    同段词元互为别名组 (label 命中任一词元 → 匹配档案中该段任意别名)。
    调用方不可编辑(research/__init__、build_model 均硬调 parse/annotate 且不传 cfg),
    故以模块级注入提供; 未注入时退化为通用财务词 + annotate 的 label 自身分词兜底。
    传 None/空 则清空回退。"""
    global _seg_domain, _seg_label
    _seg_domain, _seg_label = [], {}
    for seg in (segments or []):
        if not isinstance(seg, dict):
            continue
        toks = []
        for field in ('name', 'short'):
            for t in _split_tokens(seg.get(field) or ''):
                if t not in toks:
                    toks.append(t)
        for t in toks:
            if t not in _seg_domain:
                _seg_domain.append(t)
            group = _seg_label.setdefault(t, [])
            for other in toks:
                if other not in group:
                    group.append(other)


def _domain_kw():
    """当前生效的 DOMAIN 关键词 = 通用财务词 + segments 派生词 (调用时合并, 注入即生效)"""
    return _DOMAIN_BASE + [k for k in _seg_domain if k not in _DOMAIN_BASE]


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
    dkw = _domain_kw()
    for s in sections:
        for ln in s['lines']:
            if len(ln) < 8 or ln.startswith('|') and ln.count('|') < 3:
                continue
            if _QUANT_RE.search(ln) and any(k in ln for k in dkw):
                findings.append({'section': s['no'], 'title': s['title'], 'text': ln,
                                 'keywords': [k for k in dkw if k in ln]})
    return {'path': path, 'text': text, 'sections': sections, 'findings': findings}


def annotate(findings, label, max_hits=2):
    """对Assumptions行label匹配档案量化结论 → ['dr档案§N: ...', ...] (最多max_hits条)
    关键词三来源: 通用财务词映射 + set_segments注入的业务别名组 + label自身分词兜底
    (label 文本本就来自 cfg segments short/科目名, 分词即天然的运行时动态关键词)。
    按命中关键词数打分, 优先返回与label最相关的结论 (如业务分部行, 优先取含该分部名的段落)"""
    kws = []
    for mapping in (_LABEL_KW_BASE, _seg_label):
        for lk, ks in mapping.items():
            if lk in label:
                kws.extend(ks)
    kws.extend(_split_tokens(label))        # label自身词元兜底 (含分部名/括号别名)
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
