# -*- coding: utf-8 -*-
"""
通用数据获取层 + 兜底配置生成 (fetch_data.py)

数据源纪律:
  - 东财F10 (A股三表/主营构成): 一律走系统 curl 子进程 (python 不直连东财, 避免反爬)
  - 东财HKF10 (港股三表, datacenter长表接口): 同样系统curl; IFRS科目映射, 金额为财报
    币种(如中芯国际=美元); 现金流量表仅人民币口径, 按"期末现金/BS现金"隐含汇率折算
  - 腾讯行情 (现价/总市值/PE-TTM): 系统 curl qt.gtimg.cn; A股与港股hkXXXXX同版式
  - gildata/聚源一致预期: 本工具不直连; 由用户手工填入 configs/<code>.yaml 的
    consensus 段 (有 MCP 环境时可查后回填), 缺省时用"自动推导"占位并在依据中注明

用法:
  python fetch_data.py --code 002463                 # 拉取公开数据 → configs/002463.yaml
  python fetch_data.py --code 00981                  # 港股: 腾讯hk行情+东财HKF10 → configs/00981.yaml
  python fetch_data.py --code 002463 --stdout        # 打印生成的配置而不写文件
  python fetch_data.py --code 002463 --quote-only    # 只打印行情快照

兜底推导规则 (全部在"依据"字段中注明, 精度局限见 README):
  - 分部: 取东财F10主营构成(按产品)最近年报前3项, 其余并入"其他"; 各段增速=最近年报
    分部收入YoY(截断[-15%,35%])逐年退坡(×1/.75/.55/.40/.30), 毛利率=最近年报分部毛利率持平;
  - 费用率/税率/营运资本天数/Capex占比: 取最近年报实际比率, 5年持平(Capex退坡);
  - 股利30%/盈余公积10%/最低现金2%/四档利率(LPR附近)/现金收益1.2%: 通用默认;
  - 计划还款: 一年内到期/长借/租赁各按余额1/5逐年摊还, 短借(revolver)0;
  - 一致预期: 用最近年报收入×(1+总收入YoY)与净利率外推占位(依据注明"自动推导");
  - 可比公司: 兜底默认空清单, WACC的βu转为蓝色输入(默认1.0), 建议手工补 comps。
"""
import argparse
import datetime
import json
import os
import re
import subprocess
import sys

UA = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36'
EM_BASE = 'https://emweb.securities.eastmoney.com/PC_HSF10'


def curl_get(url, referer='https://emweb.securities.eastmoney.com/', timeout=30):
    """系统curl取数 (不用python直连, 避免反爬)"""
    cmd = ['curl', '-s', '--compressed', '--max-time', str(timeout),
           '-H', f'User-Agent: {UA}', '-H', f'Referer: {referer}', url]
    out = subprocess.run(cmd, check=True, capture_output=True, timeout=timeout + 10)
    return out.stdout


def em_code(code):
    """6位代码 → 东财代码 (SZ/SH/BJ)"""
    code = str(code).strip()
    if code.startswith(('60', '68', '90')):
        return f'SH{code}'
    if code.startswith(('00', '30', '20')):
        return f'SZ{code}'
    if code.startswith(('4', '8', '92')):
        return f'BJ{code}'
    raise ValueError(f'无法识别交易所: {code}')


def is_hk(code):
    """5位数字 = 港股代码 (如 00981)"""
    return bool(re.fullmatch(r'\d{5}', str(code).strip()))


def tx_code(code):
    """证券代码 → 腾讯行情代码 (小写; 港股为 hkXXXXX)"""
    code = str(code).strip()
    if is_hk(code):
        return f'hk{code}'
    return em_code(code).lower()


def fetch_quote(code):
    """腾讯行情快照: 现价/总市值(亿)/PE-TTM/名称
    A股(sh/sz/bj)与港股(hkXXXXX)版式的关键字段位置一致: f[3]现价 f[39]PE-TTM f[45]总市值(亿);
    港股现价/总市值为港元口径 (实测 00981/00700 字段对齐)"""
    raw = curl_get(f'https://qt.gtimg.cn/q={tx_code(code)}', referer='https://gu.qq.com/')
    txt = raw.decode('gbk', errors='replace')
    m = re.search(r'="(.*)"', txt)
    if not m:
        raise RuntimeError(f'腾讯行情解析失败: {txt[:120]}')
    f = m.group(1).split('~')
    price = float(f[3])
    pe_ttm = float(f[39])
    mcap_yi = float(f[45])          # 总市值, 亿元(A股) / 亿港元(港股)
    name = f[1]
    shares_m = round(mcap_yi * 100 / price, 1)   # 总股本(百万股) = 市值/价
    out = {'name': name, 'price': price, 'pe_ttm': pe_ttm,
           'mcap_yi': mcap_yi, 'shares_m': shares_m}
    if is_hk(code):
        out['quote_ccy'] = '港元'
    return out


def _dates_param(years):
    return '%2C'.join(f'{y}-12-31' for y in sorted(years, reverse=True))


def fetch_f10_statement(code, kind, years):
    """东财F10三表: kind ∈ lrb/zcfzb/xjllb; 返回 {year: row dict}"""
    url = (f'{EM_BASE}/NewFinanceAnalysis/{kind}AjaxNew?companyType=4'
           f'&reportDateType=0&reportType=1&dates={_dates_param(years)}&code={em_code(code)}')
    data = json.loads(curl_get(url))['data']
    out = {}
    for row in data:
        y = int(row['REPORT_DATE'][:4])
        out[y] = row
    missing = [y for y in years if y not in out]
    if missing:
        raise RuntimeError(f'F10 {kind} 缺少年报: {missing}')
    return out


def fetch_composition(code):
    """东财F10主营构成: 返回 {year: [按产品分项 dict]}"""
    url = f'{EM_BASE}/BusinessAnalysis/PageAjax?code={em_code(code)}'
    data = json.loads(curl_get(url))['zygcfx']
    out = {}
    for row in data:
        if row['MAINOP_TYPE'] != '2':       # 2=按产品
            continue
        if not row['REPORT_DATE'][5:10] == '12-31':   # 只保留年报, 防中报/季报混入
            continue
        y = int(row['REPORT_DATE'][:4])
        out.setdefault(y, []).append(row)
    for y in out:
        out[y].sort(key=lambda r: r.get('RANK') or 99)
    return out


# ============================================================
# 港股 (5位代码): 东财HKF10(datacenter API, 系统curl) + 腾讯hk行情
#   - 三表为 IFRS 科目、财报币种(如中芯国际=美元); 现金流量表为人民币口径,
#     按"CF期末现金/BS现金及等价物"的隐含汇率逐年折算回财报币种(依据注明);
#   - 无东财"主营构成(按产品)"披露 → 单段简化收入驱动;
#   - IFRS 无税金及附加/法定盈余公积等科目 → 置0并在依据注明。
# ============================================================

def fetch_hk_long_table(rpt, code, item_key):
    """东财HKF10长表 → {year: {科目: 金额}} (仅年报 DATE_TYPE_CODE=001)"""
    url = (f'https://datacenter.eastmoney.com/securities/api/data/v1/get?reportName={rpt}'
           f'&columns=ALL&filter=(SECUCODE%3D%22{code}.HK%22)&pageNumber=1&pageSize=500'
           '&sortTypes=-1&sortColumns=REPORT_DATE&source=HSF10&client=PC')
    data = json.loads(curl_get(url, referer='https://data.eastmoney.com/'))['result']['data']
    out, ccy = {}, None
    for row in data:
        if row.get('DATE_TYPE_CODE') != '001':
            continue
        y = int(row['REPORT_DATE'][:4])
        out.setdefault(y, {})[row[item_key]] = row['AMOUNT']
        ccy = ccy or row.get('CURRENCY')
    return out, ccy


def build_hist_hk(code, years):
    """东财HKF10三表 → 模型 hist 结构 (单位: 财报币种百万, IFRS科目映射)"""
    inc, ccy = fetch_hk_long_table('RPT_HKF10_FN_INCOME', code, 'ITEM_NAME')
    bal, _ = fetch_hk_long_table('RPT_HKF10_FN_BALANCE', code, 'ITEM_NAME')
    cfs, _ = fetch_hk_long_table('RPT_HKF10_FN_CASHFLOW_PC', code, 'STD_ITEM_NAME')
    missing = [y for y in years if y not in inc or y not in bal or y not in cfs]
    if missing:
        raise RuntimeError(f'HKF10 缺少年报: {missing}')

    def m(v, nd=1):
        return 0.0 if v is None else round(v / 1e6, nd)

    is_d = {k: [] for k in ('rev', 'cost', 'taxadd', 'sale', 'adm', 'rd', 'fin',
                            'othop', 'nonop', 'taxexp', 'np_p')}
    bs_d = {k: [] for k in ('cash', 'tfa', 'ar', 'pre', 'orec', 'inv', 'oca', 'ppe', 'rou',
                            'ia', 'gw', 'lpe', 'dta', 'oei', 'onca', 'stl', 'ap', 'contract',
                            'staff', 'taxp', 'opay', 'cur1y', 'ocl', 'ltl', 'lease', 'oncl',
                            'sc', 'cr', 'ts', 'oci', 'sr', 're', 'oeq', 'mi')}
    cf_d = {k: [] for k in ('np', 'da', 'ocf', 'capex', 'icf', 'fcf', 'cash1')}
    for y in sorted(years):
        L, B = inc[y], bal[y]

        def li(k):
            return L.get(k) or 0.0

        def bi(k):
            return B.get(k) or 0.0

        rev, cost = li('营业额'), li('销售成本')
        sale, adm, rd = li('销售及分销费用'), li('行政开支'), li('研发费用')
        fin = li('融资成本') - li('利息收入')          # 净融资成本 (IFRS口径)
        op, tp = li('经营溢利'), li('除税前溢利')
        is_d['rev'].append(m(rev)); is_d['cost'].append(m(cost))
        is_d['taxadd'].append(0.0)                    # IFRS无"税金及附加"科目
        is_d['sale'].append(m(sale)); is_d['adm'].append(m(adm)); is_d['rd'].append(m(rd))
        is_d['fin'].append(m(fin))
        is_d['othop'].append(m(op - (rev - cost - sale - adm - rd - fin)))   # 轧差(含其他收入/应占联营)
        is_d['nonop'].append(m(tp - op))
        is_d['taxexp'].append(m(li('税项')))
        is_d['np_p'].append(m(li('股东应占溢利')))

        cash = bi('现金及等价物')
        tfa = bi('短期投资') + bi('指定以公允价值记账之金融资产(流动)')
        ar, pre, inv = bi('应收帐款'), bi('预付款按金及其他应收款'), bi('存货')
        tca = bi('流动资产合计')
        ppe, ia = bi('物业厂房及设备'), bi('无形资产')
        dta, oei = bi('递延税项资产'), bi('长期投资')
        tnca = bi('非流动资产合计')
        stl, ap = bi('短期贷款'), bi('应付帐款')
        contract = bi('合同负债') + bi('预收款项')
        taxp, cur1y = bi('应付税项'), bi('融资租赁负债(流动)')
        tcl = bi('流动负债合计')
        ltl, lease = bi('长期贷款'), bi('融资租赁负债(非流动)')
        tncl = bi('非流动负债合计')
        sc, cr, oci = bi('股本'), bi('股本溢价'), bi('其他储备')
        re_ = bi('保留溢利(累计亏损)')
        tpe = bi('股东权益')
        bs_d['cash'].append(m(cash)); bs_d['tfa'].append(m(tfa))
        bs_d['ar'].append(m(ar)); bs_d['pre'].append(m(pre)); bs_d['orec'].append(0.0)
        bs_d['inv'].append(m(inv))
        bs_d['oca'].append(m(tca - cash - tfa - ar - pre - inv))     # 轧差
        bs_d['ppe'].append(m(ppe)); bs_d['rou'].append(0.0)          # IFRS使用权资产未单列
        bs_d['ia'].append(m(ia)); bs_d['gw'].append(0.0); bs_d['lpe'].append(0.0)
        bs_d['dta'].append(m(dta)); bs_d['oei'].append(m(oei))
        bs_d['onca'].append(m(tnca - ppe - ia - dta - oei))          # 轧差(含联营公司权益)
        bs_d['stl'].append(m(stl)); bs_d['ap'].append(m(ap))
        bs_d['contract'].append(m(contract)); bs_d['staff'].append(0.0)
        bs_d['taxp'].append(m(taxp)); bs_d['opay'].append(0.0)
        bs_d['cur1y'].append(m(cur1y))
        bs_d['ocl'].append(m(tcl - stl - ap - contract - taxp - cur1y))   # 轧差
        bs_d['ltl'].append(m(ltl)); bs_d['lease'].append(m(lease))
        bs_d['oncl'].append(m(tncl - ltl - lease))                   # 轧差(含递延收入/递延税项负债)
        bs_d['sc'].append(m(sc)); bs_d['cr'].append(m(cr)); bs_d['ts'].append(0.0)
        bs_d['oci'].append(m(oci)); bs_d['sr'].append(0.0)           # IFRS无法定盈余公积
        bs_d['re'].append(m(re_))
        bs_d['oeq'].append(m(tpe - sc - cr - oci - re_))             # 轧差
        bs_d['mi'].append(m(bi('少数股东权益')))

        C = cfs[y]
        # 现金流表为人民币口径 → 按隐含汇率(期末现金/BS现金)折算回财报币种
        fx = (C.get('期末现金') or 0.0) / cash if cash else 1.0
        if abs(fx - 1.0) < 0.05:
            fx = 1.0

        def ci(k):
            return (C.get(k) or 0.0) / fx

        cf_d['np'].append(is_d['np_p'][-1])
        cf_d['da'].append(m(ci('加:折旧及摊销')))
        cf_d['ocf'].append(m(ci('经营业务现金净额')))
        cf_d['capex'].append(m(-(ci('购建固定资产') + ci('购建无形资产及其他资产'))))
        cf_d['icf'].append(m(ci('投资业务现金净额')))
        cf_d['fcf'].append(m(ci('融资业务现金净额')))
        cf_d['cash1'].append(m(ci('期末现金')))       # = BS现金及等价物 (隐含汇率口径)
    y0 = sorted(years)[-1]
    ppe_split = {'fa_net': m(bal[y0].get('物业厂房及设备')),
                 'cip': 0.0}                            # HKF10未拆分在建工程
    return {'is': is_d, 'bs': bs_d, 'cf': cf_d, 'ppe_split': ppe_split,
            'currency': ccy or '美元', 'cf_fx_note': '现金流量表(人民币口径)按隐含汇率折算'}


def build_segments_hk(hist, years):
    """港股无东财主营构成披露 → 单段(整体)简化: 增速=总收入YoY退坡, 毛利率=历史实际"""
    rev = hist['is']['rev']
    gm = []
    for i, y in enumerate(sorted(years)):
        g = (rev[i] - hist['is']['cost'][i]) / rev[i] if rev[i] else 0.30
        gm.append(round(g, 4))
    gy = rev[-1] / rev[-2] - 1 if len(rev) >= 2 and rev[-2] else 0.10
    g0 = min(max(gy, -0.15), 0.35)
    growth = [round(g0 * f, 4) for f in (1.0, 0.75, 0.55, 0.40, 0.30)]
    return [{
        'key': 'seg1', 'name': '整体业务(单段简化)', 'short': '整体', 'driver': 'growth',
        'hist_share': [1.0] * len(years), 'hist_gm': gm,
        'share_basis': '港股无东财主营构成(按产品)披露, 单段=总收入100%',
        'vol': {'values': growth,
                'basis': f'自动推导: 最近年报总收入YoY={gy:+.1%}(截断[-15%,35%])逐年退坡×1/.75/.55/.40/.30; 建议按研究修正'},
        'gm': {'values': [gm[-1]] * 5,
               'basis': f'自动推导: 最近年报综合毛利率{gm[-1]:.1%}持平; 建议按研究修正'},
        'logic': '港股单段简化(东财HKF10无主营构成披露); 历史毛利率=1-销售成本/营业额(IFRS)',
    }]


def build_fallback_config_hk(code):
    """港股全自动兜底: 腾讯hk行情 + 东财HKF10三表(IFRS) → 完整配置dict"""
    code = str(code)
    quote = fetch_quote(code)
    today = datetime.date.today().isoformat()
    years = [2023, 2024, 2025]
    hist = build_hist_hk(code, years)
    ccy = hist['currency']
    segs = build_segments_hk(hist, years)
    fcst = list(range(years[-1] + 1, years[-1] + 6))
    ass = derive_assumptions(hist, fcst)
    ass['dividend']['surplus_rate'] = {'value': 0.0,
                                       'basis': 'IFRS无法定盈余公积科目, 计提率置0'}
    fx_hkd = 7.80 if ccy == '美元' else 1.0
    price_m = round(quote['price'] / fx_hkd, 4)       # 模型价(财报币种) = 港元价 / fx
    rev0, np0 = hist['is']['rev'][-1], hist['is']['np_p'][-1]
    gy = rev0 / hist['is']['rev'][-2] - 1 if len(hist['is']['rev']) >= 2 else 0.1
    g_path = [gy * f for f in (1.0, 0.75, 0.55)]
    cons_rev = [round(rev0 * (1 + gp), 1) for gp in g_path]
    npm0 = np0 / rev0 if rev0 else 0.1
    cons_np = [round(rv * npm0, 1) for rv in cons_rev]
    cfg = {
        'company': {'code': code, 'code_full': f'{code}.HK', 'name': quote['name'],
                    'currency_note': f'{ccy}(财报口径) / 百万元 (每股数据为{ccy}; 港元现价按fx折算)',
                    'unit': f'{ccy}百万元(财报口径)'},
        'model': {'hist_years': years, 'fcst_years': fcst,
                  'valuation_date': today, 'build_date': today},
        'market': {
            'price': {'value': price_m,
                      'basis': f'{today}腾讯行情港元现价{quote["price"]} / fx{fx_hkd}={price_m}{ccy}'},
            'shares': {'value': quote['shares_m'],
                       'basis': f'总市值{quote["mcap_yi"]:.1f}亿港元/港元现价(腾讯行情 {today}); 港股总市值=全部股本×港元价, 与A+H实际口径有差异'},
            'pe_ttm': quote['pe_ttm'],
            'pe_ttm_basis': f'腾讯行情PE-TTM(港元口径, {today})',
            'fx_hkd': {'value': fx_hkd,
                       'basis': ('手工输入: 港元/财报币种汇率(美元联系汇率区间7.75-7.85中枢); 请按实际更新'
                                 if ccy == '美元' else '财报币种=人民币, 与港元接近1:1折算, 请按实际更新')},
        },
        'consensus': {
            'rev': {'values': cons_rev, 'basis': f'自动推导(无一致预期输入): 最近年报收入×(1+YoY{gy:+.1%})退坡; 建议手工填入'},
            'np': {'values': cons_np, 'basis': f'自动推导(无一致预期输入): 收入×最近年报净利率{npm0:.1%}; 建议手工填入'},
        },
        'segments': segs,
        **ass,
        'wacc': {
            'rf': {'value': 0.04, 'basis': '通用默认(港股): 10年期美债收益率约4%(财报币种口径)'},
            'erp': {'value': 0.055, 'basis': '通用默认: ERP中枢5.5%; 港股口径请按研究修正'},
            'srp': {'value': 0.0, 'basis': '通用默认: 不取个股溢价(避免与β/情景重复计价)'},
            'kd': {'value': 0.034, 'basis': '通用默认: ≈债务加权平均利率'},
            'tg': {'value': 0.03, 'basis': '通用默认: 长期名义增速中枢下沿'},
            'override': None,
            'beta_unlevered_input': 1.0,
            'beta_unlevered_basis': '兜底默认: 未配置可比公司, βu=1.0蓝色输入; 建议补relative_val.comps后自动取中位数',
        },
        'scenarios': {
            'bear': {'rev_adj': -0.15, 'npm_adj': -0.03, 'logic': '通用默认: 分部增速-15pct, 净利率-3pct'},
            'base': {'rev_adj': 0.0, 'npm_adj': 0.0, 'logic': '基准: 不调整'},
            'bull': {'rev_adj': 0.10, 'npm_adj': 0.015, 'logic': '通用默认: 分部增速+10pct, 净利率+1.5pct'},
        },
        'hist': {k: v for k, v in hist.items() if k in ('is', 'bs', 'cf', 'ppe_split')},
        'relative_val': {
            'target_pe_lo': {'value': 25.0, 'basis': '兜底默认: 25x(无可比清单); 建议补comps后取可比中位数为上界'},
            'comps': [],
        },
        'cover': {'subtitle': f'{quote["name"]} | 港股自动兜底建模(东财HKF10+腾讯行情; 财报币种{ccy})'},
    }
    return cfg


def _m(v, nd=1):
    """元 → 百万, 四舍五入; None → 0.0"""
    if v is None:
        return 0.0
    return round(v / 1e6, nd)


def _g(row, key):
    v = row.get(key)
    return v if isinstance(v, (int, float)) else None


def build_hist(code, years):
    """拉取并映射三表为模型 hist 结构 (单位: 百万)"""
    lrb = fetch_f10_statement(code, 'lrb', years)
    zcfzb = fetch_f10_statement(code, 'zcfzb', years)
    xjllb = fetch_f10_statement(code, 'xjllb', years)
    is_d = {k: [] for k in ('rev', 'cost', 'taxadd', 'sale', 'adm', 'rd', 'fin',
                            'othop', 'nonop', 'taxexp', 'np_p')}
    bs_d = {k: [] for k in ('cash', 'tfa', 'ar', 'pre', 'orec', 'inv', 'oca', 'ppe', 'rou',
                            'ia', 'gw', 'lpe', 'dta', 'oei', 'onca', 'stl', 'ap', 'contract',
                            'staff', 'taxp', 'opay', 'cur1y', 'ocl', 'ltl', 'lease', 'oncl',
                            'sc', 'cr', 'ts', 'oci', 'sr', 're', 'oeq', 'mi')}
    cf_d = {k: [] for k in ('np', 'da', 'ocf', 'capex', 'icf', 'fcf', 'cash1')}
    for y in sorted(years):
        L, B, Cf = lrb[y], zcfzb[y], xjllb[y]
        rev = _g(L, 'TOTAL_OPERATE_INCOME') or 0.0
        cost = _g(L, 'OPERATE_COST') or 0.0
        taxadd = _g(L, 'OPERATE_TAX_ADD') or 0.0
        sale = _g(L, 'SALE_EXPENSE') or 0.0
        adm = _g(L, 'MANAGE_EXPENSE') or 0.0
        rd = _g(L, 'RESEARCH_EXPENSE') or 0.0
        fin = _g(L, 'FINANCE_EXPENSE') or 0.0
        op = _g(L, 'OPERATE_PROFIT') or 0.0
        tp = _g(L, 'TOTAL_PROFIT') or 0.0
        is_d['rev'].append(_m(rev)); is_d['cost'].append(_m(cost))
        is_d['taxadd'].append(_m(taxadd)); is_d['sale'].append(_m(sale))
        is_d['adm'].append(_m(adm)); is_d['rd'].append(_m(rd))
        is_d['fin'].append(_m(fin))
        is_d['othop'].append(_m(op - (rev - cost - taxadd - sale - adm - rd - fin)))  # 轧差
        is_d['nonop'].append(_m(tp - op))
        is_d['taxexp'].append(_m(_g(L, 'INCOME_TAX') or 0.0))
        is_d['np_p'].append(_m(_g(L, 'PARENT_NETPROFIT') or 0.0))

        def b(key):
            return _g(B, key) or 0.0
        bs_d['cash'].append(_m(b('MONETARYFUNDS')))
        bs_d['tfa'].append(_m(b('TRADE_FINASSET_NOTFVTPL')))
        bs_d['ar'].append(_m(b('NOTE_ACCOUNTS_RECE')))
        bs_d['pre'].append(_m(b('PREPAYMENT')))
        bs_d['orec'].append(_m(b('TOTAL_OTHER_RECE')))
        bs_d['inv'].append(_m(b('INVENTORY')))
        tca = b('TOTAL_CURRENT_ASSETS')
        bs_d['oca'].append(_m(tca - (b('MONETARYFUNDS') + b('TRADE_FINASSET_NOTFVTPL')
                                     + b('NOTE_ACCOUNTS_RECE') + b('PREPAYMENT')
                                     + b('TOTAL_OTHER_RECE') + b('INVENTORY'))))  # 轧差
        ppe = b('FIXED_ASSET') + b('CIP')
        bs_d['ppe'].append(_m(ppe))
        bs_d['rou'].append(_m(b('USERIGHT_ASSET')))
        bs_d['ia'].append(_m(b('INTANGIBLE_ASSET')))
        bs_d['gw'].append(_m(b('GOODWILL')))
        bs_d['lpe'].append(_m(b('LONG_PREPAID_EXPENSE')))
        bs_d['dta'].append(_m(b('DEFER_TAX_ASSET')))
        bs_d['oei'].append(_m(b('OTHER_EQUITY_INVEST')))
        tnca = b('TOTAL_NONCURRENT_ASSETS')
        bs_d['onca'].append(_m(tnca - (ppe + b('USERIGHT_ASSET') + b('INTANGIBLE_ASSET')
                                       + b('GOODWILL') + b('LONG_PREPAID_EXPENSE')
                                       + b('DEFER_TAX_ASSET') + b('OTHER_EQUITY_INVEST'))))  # 轧差
        bs_d['stl'].append(_m(b('SHORT_LOAN')))
        bs_d['ap'].append(_m(b('NOTE_ACCOUNTS_PAYABLE')))
        bs_d['contract'].append(_m(b('CONTRACT_LIAB')))
        bs_d['staff'].append(_m(b('STAFF_SALARY_PAYABLE')))
        bs_d['taxp'].append(_m(b('TAX_PAYABLE')))
        bs_d['opay'].append(_m(b('TOTAL_OTHER_PAYABLE')))
        bs_d['cur1y'].append(_m(b('NONCURRENT_LIAB_1YEAR')))
        tcl = b('TOTAL_CURRENT_LIAB')
        bs_d['ocl'].append(_m(tcl - (b('SHORT_LOAN') + b('NOTE_ACCOUNTS_PAYABLE')
                                     + b('CONTRACT_LIAB') + b('STAFF_SALARY_PAYABLE')
                                     + b('TAX_PAYABLE') + b('TOTAL_OTHER_PAYABLE')
                                     + b('NONCURRENT_LIAB_1YEAR'))))  # 轧差
        bs_d['ltl'].append(_m(b('LONG_LOAN')))
        bs_d['lease'].append(_m(b('LEASE_LIAB')))
        tncl = b('TOTAL_NONCURRENT_LIAB')
        bs_d['oncl'].append(_m(tncl - b('LONG_LOAN') - b('LEASE_LIAB')))  # 轧差
        bs_d['sc'].append(_m(b('SHARE_CAPITAL')))
        bs_d['cr'].append(_m(b('CAPITAL_RESERVE')))
        bs_d['ts'].append(_m(b('TREASURY_SHARES')))
        bs_d['oci'].append(_m(b('OTHER_COMPRE_INCOME')))
        bs_d['sr'].append(_m(b('SURPLUS_RESERVE')))
        bs_d['re'].append(_m(b('UNASSIGN_RPOFIT')))
        tpe = b('TOTAL_PARENT_EQUITY')
        bs_d['oeq'].append(_m(tpe - (b('SHARE_CAPITAL') + b('CAPITAL_RESERVE')
                                     - b('TREASURY_SHARES') + b('OTHER_COMPRE_INCOME')
                                     + b('SURPLUS_RESERVE') + b('UNASSIGN_RPOFIT'))))  # 轧差
        bs_d['mi'].append(_m(b('MINORITY_EQUITY')))

        cf_d['np'].append(is_d['np_p'][-1])
        da = (_g(Cf, 'FA_IR_DEPR') or 0.0) + (_g(Cf, 'IA_AMORTIZE') or 0.0) \
            + (_g(Cf, 'LPE_AMORTIZE') or 0.0)
        cf_d['da'].append(_m(da))
        cf_d['ocf'].append(_m(_g(Cf, 'NETCASH_OPERATE') or 0.0))
        cf_d['capex'].append(_m(-(_g(Cf, 'CONSTRUCT_LONG_ASSET') or 0.0)))
        cf_d['icf'].append(_m(_g(Cf, 'NETCASH_INVEST') or 0.0))
        cf_d['fcf'].append(_m(_g(Cf, 'NETCASH_FINANCE') or 0.0))
        cf_d['cash1'].append(bs_d['cash'][-1])   # 与BS口径一致
    ppe_split = {'fa_net': _m(_g(zcfzb[sorted(years)[-1]], 'FIXED_ASSET') or 0.0),
                 'cip': _m(_g(zcfzb[sorted(years)[-1]], 'CIP') or 0.0)}
    return {'is': is_d, 'bs': bs_d, 'cf': cf_d, 'ppe_split': ppe_split}


def build_segments(comp, years):
    """主营构成 → 分段配置 (兜底): 最近年报前3项(按RANK), 其余并'其他';
    跨年按RANK对齐(扛名称变更), '其中:'子项先行剔除防重复计数"""
    comp = {y: [it for it in items if not it['ITEM_NAME'].startswith('其中')]
            for y, items in comp.items()}
    y0 = max(y for y in comp if y <= max(years))
    items = comp[y0]
    top = items[:3]
    names = [it['ITEM_NAME'] for it in top]
    has_rest = len(items) > 3
    segs = []
    groups = list(enumerate(names))          # (rank_idx, name)
    if has_rest:
        groups.append((None, '其他'))
    for gi, (rank_idx, gname) in enumerate(groups):
        share, gm = [], []
        for y in sorted(years):
            yitems = comp.get(y) or []
            ytot = sum((_g(it, 'MAIN_BUSINESS_INCOME') or 0.0) for it in yitems)
            if rank_idx is None:
                topval = sum((_g(it, 'MAIN_BUSINESS_INCOME') or 0.0) for it in yitems[:3])
                sinc = max(ytot - topval, 0.0)
                sgm = None
            else:
                it = yitems[rank_idx] if len(yitems) > rank_idx else None
                sinc = _g(it, 'MAIN_BUSINESS_INCOME') if it else None
                sgm = _g(it, 'GROSS_RPOFIT_RATIO') if it else None
            share.append(round((sinc or 0.0) / ytot, 4) if ytot else 0.0)
            gm.append(round(sgm, 4) if sgm is not None else 0.30)
        # 最近一年分部YoY (RANK对齐)
        total_inc = sum((_g(it, 'MAIN_BUSINESS_INCOME') or 0.0) for it in items)
        if len(years) >= 2:
            yprev = comp.get(sorted(years)[-2]) or []
            ytot_p = sum((_g(it, 'MAIN_BUSINESS_INCOME') or 0.0) for it in yprev)
            if rank_idx is None:
                p = max(ytot_p - sum((_g(it, 'MAIN_BUSINESS_INCOME') or 0.0) for it in yprev[:3]), 0.0)
            else:
                p = _g(yprev[rank_idx], 'MAIN_BUSINESS_INCOME') if len(yprev) > rank_idx else 0.0
            cur = share[-1] * total_inc
            yoy = (cur / p - 1) if p else 0.10
        else:
            yoy = 0.10
        g0 = min(max(yoy, -0.15), 0.35)
        growth = [round(g0 * f, 4) for f in (1.0, 0.75, 0.55, 0.40, 0.30)]
        key = f'seg{gi + 1}' if rank_idx is not None else 'oth'
        segs.append({
            'key': key, 'name': gname, 'short': gname, 'driver': 'growth',
            'hist_share': share, 'hist_gm': gm,
            'share_basis': f'占比: 东财F10主营构成(按产品, {y0}年报)',
            'vol': {'values': growth,
                    'basis': f'自动推导: 最近年报分部收入YoY={yoy:+.1%}(截断[-15%,35%])逐年退坡×1/.75/.55/.40/.30; 建议按研究修正'},
            'gm': {'values': [gm[-1]] * 5,
                   'basis': f'自动推导: {y0}年报分部毛利率{gm[-1]:.1%}持平; 建议按研究修正'},
            'logic': f'兜底自动分段(东财F10主营构成按产品); {y0}年报收入占比{share[-1]:.1%}',
        })
    return segs


def _flat(v, nd=4):
    return [round(v, nd)] * 5


def derive_assumptions(hist, fcst_years):
    """从历史三表推导默认假设 (兜底模式)"""
    rev, cost = hist['is']['rev'], hist['is']['cost']
    h0 = len(rev) - 1
    eps = 1e-9
    tprofit = hist['is']['np_p'][-1] + hist['is']['taxexp'][-1]
    tax_eff = min(max(hist['is']['taxexp'][-1] / max(tprofit, eps), 0.05), 0.25)
    bs = hist['bs']
    dso = bs['ar'][-1] / max(rev[-1], eps) * 365
    dio = bs['inv'][-1] / max(cost[-1], eps) * 365
    dpo = bs['ap'][-1] / max(cost[-1], eps) * 365
    capex_rt = abs(hist['cf']['capex'][-1]) / max(rev[-1], eps)
    capex_path = [round(min(capex_rt, 0.35) * f, 4) for f in (1.0, 0.85, 0.70, 0.60, 0.50)]
    rep = lambda bal: [round(bal / 5, 1)] * 5
    return {
        'opex': {
            'tax_add_rate': {'values': _flat(hist['is']['taxadd'][-1] / max(rev[-1], eps)),
                             'basis': f"自动推导: 最近年报税金及附加率{hist['is']['taxadd'][-1]/max(rev[-1],eps):.2%}持平"},
            'sale_rate': {'values': _flat(hist['is']['sale'][-1] / max(rev[-1], eps)),
                          'basis': '自动推导: 最近年报销售费用率持平'},
            'adm_rate': {'values': _flat(hist['is']['adm'][-1] / max(rev[-1], eps)),
                         'basis': '自动推导: 最近年报管理费用率持平'},
            'rd_rate': {'values': _flat(hist['is']['rd'][-1] / max(rev[-1], eps)),
                        'basis': '自动推导: 最近年报研发费用率持平'},
            'oth_op': {'values': _flat(hist['is']['othop'][-1], 1),
                       'basis': '自动推导: 最近年报其他经营损益(轧差项)持平'},
            'nonop': {'values': _flat(hist['is']['nonop'][-1], 1),
                      'basis': '自动推导: 最近年报营业外收支净额持平'},
            'tax_rate': {'values': _flat(tax_eff),
                         'basis': f'自动推导: 最近年报有效税率{tax_eff:.1%}(截断[5%,25%])持平'},
            'oth_rate': {'value': round(hist['is']['fin'][-1] / max(rev[-1], eps), 4),
                         'basis': '自动推导: 最近年报财务费用率, 供Scenarios页简化净利率计算'},
        },
        'working_capital': {
            'dso': {'values': _flat(dso, 0), 'basis': f'自动推导: 最近年报DSO {dso:.0f}天持平'},
            'dio': {'values': _flat(dio, 0), 'basis': f'自动推导: 最近年报DIO {dio:.0f}天持平'},
            'dpo': {'values': _flat(dpo, 0), 'basis': f'自动推导: 最近年报DPO {dpo:.0f}天持平'},
            'pre_rate': {'values': _flat(bs['pre'][-1] / max(rev[-1], eps)), 'basis': '自动推导: 最近年报预付/收入持平'},
            'staff_rate': {'values': _flat(bs['staff'][-1] / max(rev[-1], eps)), 'basis': '自动推导: 最近年报应付职工薪酬/收入持平'},
            'taxp_rate': {'values': _flat(bs['taxp'][-1] / max(rev[-1], eps)), 'basis': '自动推导: 最近年报应交税费/收入持平'},
        },
        'capex': {
            'capex_rate': {'values': capex_path,
                           'basis': f'自动推导: 最近年报Capex/收入{capex_rt:.1%}(上限35%)逐年退坡×1/.85/.70/.60/.50'},
            'trans_rate': {'value': 0.75, 'basis': '通用默认: 当年转固=(期初在建+当年Capex)×75%'},
            'dep_rate': {'value': 0.12, 'basis': '通用默认: 固定资产净值口径12%'},
            'dep_new_rate': {'value': 0.04, 'basis': '通用默认: 转固按约半年投产计提'},
            'disp_rate': {'value': 0.01, 'basis': '通用默认: 处置率1%'},
            'amort_rate': {'value': 0.08, 'basis': '通用默认: 无形资产摊销率8%'},
        },
        'dividend': {
            'payout': {'values': [0.30] * 5, 'basis': '通用默认: 股利支付率30%(按上年归母)'},
            'surplus_rate': {'value': 0.10, 'basis': '通用默认: 法定盈余公积按归母10%计提'},
        },
        'financing': {
            'min_cash_pct': {'value': 0.02, 'basis': '通用默认: 最低现金=收入×2%'},
            'rep_st': {'values': [0.0] * 5, 'basis': '通用默认: 短借为revolver, 无计划还款'},
            'rep_cur': {'values': rep(bs['cur1y'][-1]), 'basis': f'自动推导: 最近年报末余额{bs["cur1y"][-1]:.0f}百万按1/5逐年摊还'},
            'rep_lt': {'values': rep(bs['ltl'][-1]), 'basis': f'自动推导: 最近年报末余额{bs["ltl"][-1]:.0f}百万按1/5逐年摊还'},
            'rep_lease': {'values': rep(bs['lease'][-1]), 'basis': f'自动推导: 最近年报末余额{bs["lease"][-1]:.1f}百万按1/5逐年摊还'},
            'rate_st': {'value': 0.030, 'basis': '通用默认: 一年期LPR附近'},
            'rate_cur': {'value': 0.034, 'basis': '通用默认: 参照综合借款成本'},
            'rate_lt': {'value': 0.037, 'basis': '通用默认: 项目贷略高'},
            'rate_lease': {'value': 0.045, 'basis': '通用默认: 租赁内含利率4-5%'},
            'cash_yield': {'value': 0.012, 'basis': '通用默认: 活期+短期存款1.2%'},
            'fin_oth': {'value': 30.0, 'basis': '通用默认: 手续费/汇兑等30百万'},
        },
    }


def build_fallback_config(code):
    """全自动兜底: 拉公开数据 + 推导假设 → 完整配置dict (港股走HKF10路径)"""
    code = str(code)
    if is_hk(code):
        return build_fallback_config_hk(code)
    quote = fetch_quote(code)
    today = datetime.date.today().isoformat()
    years = [2023, 2024, 2025]
    hist = build_hist(code, years)
    comp = fetch_composition(code)
    segs = build_segments(comp, years) if comp else []
    fcst = list(range(years[-1] + 1, years[-1] + 6))
    ass = derive_assumptions(hist, fcst)
    rev0, np0 = hist['is']['rev'][-1], hist['is']['np_p'][-1]
    gy = rev0 / hist['is']['rev'][-2] - 1 if len(hist['is']['rev']) >= 2 else 0.1
    g_path = [gy * f for f in (1.0, 0.75, 0.55)]
    cons_rev = [round(rev0 * (1 + gp), 1) for gp in g_path]
    npm0 = np0 / rev0 if rev0 else 0.1
    cons_np = [round(rv * npm0, 1) for rv in cons_rev]
    cfg = {
        'company': {'code': code, 'code_full': f'{code}.{em_code(code)[:2]}',
                    'name': quote['name'], 'currency_note': '人民币 / 百万元 (每股数据为元)'},
        'model': {'hist_years': years, 'fcst_years': fcst,
                  'valuation_date': today, 'build_date': today},
        'market': {
            'price': {'value': quote['price'], 'basis': f'{today}腾讯行情(qt.gtimg.cn)'},
            'shares': {'value': quote['shares_m'], 'basis': f'总市值{quote["mcap_yi"]:.1f}亿/现价(腾讯行情 {today})'},
            'pe_ttm': quote['pe_ttm'],
            'pe_ttm_basis': f'腾讯行情PE-TTM({today})',
        },
        'consensus': {
            'rev': {'values': cons_rev, 'basis': f'自动推导(无一致预期输入): 最近年报收入×(1+YoY{gy:+.1%})退坡; 建议手工填入gildata/聚源一致预期'},
            'np': {'values': cons_np, 'basis': f'自动推导(无一致预期输入): 收入×最近年报净利率{npm0:.1%}; 建议手工填入'},
        },
        'segments': segs,
        **ass,
        'wacc': {
            'rf': {'value': 0.02, 'basis': '通用默认: 10年期国债收益率约2.0%'},
            'erp': {'value': 0.055, 'basis': '通用默认: A股ERP中枢5.5%'},
            'srp': {'value': 0.0, 'basis': '通用默认: 不取个股溢价(避免与β/情景重复计价)'},
            'kd': {'value': 0.034, 'basis': '通用默认: ≈四档债务加权平均利率'},
            'tg': {'value': 0.03, 'basis': '通用默认: 长期名义增速中枢下沿'},
            'override': None,
            'beta_unlevered_input': 1.0,
            'beta_unlevered_basis': '兜底默认: 未配置可比公司, βu=1.0蓝色输入; 建议补relative_val.comps后自动取中位数',
        },
        'scenarios': {
            'bear': {'rev_adj': -0.15, 'npm_adj': -0.03, 'logic': '通用默认: 分部增速-15pct, 净利率-3pct'},
            'base': {'rev_adj': 0.0, 'npm_adj': 0.0, 'logic': '基准: 不调整'},
            'bull': {'rev_adj': 0.10, 'npm_adj': 0.015, 'logic': '通用默认: 分部增速+10pct, 净利率+1.5pct'},
        },
        'hist': hist,
        'relative_val': {
            'target_pe_lo': {'value': 25.0, 'basis': '兜底默认: 25x(无可比清单); 建议补comps后取可比中位数为上界'},
            'comps': [],
        },
        'cover': {'subtitle': f'{quote["name"]} | 自动兜底建模(东财F10+腾讯行情)'},
    }
    return cfg


def apply_fallback(raw, code):
    """就地补全缺失配置的兜底入口 (供build_model.py调用)"""
    auto = build_fallback_config(code)
    for k, v in auto.items():
        raw.setdefault(k, v)
    raw.setdefault('hist', auto['hist'])
    if not raw.get('segments'):
        raw['segments'] = auto['segments']


def main():
    ap = argparse.ArgumentParser(description='A股/港股模型配置自动获取 (东财F10·HKF10/腾讯行情, 系统curl)')
    ap.add_argument('--code', required=True, help='6位A股代码 或 5位港股代码(如 00981)')
    ap.add_argument('--out', help='配置输出路径 (默认 configs/<code>.yaml)')
    ap.add_argument('--stdout', action='store_true', help='打印到stdout而不写文件')
    ap.add_argument('--quote-only', action='store_true', help='只打印行情快照')
    args = ap.parse_args()
    if args.quote_only:
        print(json.dumps(fetch_quote(args.code), ensure_ascii=False, indent=1))
        return
    cfg = build_fallback_config(args.code)
    import yaml
    text = ('# 自动生成(fetch_data.py兜底模式) — 所有假设的"依据"字段注明推导规则, '
            '请按研究修正; 一致预期/可比公司建议手工补充\n'
            + yaml.safe_dump(cfg, allow_unicode=True, sort_keys=False, width=120))
    if args.stdout:
        print(text)
        return
    out = args.out or os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                   'configs', f'{args.code}.yaml')
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, 'w', encoding='utf-8') as f:
        f.write(text)
    print('SAVED', out)
    print(json.dumps({'name': cfg['company']['name'], 'price': cfg['market']['price']['value'],
                      'shares_m': cfg['market']['shares']['value'],
                      'segments': [s['name'] for s in cfg['segments']]}, ensure_ascii=False))


if __name__ == '__main__':
    main()
