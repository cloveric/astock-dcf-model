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
  - 一致预期: 最近年报收入按YoY×1/.75/.55退坡逐年复利, 净利率持平外推占位(依据注明"自动推导");
  - 可比公司: 兜底默认空清单, WACC的βu转为蓝色输入(默认1.0), 建议手工补 comps。

中报/季报锚定 (拐点校准, 兜底增强):
  - A股兜底生成时经 fetch_interim 抓当年最新一期已披露中报/季报(lrbAjaxNew, 报告期累计),
    存在则写入配置 interim 段(anchor: true), 并按其年化实绩校准首年收入增速(clamp[-50%,300%],
    g1>60%视为拐点爆发走加速退坡)、首年毛利率(营业成本可得时, 5年滑向周期中枢)与兜底一致预期;
  - 港股尽力而为: 复用收益表长表已拉回的中期行(不加请求)做同样的收入锚定;
  - 配置无 interim 段(含全部手工配置)→ 行为与旧版完全一致, 输出逐字节不变。
"""
import argparse
import datetime
import hashlib
import json
import math
import os
import re
import subprocess
import sys
import threading
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

UA = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36'
EM_BASE = 'https://emweb.securities.eastmoney.com/PC_HSF10'
_FETCH_MANIFEST = []
_FETCH_MANIFEST_LOCK = threading.Lock()
_RAW_SNAPSHOT_DIR = None


def _public_path_reference(path):
    """Return a project-relative path or basename for logs and manifests."""
    resolved = Path(path).resolve()
    project_root = Path(__file__).resolve().parent
    try:
        return resolved.relative_to(project_root).as_posix()
    except ValueError:
        return resolved.name


def _utc_timestamp():
    """Return an RFC 3339 UTC timestamp for source-lineage records."""
    return datetime.datetime.now(datetime.timezone.utc).isoformat().replace('+00:00', 'Z')


def reset_fetch_manifest(raw_dir=None):
    """Clear lineage and optionally enable content-addressed raw snapshots."""
    global _RAW_SNAPSHOT_DIR
    with _FETCH_MANIFEST_LOCK:
        _FETCH_MANIFEST.clear()
        _RAW_SNAPSHOT_DIR = Path(raw_dir).resolve() if raw_dir is not None else None
        if _RAW_SNAPSHOT_DIR is not None:
            _RAW_SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)


def get_fetch_manifest():
    """Return a defensive copy of successful public-data request lineage."""
    with _FETCH_MANIFEST_LOCK:
        return [dict(item) for item in _FETCH_MANIFEST]


def curl_get(url, referer='https://emweb.securities.eastmoney.com/', timeout=30):
    """系统curl取数，并记录可审计的时间、响应大小与内容哈希。"""
    cmd = ['curl', '-s', '--compressed', '--max-time', str(timeout),
           '-H', f'User-Agent: {UA}', '-H', f'Referer: {referer}', url]
    out = subprocess.run(cmd, check=True, capture_output=True, timeout=timeout + 10)
    record = {
        'url': url,
        'fetched_at': _utc_timestamp(),
        'bytes': len(out.stdout),
        'sha256': hashlib.sha256(out.stdout).hexdigest(),
    }
    with _FETCH_MANIFEST_LOCK:
        if _RAW_SNAPSHOT_DIR is not None:
            url_hash = hashlib.sha256(url.encode('utf-8')).hexdigest()[:12]
            snapshot = _RAW_SNAPSHOT_DIR / f'{url_hash}_{record["sha256"]}.bin'
            try:
                with snapshot.open('xb') as fh:
                    fh.write(out.stdout)
            except FileExistsError:
                if hashlib.sha256(snapshot.read_bytes()).hexdigest() != record['sha256']:
                    raise RuntimeError(f'原始取数快照哈希冲突: {snapshot}')
            except OSError as exc:
                raise RuntimeError(f'保存原始取数快照失败: {snapshot}: {exc}') from exc
            project_root = Path(__file__).resolve().parent
            try:
                snapshot_ref = snapshot.relative_to(project_root)
                snapshot_scope = 'project'
                snapshot_root_id = 'project'
            except ValueError:
                snapshot_ref = Path(snapshot.name)
                snapshot_scope = 'raw_dir'
                snapshot_root_id = hashlib.sha256(
                    str(_RAW_SNAPSHOT_DIR).encode('utf-8')
                ).hexdigest()[:12]
            record['snapshot'] = snapshot_ref.as_posix()
            record['snapshot_scope'] = snapshot_scope
            record['snapshot_root_id'] = snapshot_root_id
            record['snapshot_local_only'] = True
            record['snapshot_committed'] = False
        _FETCH_MANIFEST.append(record)
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


def default_hist_years(today=None):
    """按当前月份推最新年报年: 年报披露截止4月30日 → 5月起最新年报=上年, 否则=前年;
    返回近3个年报年 [y1-2, y1-1, y1]"""
    today = today or datetime.date.today()
    y1 = today.year - 1 if today.month >= 5 else today.year - 2
    return [y1 - 2, y1 - 1, y1]


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
    if price <= 0:
        raise RuntimeError(f'腾讯行情现价为0({code}停牌或无有效行情), 拒绝兜底建模; 请复盘后重试或手工配置')
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
    payload = json.loads(curl_get(url))
    data = payload.get('data') if isinstance(payload, dict) else None
    if not data:
        # D5: companyType=4为一般工商业模板, 银行/券商/保险等金融类返回data=None
        raise RuntimeError('东财F10返回空数据(companyType=4 模板不适用,'
                           '疑似银行/券商/保险等特殊报表模板),暂不支持该类标的')
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
    payload = json.loads(curl_get(url))
    data = payload.get('zygcfx') if isinstance(payload, dict) else None
    if data is None:
        # D5: 金融类等特殊模板无zygcfx节点, 显式报错而非裸TypeError/KeyError
        raise RuntimeError('东财F10返回空数据(companyType=4 模板不适用,'
                           '疑似银行/券商/保险等特殊报表模板),暂不支持该类标的')
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


def fetch_interim(code, last_annual_year, status_out=None):
    """A股中报/季报抓取 (拐点锚定用): 依次尝试当年(=last_annual_year+1)的
    09-30/06-30/03-31利润表, 取最新已披露的一期, 并带上年同期做对照.
    lrbAjaxNew 的 dates 参数支持多期并列且只返回已存在的报告期(实测688825:
    dates带2026-09-30/06-30/03-31只回03-31), 故当年3期+上年同期3期合并为1次请求.
    科目: 营收=TOTAL_OPERATE_INCOME, 归母=PARENT_NETPROFIT, 净利润=NETPROFIT,
    营业成本=OPERATE_COST(实测688825中期响应该字段即"营业成本"105.9亿;
    TOTAL_OPERATE_COST为"营业总成本"148.2亿含期间费用, 不取).
    返回 dict {label('2026Q1'/'2026H1'/'2026Q3'), date, months(3/6/9), rev, np_p, np,
    cost(可None), rev_prior, np_p_prior} — 单位统一百万(复用_m口径), 上年同期缺失时
    prior字段为None; 当年任何一期都不存在→返回None(正常, 不报错).
    锚定属增强路径: 无披露返回None并标记not_available；网络/解析异常也不阻断主流程，
    但会标记error并输出告警，避免把“取数失败”伪装成“没有中期报告”."""
    status = {'source': 'eastmoney_interim'}
    if status_out is not None:
        status_out.clear()
        status_out.update(status)
    yr = int(last_annual_year) + 1
    mmdd = ('09-30', '06-30', '03-31')
    dates = [f'{yr}-{md}' for md in mmdd] + [f'{yr - 1}-{md}' for md in mmdd]
    url = (f'{EM_BASE}/NewFinanceAnalysis/lrbAjaxNew?companyType=4&reportDateType=0'
           f'&reportType=1&dates={"%2C".join(dates)}&code={em_code(code)}')
    try:
        payload = json.loads(curl_get(url))
        data = payload.get('data') if isinstance(payload, dict) else None
    except Exception as exc:
        status.update({'status': 'error', 'error': f'{type(exc).__name__}: {exc}'})
        if status_out is not None:
            status_out.update(status)
        print(f'WARNING: 中期数据抓取失败({code}): {exc}', file=sys.stderr)
        return None
    if not data:
        status['status'] = 'not_available'
        if status_out is not None:
            status_out.update(status)
        return None
    rows = {str(r.get('REPORT_DATE', ''))[:10]: r for r in data}
    for md, months, qlab in (('09-30', 9, 'Q3'), ('06-30', 6, 'H1'), ('03-31', 3, 'Q1')):
        row = rows.get(f'{yr}-{md}')
        if row is None:
            continue
        rev = _g(row, 'TOTAL_OPERATE_INCOME')
        if not rev or rev <= 0:
            continue                    # 营收缺失/为0无法年化 → 视同该期不可用, 试更早一期
        cost = _g(row, 'OPERATE_COST')  # "营业成本"口径 (非TOTAL_OPERATE_COST营业总成本)
        prow = rows.get(f'{yr - 1}-{md}') or {}
        rev_p = _g(prow, 'TOTAL_OPERATE_INCOME')
        np_p_p = _g(prow, 'PARENT_NETPROFIT')
        result = {
            'label': f'{yr}{qlab}', 'date': f'{yr}-{md}', 'months': months,
            'rev': _m(rev),
            'np_p': _m(_g(row, 'PARENT_NETPROFIT') or 0.0),
            'np': _m(_g(row, 'NETPROFIT') or 0.0),
            'cost': _m(cost) if cost is not None else None,
            'rev_prior': _m(rev_p) if rev_p is not None else None,
            'np_p_prior': _m(np_p_p) if np_p_p is not None else None,
        }
        status.update({'status': 'ok', 'report_date': result['date']})
        if status_out is not None:
            status_out.update(status)
        return result
    status['status'] = 'not_available'
    if status_out is not None:
        status_out.update(status)
    return None


# ============================================================
# 共享推导/兜底 helper (A股与港股两条路径共用, 消除逐字重复块)
# ============================================================

def _parallel(jobs, max_workers=5):
    """并发执行无参job列表, 按原顺序返回结果 (ThreadPoolExecutor包curl_get子进程调用,
    curl_get签名不变仅编排并发; 兜底整套抓取由串行5次curl变并发)"""
    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        futs = [ex.submit(j) for j in jobs]
        return [f.result() for f in futs]


def _derive_growth_path(rev_hist):
    """共享增长路径推导: 最近两年营收→YoY(前一年为0/缺失→默认0.10起步),
    截断[-15%,35%]后按×1/.75/.55/.40/.30逐年退坡; 返回 (yoy, [5年增速])"""
    if len(rev_hist) >= 2 and rev_hist[-2]:
        gy = rev_hist[-1] / rev_hist[-2] - 1
    else:
        gy = 0.10
    g0 = min(max(gy, -0.15), 0.35)
    return gy, [round(g0 * f, 4) for f in (1.0, 0.75, 0.55, 0.40, 0.30)]


def _interim_anchor(interim, rev0, f0_year):
    """中期实绩收入锚定 (拐点校准): 仅当 interim 年份==首预测年时生效.
    年化营收 rev_f0_anchor = rev×12/months; g1 = clamp(年化/最近年报营收−1, [−0.5, 3.0]);
    g1>0.6 视为拐点爆发→加速退坡×1/.35/.20/.12/.08, 否则沿用×1/.75/.55/.40/.30;
    多分部时各分部同一factor路径(总收入口径, 首年合计增速==g1).
    锚定条件不满足(无interim/年份不匹配/中期或年报营收缺失)→返回None, 调用方走原有逻辑."""
    if not interim or not rev0 or rev0 <= 0:
        return None
    try:
        year = int(str(interim.get('date', ''))[:4])
    except (TypeError, ValueError):
        return None
    months, rev = interim.get('months'), interim.get('rev')
    if year != f0_year or months not in (3, 6, 9) or not rev or rev <= 0:
        return None
    rev_annual = rev * 12.0 / months
    g1 = min(max(rev_annual / rev0 - 1.0, -0.5), 3.0)
    if g1 > 0.6:
        # 拐点爆发(多为周期涨价年化): 加速退坡之外, 第2年起再加绝对封顶——
        # 价格尖峰的年化增速若按比例复利外推, 会推出脱离行业容量的收入(如DRAM周期)
        factors, caps = (1.0, 0.35, 0.20, 0.12, 0.08), (3.0, 0.25, 0.15, 0.10, 0.08)
        growth = [round(min(g1 * f, c), 4) for f, c in zip(factors, caps)]
        decay = '爆发期加速退坡×1/.35/.20/.12/.08, 第2年起封顶25%/15%/10%/8%(周期年化不复利外推)'
    else:
        growth = [round(g1 * f, 4) for f in (1.0, 0.75, 0.55, 0.40, 0.30)]
        decay = '逐年退坡×1/.75/.55/.40/.30'
    label = interim.get('label') or '中期'
    return {
        'label': label, 'g1': g1, 'rev_annual': rev_annual,
        'growth': growth,
        'vol_basis': (f'按{label}实绩年化校准: 年化营收≈{rev_annual:.0f}百万, '
                      f'首年增速g1={g1:+.1%}(clamp[-50%,300%]), {decay}'),
    }


def _interim_gm_path(interim, hist_gm):
    """中期实绩毛利率锚定: 年1 = gm_interim = 1−营业成本/营收(中期实际),
    5年线性滑向 (gm_interim+该分部历史毛利率均值)/2 的周期中枢;
    cost不可得→返回None(毛利率沿用现有历史口径)."""
    cost, rev = interim.get('cost'), interim.get('rev')
    if cost is None or not rev or rev <= 0:
        return None
    gm_i = 1.0 - cost / rev
    if not math.isfinite(gm_i) or not -1.0 <= gm_i <= 1.0:
        raise ValueError(
            f'{interim.get("label") or "中期"}毛利率{gm_i:.1%}超出[-100%,100%]，'
            '疑似营业成本字段映射或单位错误，拒绝用于预测锚定')
    mid = (gm_i + sum(hist_gm) / len(hist_gm)) / 2.0
    vals = [round(gm_i + (mid - gm_i) * i / 4.0, 4) for i in range(5)]
    basis = (f'首年按{interim.get("label") or "中期"}实际毛利率{gm_i:.1%}, '
             f'5年滑向周期中枢{mid:.1%}(自动推导: (中期毛利率+历史毛利率均值)/2)')
    return vals, basis


def _guard_hk_residual(name, value, total, year, abs_tol=1.0, rel_tol=1e-4):
    """Reject material negative HKF10 residual buckets.

    A tiny negative caused by source rounding is snapped to zero by the caller.
    A material negative means mapped components exceed the reported subtotal and
    must be investigated instead of being labelled as an ordinary balancing item.
    """
    if not all(math.isfinite(float(v)) for v in (value, total)):
        raise RuntimeError(f'HKF10 {year}年{name}轧差不是有限数，拒绝建模')
    tolerance = max(abs_tol, abs(float(total)) * rel_tol)
    if value < -tolerance:
        ratio = abs(value) / max(abs(float(total)), 1.0)
        raise RuntimeError(
            f'HKF10 {year}年{name}出现重大负轧差: {value:.1f}，'
            f'占对应合计{ratio:.1%}；疑似核心科目缺失或映射错误，拒绝“假配平”')
    return value


def _derive_consensus(rev_hist, np_hist, fcst_years, interim=None):
    """共享兜底一致预期 (A4): 增速按YoY×1/.75/.55退坡且逐年复利推营收
    (rev[i]=rev[i-1]×(1+g_path[i]), 不再全部从rev0单利外推);
    净利按最近年净利率×对应年营收; 前一年营收0/缺失→gy=0.10, 最近年营收0→npm=0.10.
    interim(可选): 中期实绩锚定成立时, 营收路径改为与分部收入锚定后的总收入一致
    (rev0沿g1×factors逐年复利), 归母按npm滑动推(年1 npm=中期归母/营收, 在预测区间内
    线性滑向(中期npm+最近年报npm)/2), 依据注明校准来源; 无interim时签名与行为不变.
    返回 {'rev': {'values': [...], 'basis': ...}, 'np': {'values': [...], 'basis': ...}}"""
    rev0 = rev_hist[-1] if rev_hist else 0.0
    np0 = np_hist[-1] if np_hist else 0.0
    anchor = (_interim_anchor(interim, rev0, fcst_years[0] if fcst_years else None)
              if interim else None)
    if anchor is not None:
        n = min(3, len(fcst_years)) if fcst_years else 3
        cons_rev, prev = [], rev0
        for g in anchor['growth'][:n]:
            prev = prev * (1.0 + g)               # 与分部锚定路径同口径逐年复利
            cons_rev.append(round(prev, 1))
        npm0 = np0 / rev0 if rev0 else 0.10
        npm_i = (interim['np_p'] / interim['rev']
                 if interim.get('np_p') is not None else npm0)   # 归母缺失→退回年报npm平推
        npm_mid = (npm_i + npm0) / 2.0
        npms = [npm_i + (npm_mid - npm_i) * (i / (n - 1.0) if n > 1 else 0.0)
                for i in range(n)]
        cons_np = [round(rv * m, 1) for rv, m in zip(cons_rev, npms)]
        lab = anchor['label']
        return {
            'rev': {'values': cons_rev,
                    'basis': (f'自动推导(按{lab}实绩年化校准): 首年增速g1={anchor["g1"]:+.1%}'
                              '(clamp[-50%,300%]), 营收路径与分部收入锚定一致; '
                              '建议手工填入gildata/聚源一致预期')},
            'np': {'values': cons_np,
                   'basis': (f'自动推导(按{lab}实绩校准): 年1净利率{npm_i:.1%}(={lab}归母/营收), '
                             f'线性滑向(中期npm+最近年报npm{npm0:.1%})/2={npm_mid:.1%}; '
                             '建议手工填入')},
        }
    if len(rev_hist) >= 2 and rev_hist[-2]:
        gy = rev_hist[-1] / rev_hist[-2] - 1
    else:
        gy = 0.10
    factors = (1.0, 0.75, 0.55)
    n = min(len(factors), len(fcst_years)) if fcst_years else len(factors)
    cons_rev, prev = [], rev0
    for f in factors[:n]:
        prev = prev * (1.0 + gy * f)              # 逐年复利
        cons_rev.append(round(prev, 1))
    npm0 = np0 / rev0 if rev0 else 0.10
    cons_np = [round(rv * npm0, 1) for rv in cons_rev]
    return {
        'rev': {'values': cons_rev,
                'basis': f'自动推导(无一致预期输入): 最近年报收入按YoY{gy:+.1%}×1/.75/.55退坡逐年复利; 建议手工填入gildata/聚源一致预期'},
        'np': {'values': cons_np,
               'basis': f'自动推导(无一致预期输入): 逐年复利收入×最近年报净利率{npm0:.1%}; 建议手工填入'},
    }


def _common_defaults(rf, rf_basis):
    """共享wacc/scenarios/relative_val兜底默认块 (A股与港股仅无风险利率入参不同)"""
    return {
        'wacc': {
            'rf': {'value': rf, 'basis': rf_basis},
            'erp': {'value': 0.055, 'basis': '通用默认: ERP中枢5.5%; 请按市场口径修正'},
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
        'relative_val': {
            'target_pe_lo': {'value': 25.0, 'basis': '兜底默认: 25x(无可比清单); 建议补comps后取可比中位数为上界'},
            'comps': [],
        },
    }


def _fallback_single_segment(hist, years, interim=None):
    """D4: 无按产品主营构成披露(或全部无法构成分部)时, 合成整体单分部兜底
    (字段结构与HK'整体'分部一致: hist_share全1.0, 毛利率=整体历史毛利率, 增速=总营收YoY退坡);
    interim(可选): 中期实绩锚定成立时增速/毛利率路径按_interim_anchor/_interim_gm_path
    校准(拐点型标的不再被年报锚死), 无interim→与旧版逐字节一致"""
    rev, cost = hist['is']['rev'], hist['is']['cost']
    gm = []
    for i in range(len(years)):
        g = (rev[i] - cost[i]) / rev[i] if rev[i] else 0.30
        gm.append(round(g, 4))
    gy, growth = _derive_growth_path(rev)
    vol_basis = f'自动推导: 最近年报总收入YoY={gy:+.1%}(截断[-15%,35%])逐年退坡×1/.75/.55/.40/.30; 建议按研究修正'
    gm_vals = [gm[-1]] * 5
    gm_basis = f'自动推导: 最近年报综合毛利率{gm[-1]:.1%}持平; 建议按研究修正'
    anchor = _interim_anchor(interim, rev[-1] if rev else 0.0, max(years) + 1) if interim else None
    if anchor is not None:
        growth, vol_basis = anchor['growth'], anchor['vol_basis']
        gmp = _interim_gm_path(interim, gm)
        if gmp is not None:
            gm_vals, gm_basis = gmp
    return [{
        'key': 'main', 'name': '整体业务', 'short': '整体', 'driver': 'growth',
        'hist_share': [1.0] * len(years), 'hist_gm': gm,
        'share_basis': '无按产品构成披露, 按整体单分部建模(兜底); 单段=总收入100%',
        'vol': {'values': growth, 'basis': vol_basis},
        'gm': {'values': gm_vals, 'basis': gm_basis},
        'logic': '无按产品构成披露, 按整体单分部建模(兜底); 历史毛利率=1-营业成本/营业收入',
    }]


def _interim_cfg(interim, basis):
    """interim存在→生成配置interim段 {'interim': 实绩字段+basis依据+anchor标记};
    不存在→{} (配置不出现该段, 手工配置无该段时一切行为不变)"""
    if not interim:
        return {}
    seg = dict(interim)
    seg['basis'] = basis
    seg['anchor'] = True
    return {'interim': seg}


def _market_section(quote, today):
    """A股market段 (build_fallback_config与apply_fallback共用)"""
    return {
        'price': {'value': quote['price'], 'basis': f'{today}腾讯行情(qt.gtimg.cn)'},
        'shares': {'value': quote['shares_m'],
                   'basis': f'总市值{quote["mcap_yi"]:.1f}亿/现价(腾讯行情 {today})'},
        'pe_ttm': quote['pe_ttm'],
        'pe_ttm_basis': f'腾讯行情PE-TTM({today})',
    }


def _market_section_hk(quote, ccy, today):
    """港股market段 (A5): 按财报币种映射港元汇率, 保留原始港元现价price_hkd.
    接口约定: build_model同时见到market.price_hkd与market.fx_hkd时, 应按
    price = price_hkd / fx_hkd 重折算模型价 (price字段仅为生成时的折算快照)"""
    ccy = ccy or '未知'
    fx_map = {'美元': 7.80, '人民币': 1.085, '港元': 1.0}
    fx_hkd = fx_map.get(ccy, 1.0)
    if ccy == '美元':
        fx_basis = '手工输入: 港元/财报币种汇率(美元联系汇率区间7.75-7.85中枢); 请按实际更新'
    elif ccy == '人民币':
        fx_basis = '手工输入: 财报币种=人民币, 按HKD/CNY≈1.085折算; 请按实际汇率更新'
    elif ccy == '港元':
        fx_basis = '财报币种=港元, 与现价同币种, fx=1.0'
    else:
        fx_basis = f'未识别财报币种({ccy}), 暂按fx=1.0处理(待确认); 请按实际汇率更新'
    price_m = round(quote['price'] / fx_hkd, 4)      # 模型价(财报币种) = 港元价 / fx
    return {
        'price': {'value': price_m,
                  'basis': f'{today}腾讯行情港元现价{quote["price"]} / fx{fx_hkd}={price_m}{ccy}'},
        'price_hkd': {'value': round(quote['price'], 4),
                      'basis': f'{today}腾讯行情港元现价(原始口径); build_model按price_hkd/fx_hkd重折算模型价'},
        'shares': {'value': quote['shares_m'],
                   'basis': f'总市值{quote["mcap_yi"]:.1f}亿港元/港元现价(腾讯行情 {today}); 港股总市值=全部股本×港元价, 与A+H实际口径有差异'},
        'pe_ttm': quote['pe_ttm'],
        'pe_ttm_basis': f'腾讯行情PE-TTM(港元口径, {today})',
        'fx_hkd': {'value': fx_hkd, 'basis': fx_basis},
    }


def _guess_hk_ccy(raw):
    """从既有配置文本猜财报币种 (apply_fallback部分补全market时用; 猜不到返回None→按待确认处理).
    注意判断顺序: 港元现价字样几乎必现于currency_note, 故港元放最后"""
    comp = raw.get('company') or {}
    txt = str(comp.get('unit') or '') + str(comp.get('currency_note') or '')
    for name in ('美元', '人民币', '港元'):
        if name in txt:
            return name
    return None


# ============================================================
# 港股 (5位代码): 东财HKF10(datacenter API, 系统curl) + 腾讯hk行情
#   - 三表为 IFRS 科目、财报币种(如中芯国际=美元); 现金流量表为人民币口径,
#     按"CF期末现金/BS现金及等价物"的隐含汇率逐年折算回财报币种(依据注明);
#   - 无东财"主营构成(按产品)"披露 → 单段简化收入驱动;
#   - IFRS 无税金及附加/法定盈余公积等科目 → 置0并在依据注明。
# ============================================================

def fetch_hk_long_table(rpt, code, item_key, min_year=None, extra_out=None):
    """东财HKF10长表 → {year: {科目: 金额}} (仅年报 DATE_TYPE_CODE=001)
    分页拉全: 单页最多pageSize行, 00981资产负债表实测8页3641行, 只取第1页会丢最早年科目;
    D2: 记住第1页声明的pages总数, 后续页返回空/异常直接报错(疑似WAF), 不再静默break截断;
    min_year: 数据按REPORT_DATE倒序, 某页最旧一行年份<min_year时提前结束(正常早停, 省请求);
    extra_out(可选dict): 顺带收集非年报行(DATE_TYPE_CODE!='001', 即中报/季报)
    → {日期: {'DATE_TYPE_CODE':..., 'FISCAL_YEAR':..., 科目:金额}}; 倒序首页天然含最新
    中期行, 零额外请求、不改分页逻辑, 供HK中期锚定(_hk_interim)尽力而为使用"""
    out, ccy = {}, None
    page, total_pages = 1, None
    while True:
        url = (f'https://datacenter.eastmoney.com/securities/api/data/v1/get?reportName={rpt}'
               f'&columns=ALL&filter=(SECUCODE%3D%22{code}.HK%22)&pageNumber={page}&pageSize=500'
               '&sortTypes=-1&sortColumns=REPORT_DATE&source=HSF10&client=PC')
        res = json.loads(curl_get(url, referer='https://data.eastmoney.com/')).get('result') or {}
        data = res.get('data') or []
        if total_pages is None:
            total_pages = res.get('pages') or 1       # 记住第1页声明的总页数
        elif not data:
            raise RuntimeError(
                f'HKF10 {rpt} 分页中断: 第{page}页/共{total_pages}页返回空数据或格式异常, '
                '疑似东财WAF拦截; 拒绝静默截断(会丢早年科目), 请稍后重试或检查接口')
        for row in data:
            if row.get('DATE_TYPE_CODE') != '001':
                if extra_out is not None:             # 顺带收集中报/季报行(不加请求)
                    d = str(row.get('REPORT_DATE', ''))[:10]
                    rec = extra_out.setdefault(d, {
                        'DATE_TYPE_CODE': row.get('DATE_TYPE_CODE'),
                        'FISCAL_YEAR': row.get('FISCAL_YEAR')})
                    rec[row[item_key]] = row['AMOUNT']
                continue
            y = int(row['REPORT_DATE'][:4])
            out.setdefault(y, {})[row[item_key]] = row['AMOUNT']
            ccy = ccy or row.get('CURRENCY')
        if min_year is not None and data:
            oldest = int(data[-1]['REPORT_DATE'][:4])
            if oldest < min_year:
                break                                 # 早停: 倒序排列下后续页只会更旧
        if page >= total_pages or not data:
            break
        page += 1
    return out, ccy


def _hk_interim(extra, f0_year):
    """HK中期锚定尽力而为: 从收益表长表顺带收集的非年报行(extra_out)提取
    f0_year(首预测年)最新一期中报/季报, 结构与A股fetch_interim返回一致(单位百万, 财报币种).
    仅认: ① FISCAL_YEAR='12-31'(12月末财年, 与年报按自然年取数口径一致, 否则months
    换算不成立→跳过); ② 日期为03-31/06-30/09-30(东财HKF10金额为报告期累计口径,
    实测00981 2026-03-31 DATE_TYPE_CODE=003即Q1累计3个月, 06-30 002=H1累计6个月);
    ③ '营业额'存在且>0. cost取'销售成本'(与年报IFRS路径同科目, 可得则顺带毛利率锚定).
    全部不满足→返回None(退回纯年报锚定, 属正常)."""
    if not extra:
        return None
    month_map = {'03-31': (3, 'Q1'), '06-30': (6, 'H1'), '09-30': (9, 'Q3')}
    best = None
    for d, items in extra.items():
        d = str(d)
        if not d.startswith(str(f0_year)):
            continue
        mm = month_map.get(d[5:10])
        if mm is None:
            continue
        if items.get('FISCAL_YEAR') not in (None, '12-31'):
            continue                                  # 非12月末财年 → 不做锚定
        rev = items.get('营业额')
        if not isinstance(rev, (int, float)) or rev <= 0:
            continue
        if best is None or d > best[0]:
            best = (d, mm, items)
    if best is None:
        return None
    d, (months, qlab), items = best
    prior = extra.get(f'{f0_year - 1}-{d[5:10]}') or {}
    np_p, np_t, cost = items.get('股东应占溢利'), items.get('除税后溢利'), items.get('销售成本')
    rev_p, np_p_p = prior.get('营业额'), prior.get('股东应占溢利')
    return {
        'label': f'{f0_year}{qlab}', 'date': d, 'months': months,
        'rev': _m(items['营业额']),
        'np_p': _m(np_p) if isinstance(np_p, (int, float)) else None,
        'np': _m(np_t) if isinstance(np_t, (int, float)) else None,
        'cost': _m(cost) if isinstance(cost, (int, float)) else None,
        'rev_prior': _m(rev_p) if isinstance(rev_p, (int, float)) else None,
        'np_p_prior': _m(np_p_p) if isinstance(np_p_p, (int, float)) else None,
    }


def build_hist_hk(code, years, interim_out=None):
    """东财HKF10三表 → 模型 hist 结构 (单位: 财报币种百万, IFRS科目映射)
    interim_out(可选dict): 透传给收益表长表的extra_out, 顺带收集中报/季报行(零额外请求)"""
    min_y = min(years)
    inc, ccy = fetch_hk_long_table('RPT_HKF10_FN_INCOME', code, 'ITEM_NAME', min_year=min_y,
                                   extra_out=interim_out)
    bal, _ = fetch_hk_long_table('RPT_HKF10_FN_BALANCE', code, 'ITEM_NAME', min_year=min_y)
    cfs, _ = fetch_hk_long_table('RPT_HKF10_FN_CASHFLOW_PC', code, 'STD_ITEM_NAME',
                                 min_year=min_y)
    missing = [y for y in years if y not in inc or y not in bal or y not in cfs]
    if missing:
        raise RuntimeError(f'HKF10 缺少年报: {missing}')
    # D2b 核心科目守卫: 每个请求年份必须存在以下科目; li()/bi()的缺失回退0仅限次要科目
    core_req = (('利润表', inc, ('营业额',)),
                ('资产负债表', bal, ('股东权益', '流动资产合计')),
                ('现金流量表', cfs, ('期末现金',)))
    for stmt, table, keys in core_req:
        for y in years:
            miss = [k for k in keys if table[y].get(k) is None]
            if miss:
                raise RuntimeError(
                    f'HKF10 {stmt} {y}年报核心科目缺失: {miss}; 科目缺失, '
                    '疑似金融类报表模板或接口科目变更, 拒绝按0建模')

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
        is_d['rev'].append(_m(rev)); is_d['cost'].append(_m(cost))
        is_d['taxadd'].append(0.0)                    # IFRS无"税金及附加"科目
        is_d['sale'].append(_m(sale)); is_d['adm'].append(_m(adm)); is_d['rd'].append(_m(rd))
        is_d['fin'].append(_m(fin))
        is_d['othop'].append(_m(op - (rev - cost - sale - adm - rd - fin)))  # 轧差(含其他收入/应占联营)
        is_d['nonop'].append(_m(tp - op))
        is_d['taxexp'].append(_m(li('税项')))
        is_d['np_p'].append(_m(li('股东应占溢利')))

        cash_raw = B.get('现金及等价物')
        if not cash_raw:
            raise RuntimeError(
                f'HKF10资产负债表缺{y}年报"现金及等价物"科目(接口分页缺失或科目变更), '
                '无法折算CF币种, 拒绝静默按0处理')
        cash = cash_raw
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
        oca = _guard_hk_residual('oca', tca - cash - tfa - ar - pre - inv, tca, y)
        onca = _guard_hk_residual('onca', tnca - ppe - ia - dta - oei, tnca, y)
        ocl = _guard_hk_residual('ocl', tcl - stl - ap - contract - taxp - cur1y, tcl, y)
        oncl = _guard_hk_residual('oncl', tncl - ltl - lease, tncl, y)
        oeq = _guard_hk_residual('oeq', tpe - sc - cr - oci - re_, tpe, y)
        bs_d['cash'].append(_m(cash)); bs_d['tfa'].append(_m(tfa))
        bs_d['ar'].append(_m(ar)); bs_d['pre'].append(_m(pre)); bs_d['orec'].append(0.0)
        bs_d['inv'].append(_m(inv))
        bs_d['oca'].append(_m(oca))                                  # 轧差(负数受守卫)
        bs_d['ppe'].append(_m(ppe)); bs_d['rou'].append(0.0)         # IFRS使用权资产未单列
        bs_d['ia'].append(_m(ia)); bs_d['gw'].append(0.0); bs_d['lpe'].append(0.0)
        bs_d['dta'].append(_m(dta)); bs_d['oei'].append(_m(oei))
        bs_d['onca'].append(_m(onca))                                # 轧差(含联营公司权益)
        bs_d['stl'].append(_m(stl)); bs_d['ap'].append(_m(ap))
        bs_d['contract'].append(_m(contract)); bs_d['staff'].append(0.0)
        bs_d['taxp'].append(_m(taxp)); bs_d['opay'].append(0.0)
        bs_d['cur1y'].append(_m(cur1y))
        bs_d['ocl'].append(_m(ocl))                                  # 轧差
        bs_d['ltl'].append(_m(ltl)); bs_d['lease'].append(_m(lease))
        bs_d['oncl'].append(_m(oncl))                                # 轧差(含递延收入/递延税项负债)
        bs_d['sc'].append(_m(sc)); bs_d['cr'].append(_m(cr)); bs_d['ts'].append(0.0)
        bs_d['oci'].append(_m(oci)); bs_d['sr'].append(0.0)          # IFRS无法定盈余公积
        bs_d['re'].append(_m(re_))
        bs_d['oeq'].append(_m(oeq))                                  # 轧差
        bs_d['mi'].append(_m(bi('少数股东权益')))

        C = cfs[y]
        # 现金流表为人民币口径 → 按隐含汇率(CF期末现金/BS现金)折算回财报币种;
        # D1: 按CURRENCY分币种校验隐含比率区间, 不再只认fx≈1与[3,12](此前港元主体≈0.92被误拒)
        fx = (C.get('期末现金') or 0.0) / cash
        ccy_name = ccy or '美元'
        if ccy_name == '人民币':
            if 0.95 < fx < 1.05:
                fx = 1.0                          # 人民币主体: snap到1, 不折算
            else:
                raise RuntimeError(
                    f'HKF10 {y}年隐含汇率异常: 检测币种=人民币, CF期末现金/BS现金={fx:.4f} '
                    '落在(0.95,1.05)外(人民币主体应≈1), 拒绝静默回退导致该年CF口径错乱; '
                    '请检查接口数据或手工配置hist')
        elif ccy_name == '港元':
            if not (0.80 < fx < 1.00):            # RMB/HKD≈0.92, 按实测比率折算
                raise RuntimeError(
                    f'HKF10 {y}年隐含汇率异常: 检测币种=港元, CF期末现金/BS现金={fx:.4f} '
                    '落在(0.80,1.00)外(港元主体RMB/HKD应≈0.92), 拒绝静默回退; '
                    '请检查接口数据或手工配置hist')
        elif ccy_name == '美元':
            if not (6.0 < fx < 9.0):
                raise RuntimeError(
                    f'HKF10 {y}年隐含汇率异常: 检测币种=美元, CF期末现金/BS现金={fx:.4f} '
                    '落在(6.0,9.0)外(美元主体RMB/USD应≈6-9), '
                    '拒绝静默回退fx=1.0导致该年CF混入人民币口径; 请检查接口数据或手工配置hist')
        else:
            if not (0.5 < fx < 12.0):
                raise RuntimeError(
                    f'HKF10 {y}年隐含汇率异常: 检测币种={ccy_name}(非人民币/港元/美元的其他币种, '
                    f'仅做宽松校验), CF期末现金/BS现金={fx:.4f} 落在(0.5,12)兜底区间外, '
                    '拒绝静默回退; 请检查接口数据或手工配置hist')

        def ci(k):
            return (C.get(k) or 0.0) / fx

        cf_d['np'].append(is_d['np_p'][-1])
        cf_d['da'].append(_m(ci('加:折旧及摊销')))
        cf_d['ocf'].append(_m(ci('经营业务现金净额')))
        cf_d['capex'].append(_m(-(ci('购建固定资产') + ci('购建无形资产及其他资产'))))
        cf_d['icf'].append(_m(ci('投资业务现金净额')))
        cf_d['fcf'].append(_m(ci('融资业务现金净额')))
        cf_d['cash1'].append(_m(ci('期末现金')))      # = BS现金及等价物 (隐含汇率口径)
    y0 = sorted(years)[-1]
    ppe_split = {'fa_net': _m(bal[y0].get('物业厂房及设备')),
                 'cip': 0.0}                            # HKF10未拆分在建工程
    return {'is': is_d, 'bs': bs_d, 'cf': cf_d, 'ppe_split': ppe_split,
            'currency': ccy or '美元', 'cf_fx_note': '现金流量表(人民币口径)按隐含汇率折算'}


def build_segments_hk(hist, years, interim=None):
    """港股无东财主营构成披露 → 单段(整体)简化: 增速=总收入YoY退坡, 毛利率=历史实际;
    interim(可选): 中期实绩锚定成立时增速(及销售成本可得时的毛利率)按锚定路径校准,
    无interim→与旧版逐字节一致"""
    rev = hist['is']['rev']
    gm = []
    for i, y in enumerate(sorted(years)):
        g = (rev[i] - hist['is']['cost'][i]) / rev[i] if rev[i] else 0.30
        gm.append(round(g, 4))
    gy, growth = _derive_growth_path(rev)        # 共享增长路径推导(含rev[-2]==0守卫)
    vol_basis = f'自动推导: 最近年报总收入YoY={gy:+.1%}(截断[-15%,35%])逐年退坡×1/.75/.55/.40/.30; 建议按研究修正'
    gm_vals = [gm[-1]] * 5
    gm_basis = f'自动推导: 最近年报综合毛利率{gm[-1]:.1%}持平; 建议按研究修正'
    anchor = _interim_anchor(interim, rev[-1] if rev else 0.0, max(years) + 1) if interim else None
    if anchor is not None:
        growth, vol_basis = anchor['growth'], anchor['vol_basis']
        gmp = _interim_gm_path(interim, gm)
        if gmp is not None:
            gm_vals, gm_basis = gmp
    return [{
        'key': 'seg1', 'name': '整体业务(单段简化)', 'short': '整体', 'driver': 'growth',
        'hist_share': [1.0] * len(years), 'hist_gm': gm,
        'share_basis': '港股无东财主营构成(按产品)披露, 单段=总收入100%',
        'vol': {'values': growth, 'basis': vol_basis},
        'gm': {'values': gm_vals, 'basis': gm_basis},
        'logic': '港股单段简化(东财HKF10无主营构成披露); 历史毛利率=1-销售成本/营业额(IFRS)',
    }]


def build_fallback_config_hk(code, raw_dir=None, _manifest_initialized=False):
    """港股全自动兜底: 腾讯hk行情 + 东财HKF10三表(IFRS) → 完整配置dict
    中期锚定尽力而为: 复用收益表长表已拉回的中报/季报行(零额外请求)做收入锚定"""
    if not _manifest_initialized:
        reset_fetch_manifest(raw_dir=raw_dir)
    code = str(code)
    quote = fetch_quote(code)
    today = datetime.date.today().isoformat()
    years = default_hist_years()
    hk_extra = {}
    hist = build_hist_hk(code, years, interim_out=hk_extra)
    ccy = hist['currency']
    interim = _hk_interim(hk_extra, years[-1] + 1)
    segs = build_segments_hk(hist, years, interim=interim)
    fcst = list(range(years[-1] + 1, years[-1] + 6))
    ass = derive_assumptions(hist, fcst, interim=interim)
    ass['dividend']['surplus_rate'] = {'value': 0.0,
                                       'basis': 'IFRS无法定盈余公积科目, 计提率置0'}
    common = _common_defaults(0.04, '通用默认(港股): 10年期美债收益率约4%(财报币种口径)')
    interim_basis = (f'东财HKF10收益表(RPT_HKF10_FN_INCOME) {interim["date"]}中期实绩'
                     f'(财报币种{ccy}, 报告期累计); 用于兜底首年收入锚定(年化外推, 拐点校准)'
                     if interim else '')
    cfg = {
        'company': {'code': code, 'code_full': f'{code}.HK', 'name': quote['name'],
                    'currency_note': f'{ccy}(财报口径) / 百万元 (每股数据为{ccy}; 港元现价按fx折算)',
                    'unit': f'{ccy}百万元(财报口径)'},
        'model': {'hist_years': years, 'fcst_years': fcst,
                  'valuation_date': today, 'build_date': today},
        'market': _market_section_hk(quote, ccy, today),   # A5: 分币种fx映射+price_hkd
        'consensus': _derive_consensus(hist['is']['rev'], hist['is']['np_p'], fcst,
                                       interim=interim),
        **_interim_cfg(interim, interim_basis),
        'segments': segs,
        **ass,
        'wacc': common['wacc'],
        'scenarios': common['scenarios'],
        'hist': {k: v for k, v in hist.items() if k in ('is', 'bs', 'cf', 'ppe_split')},
        'relative_val': common['relative_val'],
        'data_lineage': {'generated_at_utc': _utc_timestamp(),
                         'raw_snapshots_committed': False,
                         'raw_snapshot_mode': ('local_only' if _RAW_SNAPSHOT_DIR else 'disabled'),
                         'requests': get_fetch_manifest(),
                         'interim': {'source': 'eastmoney_hkf10_income',
                                     'status': 'ok' if interim else 'not_available'}},
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


def _cashflow_da(row):
    """Return reported D&A, preserving an unavailable source value as None."""
    values = [_g(row, key) for key in ('FA_IR_DEPR', 'IA_AMORTIZE', 'LPE_AMORTIZE')]
    if all(value is None for value in values):
        return None
    return sum(value or 0.0 for value in values)


def build_hist(code, years, tables=None):
    """拉取并映射三表为模型 hist 结构 (单位: 百万)
    tables: 可选预取的(lrb, zcfzb, xjllb)三元组 (供并发抓取后复用, 避免重复请求)"""
    if tables is None:
        lrb = fetch_f10_statement(code, 'lrb', years)
        zcfzb = fetch_f10_statement(code, 'zcfzb', years)
        xjllb = fetch_f10_statement(code, 'xjllb', years)
    else:
        lrb, zcfzb, xjllb = tables
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
        da = _cashflow_da(Cf)
        cf_d['da'].append(None if da is None else _m(da))
        cf_d['ocf'].append(_m(_g(Cf, 'NETCASH_OPERATE') or 0.0))
        cf_d['capex'].append(_m(-(_g(Cf, 'CONSTRUCT_LONG_ASSET') or 0.0)))
        cf_d['icf'].append(_m(_g(Cf, 'NETCASH_INVEST') or 0.0))
        cf_d['fcf'].append(_m(_g(Cf, 'NETCASH_FINANCE') or 0.0))
        cf_d['cash1'].append(bs_d['cash'][-1])   # 与BS口径一致
    ppe_split = {'fa_net': _m(_g(zcfzb[sorted(years)[-1]], 'FIXED_ASSET') or 0.0),
                 'cip': _m(_g(zcfzb[sorted(years)[-1]], 'CIP') or 0.0)}
    return {'is': is_d, 'bs': bs_d, 'cf': cf_d, 'ppe_split': ppe_split}


def build_segments(comp, years, interim=None, rev0=None):
    """主营构成 → 分段配置 (兜底): 最近年报前3项(按RANK), 其余并'其他';
    跨年按RANK对齐(扛名称变更), '其中:'子项先行剔除防重复计数;
    D3: 某历史年缺构成时回退最近一个有构成年份的份额结构(不再记0),
    最新历史年份额合计仍为0时显式报错;
    interim/rev0(可选): 中期实绩锚定成立时(rev0=最近年报总营收, 供g1计算),
    各分部增速统一替换为锚定路径(同factor), 毛利率按_interim_gm_path滑动
    (整体中期毛利率×各分部自身历史均值中枢); 无interim→与旧版逐字节一致"""
    anchor = (_interim_anchor(interim, rev0, sorted(years)[-1] + 1)
              if interim else None)
    comp = {y: [it for it in items if not it['ITEM_NAME'].startswith('其中')]
            for y, items in comp.items()}
    avail = sorted(y for y in comp if comp[y])
    if not avail:
        return []          # 全部无法构成分部 → 由调用方走整体单分部兜底(D4)
    cand = [y for y in avail if y <= max(years)]
    y0 = max(cand) if cand else max(avail)
    items = comp[y0]
    # 各历史年实际使用的构成年份: 缺失年回退最近一个有构成的年份(优先更早的相邻年)
    year_src = {}
    for y in sorted(years):
        if comp.get(y):
            year_src[y] = y
        else:
            earlier = [ay for ay in avail if ay < y]
            year_src[y] = earlier[-1] if earlier else min(ay for ay in avail if ay > y)
    y_last = sorted(years)[-1]
    fb_note = (f'; {y_last}年报构成缺失, 份额回退{year_src[y_last]}年报结构'
               if year_src[y_last] != y_last else '')
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
            yitems = comp.get(year_src[y]) or []
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
        # 最近一年分部YoY (RANK对齐; 上年构成缺失时经_derive_growth_path回退默认0.10起步)
        total_inc = sum((_g(it, 'MAIN_BUSINESS_INCOME') or 0.0) for it in items)
        cur = share[-1] * total_inc
        if len(years) >= 2:
            yprev = comp.get(sorted(years)[-2]) or []
            ytot_p = sum((_g(it, 'MAIN_BUSINESS_INCOME') or 0.0) for it in yprev)
            if rank_idx is None:
                p = max(ytot_p - sum((_g(it, 'MAIN_BUSINESS_INCOME') or 0.0) for it in yprev[:3]), 0.0)
            else:
                p = (_g(yprev[rank_idx], 'MAIN_BUSINESS_INCOME') or 0.0) if len(yprev) > rank_idx else 0.0
            yoy, growth = _derive_growth_path([p, cur])   # 共享增长路径(含前值0守卫)
        else:
            yoy, growth = _derive_growth_path([cur])
        vol_basis = f'自动推导: 最近年报分部收入YoY={yoy:+.1%}(截断[-15%,35%])逐年退坡×1/.75/.55/.40/.30; 建议按研究修正'
        gm_vals = [gm[-1]] * 5
        gm_basis = f'自动推导: {y0}年报分部毛利率{gm[-1]:.1%}持平; 建议按研究修正'
        if anchor is not None:            # 中期锚定: 各分部同一factor路径(总收入口径)
            growth, vol_basis = anchor['growth'], anchor['vol_basis']
            gmp = _interim_gm_path(interim, gm)
            if gmp is not None:
                gm_vals, gm_basis = gmp
        key = f'seg{gi + 1}' if rank_idx is not None else 'oth'
        segs.append({
            'key': key, 'name': gname, 'short': gname, 'driver': 'growth',
            'hist_share': share, 'hist_gm': gm,
            'share_basis': f'占比: 东财F10主营构成(按产品, {y0}年报)' + fb_note,
            'vol': {'values': growth, 'basis': vol_basis},
            'gm': {'values': gm_vals, 'basis': gm_basis},
            'logic': f'兜底自动分段(东财F10主营构成按产品); {y0}年报收入占比{share[-1]:.1%}',
        })
    if not any(s['hist_share'][-1] for s in segs):
        raise RuntimeError(
            f'东财F10主营构成缺最新历史年({y_last})数据: 各分部最新年份额合计为0 '
            '(回退最近有构成年份后仍为0, 构成收入缺失或全为0), 拒绝按0份额建模; '
            '请手工配置segments或修正构成数据')
    return segs


def _flat(v, nd=4):
    return [round(v, nd)] * 5


def derive_assumptions(hist, fcst_years, interim=None):
    """从历史三表推导默认假设 (兜底模式); interim(可选)用于拐点锚定时的capex口径切换"""
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
    anchor = (_interim_anchor(interim, rev[-1], fcst_years[0] if fcst_years else None)
              if interim else None)
    if anchor:
        # 拐点锚定时capex改按上年绝对额×1.1/1.0/0.9/0.8/0.7缓降(再折算为收入占比):
        # 收入被中期年化推升后, 历史建设期的高capex/收入占比会外推出脱离实际的资本开支
        capex_abs = abs(hist['cf']['capex'][-1])
        rev_est, _r = [], rev[-1]
        for _g in anchor['growth']:
            _r *= (1.0 + _g)
            rev_est.append(_r)
        capex_path = [round(min(capex_abs * m / max(rev_est[i], eps), 0.35), 4)
                      for i, m in enumerate((1.1, 1.0, 0.9, 0.8, 0.7))]
        capex_basis = (f'自动推导: 拐点锚定期Capex按上年绝对额{capex_abs:.0f}百万'
                       '×1.1/1.0/0.9/0.8/0.7缓降折算为收入占比(上限35%), '
                       '避免收入年化后按历史占比外推')
    else:
        capex_path = [round(min(capex_rt, 0.35) * f, 4) for f in (1.0, 0.85, 0.70, 0.60, 0.50)]
        capex_basis = f'自动推导: 最近年报Capex/收入{capex_rt:.1%}(上限35%)逐年退坡×1/.85/.70/.60/.50'
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
            'capex_rate': {'values': capex_path, 'basis': capex_basis},
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


def build_fallback_config(code, raw_dir=None):
    """全自动兜底: 拉公开数据 + 推导假设 → 完整配置dict (港股走HKF10路径)"""
    reset_fetch_manifest(raw_dir=raw_dir)
    code = str(code)
    if is_hk(code):
        return build_fallback_config_hk(code, _manifest_initialized=True)
    today = datetime.date.today().isoformat()
    years = default_hist_years()
    # 效率: 行情+三表+主营构成+当年中期实绩共6个curl并发抓取
    interim_status = {}
    quote, lrb, zcfzb, xjllb, comp, interim = _parallel([
        lambda: fetch_quote(code),
        lambda: fetch_f10_statement(code, 'lrb', years),
        lambda: fetch_f10_statement(code, 'zcfzb', years),
        lambda: fetch_f10_statement(code, 'xjllb', years),
        lambda: fetch_composition(code),
        lambda: fetch_interim(code, years[-1], status_out=interim_status),
    ])
    hist = build_hist(code, years, tables=(lrb, zcfzb, xjllb))
    segs = (build_segments(comp, years, interim=interim, rev0=hist['is']['rev'][-1])
            if comp else [])
    if not segs:
        segs = _fallback_single_segment(hist, years, interim=interim)  # D4: 整体单分部
    fcst = list(range(years[-1] + 1, years[-1] + 6))
    ass = derive_assumptions(hist, fcst, interim=interim)
    common = _common_defaults(0.02, '通用默认: 10年期国债收益率约2.0%')
    interim_basis = (f'东财F10利润表(lrbAjaxNew) {interim["date"]}中期实绩(报告期累计); '
                     '用于兜底首年收入/毛利率/一致预期锚定(年化外推, 拐点校准)'
                     if interim else '')
    cfg = {
        'company': {'code': code, 'code_full': f'{code}.{em_code(code)[:2]}',
                    'name': quote['name'], 'currency_note': '人民币 / 百万元 (每股数据为元)'},
        'model': {'hist_years': years, 'fcst_years': fcst,
                  'valuation_date': today, 'build_date': today},
        'market': _market_section(quote, today),
        'consensus': _derive_consensus(hist['is']['rev'], hist['is']['np_p'], fcst,
                                       interim=interim),
        **_interim_cfg(interim, interim_basis),
        'segments': segs,
        **ass,
        'wacc': common['wacc'],
        'scenarios': common['scenarios'],
        'hist': hist,
        'relative_val': common['relative_val'],
        'data_lineage': {'generated_at_utc': _utc_timestamp(),
                         'raw_snapshots_committed': False,
                         'raw_snapshot_mode': ('local_only' if _RAW_SNAPSHOT_DIR else 'disabled'),
                         'requests': get_fetch_manifest(),
                         'interim': interim_status},
        'cover': {'subtitle': f'{quote["name"]} | 自动兜底建模(东财F10+腾讯行情)'},
    }
    return cfg


def apply_fallback(raw, code):
    """就地补全缺失配置的兜底入口 (供build_model.py调用)
    效率修复: 先检查缺什么再按需抓取, 不再无条件整套重拉 —
      - raw已完整(如刚由build_fallback_config生成) → 直接返回, 零网络请求;
      - hist/segments/market全缺(近似空配置) → 整套兜底一次(内部并发抓取);
      - 部分缺失 → 只抓对应数据源: 只缺segments仅拉主营构成, 只缺hist仅拉三表,
        缺market/company/cover才拉行情; 其余假设段由hist本地推导, 不发请求;
    interim: 配置已带interim段(如fetch_data生成或手工填)时, 推导segments/consensus
    沿用其中期锚定; 无该段→不额外抓取, 一切行为与旧版一致"""
    code = str(code)
    need_hist = not raw.get('hist')
    need_segs = not raw.get('segments')
    need_market = not raw.get('market')
    derived_keys = ('company', 'model', 'consensus', 'opex', 'working_capital', 'capex',
                    'dividend', 'financing', 'wacc', 'scenarios', 'relative_val', 'cover')
    need_derived = [k for k in derived_keys if k not in raw]
    if not (need_hist or need_segs or need_market or need_derived):
        return          # 配置完整 → 跳过 (修复build_fallback_config刚跑完又被整套重抓)
    if need_hist and need_segs and need_market:
        auto = build_fallback_config(code)             # 近似空配置 → 整套兜底一次
        for k, v in auto.items():
            raw.setdefault(k, v)
        if not raw.get('segments'):
            raw['segments'] = auto['segments']
        return
    hk = is_hk(code)
    today = datetime.date.today().isoformat()
    years = list((raw.get('model') or {}).get('hist_years') or default_hist_years())
    fcst = list((raw.get('model') or {}).get('fcst_years') or
                range(max(years) + 1, max(years) + 6))
    hist, hist_ccy = raw.get('hist'), None
    interim = raw.get('interim') if isinstance(raw.get('interim'), dict) else None
    if need_hist:                                      # 只缺hist → 仅抓三表
        if hk:
            full = build_hist_hk(code, years)
            hist_ccy = full.get('currency')
            hist = {k: v for k, v in full.items() if k in ('is', 'bs', 'cf', 'ppe_split')}
        else:
            hist = build_hist(code, years)
        raw['hist'] = hist
    if need_segs:                                      # 只缺segments → 仅拉主营构成
        if hk:
            raw['segments'] = build_segments_hk(hist, years, interim=interim)  # 无额外请求
        else:
            comp = fetch_composition(code)
            rev0 = hist['is']['rev'][-1] if hist.get('is') else None
            segs = build_segments(comp, years, interim=interim, rev0=rev0) if comp else []
            raw['segments'] = segs or _fallback_single_segment(hist, years, interim=interim)
    quote = None
    if need_market or 'company' in need_derived or 'cover' in need_derived:
        quote = fetch_quote(code)                      # 缺market/company/cover才抓行情
    if need_market:
        raw['market'] = (_market_section_hk(quote, hist_ccy or _guess_hk_ccy(raw), today)
                         if hk else _market_section(quote, today))
    if need_derived:                                   # 其余假设段由hist本地推导
        ass = derive_assumptions(hist, fcst, interim=raw.get('interim') if isinstance(raw.get('interim'), dict) else None)
        if hk:
            ass['dividend']['surplus_rate'] = {'value': 0.0,
                                               'basis': 'IFRS无法定盈余公积科目, 计提率置0'}
        for k in ('opex', 'working_capital', 'capex', 'dividend', 'financing'):
            raw.setdefault(k, ass[k])
        common = (_common_defaults(0.04, '通用默认(港股): 10年期美债收益率约4%(财报币种口径)')
                  if hk else _common_defaults(0.02, '通用默认: 10年期国债收益率约2.0%'))
        for k in ('wacc', 'scenarios', 'relative_val'):
            raw.setdefault(k, common[k])
        raw.setdefault('consensus',
                       _derive_consensus(hist['is']['rev'], hist['is']['np_p'], fcst,
                                         interim=interim))
        raw.setdefault('model', {'hist_years': years, 'fcst_years': fcst,
                                 'valuation_date': today, 'build_date': today})
        if quote is not None:
            if hk:
                ccy = hist_ccy or _guess_hk_ccy(raw) or '未知'
                raw.setdefault('company', {
                    'code': code, 'code_full': f'{code}.HK', 'name': quote['name'],
                    'currency_note': f'{ccy}(财报口径) / 百万元 (每股数据为{ccy}; 港元现价按fx折算)',
                    'unit': f'{ccy}百万元(财报口径)'})
                raw.setdefault('cover', {
                    'subtitle': f'{quote["name"]} | 港股自动兜底建模(东财HKF10+腾讯行情; 财报币种{ccy})'})
            else:
                raw.setdefault('company', {
                    'code': code, 'code_full': f'{code}.{em_code(code)[:2]}',
                    'name': quote['name'], 'currency_note': '人民币 / 百万元 (每股数据为元)'})
                raw.setdefault('cover', {
                    'subtitle': f'{quote["name"]} | 自动兜底建模(东财F10+腾讯行情)'})


def main():
    ap = argparse.ArgumentParser(description='A股/港股模型配置自动获取 (东财F10·HKF10/腾讯行情, 系统curl)')
    ap.add_argument('--code', required=True, help='6位A股代码 或 5位港股代码(如 00981)')
    ap.add_argument('--out', help='配置输出路径 (默认 configs/<code>.yaml)')
    ap.add_argument('--stdout', action='store_true', help='打印到stdout而不写文件')
    ap.add_argument('--quote-only', action='store_true', help='只打印行情快照')
    ap.add_argument('--raw-dir', help='原始响应不可变快照目录(默认 data/raw/<code>/<UTC时间>)')
    ap.add_argument('--no-raw-snapshot', action='store_true',
                    help='不保存原始响应；仍会在配置 data_lineage 中记录URL/时间/大小/SHA-256')
    args = ap.parse_args()
    if args.quote_only:
        print(json.dumps(fetch_quote(args.code), ensure_ascii=False, indent=1))
        return
    raw_dir = None
    if not args.no_raw_snapshot:
        raw_dir = args.raw_dir or os.path.join(
            os.path.dirname(os.path.abspath(__file__)), 'data', 'raw', str(args.code),
            _utc_timestamp().replace(':', '').replace('-', ''))
    cfg = build_fallback_config(args.code, raw_dir=raw_dir)
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
    print('SAVED', _public_path_reference(out))
    print(json.dumps({'name': cfg['company']['name'], 'price': cfg['market']['price']['value'],
                      'shares_m': cfg['market']['shares']['value'],
                      'segments': [s['name'] for s in cfg['segments']]}, ensure_ascii=False))


if __name__ == '__main__':
    main()
