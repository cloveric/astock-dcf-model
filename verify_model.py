# -*- coding: utf-8 -*-
"""
模型验收脚本 (通用版): LibreOffice headless 重算 + 逐项校验

用法:
  python verify_model.py --code 300476                 # 找 out/ 下xlsx+addr, 重算并校验
  python verify_model.py out/300476_胜宏科技_估值模型.xlsx
  python verify_model.py <xlsx> --no-recalc            # 直接读已重算文件
校验项: Checks布尔总闸 / BS配平 / 现金=最低现金 / FIN迭代残差 / CF-BS现金勾稽 /
        WACC采用值=计算值 / 年中折现t序 / 基准情景=主模型 / 敏感性矩阵中心=DCF /
        黄金值断言(基准标的DCF每股复现, 跨LO版本浮点容差 + config指纹门, 见GOLDEN_DCF_PS)。
"""
import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import time

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# B3 黄金值: 基准标的构建后DCF每股应复现
# - 断言用容差而非精确==: abs(ps-golden) <= max(1e-6, abs(golden)*1e-9) (跨LibreOffice版本浮点容差)
# - config指纹门: 冻结value时同时冻结 configs/<code>.yaml 的sha256;
#   verify时若当前配置指纹与冻结不一致 → 打印INFO跳过断言(不判FAIL), 待重建后统一重冻结。
GOLDEN_DCF_PS = {
    '300476': {
        'value': 340.415208186145,
        'config_sha256': '57b10d0620e96fdc42fe2d1452404db344b01cca5d1b64662712c22a31a27e2e',
    },
}


def _num(v):
    """E2 守卫: data_only 读回的标量, 仅真数值(int/float, 排除bool)→float;
    None / 错误串('#DIV/0!'等) / 其他类型 → None, 由调用方记FAIL, 不抛TypeError。"""
    if isinstance(v, bool) or not isinstance(v, (int, float)):
        return None
    return float(v)


def fmtv(v, spec='.2f'):
    """E2 守卫式格式化: 数值按spec格式化; 非数值(None/错误串)原样repr输出, 不崩溃"""
    n = _num(v)
    return format(n, spec) if n is not None else repr(v)


def _sha256_file(path):
    h = hashlib.sha256()
    with open(path, 'rb') as fh:
        for chunk in iter(lambda: fh.read(65536), b''):
            h.update(chunk)
    return h.hexdigest()


def _golden_value(code, cfg_path):
    """B3 指纹门: 该code有冻结记录且当前config指纹与冻结一致 → 返回冻结value;
    无冻结记录 → None; 指纹不一致或配置文件缺失 → 打印INFO跳过, 返回None (不判FAIL)。"""
    g = GOLDEN_DCF_PS.get(str(code))
    if not g:
        return None
    cur = _sha256_file(cfg_path) if (cfg_path and os.path.exists(cfg_path)) else None
    if cur != g.get('config_sha256'):
        print('  [INFO] 配置已变化, 跳过黄金值断言')
        return None
    return g['value']


def lo_recalc(src, workdir):
    """LibreOffice headless 重算 (v3模型无循环引用, 一次convert即得计算值)

    E1 加固: 独立UserInstallation profile(避免桌面LibreOffice占用默认profile导致
    --convert-to 静默no-op) + 转换前删旧输出 + 转换后断言产出了新文件。"""
    os.makedirs(workdir, exist_ok=True)
    work = os.path.join(workdir, os.path.basename(src))
    shutil.copy(src, work)
    outdir = os.path.join(workdir, 'out')
    os.makedirs(outdir, exist_ok=True)
    out_path = os.path.join(outdir, os.path.basename(src))
    if os.path.exists(out_path):  # E1: 删除已存在的目标输出, 防止转换no-op时误读陈旧文件假PASS
        os.remove(out_path)
    t0 = time.time()
    profile = 'file:///tmp/lo_verify_profile_{0}'.format(os.getpid())
    subprocess.run(['soffice', '--headless', '-env:UserInstallation=' + profile,
                    '--convert-to', 'xlsx', '--outdir', outdir, work],
                   check=True, capture_output=True, timeout=300)
    # E1: 输出必须存在且 mtime >= 转换开始时间 (留2s裕量容忍粗粒度文件系统mtime)
    if not os.path.exists(out_path) or os.path.getmtime(out_path) < t0 - 2:
        raise RuntimeError('LibreOffice 转换未产出新文件')
    return out_path


def main():
    ap = argparse.ArgumentParser(description='A股估值模型验收 (LO重算+Checks)')
    ap.add_argument('xlsx', nargs='?', help='模型xlsx路径 (或用--code)')
    ap.add_argument('--code', help='读取 out/<code>_*_估值模型.xlsx')
    ap.add_argument('--addr', help='addr json 路径 (默认随xlsx同名)')
    ap.add_argument('--no-recalc', action='store_true', help='不做LO重算, 直接读文件')
    ap.add_argument('--workdir', default=None, help='重算工作目录 (默认 .cache/lo_<code>)')
    args = ap.parse_args()

    path = args.xlsx
    if not path:
        if not args.code:
            ap.error('需提供 xlsx 路径或 --code')
        base = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'out')
        cands = [f for f in os.listdir(base) if f.startswith(args.code) and f.endswith('.xlsx')]
        if not cands:
            ap.error(f'out/ 下未找到 {args.code} 的xlsx')
        # E8: 多个候选时按mtime取最新, 不再按字典序取旧文件
        cands.sort(key=lambda f: os.path.getmtime(os.path.join(base, f)), reverse=True)
        path = os.path.join(base, cands[0])
        print(f'[INFO] --code 匹配 {len(cands)} 个候选, 选用最新: {cands[0]}')
    addr_path = args.addr or os.path.splitext(path)[0] + '.addr.json'
    addr = json.load(open(addr_path, encoding='utf-8'))
    meta = addr.get('meta', {})
    hist_years = meta.get('hist_years')
    fcst_years = meta.get('fcst_years')
    if not hist_years or not fcst_years:
        # E7: 不再用写死年份兜底, 元数据缺失直接报错
        raise RuntimeError('addr.json 缺少年份元数据, 请用配套 build 重新生成')
    YRS = list(hist_years) + list(fcst_years)
    FCST = list(fcst_years)
    COLS = {y: get_column_letter(2 + i) for i, y in enumerate(YRS)}

    if args.no_recalc:
        recalced = path
    else:
        workdir = args.workdir or os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                               '.cache', f"lo_{meta.get('code', 'x')}")
        recalced = lo_recalc(path, workdir)
    WB = load_workbook(recalced, data_only=True)

    def V(sheet, cell):
        return WB[sheet][cell].value

    def series(sheet, row, years=FCST):
        return [V(sheet, f'{COLS[y]}{row}') for y in years]

    def fmt(v, nd=1):
        if v is None:
            return 'None'
        if isinstance(v, str):
            return v
        return f'{v:,.{nd}f}'

    ok = True

    def check(name, cond, detail=''):
        nonlocal ok
        ok = ok and bool(cond)
        print(f"  [{'PASS' if cond else 'FAIL'}] {name} {detail}")

    def numguard(name, **kv):
        """E2: check前数值守卫. 任一值非数值→该check直接记FAIL并打印实际值, 返回None;
        全为数值→返回float字典, 调用方再做真正的条件判断。"""
        bad = {k: v for k, v in kv.items() if _num(v) is None}
        if bad:
            check(name, False, '非数值: ' + ', '.join('{0}={1!r}'.format(k, v) for k, v in bad.items()))
            return None
        return {k: _num(v) for k, v in kv.items()}

    print('=' * 70)
    print(f"文件: {recalced}")
    print(f"标的: {meta.get('name')}({meta.get('code_full')})  基准日: {meta.get('valuation_date')}")
    print('=' * 70)

    print('\n[Checks 总闸]')
    sumv = V('Checks', addr['checks_sum'].split('!')[1])
    boolv = V('Checks', addr['checks_bool'].split('!')[1])
    print('  汇总单元格:', sumv, '| 布尔:', boolv)
    check('Checks全部通过', boolv is True or (isinstance(boolv, (int, float)) and boolv == 1), f'bool={boolv}')

    print('\n[FIN 债务与融资调度]')
    diffs = series('BS', addr['bs_chk_row'], YRS)
    print('  BS配平差额 全期:', [f'{d:.6f}' if isinstance(d, (int, float)) else str(d) for d in diffs])
    bs_maxdiff = max((abs(d) for d in diffs if isinstance(d, (int, float))), default=0)
    check('① BS差额全期|diff|<0.01', all(isinstance(d, (int, float)) and abs(d) < 0.01 for d in diffs),
          f'max|diff|={bs_maxdiff:.2e}')

    cash = series('BS', addr['bs_cash_row'])
    minc = series('FIN', addr['fin_mincash_row'])
    check('② 全期现金=最低现金', all(isinstance(c, (int, float)) and isinstance(m, (int, float))
                                   and abs(c - m) < 0.01 for c, m in zip(cash, minc)),
          f'{FCST[-1]}E cash={fmt(cash[-1])} vs mincash={fmt(minc[-1])}')
    intexp = series('FIN', addr['fin_intexp_row'])
    print('  利息支出:', [fmt(v) for v in intexp])
    dtot = series('FIN', addr['fin_dtot1_row'])
    check('④ 债务余额>=0', all(isinstance(v, (int, float)) and v >= -0.01 for v in dtot))

    rmax = series('FIN', addr['fin_resid_max_row'])
    print('  残差上限:', [f'{v:.6f}' if isinstance(v, (int, float)) else str(v) for v in rmax])
    resid_max = max((v for v in rmax if isinstance(v, (int, float))), default=0)
    check('⑤ FIN迭代残差<0.01', all(isinstance(v, (int, float)) and v < 0.01 for v in rmax),
          f'max={resid_max:.2e}')
    cfc = series('CF', addr['cf_cash1_row'])
    # B1: 先算maxdiff变量再进f-string (f-string表达式内换行为3.12+特性, 需3.9兼容)
    cf_pairs = [(a, b) for a, b in zip(cfc, cash)
                if isinstance(a, (int, float)) and isinstance(b, (int, float))]
    cf_maxdiff = max((abs(a - b) for a, b in cf_pairs), default=0)
    check('⑥ CF期末现金=BS货币资金',
          len(cf_pairs) == len(cfc) and all(abs(a - b) < 0.01 for a, b in cf_pairs),
          f'maxdiff={cf_maxdiff:.2e}')

    print('\n[WACC]')
    wcalc = V('Assumptions', addr['wacc_calc'].split('!')[1])
    wused = V('Assumptions', addr['wacc_used'].split('!')[1])
    ke = V('Assumptions', addr['ke'].split('!')[1])
    bu = V('Assumptions', addr['beta_u'].split('!')[1])
    bl = V('Assumptions', addr['beta_l'].split('!')[1])
    wd = V('Assumptions', addr['wd'].split('!')[1])
    print(f"  βu中位={fmtv(bu, '.4f')}  βl(relever)={fmtv(bl, '.4f')}  Wd={fmtv(wd, '.4%')}  Ke={fmtv(ke, '.4%')}")
    print(f"  WACC计算值={fmtv(wcalc, '.4%')}  采用值={fmtv(wused, '.4%')}")
    w = numguard('采用值=计算值(override留空)', wacc_calc=wcalc, wacc_used=wused)
    if w:
        check('采用值=计算值(override留空)', abs(w['wacc_calc'] - w['wacc_used']) < 1e-9)

    print('\n[DCF]')
    ev = V('DCF', addr['dcf_ev'].split('!')[1]); eq = V('DCF', addr['dcf_eq'].split('!')[1])
    ps_raw = V('DCF', addr['dcf_ps'].split('!')[1]); up = V('DCF', addr['dcf_upside'].split('!')[1])
    print(f"  EV={fmt(ev)}  股权价值={fmt(eq)}  每股价值={fmtv(ps_raw, '.4f')}元  隐含涨跌幅={fmtv(up, '+.1%')}")
    ps = _num(ps_raw)
    if ps is None:
        check('DCF每股为有效数值', False, f'实际值={ps_raw!r}')
    cfg_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'configs',
                            '{0}.yaml'.format(meta.get('code')))
    golden = _golden_value(meta.get('code'), cfg_path)
    if golden is not None:
        g = numguard('黄金值: DCF每股复现', dcf_ps=ps_raw)
        if g:
            # B3: 跨LO版本浮点容差 abs(ps-golden) <= max(1e-6, |golden|*1e-9), 不再精确==
            tol = max(1e-6, abs(golden) * 1e-9)
            check('黄金值: DCF每股复现(容差max(1e-6,|golden|·1e-9))',
                  abs(g['dcf_ps'] - golden) <= tol, f"{g['dcf_ps']!r} vs 基准{golden!r}")
    # E3: t序所在行从 addr['dcf_tidx'] (形如 "DCF!C11") 解析, 不再硬编码
    tidx_row = 11
    tidx_addr = addr.get('dcf_tidx')
    if tidx_addr:
        m = re.search(r'(\d+)\s*$', str(tidx_addr).split('!')[-1].replace('$', ''))
        if m:
            tidx_row = int(m.group(1))
        else:
            print(f'[WARN] dcf_tidx 无法解析行号({tidx_addr!r}), 按行11回退')
    else:
        print('[WARN] addr 缺 dcf_tidx, 按行11回退')
    tidx = [V('DCF', f'{COLS[y]}{tidx_row}') for y in FCST]
    check('年中折现t=0.5..4.5', all(isinstance(v, (int, float)) and abs(v - (0.5 + i)) < 1e-9
                                  for i, v in enumerate(tidx)), str(tidx))

    if addr.get('fcfe_ps'):
        print('\n[FCFE 双视图]')
        fps = V('FCFE', addr['fcfe_ps'].split('!')[1])
        feq = V('FCFE', addr['fcfe_eq'].split('!')[1])
        fdiff = V('FCFE', addr['fcfe_diff'].split('!')[1])
        print(f"  FCFE股权价值={fmt(feq)}  FCFE每股={fmtv(fps, '.4f')}元  vs FCFF {fmtv(ps_raw, '.4f')}元 ({fmtv(fdiff, '+.1%')})")
        g = numguard('FCFE每股>0', fcfe_ps=fps)
        if g:
            check('FCFE每股>0', g['fcfe_ps'] > 0, '{0:.2f}元'.format(g['fcfe_ps']))
        fd = _num(fdiff)
        if fd is not None and abs(fd) >= 0.30:
            print(f'  [INFO] 两法偏差{fd:+.1%}较大: 多见于融资计划大幅去杠杆的标的(FCFE含净新增借款), '
                  '机制性差异说明见FCFE页底部注记')
    if addr.get('named_ranges'):
        print(f"  named ranges: {len(addr['named_ranges'])}个 ({', '.join(list(addr['named_ranges'])[:6])}...)")

    print('\n[三情景]')
    for k in ['bear', 'base', 'bull']:
        p = V('Scenarios', addr[f'scen_{k}_ps'].split('!')[1])
        rl = V('Scenarios', addr[f'scen_{k}_rel'].split('!')[1])
        print(f"  {k}: DCF每股={fmtv(p, '.2f')}元  相对估值={fmtv(rl, '.2f')}元")
    sb = numguard('基准情景DCF=主模型DCF',
                  scen_base=V('Scenarios', addr['scen_base_ps'].split('!')[1]), dcf_ps=ps_raw)
    if sb:
        check('基准情景DCF=主模型DCF', abs(sb['scen_base'] - sb['dcf_ps']) < 0.01)

    lo = V('Relative_Val', addr['rel_lo'].split('!')[1]); hi = V('Relative_Val', addr['rel_hi'].split('!')[1])
    med = addr['rel_med_pe'].split('!')
    medpe = V(med[0], med[1]) if len(med) == 2 else None
    print(f"\n相对估值区间=[{fmtv(lo, '.2f')}, {fmtv(hi, '.2f')}]元  可比中位PE={medpe if not isinstance(medpe, float) else round(medpe, 1)}")

    sens = V('Sensitivity', addr['sens_center'].split('!')[1])
    sg = numguard('敏感矩阵中心=DCF', sens_center=sens, dcf_ps=ps_raw)
    if sg:
        check('敏感矩阵中心=DCF', abs(sg['sens_center'] - sg['dcf_ps']) < 0.01,
              '{0:.2f} vs {1:.2f}'.format(sg['sens_center'], sg['dcf_ps']))
    dpo_c = V('Sensitivity', addr['dpo_center'].split('!')[1])
    dg = numguard('DPO矩阵δ=0=DCF基准', dpo_center=dpo_c, dcf_ps=ps_raw)
    if dg:
        check('DPO矩阵δ=0=DCF基准', abs(dg['dpo_center'] - dg['dcf_ps']) < 0.01,
              '{0:.2f}'.format(dg['dpo_center']))
    env = V('Summary', addr['summary_env'].split('!')[1]); env_hi = V('Summary', addr['summary_env_hi'].split('!')[1])
    print(f"Summary综合参考区间=[{fmtv(env, '.2f')}, {fmtv(env_hi, '.2f')}]元")

    print('\n[主模型输出]')
    for y in FCST[:3]:
        rv = V('IS', addr['is_rev'][str(y)].split('!')[1])
        npv = V('IS', addr['is_np'][str(y)].split('!')[1])
        print(f'  {y}E: 营收={fmt(rv)}  归母={fmt(npv)}')

    print('\n' + '=' * 70)
    print('总体:', 'ALL PASS' if ok else '存在FAIL项')
    print('=' * 70)
    sys.exit(0 if ok else 1)


if __name__ == '__main__':
    main()
