# -*- coding: utf-8 -*-
"""
模型验收脚本 (通用版): LibreOffice headless 重算 + 逐项校验

用法:
  python verify_model.py --code 300476                 # 找 out/ 下xlsx+addr, 重算并校验
  python verify_model.py out/300476_胜宏科技_估值模型.xlsx
  python verify_model.py <xlsx> --no-recalc            # 直接读已重算文件
校验项: Checks布尔总闸 / BS配平 / 现金=最低现金 / FIN迭代残差 / CF-BS现金勾稽 /
        WACC采用值=计算值 / 年中折现t序 / 基准情景=主模型 / 敏感性矩阵中心=DCF /
        黄金值断言(基准标的DCF每股精确复现, 容差0, 见GOLDEN_DCF_PS)。
"""
import argparse
import json
import os
import shutil
import subprocess
import sys

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# 黄金值 (容差0): 基准标的构建后DCF每股必须精确复现, 防回归
GOLDEN_DCF_PS = {'300476': 340.415208186145}


def lo_recalc(src, workdir):
    """LibreOffice headless 重算 (v3模型无循环引用, 一次convert即得计算值)"""
    os.makedirs(workdir, exist_ok=True)
    work = os.path.join(workdir, os.path.basename(src))
    shutil.copy(src, work)
    outdir = os.path.join(workdir, 'out')
    os.makedirs(outdir, exist_ok=True)
    subprocess.run(['soffice', '--headless', '--convert-to', 'xlsx', '--outdir', outdir, work],
                   check=True, capture_output=True, timeout=300)
    return os.path.join(outdir, os.path.basename(src))


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
        path = os.path.join(base, sorted(cands)[0])
    addr_path = args.addr or os.path.splitext(path)[0] + '.addr.json'
    addr = json.load(open(addr_path, encoding='utf-8'))
    meta = addr.get('meta', {})
    YRS = meta.get('hist_years', [2023, 2024, 2025]) + meta.get('fcst_years', [2026, 2027, 2028, 2029, 2030])
    FCST = meta.get('fcst_years', YRS[-5:])
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
    check('① BS差额全期|diff|<0.01', all(isinstance(d, (int, float)) and abs(d) < 0.01 for d in diffs),
          f'max|diff|={max((abs(d) for d in diffs if isinstance(d, (int, float))), default=0):.2e}')

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
    check('⑤ FIN迭代残差<0.01', all(isinstance(v, (int, float)) and v < 0.01 for v in rmax),
          f'max={max((v for v in rmax if isinstance(v, (int, float))), default=0):.2e}')
    cfc = series('CF', addr['cf_cash1_row'])
    check('⑥ CF期末现金=BS货币资金', all(isinstance(a, (int, float)) and isinstance(b, (int, float))
                                       and abs(a - b) < 0.01 for a, b in zip(cfc, cash)),
          f'maxdiff={max((abs(a - b) for a, b in zip(cfc, cash)
                          if isinstance(a, (int, float)) and isinstance(b, (int, float))), default=0):.2e}')

    print('\n[WACC]')
    wcalc = V('Assumptions', addr['wacc_calc'].split('!')[1])
    wused = V('Assumptions', addr['wacc_used'].split('!')[1])
    ke = V('Assumptions', addr['ke'].split('!')[1])
    bu = V('Assumptions', addr['beta_u'].split('!')[1])
    bl = V('Assumptions', addr['beta_l'].split('!')[1])
    wd = V('Assumptions', addr['wd'].split('!')[1])
    print(f'  βu中位={bu:.4f}  βl(relever)={bl:.4f}  Wd={wd:.4%}  Ke={ke:.4%}')
    print(f'  WACC计算值={wcalc:.4%}  采用值={wused:.4%}')
    check('采用值=计算值(override留空)', abs(wcalc - wused) < 1e-9)

    print('\n[DCF]')
    ev = V('DCF', addr['dcf_ev'].split('!')[1]); eq = V('DCF', addr['dcf_eq'].split('!')[1])
    ps = V('DCF', addr['dcf_ps'].split('!')[1]); up = V('DCF', addr['dcf_upside'].split('!')[1])
    print(f'  EV={fmt(ev)}  股权价值={fmt(eq)}  每股价值={ps:.4f}元  隐含涨跌幅={up:+.1%}')
    golden = GOLDEN_DCF_PS.get(str(meta.get('code')))
    if golden is not None:
        check('黄金值: DCF每股精确复现(容差0)', ps == golden, f'{ps!r} vs 基准{golden!r}')
    tidx = [V('DCF', f'{COLS[y]}11') for y in FCST]
    check('年中折现t=0.5..4.5', all(isinstance(v, (int, float)) and abs(v - (0.5 + i)) < 1e-9
                                  for i, v in enumerate(tidx)), str(tidx))

    if addr.get('fcfe_ps'):
        print('\n[FCFE 双视图]')
        fps = V('FCFE', addr['fcfe_ps'].split('!')[1])
        feq = V('FCFE', addr['fcfe_eq'].split('!')[1])
        fdiff = V('FCFE', addr['fcfe_diff'].split('!')[1])
        print(f'  FCFE股权价值={fmt(feq)}  FCFE每股={fps:.4f}元  vs FCFF {ps:.4f}元 ({fdiff:+.1%})')
        check('FCFE每股>0', fps > 0, f'{fps:.2f}元')
        if abs(fdiff) >= 0.30:
            print(f'  [INFO] 两法偏差{fdiff:+.1%}较大: 多见于融资计划大幅去杠杆的标的(FCFE含净新增借款), '
                  '机制性差异说明见FCFE页底部注记')
    if addr.get('named_ranges'):
        print(f"  named ranges: {len(addr['named_ranges'])}个 ({', '.join(list(addr['named_ranges'])[:6])}...)")

    print('\n[三情景]')
    for k in ['bear', 'base', 'bull']:
        p = V('Scenarios', addr[f'scen_{k}_ps'].split('!')[1])
        rl = V('Scenarios', addr[f'scen_{k}_rel'].split('!')[1])
        print(f'  {k}: DCF每股={p:.2f}元  相对估值={rl:.2f}元')
    check('基准情景DCF=主模型DCF', abs(V('Scenarios', addr['scen_base_ps'].split('!')[1]) - ps) < 0.01)

    lo = V('Relative_Val', addr['rel_lo'].split('!')[1]); hi = V('Relative_Val', addr['rel_hi'].split('!')[1])
    med = addr['rel_med_pe'].split('!')
    medpe = V(med[0], med[1]) if len(med) == 2 else None
    print(f'\n相对估值区间=[{lo:.2f}, {hi:.2f}]元  可比中位PE={medpe if not isinstance(medpe, float) else round(medpe, 1)}')

    sens = V('Sensitivity', addr['sens_center'].split('!')[1])
    check('敏感矩阵中心=DCF', abs(sens - ps) < 0.01, f'{sens:.2f} vs {ps:.2f}')
    dpo_c = V('Sensitivity', addr['dpo_center'].split('!')[1])
    check('DPO矩阵δ=0=DCF基准', abs(dpo_c - ps) < 0.01, f'{dpo_c:.2f}')
    env = V('Summary', addr['summary_env'].split('!')[1]); env_hi = V('Summary', addr['summary_env_hi'].split('!')[1])
    print(f'Summary综合参考区间=[{env:.2f}, {env_hi:.2f}]元')

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
