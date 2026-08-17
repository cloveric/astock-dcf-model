# -*- coding: utf-8 -*-
"""研究层 (research layer): 全部来源默认关闭, 仅经CLI开关激活

来源优先级 (高→低):
  1. --dr <档案.md>      深度研究档案: 量化结论回填Assumptions依据列(标注dr档案§章节),
                         新增 DR研究(全文存档)/研究摘要(假设×采用值×来源×置信度)两页
  2. --consensus <file>  聚源/gildata一致预期(json/csv): 覆盖consensus段并写入研究摘要,
                         Relative_Val/Checks各加一行一致性对照
  3. --announcements     东财业绩预告/业绩快报(系统curl, python不直连东财)最新一期要点
  4. --llm auto|claude|codex|off  本机存在对应CLI时生成研究备忘录写入研究摘要页眉
任何来源不可用时优雅降级(记note, 不报错中断建模)。
"""
from . import announcements as _ann_mod
from . import consensus as _cons_mod
from . import dr_report as _dr_mod
from . import llm as _llm_mod


def collect_research(cfg, dr_path=None, consensus_path=None,
                     want_announcements=False, llm='off'):
    """汇总各研究来源 → ctx dict; 任一来源失败仅记note不中断"""
    ctx = {'notes': []}

    if dr_path:
        try:
            ctx['dr'] = _dr_mod.parse(dr_path)
            print(f"研究层: dr档案 {dr_path} "
                  f"({len(ctx['dr']['sections'])}节, {len(ctx['dr']['findings'])}条量化结论)")
        except Exception as e:
            ctx['notes'].append(f'dr档案读取失败: {e}')
            print(f'研究层: dr档案读取失败({e}), 已跳过')

    if consensus_path:
        try:
            cons = _cons_mod.load(consensus_path)
            _cons_mod.apply_to_cfg(cfg.raw, cons)
            ctx['consensus'] = cons
            print(f"研究层: 一致预期文件 {consensus_path} "
                  f"(rev {sorted(cons['rev'])}, np {sorted(cons['np'])}, 目标价 {cons['target_price']})")
        except Exception as e:
            ctx['notes'].append(f'一致预期文件读取失败: {e}')
            print(f'研究层: 一致预期文件读取失败({e}), 已跳过')

    if want_announcements:
        code = cfg.get('company.code')
        try:
            ctx['announcements'] = _ann_mod.fetch(code)
            n = len(ctx['announcements'].get('items', []))
            print(f'研究层: 东财公告要点 {n} 条' if n else
                  f"研究层: 东财公告无最新要点({ctx['announcements'].get('note', '')})")
        except Exception as e:
            ctx['announcements'] = {'items': [], 'note': str(e)}
            ctx['notes'].append(f'公告抓取失败: {e}')
            print(f'研究层: 公告抓取失败({e}), 已跳过')

    if llm != 'off':
        prompt = _memo_prompt(cfg, ctx)
        memo, note = _llm_mod.gen_memo(llm, prompt)
        ctx['memo'] = memo
        ctx['notes'].append(note)
        print(f'研究层: {note}')

    active = any(ctx.get(k) for k in ('dr', 'consensus', 'announcements', 'memo'))
    return ctx if active or ctx['notes'] else None


def _memo_prompt(cfg, ctx):
    co = cfg.raw.get('company', {})
    lines = [f"标的: {co.get('name')}({co.get('code_full')})"]
    if ctx.get('dr'):
        lines.append('深度研究档案要点:')
        for f in ctx['dr']['findings'][:12]:
            lines.append(f"- §{f['section']} {f['text']}")
    if ctx.get('consensus'):
        c = ctx['consensus']
        lines.append(f"一致预期: 营收{c['rev']} 归母{c['np']} 目标价{c['target_price']} (来源{c['source']})")
    lines.append('请写一段不超过250字的研究备忘录: 核心投资逻辑、关键假设中最需复核的3项、主要风险。'
                 '纯文本, 不要markdown标题。')
    return '\n'.join(lines)
