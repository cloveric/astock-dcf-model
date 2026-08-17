# -*- coding: utf-8 -*-
"""东财公告要点提取: 业绩预告 / 业绩快报 (一律系统curl, python不直连东财)

数据源: datacenter-web.eastmoney.com 数据中心
  业绩预告 reportName=RPT_PUBLIC_OP_PREDICT
  业绩快报 reportName=RPT_FCI_PERFORMANCEE
取最新一期(按公告日倒序), 输出要点文本列表; 任何失败优雅降级为note。
"""
import json

from fetch_data import curl_get

_DC = 'https://datacenter-web.eastmoney.com/api/data/v1/get'


def _get(report_name, code, size=3):
    url = (f'{_DC}?reportName={report_name}&columns=ALL'
           f'&filter=(SECURITY_CODE%3D%22{code}%22)&pageNumber=1&pageSize={size}'
           f'&sortColumns=NOTICE_DATE&sortTypes=-1&source=DataCenter')
    raw = curl_get(url, referer='https://data.eastmoney.com/')
    d = json.loads(raw.decode('utf-8', errors='replace'))
    return (d.get('result') or {}).get('data') or []


def _yi(v):
    try:
        return f'{float(v) / 1e8:.2f}亿'
    except (TypeError, ValueError):
        return '—'


def _pct(v):
    try:
        return f'{float(v):+.1f}%'
    except (TypeError, ValueError):
        return '—'


def fetch(code):
    """→ {'items': [要点文本...], 'note': str}"""
    items = []
    notes = []

    try:
        rows = _get('RPT_PUBLIC_OP_PREDICT', code, size=1)
        if rows:
            r0 = rows[0]
            typ = r0.get('FORECASTTYPE') or r0.get('PREDICT_TYPE') or ''
            rdate = (r0.get('REPORTDATE') or r0.get('REPORT_DATE') or '')[:10]
            ndate = (r0.get('NOTICE_DATE') or '')[:10]
            lo, hi = r0.get('FORECASTL'), r0.get('FORECASTT')
            glo, ghi = r0.get('INCREASEL'), r0.get('INCREASET')
            content = (r0.get('FORECASTCONTENT') or r0.get('PREDICT_CONTENT') or '').strip().replace('\n', ' ')
            if len(content) > 120:
                content = content[:117] + '...'
            seg = f"业绩预告({rdate}期, {ndate}披露, {typ})"
            if lo is not None and hi is not None:
                seg += f": 归母 {_yi(lo)}~{_yi(hi)}"
                if glo is not None and ghi is not None:
                    seg += f" (同比{_pct(glo)}~{_pct(ghi)})"
            if content:
                seg += f"; 摘要: {content}"
            items.append(seg)
        else:
            notes.append('最新一期无业绩预告')
    except Exception as e:
        notes.append(f'业绩预告抓取失败({e})')

    try:
        rows = _get('RPT_FCI_PERFORMANCEE', code, size=1)
        if rows:
            r0 = rows[0]
            rdate = (r0.get('REPORT_DATE') or '')[:10]
            ndate = (r0.get('NOTICE_DATE') or '')[:10]
            items.append(f"业绩快报({rdate}期, {ndate}披露): 营收{_yi(r0.get('TOTAL_OPERATE_INCOME'))}"
                         f"(同比{_pct(r0.get('YSTZ'))}), 归母{_yi(r0.get('PARENT_NETPROFIT'))}"
                         f"(同比{_pct(r0.get('SJLTZ'))})")
        else:
            notes.append('最新一期无业绩快报')
    except Exception as e:
        notes.append(f'业绩快报抓取失败({e})')

    return {'items': items, 'note': '; '.join(notes)}
