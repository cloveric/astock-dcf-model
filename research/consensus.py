# -*- coding: utf-8 -*-
"""聚源/gildata一致预期文件输入 (json/csv)

json格式:
  {"rev": {"2026": 33370.0, "2027": 53930.0, "2028": 77430.0},
   "np":  {"2026": 9120.0, "2027": 15410.0, "2028": 22970.0},
   "target_price": 350.0, "source": "gildata/聚源", "date": "2026-08-06"}
csv格式 (表头 year,rev,np[,target_price,source,date]):
  year,rev,np,target_price,source,date
  2026,33370.0,9120.0,350.0,gildata,2026-08-06
单位: 营收/归母=人民币百万元; 目标价=元/股。缺少年份回退到配置已有值。
环境有MCP gildata工具时可由调用方自动拉取后写成本文件再传入; 工具本身不直连付费源。
"""
import csv
import json
import os


def load(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == '.json':
        with open(path, encoding='utf-8') as f:
            d = json.load(f)
        rev = {int(k): float(v) for k, v in (d.get('rev') or {}).items()}
        np_ = {int(k): float(v) for k, v in (d.get('np') or {}).items()}
        tp = d.get('target_price')
        src, dt = d.get('source'), d.get('date')
    else:
        rev, np_, tp, src, dt = {}, {}, None, None, None
        with open(path, encoding='utf-8-sig') as f:
            for row in csv.DictReader(f):
                y = int(str(row['year']).strip())
                if row.get('rev'):
                    rev[y] = float(row['rev'])
                if row.get('np'):
                    np_[y] = float(row['np'])
                if row.get('target_price') and tp is None:
                    tp = float(row['target_price'])
                src = src or (row.get('source') or None)
                dt = dt or (row.get('date') or None)
    if not rev and not np_:
        raise ValueError(f'一致预期文件无有效数据: {path}')
    return {'rev': rev, 'np': np_, 'target_price': tp,
            'source': src or os.path.basename(path), 'date': dt or '', 'path': path}


def apply_to_cfg(raw, cons):
    """用文件一致预期覆盖配置 consensus 段 (依据字段如实溯源, 不做复制填充造假)

    按年份对齐: 槽位 = 配置 model.fcst_years 前 3 个预测年, 逐槽三种情形:
      ① 文件覆盖该年 → 写文件值;
      ② 文件未覆盖但配置有原值 → 保留配置原值 (不用文件值覆盖填充);
      ③ 文件与配置均无该槽 → 用最近文件年的值平推兜底, 并在 basis 中如实标注;
    文件年份与预测序列完全无交集则报错 (防位置错位写错槽)。
    basis 按实际情形拼接, 如:
      '聚源/gildata一致预期文件(gildata, 2026-08-06, 覆盖2027/2028); 2026沿用原配置'
      '聚源/gildata一致预期文件(..., 覆盖2027/2028); 2026按2027值平推(文件未覆盖)'"""
    meta = f"{cons['source']}{', ' + cons['date'] if cons['date'] else ''}"
    node = raw.setdefault('consensus', {})
    fcst = (raw.get('model') or {}).get('fcst_years') or []
    slots = {int(y): i for i, y in enumerate(fcst[:3])}

    def _merge(key, series):
        if not series:
            return
        hit = {y: v for y, v in series.items() if y in slots}
        if not hit:
            raise ValueError(
                f'一致预期[{key}]年份{sorted(series)}与配置预测年前3年{sorted(slots)}无交集')
        ent = node.get(key) or {}
        vals = ent.get('values') or ent.get('value') or []
        if not isinstance(vals, list):
            vals = [vals]
        out = list(vals)
        while len(out) < len(slots):
            out.append(None)                # 先占位, 下面按槽位逐年决定
        kept, padded = [], []               # 沿用原配置的年 / 平推兜底的(年, 取值来源年)
        for y, i in sorted(slots.items(), key=lambda kv: kv[1]):
            if y in hit:                    # ① 文件覆盖 → 写文件值
                out[i] = hit[y]
            elif i < len(vals) and vals[i] is not None:
                kept.append(y)              # ② 文件未覆盖 → 保留配置原值
            else:                           # ③ 双方均无 → 最近文件年值平推兜底
                near = min(hit, key=lambda fy: (abs(fy - y), fy))
                out[i] = hit[near]
                padded.append((y, near))
        cov = '/'.join(str(y) for y in sorted(hit))
        parts = [f'聚源/gildata一致预期文件({meta}, 覆盖{cov})']
        if kept:
            parts.append(f"{'/'.join(str(y) for y in kept)}沿用原配置")
        for y, near in padded:
            parts.append(f'{y}按{near}值平推(文件未覆盖)')
        node[key] = {'values': out, 'basis': '; '.join(parts)}

    _merge('rev', cons['rev'])
    _merge('np', cons['np'])
