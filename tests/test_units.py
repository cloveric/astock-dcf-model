# -*- coding: utf-8 -*-
"""离线单元测试 (pytest, 不依赖LibreOffice/外网):
   - research.consensus.apply_to_cfg 按年份对齐合并 (BUG-1回归: 三年全给/缺首年/缺中间年/全缺报错)
   - research.dr_report.annotate 打分排序
   - web.server._repo_file 路径穿越校验
运行: python -m pytest tests/test_units.py -q
"""
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from research import consensus as cons_mod
from research import dr_report


# ---------------- consensus 合并 (BUG-1 回归) ----------------

def _raw():
    return {'model': {'fcst_years': [2026, 2027, 2028, 2029, 2030]},
            'consensus': {'rev': {'values': [100.0, 110.0, 120.0], 'basis': '原配置'},
                          'np': {'values': [10.0, 11.0, 12.0], 'basis': '原配置'}}}


def _cons(rev=None, np_=None):
    return {'rev': rev or {}, 'np': np_ or {}, 'target_price': None,
            'source': 'pytest', 'date': '', 'path': 'x.json'}


def test_merge_full_three_years():
    raw = _raw()
    cons_mod.apply_to_cfg(raw, _cons(rev={2026: 1.0, 2027: 2.0, 2028: 3.0}))
    assert raw['consensus']['rev']['values'] == [1.0, 2.0, 3.0]
    assert raw['consensus']['np']['values'] == [10.0, 11.0, 12.0]   # 文件未给np, 原值保留


def test_merge_missing_first_year():
    """缺首年: 文件给2027/2028, 2026槽位必须保留原值 (旧按位置合并会把2027写进2026槽)"""
    raw = _raw()
    cons_mod.apply_to_cfg(raw, _cons(rev={2027: 2.0, 2028: 3.0}))
    assert raw['consensus']['rev']['values'] == [100.0, 2.0, 3.0]


def test_merge_missing_middle_year():
    """缺中间年: 文件给2026/2028, 2027槽位保留原值"""
    raw = _raw()
    cons_mod.apply_to_cfg(raw, _cons(rev={2026: 1.0, 2028: 3.0}))
    assert raw['consensus']['rev']['values'] == [1.0, 110.0, 3.0]


def test_merge_all_years_outside_raises():
    """文件年份与预测序列全不交集 → 显式报错"""
    raw = _raw()
    with pytest.raises(ValueError):
        cons_mod.apply_to_cfg(raw, _cons(rev={2031: 1.0, 2032: 2.0}))


def test_merge_partial_overlap_keeps_known_years():
    """部分年份在预测序列内: 写可对齐年份, 序列外年份忽略"""
    raw = _raw()
    cons_mod.apply_to_cfg(raw, _cons(rev={2028: 3.0, 2031: 9.9}))
    assert raw['consensus']['rev']['values'] == [100.0, 110.0, 3.0]


# ---------------- dr annotate 打分 ----------------

_FINDINGS = [
    {'section': '3', 'title': '业务量价', 'text': 'HDI销量2026年预计增长30%, 产能利用率高位',
     'keywords': ['HDI', '销量', '产能', '增速']},
    {'section': '4', 'title': '毛利率', 'text': '高多层板ASP提升至800元/平米',
     'keywords': ['高多层', 'ASP']},
    {'section': '5', 'title': '产能', 'text': '公司产能扩张按计划推进',
     'keywords': ['产能']},
]


def test_annotate_relevant_first():
    """多关键词命中的结论排在前面"""
    hits = dr_report.annotate(_FINDINGS, 'HDI 销量')
    assert hits and hits[0].startswith('dr档案§3:')


def test_annotate_label_without_keyword_mapping():
    """label无关键词映射 → 不回填"""
    assert dr_report.annotate(_FINDINGS, '现金流') == []


def test_annotate_max_hits_respected():
    hits = dr_report.annotate(_FINDINGS, '销量 产能', max_hits=1)
    assert len(hits) == 1


def test_annotate_truncates_long_text():
    long_text = 'HDI销量' + '长' * 100 + '30%'
    hits = dr_report.annotate(
        [{'section': '1', 'title': 't', 'text': long_text, 'keywords': ['HDI']}], 'HDI')
    assert hits[0].endswith('...') and len(hits[0]) <= len('dr档案§1: ') + 60


# ---------------- web 路径校验 ----------------

def test_repo_file_accepts_repo_relative():
    from web import server
    p = server._repo_file('configs/300476.yaml')
    assert p == os.path.join(server.REPO_ROOT, 'configs', '300476.yaml')


def test_repo_file_rejects_traversal():
    from fastapi import HTTPException
    from web import server
    with pytest.raises(HTTPException):
        server._repo_file('../outside.yaml')
    with pytest.raises(HTTPException):
        server._repo_file('/etc/passwd')


def test_repo_file_rejects_missing_file():
    from fastapi import HTTPException
    from web import server
    with pytest.raises(HTTPException):
        server._repo_file('configs/not_exist_999999.yaml')


def test_repo_file_empty_returns_none():
    from web import server
    assert server._repo_file('') is None
    assert server._repo_file(None) is None
