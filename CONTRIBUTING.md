# Contributing

感谢贡献。本仓库的核心纪律先于一切功能迭代, 违反即返工。

## 核心纪律

1. **预测期零硬编码**: 预测期所有单元格必须由假设驱动的公式生成; 硬编码仅限历史实际值(绿色)与蓝色假设输入。
2. **每条假设必须带依据**: YAML 中每个假设条目必须携带 `basis` 文字字段, 会写入 Assumptions 页备注列。无依据的假设不予合入。
3. **三表严格配平**: BS 配平差额全期必须 < 0.01, 不允许手填轧差; 配平只能由 FIN 页 revolver/sweep 机制完成。
4. **Checks 全 TRUE**: 提交前必须本地跑通 `python verify_model.py --code 300476`(需 LibreOffice), Checks 布尔汇总必须为 TRUE。
5. **基准复现**: 改动不得改变胜宏科技(300476)基准输出 —— DCF 每股 340.415208186145 元; 新增 sheet/新增行除外, 既有单元格数值必须零差异。
6. **数据源纪律**: 东财一律经系统 `curl` 子进程获取(python 不直连东财); 一致预期等付费源不直连, 走 `--consensus` 文件输入。
7. **研究层默认关闭**: `--dr/--consensus/--announcements/--llm` 不传时, 输出必须与基础模型完全一致。

## 开发流程

```bash
pip install -r requirements.txt
python build_model.py --code 300476          # 构建
python verify_model.py --code 300476         # LibreOffice 重算验收
python tests/smoke_check.py out/300476_胜宏科技_估值模型.xlsx   # 离线冒烟
```

- 分支: `feat/<主题>` / `fix/<主题>`; commit 信息用 [Conventional Commits](https://www.conventionalcommits.org/)(`feat:` / `fix:` / `docs:` / `chore:` / `test:`)。
- PR 必须通过 CI(离线语法检查 + fixture 构建 + 幂等性 + 可用时的 LibreOffice 全量验收)。
- 改动了 Checks 项数、sheet 结构或 addr.json 键, 需同步更新 `verify_model.py`、`tests/smoke_check.py` 与 README 的对应清单。

## 新增个股配置

复制 `configs/300476.yaml` 改写, 或 `python fetch_data.py --code <代码>` 自动生成兜底配置后按研究修正。配置即研究底稿, 依据字段写清来源与日期。
