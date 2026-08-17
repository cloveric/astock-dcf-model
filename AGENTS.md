# AGENTS.md — 给 AI 协作者的仓库说明

## 这是什么

A 股机构级 DCF 估值模型生成器。`build_model.py --code <6位代码>` 读取 `configs/<code>.yaml`,
用 openpyxl 生成 17 张工作表的全公式 Excel 模型(研究层激活时追加 DR研究/研究摘要);
`verify_model.py` 用 LibreOffice headless 重算并验收 Checks 页。

## 不可违反的铁律 (改动前必读)

1. **基准复现零差异**: 胜宏科技(300476)基准 DCF 每股 = **340.415208186145 元**。
   任何改动后必须重跑 `python build_model.py --code 300476 && python verify_model.py --code 300476`,
   既有 16 张表(除 Summary/Checks 新增的 FCFE/named-range 行)数值零差异, Checks 全 TRUE。
2. **预测期零硬编码**: 预测期单元格一律公式; 硬编码仅限历史实际值(绿色 kind='x')与假设输入(蓝色 kind='in')。
3. **每个假设带依据**: YAML 每个假设条目必须有 `basis` 字段; 兜底自动推导的也必须标注"自动推导"。
4. **东财只走系统 curl**: 禁止 python 直接 requests/urllib 访问东财(反爬纪律); 行情走腾讯 qt.gtimg.cn。
5. **研究层开关默认关闭**: 不传 `--dr/--consensus/--announcements/--llm` 时输出必须与默认路径完全一致。

## 结构

- `build_model.py` — 单文件构建器(~2200 行), 各 sheet 依次建骨架后回填公式; 单元格地址索引导出 addr.json。
- `fetch_data.py` — 数据层: 东财 F10 三表/主营构成、腾讯行情、兜底配置推导。
- `verify_model.py` — LibreOffice 重算验收(依赖 addr.json 寻址)。
- `research/` — 研究层: dr_report(档案解析与依据回填) / consensus(一致预期文件) / announcements(东财公告) / llm(本机CLI备忘录) / sheets(两个研究页)。
- `configs/` — 个股配置(300476=完整手工范例, 002463=全自动范例)。
- `examples/` — 成稿 xlsx + 验收日志 + research fixture; `tests/smoke_check.py` 离线冒烟。
- 关键行号/地址不落盘硬编码在文档里 —— 全部经 addr.json 传递。

## 常用命令

```bash
python build_model.py --code 300476                 # 构建(默认路径)
python verify_model.py --code 300476                # LO重算验收
python build_model.py --code 300476 --dr examples/research/dr_300476.md \
    --consensus examples/research/consensus_300476.json --announcements --llm off   # 研究层全开
python tests/smoke_check.py out/300476_胜宏科技_估值模型.xlsx   # 无LibreOffice环境的冒烟
```

## 工程约定

- Conventional Commits; 产物 `out/`、`.cache/` 不入库; examples/ 下样例除外。
- 新增 sheet/Checks 项时: 同步更新 `verify_model.py`、`tests/smoke_check.py`、README 清单与本文件。
- LibreOffice 重算是唯一验收口径; openpyxl 只写不算。
