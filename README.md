# astock-dcf-model — A股机构级 DCF 估值模型生成器

一个配置驱动的 A 股估值建模工具：给定 6 位股票代码，自动生成 **16 张工作表** 的美元基金级三表联动 + DCF 估值 Excel 模型（openpyxl 全公式，LibreOffice/Excel/WPS 可重算，自带 Checks 校验页）。

由胜宏科技(300476)定制建模脚本产品化而来；用新工具重跑胜宏配置，基准 DCF 每股价值 **340.4152 元**，与原定制脚本逐值一致。

## 功能

- **三表联动预测**：IS/BS/CF 全公式预测（默认 3 年历史 + 5 年预测，可在配置中调整），BS 由 FIN 页 revolver/现金 sweep/理财承接严格配平（配平差额恒 <0.01）
- **FIN 债务调度页**：资金缺口 → revolver 新增借款；超额现金 sweep 还债 → 溢出购买理财；货币资金=最低现金，杜绝"高现金+高借款"并存。利息=平均余额×利率的循环链在表内**迭代展开 4 轮**（全簿无循环引用，无需开启迭代计算，LibreOffice headless 一次重算即收敛，残差<0.01）
- **WACC 做实**：可比公司 βl 按各自 D/E 与税率 unlever 取中位 → 按市值口径目标结构 relever（Hamada）；`采用值=IF(override="",计算值,override)`，下游全部引用采用值
- **DCF**：FCFF + Gordon 终值，年中折现（t=0.5/1.5/.../4.5），EV→Equity 桥含货币资金/交易性金融资产/其他权益工具投资/有息负债
- **三情景**：熊/基/牛；情景开关贯穿分部增速与毛利率；基准情景直接引用主模型输出，熊/牛为简化净利率法并附桥接说明
- **相对估值**：可比公司表（市值/PE/一致预期/PEG）+ 目标 PE 带定价
- **Sensitivity**：WACC×g 矩阵（轴公式化自动对中）、归母增速×PE 矩阵、DPO 单维敏感性
- **Checks 校验页**：11 项自动校验（BS 配平/现金勾稽/权益滚动/收入一致性/毛利率区间/DCF>0/矩阵中心/WACC 逻辑/现金>0/债务≥0/迭代残差），汇总布尔
- **每个假设保留"依据"文字字段**（硬性要求）：配置中每条假设均可附 basis，写入 Assumptions 页备注列

## 16 张工作表结构

| 顺序 | Sheet | 内容 |
|---|---|---|
| 1 | Cover | 标的摘要、关键市场数据、模型输出、建模注记、数据来源、免责声明 |
| 2 | Summary | Football field：各方法估值区间 vs 现价，文本条形可视化 |
| 3 | Assumptions | 驱动总表（11 节：全局/一致预期/分部量价/分部毛利率/费用率税率/营运资本/Capex折旧/股利/债务融资/WACC/情景参数） |
| 4 | Revenue_Segments | 分业务收入拆解（量×价驱动，受情景开关调整）+ 基准情形演算 |
| 5 | IS | 利润表（历史+预测全公式） |
| 6 | BS | 资产负债表（历史尾差自动并入"其他权益项目"清零） |
| 7 | CF | 现金流量表（间接法；利息重分类至筹资；投资含理财净变动） |
| 8 | Schedules | 营运资本（DSO/DIO/DPO 天数驱动）/无形资产摊销/股利 |
| 9 | PPE | 固定资产/在建工程滚动（转固率、分档折旧率、隐含折旧年限校验） |
| 10 | FIN | 债务与融资调度（revolver/sweep/理财 + 4 轮迭代展开 + 收敛残差） |
| 11 | Equity_Roll | 所有者权益逐项滚动（盈余公积计提、稀释股数） |
| 12 | DCF | FCFF、年中折现、Gordon 终值、EV→Equity 桥、隐含倍数 |
| 13 | Relative_Val | 可比公司 + β unlever/relever（L-O 列）+ 目标 PE 定价 |
| 14 | Sensitivity | WACC×g / 增速×PE / DPO 三张敏感性矩阵 |
| 15 | Scenarios | 熊/基/牛三情景估值与汇总对比 |
| 16 | Checks | 11 项自动校验 + 汇总布尔 |

## 快速上手

```bash
pip install -r requirements.txt          # openpyxl / pyyaml

# 方式一：全手工配置（推荐, 精度最高） — 复制范例改写
cp configs/300476.yaml configs/600xxx.yaml  # 编辑公司/分部/假设/可比/情景
python build_model.py --code 600xxx
python verify_model.py --code 600xxx         # LibreOffice重算+逐项校验

# 方式二：自动兜底（一行命令, 全自动拉公开数据建粗模）
python fetch_data.py --code 002463           # 东财F10三表+主营构成+腾讯行情 → configs/002463.yaml
python build_model.py --code 002463
python verify_model.py --code 002463

# 配置文件缺失或缺 segments/hist 时, build_model.py 也会自动走兜底
python build_model.py --code 601138          # 无需任何准备, 直接出模型
```

产物默认在 `out/<代码>_<名称>_估值模型.xlsx`（另附同名 `.addr.json` 单元格地址索引，供验收/二次开发）。`examples/` 收录了胜宏科技(300476)与沪电股份(002463)两个完整样例及验收日志。

## 配置字段说明（configs/<code>.yaml）

所有"假设条目"支持三种写法：`key: 标量`、`key: [5个预测年列表]`、`key: {value(s): ..., basis: "依据文字"}`。标量自动按预测年广播。

| 段 | 关键字段 | 说明 |
|---|---|---|
| `company` | code / code_full / name | 标的标识 |
| `model` | hist_years / fcst_years / valuation_date / build_date | 年份区间与基准日（3+5 可改） |
| `market` | price / shares / pe_ttm | 现价、总股本（百万股）、PE-TTM（均带依据） |
| `consensus` | rev / np（前3个预测年） | 一致预期（gildata/聚源等；无则自动外推占位并注明） |
| `latest_quarter` | label / rev / np | 最新季报实绩（可选） |
| `segments` | 列表：`key/name/short/driver(vol_asp或growth)/hist_share/hist_gm/vol/asp/gm/logic` | 分部业务；`vol_asp`=量×价驱动，`growth`=收入增速驱动 |
| `opex` | tax_add_rate / sale_rate / adm_rate / rd_rate / oth_op / nonop / tax_rate / oth_rate | 费用率/其他损益/有效税率 |
| `working_capital` | dso / dio / dpo / pre_rate / staff_rate / taxp_rate | 营运资本天数与比率 |
| `capex` | capex_rate / trans_rate / dep_rate / dep_new_rate / disp_rate / amort_rate | 资本开支占比路径与折旧摊销参数 |
| `dividend` | payout / surplus_rate | 股利支付率（按上年归母）、盈余公积计提率 |
| `financing` | min_cash_pct / rep_st / rep_cur / rep_lt / rep_lease / rate_* / cash_yield / fin_oth | FIN 页调度输入（四档债务计划还款与利率） |
| `wacc` | rf / erp / srp / kd / tg / override / wd_basis | WACC 组件；无可比公司时用 `beta_unlevered_input` 蓝色输入兜底 |
| `scenarios` | bear / base / bull：`rev_adj / npm_adj / logic` | 情景参数（驱动情景开关与 Scenarios 页） |
| `hist` | is / bs / cf / ppe_split / notes | 历史三表（百万），`fetch_data.py` 可自动生成；BS 尾差构建时自动清零 |
| `relative_val` | target_pe_lo / comps 列表（name/code/mcap/pe_ttm/np_f0/np_f1/src/beta_l/d_e/tax） | 可比公司；`beta_l` 为空者不进 βu 中位数 |
| `sensitivity` | pe_list / np_growth_list / highlight / dpo_deltas / matrix2_note | 敏感性矩阵参数 |
| `references` / `cover` | 文本列表 | Revenue_Segments 参考信息块、Cover 注记/来源 |

## 数据口径

- **东财 F10**（利润表/资产负债表/现金流量表/主营构成）：一律经**系统 curl** 抓取（`fetch_data.py` 内 `subprocess` 调 curl，python 不直连东财，避免反爬）；单位统一换算为人民币百万元
- **腾讯行情**（`qt.gtimg.cn`）：现价/总市值/PE-TTM；总股本=总市值/现价
- **一致预期**：工具不直连付费源；有 gildata/聚源等渠道时查实后手工填入 `consensus` 段（依据字段注明来源与日期），缺省时自动按最近年报增速外推占位
- **可比公司 βl/D/E**：分析师蓝色输入（参考行情终端 β 与最新年报杠杆），可按终端数据更新
- 历史 BS 的"其他流动资产/其他非流动资产/其他流动负债/其他非流动负债/其他权益项目"为轧差项（=合计-明细），保证历史严格配平；0.1 级尾差构建时并入"其他权益项目"清零

## 已知局限

- **兜底模式精度**：自动分段依赖东财"主营构成（按产品）"披露粒度——很多公司只披露单一产品（如"PCB 制造"占 95%+），此时模型退化为单段/少段收入驱动；费用率/天数/税率取最近年报持平或简单退坡，**不等于分析师判断**，务必在配置中按研究修正（每个自动假设的依据字段都标了"自动推导"）
- 少数股东损益/权益按持平滚动，归母=净利润（有重要少数股东权益的公司需手工处理）
- 预测期资产减值/投资收益等非经常项简化为固定小额，以保证三表严格配平
- 理财净变动简化：缺口年不赎回；处置固定资产按账面值回收无损益
- 年中折现为估值基准日与财年起点的标准近似，不做 stub 调整
- 沪电示例中 2028E 一致预期绝对值取自 gildata 表（其表内 2028 同比列与绝对值口径不一致，以绝对值为准）

## 免责声明

本工具仅用于研究学习。所有预测基于公开信息与主观假设，不构成任何投资建议或证券买卖要约。历史数据虽经核对来源，仍可能存在口径差异或错误；预测存在重大不确定性。使用者应自行判断并承担风险。
