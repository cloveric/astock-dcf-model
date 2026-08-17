<div align="center">

# astock-dcf-model

**配置驱动的 A 股机构级三表联动估值模型生成器: 一行命令, 产出可用 LibreOffice 机器验收的 DCF/FCFE 双视图模型。**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](./LICENSE)
[![Python](https://img.shields.io/badge/python-3.9%2B-blue.svg)](https://www.python.org/)
[![Checks](https://img.shields.io/badge/checks-11%2F11%20gating%20passing-brightgreen.svg)](#验证记录)
[![CI](https://github.com/cloveric/astock-dcf-model/actions/workflows/ci.yml/badge.svg)](https://github.com/cloveric/astock-dcf-model/actions)

</div>

---

## 价值主张

A 股卖方与买方的 Excel 估值模型长期停留在"手工坊"状态: 利润表预测与资产负债表脱节、配平靠手填轧差、融资结构缺失导致巨额现金与巨额借款并存、WACC 凭经验拍数、预测期硬编码使模型不可复用也不可审计。

本工具将建模纪律固化为工程约束: **配置驱动**(个股差异全部收敛到一个 YAML 文件)、**预测期零硬编码**(3 年历史 + 5 年预测全部公式生成)、**每条假设强制携带依据字段**、**三表严格配平**(差额全期恒为零)、**revolver 融资闭环**(利息循环链在表内迭代展开, 无循环引用)、**WACC 可比公司实算**, 并且生成的每一张工作簿都可以用 LibreOffice headless 重算做机器验收——模型对不对, 不需要肉眼复核, Checks 页给出布尔答案(11 项门控 + 毛利率信息行)。

与手工坊的对照:

| 维度 | 手工坊 Excel | 本工具 |
|---|---|---|
| 三表关系 | 利润表单飞, BS 手填轧平 | IS/BS/CF 全公式联动, 差额恒 = 0 |
| 融资结构 | 现金与借款并存, 利息脱钩 | FIN 页 revolver + 现金 sweep + 理财承接, 利息 = 平均余额 × 利率 |
| 循环引用 | 开迭代计算或硬填利息 | 循环链表内展开 4 轮, 全簿无循环引用, LO 一次重算收敛 |
| WACC | 8%/10% 拍脑袋 | 可比 βl unlever 取中位 → 目标结构 relever, 采用值 = IF(override="", 计算值, override) |
| 预测期 | 敲死数字 | 全部由假设驱动的公式, 改一个假设全簿重算 |
| 假设依据 | 散落在分析师脑子里 | 每条假设带 basis 字段, 写入 Assumptions 备注列; 可再被 dr 档案 grounding |
| 验收 | 肉眼 + 经验 | Checks 页 11 项门控公式校验(另有毛利率信息行) + LibreOffice 重算布尔总闸 |
| 估值口径 | 单一 DCF | FCFF 与 FCFE 双视图并列, 相对估值 + 三情景 + 敏感性矩阵交叉验证 |
| 换标的 | 重做一遍 | 换一个 6 位代码 |

## 架构

```
数据层   东财F10(三表/主营构成, 系统curl)  东财HKF10(港股三表)  腾讯行情(现价/市值/PE, 含hkXXXXX)  一致预期文件(--consensus)
            │                                        │
            ▼                                        │
       fetch_data.py  ── 兜底推导(缺配置时) ──┐        │
                                            ▼        ▼
配置层   configs/<code>.yaml   公司/分部量价/费用率/税率/营运资本/Capex/股利/融资/WACC/情景
                                            │  每条假设带依据
            ┌───────────────────────────────┼─────────────────────────────┐
            ▼                               ▼                             ▼
研究层   --dr 档案.md → 量化结论回填依据列   --announcements → 公告要点      --llm → 研究备忘录
            │ (DR研究/研究摘要两页存档)      (系统curl东财)                 (本机claude/codex, 默认off)
            ▼
构建层   build_model.py  openpyxl 全公式生成 17 张工作表 + 15 个 named ranges + addr.json 地址索引
            │
            ▼
验收层   verify_model.py  LibreOffice headless 重算 → Checks 11 项门控校验逐项验收 → 布尔总闸
```

三层可独立使用: 手工编写配置可完全绕过数据层(精度最高的用法); 研究层四个开关全部默认关闭, 不传任何研究参数时输出与基础模型逐项一致。

## 工作表清单 (17 + 2)

| # | Sheet | 职能 |
|---|---|---|
| 1 | Cover | 标的摘要、关键市场数据、模型输出、建模注记、数据来源、免责声明 |
| 2 | Summary | Football field: FCFF/FCFE/三情景/相对估值区间 vs 现价, 文本条形可视化 |
| 3 | Assumptions | 驱动总表 11 节(全局/一致预期/分部量价/分部毛利率/费用率税率/营运资本/Capex 折旧/股利/债务融资/WACC/情景), 依据列可被 dr 档案回填 |
| 4 | Revenue_Segments | 分业务收入拆解(量×价驱动, 受情景开关调整) + 基准情形独立演算 |
| 5 | IS | 利润表(历史 + 预测全公式) |
| 6 | BS | 资产负债表(历史尾差自动并入"其他权益项目"清零) |
| 7 | CF | 现金流量表(间接法; 利息重分类至筹资; 投资含理财净变动) |
| 8 | Schedules | 营运资本(DSO/DIO/DPO 天数驱动)/无形资产摊销/股利 |
| 9 | PPE | 固定资产/在建工程滚动(转固率、分档折旧率、隐含折旧年限校验) |
| 10 | FIN | 债务与融资调度: revolver/sweep/理财 + 利息循环链 4 轮迭代展开 + 收敛残差 |
| 11 | Equity_Roll | 所有者权益逐项滚动(盈余公积计提、稀释股数) |
| 12 | DCF | FCFF、年中折现、Gordon 终值、EV→Equity 桥、隐含倍数 |
| 13 | FCFE | 股权自由现金流 = 归母 + D&A − ΔNWC − Capex + 净新增借款(FIN), 按 Ke 折现, 与 FCFF 并列对照; 终值取正常化FCFE(净新增借款按g封顶), 附理财净变动勾稽行与两法差异原因 |
| 14 | Relative_Val | 可比公司 + β unlever/relever + 目标 PE 定价; 研究层激活时附一致预期目标价对照行 |
| 15 | Sensitivity | WACC×g / 增速×PE / DPO 三张敏感性矩阵(轴为公式自动对中) |
| 16 | Scenarios | 熊/基/牛三情景估值与汇总对比 |
| 17 | Checks | 11 项门控校验(含 named range 存在性) + 毛利率信息行 + 汇总布尔 |
| 18* | DR研究 | `--dr` 激活: 深度研究档案全文存档, 量化结论可追溯 |
| 19* | 研究摘要 | 研究层激活: 假设项 × 采用值 × 来源 × 置信度对照表 + LLM 备忘录页眉 + 公告要点 |

## 快速上手

```bash
git clone https://github.com/cloveric/astock-dcf-model.git
cd astock-dcf-model
pip install -r requirements.txt        # 核心依赖(openpyxl锁区间/pyyaml); Web模式装 requirements-web.txt; 跑测试装 requirements-dev.txt; 验收环节需本机 LibreOffice
```

最小命令:

```bash
python build_model.py --code 601138    # 无配置时自动拉公开数据走兜底建模
```

完整范例(胜宏科技, 配置已收录):

```bash
python build_model.py --code 300476
python verify_model.py --code 300476   # LibreOffice 重算验收, 基准 DCF 每股 340.4152 元
```

换标的:

```bash
python fetch_data.py --code 002463     # 东财三表+主营构成+腾讯行情 → configs/002463.yaml
python build_model.py --code 002463
python verify_model.py --code 002463
```

精度最高的用法是复制 `configs/300476.yaml` 手工改写——配置即研究底稿, 依据字段写清来源与日期。

产物写入 `out/`: `<代码>_<名称>_估值模型.xlsx`(全公式, Excel/WPS/LibreOffice 均可重算)与同名 `.addr.json`(单元格地址索引, 供验收与二次开发定位)。`examples/` 收录胜宏、沪电两个成稿及验收日志。

## 研究层 (默认全部关闭)

四个开关按来源优先级排列, 任一不可用即优雅降级, 不影响建模主流程:

```bash
python build_model.py --code 300476 \
    --dr examples/research/dr_300476.md \          # 深度研究档案: 量化结论按关键词回填
    --consensus examples/research/consensus_300476.json \   # 聚源/gildata 一致预期(含目标价)
    --announcements \                              # 东财业绩预告/快报最新一期要点(系统curl)
    --llm auto                                     # 本机 claude/codex CLI 生成研究备忘录
```

- `--dr <档案.md>`: 章节化解析研究档案, 提取含数字与单位的量化结论, 按关键词匹配回填到 Assumptions 依据列(标注 `dr档案§章节号`, 只追加不改值); 工作簿新增 **DR研究**(全文存档)与 **研究摘要** 两页。
- `--consensus <json/csv>`: 一致预期文件(2026-28E 营收/归母/目标价, 单位: 百万元/元), 覆盖配置 `consensus` 段并在 Relative_Val 与 Checks 各加一行目标价一致性对照。工具不直连付费源: 有 MCP gildata 工具的环境可由调用方拉取后写成该文件再传入, 格式见 `examples/research/consensus_300476.json`。
- `--announcements`: 东财数据中心业绩预告/业绩快报(系统 curl, python 不直连东财), 最新一期要点进研究摘要。
- `--llm auto|claude|codex|off`: 本机存在对应 CLI 时生成一段研究备忘录写入研究摘要页眉; 不存在或调用失败仅记录降级说明, 默认 `off`。

## Web 模式

本机任务制 Web 服务(FastAPI + 自包含单页前端, 零构建、无 Node 依赖):

```bash
python -m web.server          # http://127.0.0.1:8000 (HOST/PORT 环境变量可改)
```

- 提交表单: 证券代码(6 位 A 股 / 5 位港股) + 可选配置文件路径 + 研究层四个开关(dr/consensus/announcements/llm); 任务列表自动轮询进度, 进行中任务的日志尾部自动轮询展开, 完成后出现下载按钮并展示 LibreOffice 验收摘要(本机无 soffice 时注明跳过), 失败可查看构建日志尾部;
- 接口: `POST /api/jobs` 提交, `GET /api/jobs` 历史列表, `GET /api/jobs/{id}` 详情(含日志尾部/验收摘要), `GET /api/jobs/{id}/download` 下载 xlsx, `DELETE /api/jobs/{id}` 删除任务并清理产物(进行中任务 409 拒绝);
- 实现纪律: 建模逻辑零重写——单 worker 线程串行调用 `build_model.py` 子进程; 任务状态落盘 `web/.data/jobs.json`(原子写), 产物与日志存 `web/.data/out/<任务id>/`; 任务保留上限 100 个, 超限淘汰最早完成的任务并同步清理产物目录; 服务重启时未完成任务标记为失败(中断); config/dr/consensus 仅接受仓库内已存在文件, 防路径穿越;
- **安全声明: 服务无鉴权, 仅供本机/可信网络使用**; 需要对外暴露时请自行加反向代理鉴权, 勿直接绑定公网地址。

## Docker

```bash
docker build -t astock-dcf-model .
docker run -p 8000:8000 astock-dcf-model        # 打开 http://127.0.0.1:8000
```

镜像基于 `python:3.12-slim`, 预装 curl(数据层只走系统 curl)与 libreoffice-calc(容器内可做 `verify_model.py` 重算验收); 容器内亦可直接执行 `python build_model.py --code 300476`。

## 港股支持 (5 位代码)

```bash
python fetch_data.py --code 00981     # 腾讯hk行情 + 东财HKF10三表(IFRS) → configs/00981.yaml
python build_model.py --code 00981
python verify_model.py --code 00981
```

港股路径: 行情走腾讯 `hkXXXXX`(港元现价/总市值/PE-TTM, 字段版式与 A 股一致); 财务走东财 HKF10 datacenter 接口(系统 curl, IFRS 科目映射为模型 hist 结构)。已收录 `configs/00981.yaml`(中芯国际, 自动生成后按研究修正)与 `examples/00981_verify.txt` 验收日志。

口径与局限(使用前必读):

- **币种**: 模型内部一律用财报币种百万(如中芯国际为美元); 配置同时提供 `market.price_hkd`(港元原始现价)与 `market.fx_hkd`(港元/财报币种汇率, 美元报告主体 7.80、人民币报告主体约 1.085)时, 构建现算 `price = price_hkd / fx_hkd`——**改 fx_hkd 即生效**; 仅有 `market.price` 时按生成时折算快照使用; PE-TTM 为港元行情口径, 仅供对照;
- **现金流量表折算**: 东财 HKF10 现金流量表仅人民币口径, 按"期末现金 ÷ BS 现金及等价物"的隐含汇率逐年折算回财报币种(各年依据列注明), 因此 CF 期末现金与 BS 严格勾稽, 但流量项存在期末汇率近似;
- **IFRS 科目映射**: 无税金及附加/法定盈余公积/一年内到期非流动负债单列, 使用权资产并入物业厂房及设备; 各年"其他"科目为轧差项, 历史严格配平;
- **单段收入**: 港股无东财"主营构成(按产品)"披露, 兜底为整体单段(增速=总收入 YoY 退坡), 务必按研究拆分修正;
- **市值口径**: 港股总市值 = 全部股本 × 港元价, 对 A+H 两地上市公司与实际加权市值存在差异;
- **偿债假设**: 兜底"余额 1/5 逐年摊还"会使重资产扩产标的 FCFE 机制性深负(期初存量现金被一次性 sweep 偿债亦然); `configs/00981.yaml` 已按公司实际(有息负债稳定、现金为资本开支储备)修正为滚动续作 + 最低现金 60%, 换标的时按研究修正;
- **无涨跌停**: 港股无单日涨跌停限制(A 股 ±10%/20%), 口径差异已在 Cover 页注明, 不影响模型公式。

## 配置规范

所有建模判断集中在 `configs/<code>.yaml`。铁律: **每个假设必须写依据**。条目三种写法:

```yaml
tax_rate: 0.15                                              # 标量, 按预测年广播
dso: [88, 85, 82, 80, 78]                                   # 5 个预测年列表
capex_rate: {value: 0.045, basis: "公司指引+近三年均值"}      # 推荐: 值 + 依据
```

| 段 | 关键字段 | 说明 |
|---|---|---|
| `company` | code / code_full / name | 标的标识 |
| `model` | hist_years / fcst_years / valuation_date / build_date / mi_share(可选) | 年份区间与基准日(历史 ≥2 年、预测 3–5 年, 超范围构建时报错); `mi_share` 覆盖少数股东损益占比(默认按历史均值自动推导) |
| `market` | price / shares / pe_ttm; 港股另有 price_hkd / fx_hkd | 现价、总股本(百万股)、PE-TTM(均带依据); 港股同时给出 `price_hkd` 与 `fx_hkd` 时, 构建按 `price_hkd/fx_hkd` 现算财报币种现价(改 fx 即生效) |
| `consensus` | rev / np(前 3 个预测年) | 一致预期; 可被 `--consensus` 文件覆盖; 无则自动外推占位并注明 |
| `latest_quarter` | label / rev / np | 最新季报实绩(可选) |
| `segments` | key / name / **short(必填, 行标签用简称)** / driver(`vol_asp` 或 `growth`) / hist_share / hist_gm / vol / asp / gm / logic | 分部业务驱动; 无按产品构成披露时 fetch 自动落为单一"整体"分部 |
| `opex` | sale_rate / adm_rate / rd_rate / tax_rate / oth_op / nonop 等 | 费用率/其他损益/有效税率 |
| `working_capital` | dso / dio / dpo / pre_rate / staff_rate / taxp_rate | 营运资本天数与比率 |
| `capex` | capex_rate / trans_rate / dep_rate / dep_new_rate / disp_rate / amort_rate | 资本开支与折旧摊销 |
| `dividend` | payout / surplus_rate | 股利支付率(按上年归母)、盈余公积计提率 |
| `financing` | min_cash_pct / rep_st / rep_cur / rep_lt / rep_lease / rate_* / cash_yield | FIN 页调度输入 |
| `wacc` | rf / erp / srp / kd / tg / override / wd_basis | WACC 组件; 无可比公司时 `beta_unlevered_input` 兜底 |
| `scenarios` | bear / base / bull: rev_adj / npm_adj / logic | 情景参数 |
| `checks` | gm_band(可选, 默认 [0, 0.6]) | 毛利率信息行区间(0.4.0 起该项为信息展示, 不参与 Checks 布尔闸门) |
| `hist` | is / bs / cf / ppe_split / notes | 历史三表(百万元), `fetch_data.py` 可自动生成 |
| `relative_val` | target_pe_lo / comps | 可比公司; `beta_l` 为空者不进 βu 中位数 |
| `sensitivity` | pe_list / np_growth_list / highlight / dpo_deltas | 敏感性矩阵参数 |
| `references` / `cover` | 文本列表 | 参考信息块与 Cover 注记 |

## 方法论

- **FCFF + Gordon 终值**: 企业自由现金流按 WACC 折现, 终值 = FCFF₅ × (1+g)/(WACC−g); WACC×g 双维敏感性兜底。
- **FCFE 双视图**: 股权自由现金流 = 归母净利 + D&A − ΔNWC − Capex + 净新增借款(逐年反映 FIN 页实际债务调度), 按 Ke 折现直接得归母股权价值(少数股东权益在 DCF 页桥内扣除); 终值取正常化 FCFE(净新增借款按 g 封顶, 防显性期末大额净融资被永续外推); 显性期附"理财净变动"勾稽参考行(不计入), 两法差异表内可勾稽, 差异注记写在 FCFE 页底部。
- **年中折现**(mid-year convention): t = 0.5 / 1.5 / … / 4.5, 承认现金流年内均匀发生。
- **EV → Equity 桥**: 企业价值 − 有息负债 − 少数股东权益 + 货币资金 + 交易性金融资产 + 其他权益工具投资, 逐项列示。
- **revolver 配平**: 资金缺口 → revolver 新增借款; 超额现金 sweep 还债 → 溢出购买理财; 货币资金恒等于最低现金, 杜绝"高现金 + 高借款"并存。利息 = 平均余额 × 利率的循环链在 FIN 页表内迭代展开 4 轮(实测残差 ≤ 1e-6), 全簿无循环引用。
- **WACC 实算**: 可比公司 βl 按各自 D/E 与税率 unlever 取中位, 按目标结构 relever(Hamada); 采用值 = IF(override="", 计算值, override), 全链路可追溯。
- **三情景**: 熊/基/牛开关贯穿分部增速与毛利率; 基准直接引用主模型输出, 熊/牛用简化净利率法并附桥接说明。
- **Named ranges 审计轨迹**: WACC 采用值、永续 g、Ke、Kd、rf、ERP、βu、税率、股利支付率、最低现金占比、情景开关、稀释股数、DCF/FCFE 每股共 15 个关键驱动建立命名区域, FCFE 页公式直接引用名称, Checks 末项校验其存在性。

## 数据口径

- **东财 F10**(三表/主营构成): 一律经系统 curl 子进程抓取(python 不直连东财), 单位统一换算为人民币百万元; 历史年报年按当前月份自动推导(披露截止 4/30: 5 月起最新=上年, 否则=前年);
- **东财 HKF10**(港股三表): datacenter 长表接口, 系统 curl, 循环分页拉全; IFRS 科目映射, 金额为财报币种(详见"港股支持"节的口径说明); BS 现金缺失或隐含汇率异常([3,12] 外)时显式报错, 拒绝静默折算;
- **腾讯行情**(qt.gtimg.cn): 现价/总市值/PE-TTM(A 股与港股 hkXXXXX 同版式); 总股本 = 总市值 / 现价; 现价为 0(停牌/无行情)时拒绝兜底建模;
- **一致预期**: 不直连付费源。`--consensus` 文件输入(聚源/gildata 口径), 或查实后手工填入配置 `consensus` 段(依据注明来源与日期), 缺省时按最近年报增速外推占位并标注"自动推导";
- **可比公司 βl / D/E**: 分析师输入项(参考行情终端 β 与最新年报杠杆);
- 历史 BS 的若干"其他"科目为轧差项(= 合计 − 明细), 保证历史严格配平; 0.1 级尾差构建时并入"其他权益项目"清零。

## 验证记录

验收口径: LibreOffice headless 重算全簿后逐项读取 Checks 页与关键单元格; 环境 LibreOffice 26.2.5.2 / Python 3.14 / openpyxl 3.1.5。

| 标的 | 模式 | 基准 DCF 每股 | FCFE 每股 | BS 配平差额 | FIN 迭代残差 | Checks |
|---|---|---|---|---|---|---|
| 胜宏科技 300476 | 完整手工配置 | **340.415208186145 元** | 325.4508 元(−4.4%) | 全期 = 0.00 | ≤ 1.6e-6 | **11/11 门控 TRUE** |
| 沪电股份 002463 | fetch_data 全自动 | 112.2842 元³ | 108.3978 元(−3.5%) | 全期 = 0.00 | ≤ 4.6e-7 | **11/11 门控 TRUE** |
| 工业富联 601138 | 兜底(零配置) | 52.4484 元⁵ | 44.1898 元(−15.7%¹) | 全期 = 0.00 | < 0.01 | **11/11 门控 TRUE** |
| 中芯国际 00981.HK | HKF10 全自动 + 手工修正(偿债滚动/现金储备) | 1.6233 美元⁴ | 2.1614 美元(+33.1%²) | 全期 = 0.00 | ≤ 1.8e-6 | **11/11 门控 TRUE** |

¹ 0.4.0 起 FCFE 终值净借款按 `D×g` 对称正常化(不再单边封顶), 去杠杆还款计划不再被 Gordon 永续外推; 601138 两法差由 0.3.1 的 −57.4% 收敛至 −15.7%, 剩余差异见 FCFE 页底部注记。
² 港股重资产扩产标的: 最低现金按收入 60% 保留为资本开支储备(不参与 sweep); 0.4.0 终值对称正常化后 FCFE 由 0.6769 修正为 2.1614 美元, 高于 FCFF 主要因储备现金留存表内与净现金结构, 两法口径差异为机制性而非公式错误。
³ EV→Equity 桥扣少数股东权益(002463 末年报 mi=15.4 百万); 0.4.0 起历史归母净利润改用财报披露硬数(002463 历史存在少数股东亏损, 2025A 归母 3818.7→3822.3 百万), DCF 每股在 0.01 元内不变。
⁴ HKF10 分页修复+少数股东权益入桥的口径说明见 0.3.1 记录; 0.4.0 起利润表分列"合并/归母/少数股东损益", 00981 少数股东损益占比按历史均值 27.67% 自动推导, 2026E 归母净利润由 662.5 修正为 478.3 百万美元, EPS/FCFE/研究摘要同步归母口径(FCFF/EV 桥维持合并口径不变, DCF 每股不变)。
⁵ 兜底模式行情随取数日刷新, 与 0.3.1 记录的 52.4651 元差异来自现价变动, 非公式回归。

**复现口径(0.4.0)**: 重跑 300476 与 0.3.1 成稿逐单元格比对: DCF 每股精确复现 **340.415208186145 元**(黄金值断言按配置指纹+容差机制通过); 数值变化仅限两处口径修正——FCFE 终值链(325.4138→325.4508 元)与 DPO 敏感性非中心行(首年改按全额 COGS 重定价); 其余含少数股东损益行在内的新增行对 mi=0 标的数值恒等。(0.2.0/0.3.x 的历史零差异记录见 CHANGELOG。)

**幂等性**: 同一配置连续构建 10 次, 全部单元格(公式 + 输入值)签名完全一致(SHA-256 前缀 `f0b301fd28b8b85d`)。

**研究层验证**: `--dr --consensus --announcements` 全开构建 300476, 产出 19 张工作表, 9 行假设依据被 dr 档案回填, 研究摘要含一致预期/目标价/公告要点对照, LibreOffice 重算 Checks 门控全 TRUE, DCF 每股不变。

## 已知局限

- **兜底模式精度**: 自动分段依赖东财"主营构成(按产品)"披露粒度; 费用率/天数/税率取最近年报持平或简单退坡, 不等于分析师判断——自动假设的依据字段均标注"自动推导", 务必按研究修正;
- 少数股东损益按持平滚动(归母 ≈ 净利润); 少数股东权益已入 EV→Equity 桥扣减(0.3.1 起), 少数股东权益重大的公司(如 00981)建议按研究手工处理损益归属;
- 预测期资产减值/投资收益等非经常项简化为固定小额, 以保证三表严格配平;
- 理财净变动简化: 缺口年不赎回; 处置固定资产按账面值回收无损益;
- 年中折现为估值基准日与财年起点的标准近似, 不做 stub 调整;
- FCFE 在融资计划大幅加/去杠杆的标的上与 FCFF 机制性偏离, 属口径特征而非错误;
- **港股**: IFRS 科目映射存在轧差项、现金流量表按隐含汇率折算回财报币种、单段收入简化、市值为全股本×港元价口径——完整清单见"港股支持"节, 换标的时务必按研究修正配置。

## 工程化

- CI(`.github/workflows/ci.yml`)离线可跑: 语法检查 → pytest 离线单测(consensus合并回归/dr打分/web路径校验) → 用仓库内 `configs/300476.yaml` 构建 → `tests/smoke_check.py` 结构冒烟(17 表/15 个 named ranges/Checks 结构/addr 完整) → 双构建幂等性断言 → 环境有 LibreOffice 时追加全量验收(含 300476 黄金值断言: 配置指纹一致时按容差 max(1e-6,|golden|·1e-9) 复现 340.415208186145);- 贡献准则见 [CONTRIBUTING.md](./CONTRIBUTING.md), 安全披露见 [SECURITY.md](./SECURITY.md), 版本历史见 [CHANGELOG.md](./CHANGELOG.md), AI 协作者说明见 [AGENTS.md](./AGENTS.md)。

## 免责声明

本工具仅用于研究学习。所有预测基于公开信息与主观假设, **不构成任何投资建议或证券买卖要约**。历史数据虽经核对来源, 仍可能存在口径差异或错误; 预测存在重大不确定性。使用者应自行判断并承担风险。

## License

[MIT](./LICENSE) © astock-dcf-model contributors
