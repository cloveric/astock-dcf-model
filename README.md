<div align="center">

# astock-dcf-model

**一行命令，把一家 A 股公司变成一份美元基金级的三表联动 DCF 估值模型。**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python](https://img.shields.io/badge/python-3.9%2B-blue.svg)](https://www.python.org/)
[![Checks](https://img.shields.io/badge/checks-11%2F11%20passing-brightgreen.svg)](#-实测验证)
[![Platform](https://img.shields.io/badge/platform-A%E8%82%A1-red.svg)](#-快速上手)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](./LICENSE)

</div>

---

## 📑 目录

- [为什么做这个](#-为什么做这个)
- [功能亮点](#-功能亮点)
- [工作流程](#-工作流程)
- [16 张工作表结构](#-16-张工作表结构)
- [快速上手](#-快速上手)
- [配置指南](#-配置指南)
- [估值哲学](#-估值哲学)
- [数据口径](#-数据口径)
- [实测验证](#-实测验证)
- [已知局限](#-已知局限)
- [免责声明](#-免责声明)
- [License](#-license)

---

## 💡 为什么做这个

A 股研究圈有一个尴尬的现实：**Excel 估值模型满天飞，像样的开源机构级模型几乎没有。**

- **三表不配平** —— 利润表预测得天花乱坠，资产负债表差额靠手填数字硬轧平，现金流量表形同虚设；
- **硬编码满天飞** —— 预测期收入直接敲死数字，改一个假设要重算半天，模型不可复用、不可审计；
- **融资结构缺失** —— 公司账面同时挂着巨额现金和巨额借款，没有 revolver、没有现金 sweep，利息和负债完全脱钩；
- **WACC 拍脑袋** —— 8%、10% 随手一填，没有可比公司 β 的 unlever / relever 实算过程。

本工具就是为了终结这些"手工坊"做法：**配置驱动、全公式生成、三表严格配平、融资闭环调度、每个假设必须写明依据**，且生成的 Excel 可用 LibreOffice headless 重算做机器验收——模型对不对，跑一遍 Checks 页就知道。

## ✨ 功能亮点

- **⚖️ 三表严格配平**
  IS/BS/CF 全公式联动，BS 配平差额全期恒 < 0.01（实测为 0.00），拒绝任何手填轧差。

- **🏦 FIN 循环贷款资金滚动**
  资金缺口 → revolver 自动新增借款；超额现金 sweep 还债 → 溢出购买理财；货币资金恒等于最低现金，杜绝"高现金+高借款"并存。利息循环链在表内迭代展开 4 轮，全簿无循环引用，LibreOffice 一次重算即收敛（残差 < 1e-5）。

- **📐 WACC 可比公司实算，非拍脑袋**
  可比公司 βl 按各自 D/E 与税率 unlever 取中位，再按目标结构 relever（Hamada）；`采用值 = IF(override="", 计算值, override)`，全链路可追溯。

- **🔮 预测期零硬编码**
  3 年历史 + 5 年预测全部由驱动假设公式生成，改任何一个假设，16 张表自动重算。

- **📝 假设全部带依据**
  每条假设强制携带 `basis` 文字字段，写入 Assumptions 页备注列——模型可审计，数字有出处。

- **✅ LibreOffice 可重算校验**
  `verify_model.py` 用 LibreOffice headless 重算全簿并逐项验收 Checks 页 11 项校验（BS 配平/现金勾稽/权益滚动/迭代残差……），通过与否一目了然。

- **🔁 一键换股票**
  `fetch_data.py` 自动拉取东财 F10 三表 + 主营构成 + 腾讯行情生成配置；换标的只需换一个 6 位代码。

- **🎯 估值全家桶**
  DCF（FCFF + 年中折现）+ 相对估值（可比 PE 带）+ 敏感性矩阵 ×3 + 熊/基/牛三情景，Cover 页一键成稿。

## 🔄 工作流程

```mermaid
flowchart LR
    A["📊 数据层<br/>东财F10三表 · 主营构成<br/>腾讯行情 · 一致预期"] --> B["fetch_data.py<br/>自动抓取成配置"]
    B --> C["⚙️ 配置层<br/>configs/&lt;code&gt;.yaml<br/>每个假设带依据"]
    C --> D["build_model.py<br/>openpyxl 全公式生成"]
    D --> E["📑 产出<br/>16 表估值模型 xlsx<br/>+ addr.json 地址索引"]
    E --> F["verify_model.py<br/>LibreOffice 重算验收"]
    F -->|"11 项 Checks 全过"| G["✅ 可交付的机构级模型"]
    F -->|"任一不过"| C
```

也可以绕过自动抓取，手工编写配置（精度最高）：`配置 → 建模 → 校验 → Excel` 全链路不变。

## 📚 16 张工作表结构

| 顺序 | Sheet | 干什么 |
|---|---|---|
| 1 | Cover | 标的摘要、关键市场数据、模型输出、建模注记、数据来源、免责声明 |
| 2 | Summary | Football field：各方法估值区间 vs 现价，文本条形可视化 |
| 3 | Assumptions | 驱动总表（11 节：全局/一致预期/分部量价/分部毛利率/费用率税率/营运资本/Capex 折旧/股利/债务融资/WACC/情景参数） |
| 4 | Revenue_Segments | 分业务收入拆解（量×价驱动，受情景开关调整）+ 基准情形演算 |
| 5 | IS | 利润表（历史+预测全公式） |
| 6 | BS | 资产负债表（历史尾差自动并入"其他权益项目"清零） |
| 7 | CF | 现金流量表（间接法；利息重分类至筹资；投资含理财净变动） |
| 8 | Schedules | 营运资本（DSO/DIO/DPO 天数驱动）/无形资产摊销/股利 |
| 9 | PPE | 固定资产/在建工程滚动（转固率、分档折旧率、隐含折旧年限校验） |
| 10 | FIN | 债务与融资调度（revolver/sweep/理财 + 4 轮迭代展开 + 收敛残差） |
| 11 | Equity_Roll | 所有者权益逐项滚动（盈余公积计提、稀释股数） |
| 12 | DCF | FCFF、年中折现、Gordon 终值、EV→Equity 桥、隐含倍数 |
| 13 | Relative_Val | 可比公司 + β unlever/relever + 目标 PE 定价 |
| 14 | Sensitivity | WACC×g / 增速×PE / DPO 三张敏感性矩阵 |
| 15 | Scenarios | 熊/基/牛三情景估值与汇总对比 |
| 16 | Checks | 11 项自动校验 + 汇总布尔 |

## 🚀 快速上手

### 安装

```bash
git clone https://github.com/cloveric/astock-dcf-model.git
cd astock-dcf-model
pip install -r requirements.txt    # openpyxl / pyyaml
# 验收环节需本机装有 LibreOffice（soffice 命令可用）
```

### 最小命令：一行出模型

```bash
python build_model.py --code 601138
```

无需任何准备——配置缺失时自动拉取公开数据走兜底流程，直接产出工业富联的完整 16 表模型。

### 胜宏科技（300476）完整范例

```bash
# configs/300476.yaml 已收录完整手工配置，直接构建 + 验收
python build_model.py --code 300476
python verify_model.py --code 300476
```

输出：基准 DCF 每股价值 **340.42 元**，与原定制脚本逐值零差异，Checks 11/11 通过。

### 换一只自己的股票（推荐流程）

```bash
# 方式一：全自动拉公开数据建粗模
python fetch_data.py --code 002463   # 东财三表+主营构成+腾讯行情 → configs/002463.yaml
python build_model.py --code 002463
python verify_model.py --code 002463

# 方式二：复制胜宏范例手工改写（精度最高）
cp configs/300476.yaml configs/600xxx.yaml   # 编辑公司/分部/假设/可比/情景
python build_model.py --code 600xxx
python verify_model.py --code 600xxx
```

### 输出说明

产物默认写入 `out/`：

- `<代码>_<名称>_估值模型.xlsx` —— 16 表全公式模型，Excel / WPS / LibreOffice 均可打开重算；
- `<代码>_<名称>_估值模型.addr.json` —— 单元格地址索引，供验收脚本与二次开发定位。

`examples/` 收录胜宏科技(300476)、沪电股份(002463)两个成稿及验收日志，可直接打开对照学习。

## ⚙️ 配置指南

所有建模判断集中在 `configs/<code>.yaml`。**铁律：每个假设必须写依据**——每条假设条目支持三种写法：

```yaml
tax_rate: 0.15                                    # 标量，按预测年广播
dso: [88, 85, 82, 80, 78]                         # 5 个预测年列表
capex_rate: {value: 0.045, basis: "公司指引+近三年均值"}  # 推荐写法：值 + 依据
```

`basis` 会被写入 Assumptions 页备注列，让模型永远可审计。

| 段 | 关键字段 | 说明 |
|---|---|---|
| `company` | code / code_full / name | 标的标识 |
| `model` | hist_years / fcst_years / valuation_date / build_date | 年份区间与基准日（3+5 可改） |
| `market` | price / shares / pe_ttm | 现价、总股本（百万股）、PE-TTM（均带依据） |
| `consensus` | rev / np（前 3 个预测年） | 一致预期（gildata/聚源等；无则自动外推占位并注明） |
| `latest_quarter` | label / rev / np | 最新季报实绩（可选） |
| `segments` | key / name / driver(`vol_asp` 或 `growth`) / hist_share / hist_gm / vol / asp / gm / logic | 分部业务；`vol_asp`=量×价驱动，`growth`=收入增速驱动 |
| `opex` | sale_rate / adm_rate / rd_rate / tax_rate / oth_op / nonop … | 费用率/其他损益/有效税率 |
| `working_capital` | dso / dio / dpo / pre_rate / staff_rate / taxp_rate | 营运资本天数与比率 |
| `capex` | capex_rate / trans_rate / dep_rate / dep_new_rate / disp_rate / amort_rate | 资本开支占比路径与折旧摊销参数 |
| `dividend` | payout / surplus_rate | 股利支付率（按上年归母）、盈余公积计提率 |
| `financing` | min_cash_pct / rep_st / rep_cur / rep_lt / rep_lease / rate_* / cash_yield | FIN 页调度输入（四档债务计划还款与利率） |
| `wacc` | rf / erp / srp / kd / tg / override / wd_basis | WACC 组件；无可比公司时用 `beta_unlevered_input` 兜底 |
| `scenarios` | bear / base / bull：`rev_adj / npm_adj / logic` | 情景参数（驱动情景开关与 Scenarios 页） |
| `hist` | is / bs / cf / ppe_split / notes | 历史三表（百万元），`fetch_data.py` 可自动生成 |
| `relative_val` | target_pe_lo / comps 列表（name / code / mcap / pe_ttm / np_f0 / np_f1 / beta_l / d_e / tax） | 可比公司；`beta_l` 为空者不进 βu 中位数 |
| `sensitivity` | pe_list / np_growth_list / highlight / dpo_deltas | 敏感性矩阵参数 |
| `references` / `cover` | 文本列表 | Revenue_Segments 参考信息块、Cover 注记/来源 |

## 🧭 估值哲学

本工具遵循美元基金（美元 buy-side）的标准方法论，不发明轮子，但把每一步做实：

- **FCFF + Gordon 终值**：企业自由现金流折现，终值用永续增长率 g，WACC×g 双维敏感性兜底；
- **年中折现**（mid-year convention）：t = 0.5 / 1.5 / … / 4.5，承认现金流在年内均匀发生，比年末折现更贴近实务；
- **EV → Equity 桥**：企业价值 − 有息负债 + 货币资金 + 交易性金融资产 + 其他权益工具投资，逐项列示不藏猫腻；
- **revolver 配平**：资产负债表的平衡不靠手填，而靠债务调度页的融资闭环——缺多少钱借多少钱，多多少钱还多少债，利息=平均余额×利率，迭代至收敛；
- **相对估值带**：可比公司 PE + 目标 PE 区间定价，与 DCF 互为印证，Football field 一页看全；
- **三情景**：熊/基/牛开关贯穿分部增速与毛利率，基准情景直接引用主模型输出，熊/牛用简化净利率法并附桥接说明——承认预测的不确定性，而不是假装精确。

## 🗄️ 数据口径

- **东财 F10**（利润表/资产负债表/现金流量表/主营构成）：经系统 curl 抓取（`fetch_data.py` 内 `subprocess` 调 curl，python 不直连东财，避免反爬）；单位统一换算为人民币百万元；
- **腾讯行情**（`qt.gtimg.cn`）：现价/总市值/PE-TTM；总股本 = 总市值 / 现价；
- **一致预期**：工具不直连付费源；有 gildata/聚源等渠道时查实后手工填入 `consensus` 段（依据字段注明来源与日期），缺省时自动按最近年报增速外推占位；
- **可比公司 βl / D/E**：分析师输入项（参考行情终端 β 与最新年报杠杆），可按终端数据更新；
- 历史 BS 的"其他流动资产/其他非流动资产/其他流动负债/其他非流动负债/其他权益项目"为轧差项（= 合计 − 明细），保证历史严格配平；0.1 级尾差构建时并入"其他权益项目"清零。

## 🧪 实测验证

所有数字均可用 `examples/` 中的验收日志复现（LibreOffice headless 重算后逐项比对）：

| 标的 | 模式 | 结果 |
|---|---|---|
| **胜宏科技 (300476)** | 完整手工配置 | 基准 DCF **340.4152 元**，与原定制脚本**逐值零差异**；BS 配平差额全期 = 0.00；FIN 迭代残差 ≤ 1.6e-6；CF 期末现金与 BS 货币资金差异 ≤ 1e-12；WACC 实算 7.5620%（βu 中位 1.0108）；Checks **11/11 通过** |
| **沪电股份 (002463)** | fetch_data 全自动 | 基准 DCF 112.2922 元；BS 配平差额全期 = 0.00；FIN 迭代残差 ≤ 4.6e-7；WACC 实算 7.5637%；Checks **11/11 通过** |
| **工业富联 (601138)** | 兜底（5 行最小配置） | 无预设配置直接建模，全链路跑通出表 |

## ⚠️ 已知局限

- **兜底模式精度**：自动分段依赖东财"主营构成（按产品）"披露粒度——很多公司只披露单一产品（如"PCB 制造"占 95%+），此时模型退化为单段/少段收入驱动；费用率/天数/税率取最近年报持平或简单退坡，**不等于分析师判断**，务必在配置中按研究修正（每个自动假设的依据字段都标了"自动推导"）；
- 少数股东损益/权益按持平滚动，归母 = 净利润（有重要少数股东权益的公司需手工处理）；
- 预测期资产减值/投资收益等非经常项简化为固定小额，以保证三表严格配平；
- 理财净变动简化：缺口年不赎回；处置固定资产按账面值回收无损益；
- 年中折现为估值基准日与财年起点的标准近似，不做 stub 调整。

## 📜 免责声明

本工具仅用于研究学习。所有预测基于公开信息与主观假设，**不构成任何投资建议或证券买卖要约**。历史数据虽经核对来源，仍可能存在口径差异或错误；预测存在重大不确定性。使用者应自行判断并承担风险。

## 📄 License

[MIT](./LICENSE) © astock-dcf-model contributors
