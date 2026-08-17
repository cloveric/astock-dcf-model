# Changelog

本项目遵循 [Semantic Versioning](https://semver.org/lang/zh-CN/)。

## [0.4.0] - 2026-08-18

全量审计修复版: 基于 2026-08-17 多路对抗性审计(约 40 项确认发现)一次性修复估值口径、取数健壮性、校验器可信度、Web 服务与代码卫生五大类问题。三个收录标的重建后 LibreOffice 重算全部 ALL PASS; 300476 黄金值 340.415208186145 位级复现。

### 估值口径(数值有变, 属修正)
- **归母/合并分口径(重大)**: 利润表分列"净利润(合并)/归母净利润/少数股东损益"; 历史归母改用财报披露硬数, 预测按新增假设行"少数股东损益占净利润比"(`model.mi_share` 可覆盖, 默认历史均值 clamp[0,50%])归母化; EPS/Equity_Roll/FCFE/CF 起点/研究摘要全部切归母口径, FCFF/EV 桥维持合并口径。00981 少数股东占比 27.67%, 2026E 归母 662.5→478.3 百万美元(修复前虚高 ~44%); 002463 历史归母回贴披露值(存在少数股东亏损); 300476 mi=0 数值恒等。
- **FCFE 终值对称正常化**: 终值净新增借款由 `MIN(实际, D×g)` 单边封顶改为直接 `D×g` 稳态化, 去杠杆调度不再被 Gordon 永续外推(300476 FCFE 325.4138→325.4508 元; 601138 两法差 −57.4%→−15.7%; 00981 0.6769→2.1614 美元)。
- **DPO 敏感性首年口径**: 首预测年改按全额 COGS 重定价(历史应付为既成事实), 修复此前每档低估 δ/365×COGS_H0 的问题。
- **兜底一致预期改逐年复利**并统一除零守卫(旧实现对增长公司给出逐年下滑的"一致预期"); A股/HK 共享 `_derive_consensus`。
- **港股汇率**: 现价折算按报告币种映射(美元 7.80/人民币 1.085/港元 1.0), 配置新增 `market.price_hkd`, 构建现算 `price_hkd/fx_hkd`(改 fx 即生效); 隐含汇率守卫分币种放行(港币报告主体 ~0.92 不再被误拒)。

### 取数健壮性(拒绝静默错模型)
- HKF10 分页中途空响应改为显式报错(疑似 WAF), 不再静默截断; 新增按年份早停; 核心科目(营业额/股东权益/流动资产合计/期末现金)缺失即报错拒绝按 0 建模。
- A股 F10/主营构成接口空数据(银行/券商/保险等特殊模板)给出明确 RuntimeError, 不再裸 TypeError。
- 主营构成缺最新年: 份额回退至最近可用年并注明, 最新年份额全 0 直接报错; 无"按产品"构成时兜底为单一"整体"分部(不再产出空 segments 损坏工作簿)。
- `apply_fallback` 按缺失项精准补抓(修复整套重复抓取), 抓取并发化。

### 构建与校验可信度
- 移除一致预期行 `c_np26/c_rev26/c_np27` 等 7 处 2026 硬编码, 改按 fcst_years 动态寻址(修复 2027-05 起全线 KeyError 的定时炸弹); 版面列号(Scenarios 种子/Assumptions 依据列/报表页备注列)全部动态化, build 入口显式校验历史≥2 年、预测 3–5 年。
- verify_model: 修复 Python≤3.11 f-string 语法错误(README 承诺 3.9+ 现已成立); 黄金值改"配置 sha256 指纹门 + 容差 max(1e-6,|golden|·1e-9)"机制(配置刷新自动跳过, 不再误报); LibreOffice 重算加独立用户目录+产物新鲜度断言(修复桌面 LO 占用时验证旧文件假 PASS); 单元格读取全量 None/错误串守卫(不再 TypeError 中断报告); t 序行号经 addr `dcf_tidx` 传递; addr 缺年份元数据时报错而非用冻结默认; `--code` 多候选取 mtime 最新。
- Checks: 毛利率区间改为信息行(不入布尔闸门), 区间可经 `checks.gm_band` 配置——高毛利公司不再因历史事实被误判 FAIL; 门控 11 项。
- `sensitivity.dpo_deltas` 缺 0 自动补 0(修复 addr 写出 "Sensitivity!None" 崩溃); `relative_val.target_pe_lo` 容忍列表写法; `beta_unlevered_input` 被可比中位覆盖时打印警告。
- CI: LibreOffice 安装去静默(`apt-get update && install`, 失败即红), 黄金值门不再可能被悄悄跳过; smoke_check 缺键先打印全部失败再退出。

### 研究层
- consensus 合并取消"末值顺延填充": 文件覆盖年写文件值、config 有原值沿用、两者皆无才平推且 basis 如实标注(修复溯源造假); llm 备忘录检查 returncode(CLI 报错横幅不再混入工作簿); dr_report 关键词由 config segments 动态派生(去除胜宏产品词硬编码); sheets 样式复用注入。

### Web 服务
- worker 线程异常护栏(单次 OSError 不再永久瘫痪)+任务终态兜底; 符号链接部署下路径校验修复; 下载改持锁读 bytes 消除删除竞态; 单实例文件锁(拒绝 --workers>1 脑裂); 任务 ID 碰撞重生成; 提交时即做交易所语义校验(非法代码立即 400); 日志尾部 O(4KB) 读取; 列表接口去每任务 stat; 前端轮询活跃 2s/空闲 15s/隐藏暂停。

### 代码卫生
- 删除 `wb.calculation.iterate` 残留(不再掩盖未来意外循环引用)、死代码 `arow_e`/`crow('dash')` 分支; `srow_fill` 显式传表; 分部份额单次写入; A股/HK 兜底默认块与增长路径推导去重(`_common_defaults`/`_derive_growth_path`); `_m()` 统一。

### 验证
- 300476: DCF 340.415208186145 位级复现(黄金值断言 PASS), 除 FCFE 终值链与 DPO 非中心行外与 0.3.1 成稿数值零差异; 002463/00981/601138(零配置兜底, 走全新取数链路)构建+LO 重算 ALL PASS; pytest 16/16; examples 三标的成稿与验收日志已刷新。

## [0.3.1] - 2026-08-17

### 修复
- **BUG-1 一致预期缺年错位**: `research/consensus.py` 的 `_merge` 由按位置合并改为**按年份对齐**——槽位 = 配置 `model.fcst_years` 前 3 个预测年, 只写文件中出现的年份, 缺年保留配置原值, 文件年份与预测序列完全无交集则显式报错 (旧实现在文件只给部分年份时会写错槽位); 新增 5 个回归用例(三年全给/缺首年/缺中间年/全缺报错/部分交集)。
- **BUG-2 东财 HKF10 未分页**: `fetch_hk_long_table` 循环分页拉全(00981 资产负债表实测 8 页 3,641 行, 旧实现只取第 1 页导致最早年科目缺失、隐含汇率静默回退 fx=1.0 使该年 CF 混入人民币口径); `build_hist_hk` 新增守卫: BS"现金及等价物"缺失、或隐含汇率落在 [3,12] 外(≈1 除外)时显式报错, 拒绝静默折算; 修复后 00981 的 2023 年三表完整(CF 期末现金 44,019.4 → 6,215.1 百万美元, 与 BS 严格勾稽), 重新生成 `configs/00981.yaml`。
- **SEC-1 Web XSS**: 任务列表/详情中 `j.name`/`params`/`error` 等字段统一转义(沿用 log_tail 同法), 不再原样拼接 innerHTML。

### 新增
- **EV→Equity 桥加"减: 少数股东权益"行**(取 hist mi 账面值): 无少数股东权益的标的(如 300476, mi=0)结果零变化; 中芯国际(00981)NCI 135.8 亿美元入桥后 DCF 每股 3.21 → 1.62 美元(口径修正, 非回归); FCFE 页同步注记"FCFE 基于归母净利折现为归母股权价值, 少数股东权益已在 DCF 页桥内扣除"。
- **FCFE 终值正常化**: 终值改用正常化 FCFE(净新增借款按 g 封顶: `MIN(实际调度, 期末有息负债×g)`), 防止显性期末大额净融资被 Gordon 公式永续外推; 去杠杆/平稳年维持实际调度(300476 FCFE 每股 325.4138 不变)。
- **FCFE 显性期"理财净变动"勾稽参考行**(不计入 FCFE): 现金↔理财形态转换显性列示, 两法差异表内可勾稽。
- **verify_model.py 黄金值断言**: 300476 构建必须 DCF 每股 == 340.415208186145(容差 0), CI 的 LibreOffice 步自动触发。
- **pytest 离线单测** `tests/test_units.py`(13 项): consensus 合并回归 / dr annotate 打分 / web 路径穿越校验。
- **Web 运维**: 任务保留上限 100 个(超限淘汰最早完成任务并同步清理产物目录); `DELETE /api/jobs/{id}`(进行中任务 409 拒绝); 构建成功后本机有 LibreOffice 时自动附跑验收, 摘要随任务详情返回; 前端 running 任务日志尾部自动轮询展开, 完成后展示验收摘要; README 明示"无鉴权, 仅可信网络"。

### 变更
- **fetch_data 年份自动推导**: 历史年报年按当前月份推导(年报披露截止 4/30: 5 月起最新=上年, 否则=前年), 替换写死 [2023,2024,2025]; `fetch_quote` 新增 price=0 停牌守卫(拒绝兜底建模)。
- **requirements 分层**: `requirements.txt`(核心: openpyxl>=3.1,<3.2 锁区间 / pyyaml) / `requirements-web.txt`(+fastapi/uvicorn) / `requirements-dev.txt`(+pytest); Dockerfile 装 web 全套。
- `build_model.py` 工作表排序由私有 `wb._sheets` 改为公开 API `wb.move_sheet`; `verify_model.py` 列名推算改 `get_column_letter`。

### 验证
- **铁律复验**: 300476 重跑, DCF 每股精确等于 340.415208186145 元(黄金值断言 PASS), FCFE 每股 325.4138 不变; 与 0.3.0 成稿逐单元格比对: 除 DCF 页新增"减: 少数股东权益"行、FCFE 页新增"理财净变动/正常化FCFE"行及其引用行号联动外数值零差异; LibreOffice 重算 Checks 12/12 TRUE (ALL PASS); 研究层全开(`--dr --consensus`)复跑黄金值断言仍精确通过。
- **00981 重验**: 分页修复+手工修正口径回贴后重新生成配置, 构建+LO 重算 ALL PASS; BUG-2 修复单独复跑 DCF=3.2097(≈3.2101, 微量漂移为行情刷新); 叠加少数股东权益入桥后 DCF 每股 1.6233 美元 / FCFE 0.6769 美元(−58.3%), `examples/00981_*` 已刷新。
- **002463 复跑**: DCF 112.2922 → 112.2842 元(mi 15.4 百万入桥), FCFE 107.5540 不变, ALL PASS, `examples/002463_*` 已刷新。
- **Web 端到端**: 提交 300476 → done → 验收摘要(ALL PASS)随详情返回 → 下载 200 → DELETE 后记录与产物目录同步删除。

## [0.3.0] - 2026-08-17

### 新增
- **Web 模式(`web/`)**: FastAPI 后端 + 自包含单页前端(零构建、原生 fetch 轮询); `POST /api/jobs` 提交建模任务(代码 + config/dr/consensus/announcements/llm 开关), `GET /api/jobs` 历史列表, `GET /api/jobs/{id}` 进度与日志尾部, `GET /api/jobs/{id}/download` 下载 xlsx; 单 worker 线程串行调用 `build_model.py` 子进程(建模逻辑零重写), 任务状态落盘 `web/.data/jobs.json`, 产物存 `web/.data/out/`; `python -m web.server` 单命令启动。
- **Docker**: `Dockerfile`(python:3.12-slim + curl + libreoffice-calc, EXPOSE 8000, CMD 启动 Web 服务)与 `.dockerignore`。
- **港股支持(5 位代码)**:
  - `fetch_data.py`: 腾讯行情 `hkXXXXX` 直连(现价/总市值/PE-TTM); 东财 HKF10 datacenter 接口(系统 curl)拉取 IFRS 三表并映射为模型 hist 结构(2023-2025 年报); 单段收入兜底(港股无主营构成披露);
  - 币种口径: 模型内部用财报币种百万, 港元现价按配置 `market.fx_hkd` 折算; 现金流量表(人民币口径)按隐含汇率折算回财报币种, CF-BS 现金严格勾稽;
  - `build_model.py`: `company.unit` 可覆盖各页单位注记; HK 标的 Cover 页自动注明"无涨跌停/币种折算"口径(仅 HK 触发, A 股输出零差异);
  - 收录 `configs/00981.yaml`(中芯国际, HKF10 全自动 + 手工修正偿债滚动与现金储备口径)与 `examples/00981_verify.txt`。

### 验证
- 铁律复验: 重跑 300476, 基准 DCF 每股精确等于 340.415208186145 元, 与 0.2.0 成稿逐单元格零差异, LibreOffice 重算 Checks 12/12 TRUE(ALL PASS);
- Web 端到端实测: 启动服务 → 提交 300476 → 完成 → 下载 xlsx → LibreOffice 重算 ALL PASS;
- 港股实测: 00981 中芯国际构建 + LibreOffice 重算, Checks 12/12 TRUE, verify ALL PASS(DCF 3.2101 美元 / FCFE 0.6766 美元, 两法差异为口径特征, 见 README 港股节);
- Dockerfile 经静态检查(本机无 docker 环境, 未实际 build)。

## [0.2.0] - 2026-08-17

### 新增
- **FCFE 页(第 13 张核心报表)**: FCFE = 归母净利 + D&A − ΔNWC − Capex + 净新增借款(引用 FIN 页调度), 按 Ke 年中折现, 与 FCFF 结果并列对照, 附两法差异原因注记; Summary 页新增 FCFE 行。
- **Named ranges 审计轨迹**: WACC 采用值/计算值、永续 g、Ke、Kd、rf、ERP、βu、税率、股利支付率、最低现金占比、情景开关、稀释股数、DCF/FCFE 每股共 15 个关键驱动建立命名区域; Checks 页新增第 12 项"named range 存在性"校验。
- **研究层 `research/`(全部默认关闭, 不影响旧路径)**:
  - `--dr <档案.md>`: 深度研究档案量化结论回填 Assumptions 依据列(标注 `dr档案§章节号`), 新增 **DR研究**(全文存档)与 **研究摘要**(假设×采用值×来源×置信度对照表)两页;
  - `--consensus <json/csv>`: 聚源/gildata 一致预期文件输入(营收/归母/目标价), 覆盖 consensus 段并写入研究摘要, Relative_Val 与 Checks 各增一行目标价一致性对照;
  - `--announcements`: 东财业绩预告/业绩快报(系统 curl)最新一期要点进研究摘要;
  - `--llm auto|claude|codex|off`: 本机存在对应 CLI 时生成研究备忘录写入研究摘要页眉, 缺失时优雅降级(默认 off)。
- `verify_model.py` 输出 FCFE 双视图对照与 named range 摘要; 两法偏差 ≥30% 时给出 INFO 提示(不去杠杆标的的融资计划差异属机制性)。
- 示例研究层 fixture: `examples/research/dr_300476.md`、`examples/research/consensus_300476.json`。
- 工程化: CONTRIBUTING / SECURITY / CODE_OF_CONDUCT / AGENTS / CHANGELOG; `.github/` issue 模板、PR 模板与离线 CI(语法检查 + fixture 构建 + 冒烟 + 幂等性 + 可用时 LibreOffice 全量验收)。
- `tests/smoke_check.py` 离线冒烟校验脚本。

### 变更
- 工作表由 16 张增至 17 张(FCFE 插入 DCF 之后); Summary/Checks 页各新增上述行, 既有单元格数值零变化。
- README 重写为机构白皮书风格(零 emoji、删除 mermaid、补方法论与验证记录)。

### 验证
- 胜宏科技(300476)基准 DCF = 340.415208186145 元, 与 0.1.0 零差异; LibreOffice 重算 Checks 12/12 TRUE。
- 10 次重复构建单元格签名完全一致(幂等)。

## [0.1.0] - 2026-08-17

### 新增
- 首个公开版本: 配置驱动(configs/<code>.yaml)的 A 股 16 表三表联动 DCF 模型生成器。
- `fetch_data.py`: 东财 F10 三表/主营构成 + 腾讯行情(均系统 curl), 全自动兜底建模。
- `verify_model.py`: LibreOffice headless 重算 + Checks 11 项验收。
- FIN 页 revolver/现金 sweep/理财承接, 利息循环链表内迭代展开 4 轮(全簿无循环引用)。
- WACC 可比公司 β unlever/relever 实算; DCF 年中折现; 熊/基/牛三情景; 敏感性矩阵 ×3。
- 收录 configs/300476.yaml(完整手工配置)与 configs/002463.yaml(全自动生成); examples/ 两个成稿及验收日志。
