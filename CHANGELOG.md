# Changelog

本项目遵循 [Semantic Versioning](https://semver.org/lang/zh-CN/)。

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
