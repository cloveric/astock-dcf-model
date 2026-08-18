# Security Policy

## 范围

本仓库默认作为本地建模工具运行。Web 模式会在 `web/.data/` 保存任务参数、日志和工作簿;
浏览器可选择在当前 origin 的 `localStorage` 保存 API token。主要攻击面:

- `fetch_data.py` / `research/announcements.py` 通过系统 `curl` 访问公开数据接口(东财/腾讯);
- `--dr` / `--consensus` 读取本地文件;
- `--llm` 调用本机已安装的 claude / codex CLI。

## Web 服务安全基线

- 默认 `HOST=127.0.0.1`, 本机回环访问可不配置令牌。只要 `HOST` 是非回环地址
  （包括容器内的 `0.0.0.0`）, 启动时就必须配置 `WEB_API_TOKEN`; 所有 `/api` 请求使用
  `Authorization: Bearer <token>`。建议用 `openssl rand -hex 32` 生成独立随机令牌，不要复用登录密码。
- 即使服务被错误地绑定到外部地址且未设置 `HOST`, 来自非回环客户端的 `/api` 请求仍会被拒绝。
  静态首页不包含任务数据；反向代理应启用 TLS，且不得记录 `Authorization` 请求头。
- `WEB_MAX_ACTIVE_JOBS` 默认 8，达到排队中 + 运行中上限时提交返回 HTTP 429；
  `WEB_MAX_JOBS` 默认 100，终态任务超过上限时淘汰最早记录及其产物。服务仅支持单进程，
  不要使用 `uvicorn --workers > 1`。
- API 请求体默认限制为 64 KiB（`WEB_MAX_REQUEST_BYTES`）；配置、DR 和一致预期路径最长
  512 字符，且仍必须解析为仓库内的已有普通文件。
- Web 任务的 `--llm` 默认强制为 `off`。仅在确认研究材料可以发送给对应 CLI 服务后，
  才设置 `WEB_ALLOW_LLM=1`; 该开关不会替代 API 鉴权。

### 验收状态和下载门禁

- `verified`: `verify_model.py` 退出码为 0，默认允许下载。
- `built_unverified`: 已生成工作簿，但运行环境找不到 LibreOffice；默认禁止下载。仅在明确接受
  未验收产物时设置 `WEB_ALLOW_UNVERIFIED_DOWNLOAD=1`。
- `failed_validation`: LibreOffice 验收非零退出或执行失败，禁止下载，即使开启上述未验收下载开关。
- `failed`: 构建失败，未进入验收。

Web 服务要求验收子进程退出码、`--json-summary` 的 `verdict=PASS` 和 JSON `exit_code=0`
三者一致才标记 `verified`；JSON 缺失、损坏、版本不兼容或结果不一致均为 `failed_validation`。
任务记录持久化 `verify_status`、`verify_returncode`、`verify_verdict` 和结构化验收摘要，不要只根据
工作簿文件存在或日志中的文字判断验收成功。

Docker 镜像以非 root 用户运行，内置 LibreOffice，并使用静态首页作为健康检查。镜像内默认
`HOST=0.0.0.0`，因此必须通过运行时 secret/环境变量注入 `WEB_API_TOKEN` 才能启动；不要把真实
token 写进 Dockerfile、镜像层或仓库文件。

## 报告漏洞

如发现安全问题(例如命令注入、路径穿越、不安全的反序列化), 请**不要**公开提 issue,
改用 GitHub 私有漏洞报告(Security → Report a vulnerability)或联系仓库所有者。
我们将在确认后尽快修复并披露。

## 支持版本

仅最新 `main` 分支获得安全修复。

## 使用侧安全约定

- YAML 配置仅经 `yaml.safe_load` 解析, 请勿引入自定义 tag;
- 不要向 `--dr`/`--consensus` 传入来源不可信的文件而不加审阅(其内容会写入工作簿);
- `--llm` 会把研究要点发送到本机 LLM CLI 对应的服务, 涉密研究材料请使用 `--llm off`(默认)。
