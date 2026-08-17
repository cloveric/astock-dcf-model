# Security Policy

## 范围

本仓库为本地建模工具, 不持有服务端、不存储用户凭证。主要攻击面:

- `fetch_data.py` / `research/announcements.py` 通过系统 `curl` 访问公开数据接口(东财/腾讯);
- `--dr` / `--consensus` 读取本地文件;
- `--llm` 调用本机已安装的 claude / codex CLI。

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
