# Outlook Policy Weekly Fetcher

简要说明
- 脚本 `fetch_policy_emails.py` 在 Windows/Outlook 下抓取指定邮箱文件夹的邮件，保存正文和允许的附件为文件，并将文本聚合发送到 OpenAI 生成 weekly policy report。

主要文件
- `fetch_policy_emails.py` — 主脚本
- `prompts_extractor.py` — Step 1 抽取器 SYSTEM/USER Prompt
- `prompts_report_generator.py` — Step 2 报告生成器 SYSTEM/USER Prompt
- `requirements.txt` — 依赖清单

准备与依赖
1. 在 Windows 上运行，Outlook 已配置并且能访问目标邮箱。
2. 建议创建虚拟环境并安装依赖：
```powershell
python -m venv .venv; .\.venv\Scripts\Activate.ps1; pip install -r requirements.txt
```
3. 在仓库目录创建 `.env` 或在系统环境变量中设置：
```text
OPENAI_API_KEY=sk-...
# 可选：设置默认 Outlook 邮箱与文件夹（可被 CLI 覆盖）
OUTLOOK_MAILBOX=your.mailbox@company.com
OUTLOOK_FOLDER_PATH=收件箱/2. Policy Update
```
- 说明：`OUTLOOK_FOLDER_PATH` 支持分隔符 `/`、`\`、`>`，脚本会自动拆分成层级路径。

运行与 CLI
脚本采用 two-step 工作流：先结构化抽取（输出 JSON），再从 JSON 生成 Markdown 报告。

常用命令（Windows PowerShell）：
- 默认（最近 7 天） ：
```powershell
python fetch_policy_emails.py
```
- 指定天数回溯（例如最近 14 天） ：
```powershell
python fetch_policy_emails.py --days 14
```
- 指定起止日期（格式 YYYY-MM-DD） ：
```powershell
python fetch_policy_emails.py --since 2025-08-10 --until 2025-08-21
```
- 仅执行抽取（Step 1），输出 EXTRACT_*.json ：
```powershell
python fetch_policy_emails.py --only-extract
```
- 从已有 JSON 直接生成报告（跳过 Step 1） ：
```powershell
python fetch_policy_emails.py --json-input .\20250828_policy\EXTRACT_250822-250828.json
```
- 覆盖默认模型 ：
```powershell
python fetch_policy_emails.py --model-extract gpt-5-mini --model-generate gpt-5-mini
```

Outlook 路径自定义（.env 或 CLI）
- .env（默认值） ：
  - `OUTLOOK_MAILBOX`：Outlook 邮箱的显示名或地址。
  - `OUTLOOK_FOLDER_PATH`：邮箱下的文件夹路径（如 `收件箱/2. Policy Update`）。
- CLI（优先级高于 .env） ：
  - `--mailbox`：覆盖默认邮箱。
  - `--folder`：覆盖默认文件夹路径。

辅助发现命令（便于用户确定自己的配置）
- 列出当前配置文件中所有可用邮箱：
```powershell
python fetch_policy_emails.py --list-mailboxes
```
- 列出某个邮箱下的文件夹树（默认深度 2 层，可通过 `--folders-depth` 调整） ：
```powershell
python fetch_policy_emails.py --mailbox "your@company.com" --list-folders --folders-depth 3
```

日期优先级与行为
- 如果同时给出 `--since` 则以 `--since` 为起点，`--days` 将被忽略。
- 如果提供 `--until`，脚本将以 `--until` 为结束时间，否则使用当前时间。
- 所有日期在内部被标准化为本地时间（naive local datetime）。若意外 `since > until`，脚本会自动交换两者并记录日志。

报告与输出命名规则
- 合并原始材料（用于提示词输入）：`ALL_YYMMDD-YYMMDD.txt`
- Step 1 抽取输出：`EXTRACT_YYMMDD-YYMMDD.json`
- Step 2 报告输出：`Weekly_report_YYMMDD-YYMMDD.md`
- 若未捕获到邮件，文件名回退到抓取窗口 `since/until`。

邮件时间与去重逻辑
- 时间选择优先级（实现建议） ：
  1. 邮件头 `Date`（解析为本地时间）
  2. 若缺失，则使用 Outlook `ReceivedTime`
- 去重：优先用 `InternetMessageID` 或 `EntryID`；都不可用则对 `subject|sender|received|body_snip` 做 SHA256 指纹。

环境变量
- `OPENAI_API_KEY` — 必需
- `OUTLOOK_MAILBOX` / `OUTLOOK_FOLDER_PATH` — 可选，作为默认路径
- `USE_MOCK_EMAILS` — `true` 时本地生成模拟邮件，便于离线测试
- `SKIP_OPENAI` — `true` 跳过 OpenAI 调用并输出占位报告
- `DETACH_OPENAI` — `true` 后台线程生成报告
- `OPENAI_MODEL_EXTRACT` / `OPENAI_MODEL_GENERATE` — 模型覆盖（同 CLI）

附件与 PDF 文本提取
- 默认保存常见文档（.pdf/.docx/.xlsx/.txt/.csv 等），图片/.eml/.emz 跳过。
- 如需 PDF 文本内联，请安装 pdfplumber ：
```powershell
pip install pdfplumber
```

故障排查
- 如果 Outlook Restrict 返回 0 条，但你确认该时间段有邮件：
  - 确认 `.env` 中的 `OUTLOOK_MAILBOX` 与 `OUTLOOK_FOLDER_PATH` 正确；
  - 用 `--list-mailboxes`、`--list-folders` 先核对路径；
  - 本脚本已优先采用 DASL 过滤并回退到 MAPI 字符串，同时在客户端二次过滤，通常能规避区域设置导致的过滤失效。

