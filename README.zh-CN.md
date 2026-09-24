# GHCP Proxy 使用指南

GHCP Proxy 是一个运行在本地（仅监听 `127.0.0.1`）的兼容代理：它把 Codex / Claude Code
的请求翻译后转发给 GitHub Copilot（官方 `github-copilot-sdk`），也可以把特定模型别名
转发到 ChatGPT Excel 插件后端（按 ChatGPT 订阅计费）。仪表盘提供登录、集成配置、
用量与费用估算。

## 合规声明（必读）

本项目是非官方兼容层。滥用可能违反 GitHub 服务条款或可接受使用政策，包括：
把普通请求标记为免费/代理流量、绕过计费、超限请求等，都可能导致 API 访问被暂停
甚至账号被永久封禁。请做到：

- 只使用**自己的**账号
- 遵守 GitHub 的限额与配额
- 不利用本项目规避计费或配额
- Excel 上游只使用**自己的** ChatGPT 账号会话；凭证仅存于代理内存（macOS），
  绝不回传仪表盘、绝不写入日志

GitHub 是计费与执法的最终依据，详见 readme.md 中的官方链接。

## 准备工作

- Python 3.11 或更高版本
- 一个有 GitHub Copilot 权限的 GitHub 账号
- （可选）已安装 Codex 或 Claude Code
- （可选）macOS / Windows 桌面版 Excel + 官方 ChatGPT 插件（用于 Excel 上游）

## 安装与启动

macOS / Linux：

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python proxy.py
```

Windows PowerShell：

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe .\proxy.py
```

启动后打开仪表盘：

```text
http://localhost:8000/
```

## 首次配置（仪表盘）

1. 按提示完成 GitHub 设备码登录（用有 Copilot 权限的账号）
2. 打开 **Integrations** 页面
3. 启用 Codex、Claude Code 之一或全部
4. 点击安装 shell 命令（之后可用 `start-ghproxy` / `stop-ghproxy`）
5. 可选：启用开机自启
6. 重启已经打开的 Codex / Claude Code 会话（它们只在会话启动时读取配置）

无需 Node.js、`npx`，也无需手改 `~/.codex` / `~/.claude` 配置；代理会写入并自动备份。

## 日常使用

已启用自启时，登录系统后代理自动运行；手动控制：

```bash
start-ghproxy    # 启动（后台运行）
stop-ghproxy     # 停止
```

```powershell
Start-GHProxy    # Windows PowerShell
Stop-GHProxy
```

代理只监听本机：

- OpenAI 兼容接口：`http://localhost:8000/v1`
- 仪表盘：`http://localhost:8000/`
- 日志：`~/Library/Application Support/ghcp_proxy/ghcp-proxy.stdout.log`（macOS）

## 配置 Codex / Claude Code

代理会自动把客户端指向 `http://127.0.0.1:8000/v1`。手动使用时的环境变量：

```bash
export OPENAI_BASE_URL=http://127.0.0.1:8000/v1
export OPENAI_API_KEY=anything   # 任意值即可，代理不校验
```

在 Codex 中启动新会话后，模型选择器里即可看到代理下发的模型列表
（含 `auto` 和各 `*-excel` 别名）。选普通模型走 GitHub Copilot 计费；
选 Excel 别名走 ChatGPT 订阅（见下节）。

## GPT Excel 上游

`gpt-6-astra-excel`、`gpt-5.6-luna-excel`、`gpt-5.6-terra-excel`、
`gpt-5.6-sol-excel` 是仅支持 `/v1/responses` 的实验性别名，请求发往 ChatGPT
Excel 插件后端对应模型（`gpt-6-astra` / `gpt-5.6-luna` / …），而不是 GitHub
Copilot。此通道非官方，后端可能随时变化。

### 前提：在 Excel 插件里登录一次

- Windows：桌面版 Excel 安装官方 ChatGPT 插件并登录至少一次。代理从 Office
  WebView2 的 Local Storage 读取缓存的 `bps_auth_tokens` 会话，Excel 关闭时也能读，
  凭证用当前用户的 DPAPI 密钥加密保存。
- macOS：桌面版 Excel 打开官方 ChatGPT 插件并登录。代理直接读取 Excel WebKit
  LocalStorage 的 SQLite 库（`~/Library/Containers/com.microsoft.Excel/Data/Library/WebKit/WebsiteData/...`），
  不安装证书、不改系统代理。会话只保存在代理内存，代理重启后会自动重新读取。

代理在启动时和 Excel 路由请求发现会话缺失/过期时自动重读。**token 过期时，
打开 Excel 的 ChatGPT 任务面板刷新一次即可**，无需重启代理。

### 使用

在 Codex 里选择 `gpt-6-astra-excel`（或其它 Excel 别名）即可，其余模型仍走
GitHub Copilot。四个别名都支持 `low` / `medium` / `high` / `xhigh` 推理强度
（`x-high` 作为输入别名也被接受）。命令行冒烟测试：

```bash
curl -s http://127.0.0.1:8000/v1/responses \
  -H "Authorization: Bearer anything" \
  -H "Content-Type: application/json" \
  -d '{"model":"gpt-6-astra-excel","input":"hi"}'
```

### 会话管理

查看非敏感状态（是否已配置、过期时间、读取方式）：

```bash
curl -s http://127.0.0.1:8000/api/config/excel-session
```

```powershell
Invoke-RestMethod http://127.0.0.1:8000/api/config/excel-session
```

- 强制重新读取插件缓存：`POST /api/config/excel-session`，body `{"action": "read_cached"}`
- 清除当前会话（Windows 同时删除加密副本，不必停代理）：`DELETE /api/config/excel-session`
- 从 WebKit 缓存刷新并提交给代理的现成工具：`tools/prime-excel-session-macos.py`

### 换账号

代理读取的是 Excel 插件当前登录账号写进 WebKit 的会话，所以最直接的方式就是
在 Excel 插件里退出并用新账号登录，代理会自动跟上。也可以手动向代理提交
会话 headers（body 为 `{"headers": {...}}`，需包含 `authorization: Bearer ...`
和 `chatgpt-account-id` 等字段）——只使用自己的账号。

## 手动 API 调用示例

列出可用模型：

```bash
curl -s http://127.0.0.1:8000/v1/models -H "Authorization: Bearer anything"
```

Responses 调用（普通模型走 Copilot）：

```bash
curl -s http://127.0.0.1:8000/v1/responses \
  -H "Authorization: Bearer anything" \
  -H "Content-Type: application/json" \
  -d '{"model":"auto","input":"你好"}'
```

## 计费说明

仪表盘展示近期的请求、token 用量与估算费用，**仅供参考**，以 GitHub 官方账单为准：

- Copilot 流量按模型费率折算为 GitHub AI Credits（1 Credit = $0.01）
- Excel 别名按其对应模型的费率折算为 OpenAI Credits（1 Credit = $0.04），
  未收录费率的模型（如 `gpt-6-astra-excel`）显示 "Token pricing unavailable"，
  不影响使用，费用计入你的 ChatGPT 订阅

## 自动更新

从 git 检出启动时，代理每 15 分钟检查 `origin/main` 并自动安全升级（升级后自动重启）。
默认 **user 模式**：本地未提交的修改会先 stash，升级后重新应用；无法安全重放时
跳过并在仪表盘提示。环境变量开关：

```bash
export GHCP_AUTO_UPDATE=0                        # 关闭自动更新
export GHCP_AUTO_UPDATE_MODE=developer           # 开发模式：有未提交修改时阻止升级
export GHCP_AUTO_UPDATE_INTERVAL_SECONDS=900     # 检查间隔（默认 15 分钟）
```

## 常见问题

**`start-ghproxy` 命令不存在** — 还没安装 shell 命令：到仪表盘 Integrations 页面
点安装，或新开终端先 `source ~/.zshrc`。

**打开仪表盘提示无法连接** — 代理还没启动完（后台启动有几秒延迟），稍等重试；
或端口 8000 被旧进程占用，先 `stop-ghproxy` 再启动。

**Codex / Claude Code 仍连旧服务商** — 客户端只在会话启动时读配置，完全重启
会话；无效则到 Integrations 里先禁用再启用。

**`No module named fastapi` / `uvicorn`** — 依赖装进了别的 Python，用本地虚拟
环境安装并启动（`source .venv/bin/activate` 后 `pip install -r requirements.txt`）。

**上游请求超时** — 提高超时：

```bash
export GHCP_UPSTREAM_TIMEOUT_SECONDS=300
```

**Excel 模型报会话未配置/过期** — 打开 Excel 的 ChatGPT 任务面板刷新登录，
然后重试（代理会自动重读）。

**GitHub 登录失败（DNS / 超时 / VPN）** — 先在终端验证 `curl -I https://github.com`
能通，再重新启动代理登录。

## 参考

- 完整英文说明：readme.md
- 缓存与计费调查：`docs/prompt-cache-investigation-*.md`
- GitHub 官方：Copilot 模型与定价、用量限额、API 条款（见 readme.md 链接）
