# Hotmail / Outlook.com 招聘邮件提醒

适合投简历期间使用的 Python 小脚本。只读访问自己的邮箱，不修改已读状态，不删除邮件，不发送邮件。

默认方案：Windows 登录期间每 15 分钟检查收件箱，针对面试、笔试、测评、填表、补材料、录用入职等具体待办弹出桌面提醒；北京时间每天 20:00 后第一次检查时生成当天摘要。有新通知才推送，没有时保持安静。也可配置 Telegram 手机提醒。

## 筛选怎么做

1. 黑名单优先排除；默认检查 Inbox 的重点、其他、已读和未读邮件。
2. 本地规则排除职位推荐、培训营销、单纯投递回执等明显无行动邮件。
3. 具有招聘语境和具体行动的候选邮件交给 AI 判断，提取待办和原文截止时间。每封最多向 AI 服务发送 6000 字符正文，不发附件，不访问正文链接。
4. AI 模式仅提醒置信度达到阈值的行动通知。未配置 AI 时使用较保守的规则；AI 请求失败时强规则命中照常提醒，未完成的分类下次重试。
5. SQLite 保存判断结果和每个渠道的发送状态；重复检查不会重复提醒同一邮件。新发来的改期或取消邮件仍会提醒。

**筛选不能保证零误报或零漏报。** 刚开始使用时，建议每天对照 Outlook 检查几天，把确认过的 HR 邮箱/域名加入 `ALLOW_SENDERS`，把广告来源加入 `BLOCK_SENDERS`。白名单仍需匹配招聘行动，不会把该域名所有邮件都推送。规则模式对含糊通知的识别弱于 AI。

普通拒信、无行动的进度更新默认不提醒。收件箱子文件夹不自动递归；如 Outlook 规则会把 HR 邮件移动到自定义文件夹，需把其 Graph 文件夹 ID 加入 `MAIL_FOLDERS`。默认不检查垃圾箱，可改为 `inbox,junkemail`，但垃圾箱中更需核实来源。脚本不会验证发件人真实性，提醒不代表邮件可信。

## 1. 安装环境

将项目下载或解压到任意可写目录，在包含 `README.md` 的文件夹中打开 PowerShell。推荐安装 Python 3.12 或 3.13，并安装 Python Launcher 或将 Python 加入 PATH。脚本语法兼容 Python 3.8+。

运行安装脚本，自动创建当前项目的虚拟环境、安装依赖，并在首次安装时生成 `.env`：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\setup.ps1
notepad .env
```

安装脚本会从自身所在目录定位文件，支持中文和空格路径，不依赖盘符、用户名或 PowerShell 当前工作目录。重复运行会保留已有 `.env`，不会替你登录邮箱或注册定时任务。每位使用者需要填写自己的邮箱、Microsoft 应用 ID 和可选 AI 密钥。

如果没有 Python Launcher，可直接使用加入 PATH 的 `python`；也可以指定 Python 可执行文件：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\setup.ps1 -PythonExecutable python
```

不使用安装脚本时，可以在项目目录手动安装：

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
if (-not (Test-Path -LiteralPath .env)) { Copy-Item -LiteralPath .env.example -Destination .env }
```

`requirements-lock.txt` 记录了 Windows / Python 3.13 验证通过的依赖版本；同环境重建时可用 `pip install -r requirements-lock.txt`。其他 Python 版本优先使用 `requirements.txt` 选择兼容版本。

## 2. 注册 Microsoft 只读应用（一次性）

使用 Microsoft Graph + OAuth 官方登录，不使用邮箱密码、应用密码或过时的基本认证。个人 Outlook/Hotmail 支持委托 `Mail.Read` 权限。

1. 打开 [Microsoft Entra 管理中心](https://entra.microsoft.com/)，进入 **身份 → 应用程序 → 应用注册 → 新注册**。
2. 名称可以是 `My HR Mail Reminder`。支持的账户类型选 **任何组织目录中的账户和个人 Microsoft 账户**，或支持个人 Microsoft 账户的选项。无需填写重定向 URI。
3. 在概述复制 **应用程序（客户端）ID**，写入 `.env` 的 `MS_CLIENT_ID`。
4. 在 **身份验证 → 高级设置** 将 **允许公共客户端流 / Allow public client flows** 设为 **是**。界面名称可能随管理中心更新而变化。
5. **API 权限 → 添加权限 → Microsoft Graph → 委托的权限**，添加 `Mail.Read`。只需要邮件读取，不添加 `Mail.ReadWrite`、`Mail.Send` 或应用程序权限。不需要创建客户端密钥。
6. `.env` 的 `MS_ACCOUNT` 填你要检查的完整邮箱，例如 `abc@hotmail.com`。

应用注册需要可访问的 Entra 目录及注册权限。部分纯个人微软账号进入管理中心时没有可用目录，此时先通过自己的 Azure 账户获得目录，或请可管理的目录创建一个支持个人账户的应用；目标邮箱仍可为个人 Hotmail。如果此步骤受限，脚本无法仅凭邮箱密码绕过授权。

官方参考：[注册应用](https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-register-app)、[MSAL Python 设备登录](https://learn.microsoft.com/en-us/entra/msal/python/getting-started/acquiring-tokens)、[Graph 列出邮件](https://learn.microsoft.com/en-us/graph/api/user-list-messages)、[权限说明](https://learn.microsoft.com/en-us/graph/permissions-reference)。

## 3. 配置 AI（可选，但推荐）

### 阿里千问 / DashScope

在本地 `.env` 配置下列项目（不要把真实密钥写进 `.env.example` 或发到聊天中）：

```dotenv
DASHSCOPE_API_KEY=你的百炼密钥
AI_API_KEY=
AI_BASE_URL=https://dashscope.aliyuncs.com/compatible-mode/v1
AI_MODEL=qwen-plus
```

脚本支持直接读取 `DASHSCOPE_API_KEY`，请求时关闭思考模式并要求 JSON 输出。默认使用北京地域的兼容接口；API Key 必须与接口地域一致。新加坡地域可配置 `https://dashscope-intl.aliyuncs.com/compatible-mode/v1`。百炼还支持业务空间专属域名，可直接复制自己控制台提供的兼容接口地址到 `AI_BASE_URL`。

配置后先用虚构邮件验证，不需要先登录邮箱：

```powershell
.\.venv\Scripts\python.exe hr_digest.py check-ai
```

此命令会发出一次小型模型请求，可能按百炼用量计费。只有测试通过才表示密钥、地域、模型权限以及 JSON 输出实际可用；`status` 的“已启用”仅表示配置存在。AI 故障时仍按规则回退。若此前已经以纯规则模式处理过邮件，启用 AI 不会重新分类那些已完成记录。

官方说明：[OpenAI Chat 兼容接口与地域](https://help.aliyun.com/zh/model-studio/compatibility-of-openai-with-dashscope)、[千问 JSON 结构化输出](https://help.aliyun.com/zh/model-studio/qwen-structured-output)。

### 其他兼容服务

不填即可先本地运行。已有支持 OpenAI Chat Completions 格式的服务时填写：

```dotenv
AI_BASE_URL=https://你的服务地址/v1
AI_API_KEY=你的密钥
AI_MODEL=你的服务支持的模型名称
AI_MIN_CONFIDENCE=0.80
```

服务必须支持 `/chat/completions` 和 `response_format={"type":"json_object"}`。请使用 HTTPS，不要在地址末尾再次添加 `/chat/completions`。是否收费、额度、网络连通性取决于所选服务。脚本不默认购买任何服务。

启用意味着招聘候选邮件的发件人、主题、收信时间和部分正文会发送给你配置的 AI 服务。正文不会记录在普通运行日志中，但命中的摘要、邮件内容和分类记录会保存在本机 `data`。请只在自己的 Windows 账户下运行。邮件内容按不可信数据处理，AI 只做分类，不执行邮件中的命令。

## 4. 授权、验证收信和提醒

```powershell
.\.venv\Scripts\python.exe hr_digest.py demo
.\.venv\Scripts\python.exe hr_digest.py login
.\.venv\Scripts\python.exe hr_digest.py test-notify
.\.venv\Scripts\python.exe hr_digest.py run
.\.venv\Scripts\python.exe hr_digest.py status
```

`login` 会显示 Microsoft 登录地址和一次性代码。自己打开该 Microsoft 地址，输入代码并确认是自己注册的应用，选择与 `MS_ACCOUNT` 一致的邮箱并授权。此步骤需要本人交互，其他命令不会在后台弹出登录流程。

Windows 上登录缓存由当前 Windows 用户的 DPAPI 加密，后续通过 MSAL 静默续期；更改密码、撤销授权或 Microsoft 要求重新验证时，再执行 `login`。不会回退到明文保存 token。

首次检查默认回看 7 天，历史未处理通知也可能触发提醒；日报仅汇总当天收到且被识别为待办的邮件。已读邮件同样检查，避免“点开过但忘记处理”被过滤掉。首次运行时可把 `INITIAL_LOOKBACK_DAYS` 改为 `1` 减少旧通知。

桌面提醒使用 Windows 通知区域气泡，正文可能截短，完整内容见本地摘要。Windows 勿扰、通知设置、锁屏和通知区域策略可能阻止显示；系统 API 调用成功不等于你已经看见。`test-notify` 用来确认自己的环境能看到提醒。

## 5. 安装每 15 分钟定时任务

确认手动 `run` 和提醒测试成功后，执行：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install_task.ps1
```

脚本会创建当前 Windows 用户的计划任务 `Hotmail-HR-Digest`，使用项目的虚拟环境后台运行，不弹出命令行窗口。`ExecutionPolicy Bypass` 仅作用于这一条安装命令的进程。安装不会要求保存 Windows 密码。

默认仅在当前 Windows 用户已登录且电脑开机联网时运行；睡眠、关机或退出 Windows 时不会检查。恢复后检查自上次成功时间以来的邮件。每日摘要在北京时间 `DAILY_DIGEST_TIME` 之后的第一次成功检查生成，默认 20:00；电脑当日一直离线时不会凭空准时提醒，也不会补发过往每一天的日报，但未检查的新邮件恢复后仍会筛选。

更改间隔或卸载：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install_task.ps1 -IntervalMinutes 10
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install_task.ps1 -Remove
```

修改 `.env` 会在下次任务运行生效。移动项目后需重新安装任务。安装器覆盖同名任务，因此不要将这个名字用于其他任务。

## 可选：Telegram 手机推送

使用你自己创建的 Telegram bot，把 token 和你自己的 chat ID 填入 `.env`：

```dotenv
TELEGRAM_BOT_TOKEN=
TELEGRAM_CHAT_ID=
```

先在 Telegram 中给该 bot 发消息启用会话，再执行 `test-notify`。机器需能访问 `api.telegram.org`。不会代你创建 bot 或向他人发消息。Telegram 收到的摘要同样属于向外部服务传输邮件信息。不需要桌面提醒时设置 `DESKTOP_NOTIFY=false`。

如果希望在电脑关机时也收到手机提醒，需要把程序运行在持续开机的 Windows 设备上，或另行部署服务器并配置该平台的加密缓存和定时任务；当前安装器面向 Windows。

## 分发给其他人

在项目目录运行：

```powershell
.\.venv\Scripts\python.exe build_release.py
```

生成的 `dist/hotmail-hr-digest.zip` 可以发给其他人。打包器只收录指定的源代码、安装脚本、配置模板、文档和测试，不包含 `.env`、`data`、登录缓存、虚拟环境或 Git 历史。

接收方解压到自己选择的目录，安装 Python，执行 `setup.ps1`，填写自己的 `.env`，按上述步骤完成 Microsoft 授权并安装定时任务。不要直接把正在使用的整个项目文件夹压缩发送，也不要复制 `.venv`；虚拟环境和 Windows 加密登录缓存不能在不同电脑之间直接复用。移动已安装的项目时，应在新位置重建虚拟环境并重新安装定时任务。

## 文件、排障和边界

- `data/reports/YYYY-MM-DD.txt`：每次成功检查刷新当天摘要，到日报时间推送一次；`digest` 命令可立即生成/刷新并推送。已有日报推送不会重复发送，晚到邮件仍即时提醒。
- `data/app.log`：轮转运行日志，最多约 4 MB，不记录 token 或正文。
- `data/state.sqlite3`：处理记录和待发送队列，包含邮件正文片段；这是本地敏感数据，不上传代码仓库。
- `data/token.cache`：当前 Windows 用户加密的 Microsoft 登录缓存。
- `.env`：API 密钥和账户配置，已加入 `.gitignore`。不要截图、提交或分享此文件。
- 发送失败会保留待发送记录；下次正常运行时重试。渠道分别去重，Telegram 失败不会使已成功的桌面提醒再次发送。
- Graph 请求失败不推进检查进度；API 限流/暂时错误有有限次数重试，分页和历史检查没有“只看前 50 封”的截断。
- 极端情况下，服务已收到消息但网络响应丢失，或发送后进程在记录成功之前崩溃，仍可能重复提醒，这是外部推送无法保证恰好一次的边界。
- 每封邮件只提醒一次，不会重复催办直到完成；日报是收到的招聘行动摘要，不是具备完成状态的任务管理器。截止时间来自原文，规则模式会提示自行核对，不推算时区或日期。
- 不读取附件、不识别图片中的通知。只有附件含具体行动、正文含糊时可能漏掉。

查看状态、手动生成摘要和测试：

```powershell
.\.venv\Scripts\python.exe hr_digest.py status
.\.venv\Scripts\python.exe hr_digest.py digest
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
Get-ScheduledTask -TaskName 'Hotmail-HR-Digest'
Get-Content .\data\app.log -Tail 30
```

若提示授权失效：重新 `login`；403 检查委托 `Mail.Read` 和个人账户支持；AI 不可用检查模型、JSON 输出支持、余额与网络；手机推送失败检查 bot 会话和 chat ID。关闭的渠道若已有积压记录，会保留但暂停发送，重新启用并配置正确凭据后继续发送。

当前目录内的代码和离线测试可直接运行。接入真实邮箱、真实 AI 以及定时运行，均以你填写配置、完成登录和执行安装器后的实际验证为准。
