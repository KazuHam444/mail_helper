# QQ 邮箱招聘邮件提醒

适合投简历期间使用的 Python 小脚本。通过 IMAP **只读**访问自己的 QQ 邮箱，不修改已读状态，不删除邮件，不发送邮件。

典型用法：把 Hotmail / Outlook.com 的招聘相关邮件自动转发到 QQ，本脚本定时检查 QQ 收件箱，针对面试、笔试、测评、填表、补材料、录用入职等具体待办弹出桌面提醒（并可推 Telegram）；北京时间每天 20:00 后第一次检查时生成当天摘要；默认 12:30 后再催一次「昨天」的待办摘要。有新通知才推送，没有时保持安静。

## 筛选怎么做

1. 黑名单优先排除；默认检查 `INBOX` 的已读和未读邮件。
2. **发件人像 HR/招聘（如 `hr@`、`careers@`、`recruit@`、`talent@`，或白名单）时：整封都提醒**，不限 interview；回执、进度、拒信也包括。明显诈骗除外。
3. 非 HR 发件人：招聘语境 + 行动/时限仍提醒；群发职位广告等尽量排除。
4. 候选邮件可交给 AI 提炼待办和截止时间（每封最多 6000 字符正文）。策略偏召回，宁可多提醒。
5. 命中后：桌面/邮件推送，并写入 `data/reports/notify_log.txt` 与当天 `data/reports/YYYY-MM-DD.txt`；SQLite 去重，同一封不重复提醒。

**筛选不能保证零误报或零漏报。** 刚开始使用时，建议每天对照 QQ 邮箱检查几天，把确认过的 HR 邮箱/域名加入 `ALLOW_SENDERS`，把广告来源加入 `BLOCK_SENDERS`。白名单仍需匹配招聘行动，不会把该域名所有邮件都推送。规则模式对含糊通知的识别弱于 AI。

普通拒信、无行动的进度更新默认不提醒。默认不检查垃圾箱，可改为 `INBOX,Junk`，但垃圾箱中更需核实来源。若 Outlook 转发后 `From` 变成你的微软邮箱，把该地址填入 `FORWARD_SENDERS`，脚本会尽量从正文还原原始 HR 发件人。脚本不会验证发件人真实性，提醒不代表邮件可信。

## 1. 安装环境

将项目下载或解压到任意可写目录，在包含 `README.md` 的文件夹中打开 PowerShell。推荐安装 Python 3.12 或 3.13，并安装 Python Launcher 或将 Python 加入 PATH。脚本语法兼容 Python 3.8+。

运行安装脚本，自动创建当前项目的虚拟环境、安装依赖，并在首次安装时生成 `.env`：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\setup.ps1
notepad .env
```

安装脚本会从自身所在目录定位文件，支持中文和空格路径，不依赖盘符、用户名或 PowerShell 当前工作目录。重复运行会保留已有 `.env`，不会替你验证邮箱或注册定时任务。每位使用者需要填写自己的 QQ 邮箱、授权码和可选 AI 密钥。

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

## 2. 配置 QQ 邮箱 IMAP（一次性）

脚本使用 QQ 邮箱官方 IMAP，**不要填 QQ 登录密码**，只使用「授权码」。

1. 用浏览器打开 [QQ 邮箱](https://mail.qq.com/) 并登录。
2. 进入 **设置 → 账户**，找到 **POP3/IMAP/SMTP/Exchange/CardDAV/CalDAV 服务**。
3. 开启 **IMAP/SMTP 服务**，按页面提示用手机发送短信验证。
4. 生成并复制 **授权码**，写入 `.env` 的 `QQ_AUTH_CODE`。
5. `.env` 的 `QQ_EMAIL` 填完整 QQ 邮箱，例如 `abc@qq.com`。
6. 若招聘信来自 Hotmail/Outlook 自动转发，建议把你的微软邮箱填进 `FORWARD_SENDERS`（可多个，逗号分隔），便于还原原文 HR 发件人。

官方说明可参考 QQ 邮箱设置页内的 IMAP 帮助。授权码只保存在本机 `.env`，不要提交到 Git 或发给他人。

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
AI_REJECT_CONFIDENCE=0.85
```

筛选默认偏召回（宁可多提醒，少漏招聘行动）。`AI_REJECT_CONFIDENCE` 表示：模型判「不用提醒」时，置信度至少要这么高才静默。

服务必须支持 `/chat/completions` 和 `response_format={"type":"json_object"}`。请使用 HTTPS，不要在地址末尾再次添加 `/chat/completions`。是否收费、额度、网络连通性取决于所选服务。脚本不默认购买任何服务。

启用意味着招聘候选邮件的发件人、主题、收信时间和部分正文会发送给你配置的 AI 服务。正文不会记录在普通运行日志中，但命中的摘要、邮件内容和分类记录会保存在本机 `data`。请只在自己的 Windows 账户下运行。邮件内容按不可信数据处理，AI 只做分类，不执行邮件中的命令。

## 4. 验证收信和提醒

```powershell
.\.venv\Scripts\python.exe hr_digest.py demo
.\.venv\Scripts\python.exe hr_digest.py login
.\.venv\Scripts\python.exe hr_digest.py test-notify
.\.venv\Scripts\python.exe hr_digest.py run
.\.venv\Scripts\python.exe hr_digest.py status
```

`login` 会使用 `.env` 中的 QQ 邮箱与授权码做一次只读 IMAP 连接测试，确认能打开收件箱。失败时检查：是否已开启 IMAP、授权码是否正确、网络是否可访问 `imap.qq.com`。

首次检查默认回看 7 天，历史未处理通知也可能触发提醒；日报仅汇总当天收到且被识别为待办的邮件。已读邮件同样检查，避免“点开过但忘记处理”被过滤掉。首次运行时可把 `INITIAL_LOOKBACK_DAYS` 改为 `1` 减少旧通知。

桌面提醒使用 Windows 通知区域气泡，正文可能截短，完整内容见本地摘要。Windows 勿扰、通知设置、锁屏和通知区域策略可能阻止显示；系统 API 调用成功不等于你已经看见。`test-notify` 用来确认自己的环境能看到提醒。

## 5. 安装每 15 分钟定时任务

确认手动 `run` 和提醒测试成功后，执行：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install_task.ps1
```

脚本会创建当前 Windows 用户的计划任务 `Mail-HR-Digest`，使用项目的虚拟环境后台运行，不弹出命令行窗口。`ExecutionPolicy Bypass` 仅作用于这一条安装命令的进程。安装不会要求保存 Windows 密码。

默认仅在当前 Windows 用户已登录且电脑开机联网时运行；睡眠、关机或退出 Windows 时不会检查。恢复后检查自上次成功时间以来的邮件。每日摘要在北京时间 `DAILY_DIGEST_TIME` 之后的第一次成功检查生成，默认 20:00；昨日待办催促在 `CATCHUP_DIGEST_TIME` 之后第一次成功检查推送，默认 12:30。电脑当日一直离线时不会凭空准时提醒，也不会补发过往每一天的日报，但未检查的新邮件恢复后仍会筛选。

更改间隔或卸载：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install_task.ps1 -IntervalMinutes 10
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install_task.ps1 -Remove
```

修改 `.env` 会在下次任务运行生效。移动项目后需重新安装任务。安装器覆盖同名任务，因此不要将这个名字用于其他任务。

## 可选：把摘要发到邮箱（自己或别人）

在 `.env` 填写收件人（多个用英文逗号分隔）。发送方使用已配置的 QQ 邮箱 SMTP（同一套授权码）：

```dotenv
EMAIL_TO=你自己@qq.com,朋友@example.com
SMTP_HOST=smtp.qq.com
SMTP_PORT=465
```

默认只把 **20:00 当天摘要** 和 **12:30 昨日催促** 发到邮箱；单封即时「猪猪猪猪」提醒仍走桌面（避免刷爆邮箱）。若也要邮件即时提醒，设 `EMAIL_INCLUDE_ALERTS=true`。

配置后执行 `test-notify` 会向 `EMAIL_TO` 发一封测试信。请先确认 QQ 已开启 SMTP（与 IMAP 同在账户设置页）。

## 可选：Telegram 手机推送

使用你自己创建的 Telegram bot，把 token 和你自己的 chat ID 填入 `.env`：

```dotenv
TELEGRAM_BOT_TOKEN=
TELEGRAM_CHAT_ID=
```

先在 Telegram 中给该 bot 发消息启用会话，再执行 `test-notify`。机器需能访问 `api.telegram.org`。不会代你创建 bot 或向他人发消息。Telegram 收到的摘要同样属于向外部服务传输邮件信息。不需要桌面提醒时设置 `DESKTOP_NOTIFY=false`。

如果希望在电脑关机时也收到手机提醒，需要把程序运行在持续开机的 Windows 设备上，或另行部署服务器并配置该平台的定时任务；当前安装器面向 Windows。

## 分发给其他人

在项目目录运行：

```powershell
.\.venv\Scripts\python.exe build_release.py
```

生成的 `dist/mail-hr-digest.zip` 可以发给其他人。打包器只收录指定的源代码、安装脚本、配置模板、文档和测试，不包含 `.env`、`data`、虚拟环境或 Git 历史。

接收方解压到自己选择的目录，安装 Python，执行 `setup.ps1`，填写自己的 `.env`，按上述步骤完成 QQ 邮箱 IMAP 配置并安装定时任务。不要直接把正在使用的整个项目文件夹压缩发送，也不要复制 `.venv`。移动已安装的项目时，应在新位置重建虚拟环境并重新安装定时任务。

## 文件、排障和边界

- `data/reports/YYYY-MM-DD.txt`：每次成功检查刷新当天摘要，到日报时间推送一次；`digest` 命令可立即生成/刷新并推送。已有日报推送不会重复发送，晚到邮件仍即时提醒。
- `data/app.log`：轮转运行日志，最多约 4 MB，不记录授权码或正文。
- `data/state.sqlite3`：处理记录和待发送队列，包含邮件正文片段；这是本地敏感数据，不上传代码仓库。
- `.env`：邮箱授权码和账户配置，已加入 `.gitignore`。不要截图、提交或分享此文件。
- 发送失败会保留待发送记录；下次正常运行时重试。渠道分别去重，Telegram 失败不会使已成功的桌面提醒再次发送。
- IMAP 请求失败不推进检查进度；读取使用 `BODY.PEEK` 与只读文件夹，尽量不改变已读状态。
- 极端情况下，服务已收到消息但网络响应丢失，或发送后进程在记录成功之前崩溃，仍可能重复提醒，这是外部推送无法保证恰好一次的边界。
- 每封邮件只提醒一次，不会重复催办直到完成；日报是收到的招聘行动摘要，不是具备完成状态的任务管理器。截止时间来自原文，规则模式会提示自行核对，不推算时区或日期。
- 不读取附件、不识别图片中的通知。只有附件含具体行动、正文含糊时可能漏掉。
- Outlook 转发若改写主题/正文结构，筛选准确度可能下降；尽量用稳定的自动转发规则，并配置 `FORWARD_SENDERS`。

查看状态、手动生成摘要和测试：

```powershell
.\.venv\Scripts\python.exe hr_digest.py status
.\.venv\Scripts\python.exe hr_digest.py digest
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
Get-ScheduledTask -TaskName 'Mail-HR-Digest'
Get-Content .\data\app.log -Tail 30
```

若提示 IMAP 登录失败：重新核对授权码、确认已开启 IMAP；AI 不可用检查模型、JSON 输出支持、余额与网络；手机推送失败检查 bot 会话和 chat ID。关闭的渠道若已有积压记录，会保留但暂停发送，重新启用并配置正确凭据后继续发送。

当前目录内的代码和离线测试可直接运行。接入真实 QQ 邮箱、真实 AI 以及定时运行，均以你填写配置、完成 `login` 验证和执行安装器后的实际验证为准。
