# IMAP 配置

普通用户优先运行根目录的 `start.py` 或双击 `start-windows.bat`，通过仅在本机打开的网页向导完成连接、预览和执行，不需要手工编辑配置文件。

以下方式仅供命令行用户使用。将内容保存为用户有权访问、且位于仓库目录之外的本地文件，例如 `C:\secure\invoice-mail.env`。不要将真实密码提交到版本库、打包到 Skill、截图或发送到聊天记录。

```dotenv
IMAP_HOST=imap.example.com
IMAP_PORT=993
IMAP_USERNAME=finance@example.com
IMAP_PASSWORD=replace-with-an-app-password
IMAP_MAILBOX=INBOX
IMAP_USE_SSL=true
```

`IMAP_MAILBOX` 可改成已归档的发票文件夹。常见邮箱需要在网页端启用 IMAP，并创建应用专用密码；不要使用网页登录密码。脚本不会发送邮件，也不会更改已读、星标或移动状态。

## 可选过滤规则

- `--since YYYY-MM-DD`：包含该日，默认最近 30 天。
- `--until YYYY-MM-DD`：包含该日。未指定时截至今天。
- `--keywords`：逗号分隔的邮件主题、正文或附件名匹配词。
- `--include-all-pdfs`：不以关键字过滤 PDF，仅限用户明确要求全量处理时使用。

如果服务器不支持 `SINCE` 查询，或网络认证失败，先核对主机、端口、SSL 设置和应用密码；不要用反复登录重试来掩盖认证问题。
