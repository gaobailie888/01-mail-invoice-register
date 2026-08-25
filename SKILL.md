---
name: mail-invoice-register
description: 从 IMAP 邮箱或本地文件夹批量整理发票，解析电子发票 PDF 并生成可追溯 Excel。适用于邮箱发票、进项发票登记、购方归档与汇总；不用于发送或更改邮件，也不替代财务、税务审核。
metadata:
  short-description: 下载邮件发票附件并生成可追溯 Excel 登记表
---

# 邮箱发票自动登记

## 安全边界 {#safety}

- 只使用用户有权访问的邮箱，并通过 IMAP 以 `readonly=True` 打开；不得发送、删除、移动、标记邮件或改变邮箱状态。
- 优先使用邮箱应用专用密码，不得要求用户把网页登录密码、验证码或真实凭证发到对话中。
- 密码不得写入 Excel、运行摘要、日志、截图、GitHub 或 Skill 文件；本地向导退出后不保留密码。
- 第一次正式下载前必须先做预览，由用户确认日期、邮箱文件夹和候选附件范围。
- 不猜测解析失败的字段；扫描件、OFD、XML、图片、压缩包、加密文件及异常版式应保留并标记为待人工处理。
- 重复发票默认保留并标记；只有文件哈希完全相同的后续副本可以不计入“建议合计”。号码重复但内容不同、号码缺失均保留，最终入账、报销和税务决定必须由人完成。
- 真实购方映射规则不得写入公开 Skill；使用仓库外的私有 JSON，GitHub 只放虚构示例。

## 概述 {#overview}

本 Skill 把“找发票、下载附件、逐张录入 Excel、再按公司归档”变成可追溯流程。支持邮箱与本地文件夹两种入口，输出发票汇总、项目明细和处理日志；用户还可以选择按购方公司整理原件，并生成购方分表、统计口径和金额汇总。

## 普通用户流程 {#guided-flow}

1. 从 GitHub 下载源码后，Windows 用户双击 `start-windows.bat`；其他系统运行 `python start.py`。
2. 选择来源：邮箱模式或本地文件夹模式。
3. 邮箱模式下，本地网页向导只绑定 `127.0.0.1`，由用户本人填写邮箱账号和应用专用密码；本地模式只选择发票目录。
4. 先执行预览；邮箱模式可以先测试连接。
5. 按需勾选“按购方整理原件”和“生成购方分表与金额汇总”；需要人工映射时选择仓库外的私有规则 JSON。
6. 用户核对候选数量和范围后，点击确认执行。
7. 打开生成的 Excel，筛选待人工处理、校验异常、疑似重复和购方未分类记录。

需要配置说明时，读取 `references/configuration.md`；涉及购方规则时读取 `references/buyer-organization.md`；需要核对输出表述时读取 `references/content-claim-check.md`。

## Agent 执行流程 {#agent-flow}

当环境允许直接运行脚本时：

1. 使用包内的 `scripts/download_and_register.py`；需要图形向导时使用 `scripts/setup_wizard.py` 与 `assets/wizard.html`。在工作区运行时，仅复制所需执行资源。
2. 安装 `requirements.txt` 中的依赖。
3. 使用 skill 包目录以外的凭证文件执行 `--dry-run`。
4. 向用户展示候选数量和文件类型，不展示密码、完整邮件正文或无关附件。
5. 用户确认后正式执行，并交付 Excel、附件目录和运行摘要。

## 命令行参考 {#cli}

```powershell
python scripts/download_and_register.py `
  --config C:\secure\invoice-mail.env `
  --since 2026-08-01 `
  --until 2026-08-31 `
  --output-dir C:\invoice-output `
  --dry-run
```

确认范围后移除 `--dry-run`。仅在用户明确要求时间段内所有 PDF 时使用 `--include-all-pdfs`。

## 输出与验收 {#delivery}

输出目录包含：

- `attachments/`：原始候选附件；
- `AI邮箱发票登记YYYYMMDD.xlsx`：发票汇总、项目明细、处理日志；
- `run-summary.json`：候选数、解析成功数、待人工数、重复号码数和输出路径；
- 可选 `by_company/`、`*_购方分表.xlsx` 和 `buyer-summary.json`：购方归档与汇总结果。

验收时重点检查：连接是否只读、日期范围是否正确、候选数量是否合理、异常是否保留、重复是否按保守规则标记、未分类是否交回人工、输出中是否不存在密码和无关邮件正文。
