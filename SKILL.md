---
name: sangfor-report-downloader
description: 下载 SOAR 原始数据并生成指定客户、指定时间段的本地 Excel 报告。仅在用户明确要原始报表数据、Excel 工作簿，或只要资产表、暴露面、事件表、告警表、漏洞表等单项数据时使用。不要在用户最终目标是生成 AI PPT 报告时使用；那种场景应优先使用 ai-ppt-pipeline，它会先调用本下载器再继续生成 PPT。
---

# 深信服报告下载器

仅在需要本地 SOAR 数据提取和 Excel 生成时使用本技能。

## 适用场景

- “下载某客户某时间段的报告”
- “帮我出一份 Excel 报表”
- “只要资产表/暴露面/事件表/告警表/漏洞表”
- “先把原始 response 导出来看看”

不要用本技能做端到端 AI PPT 生成。
PPT 模板选择不在本技能处理；用户要求协同运营报告/plus 或托管运营报告/simple 时，切换到 `ai-ppt-pipeline` 并传对应 `--template`。

## 前置条件

1. Chrome Cookie 插件已经安装并已写出 Cookie 文件。
   - MSS Cookie: `M:\Users\$env:USERNAME\Downloads\cookies.txt`
   - XDR Cookie: `M:\Users\$env:USERNAME\Downloads\xdr_cookies.txt`
2. 当前项目已经执行过 `npm install`。
3. 调用方已经给出明确的结构化参数，例如客户名、客户 ID、开始日期、结束日期，以及可选的报告类型。

## 核心命令

直接调用 OpenClaw workspace 下的脚本。

```powershell
node "$HOME\.openclaw\workspace\skills\sangfor-report-downloader\sangfor_downloader.js" `
  --customer "客户中文名" `
  --start "2026-05-12" `
  --end "2026-05-13" `
  --cookie-path "M:\Users\$env:USERNAME\Downloads\cookies.txt" `
  --xdr-cookie-path "M:\Users\$env:USERNAME\Downloads\xdr_cookies.txt"
```

`exposed` 会调用暴露面接口：先用 `company_id` 获取 `target_company_list`，按客户中文名匹配 `target_company` 并取 `target_company_id`；如果只有一个目标公司则直接使用该 ID。随后导出暴露面 zip，提取其中唯一的 xlsx，并把该 xlsx 的统计 sheet 写入总报告的 `暴露面` sheet。

## 输出

主输出是一份 Excel 工作簿：

```text
{customer或customer_id}_report.xlsx
```

默认保存到：

```text
$HOME\.openclaw\workspace\skills\sangfor-report-downloader
```

可用 `--output-dir` 指定其他输出目录。

如果使用 `--response-only`，则保存原始 JSON 响应文件，而不是生成工作簿。

## 重要说明

- 生成的工作簿可以作为后续 PPT 生成的 Excel 输入，但本技能本身只到 Excel 为止。
- PPT 模板由后续 `ai-ppt-pipeline` 选择：默认协同运营报告 `mss_classic_ops_2`；用户明确指定托管运营报告/simple 时使用 `mss_classic_ops_3`。
- 后续如有数据从 XDR 获取，使用 `--xdr-cookie-path` 指向 `xdr_cookies.txt`，不要复用 MSS 的 `cookies.txt`。
- 如果用户要的是 AI PPT，或者要对 PPT 进行重写，不要停留在本技能，应切换到 `ai-ppt-pipeline` 或 `ai-report-generator`。
- 如果用户明确只要某一个表，保留其显式指定的报告类型，不要默认扩成全部数据。
