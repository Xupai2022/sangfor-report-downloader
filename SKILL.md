# SKILL.md - 深信服报告下载技能

## 用途
根据用户自然语言请求，下载指定客户、指定时间段的 SOAR 报告数据。默认下载全部表，最终只生成一份 `{customer或customer_id}_report.xlsx`。

## 触发条件
用户说“下载某客户某时间段的报告/全部表/资产表/事件表/告警表/漏洞表”等类似需求时使用。

## 前置依赖
1. Chrome 浏览器插件 `M:\Users\xupai\Desktop\get_mss_cookie` 已安装并运行。
2. 插件已把 Cookie 写入文件。脚本默认读当前目录 `cookies.txt`，也支持 `--cookie-path` 指定目录或文件。
3. 生成总报告 Excel 需要安装依赖：`npm install`。

## 默认命令
```bash
node sangfor_downloader.js --customer "客户中文名" --start "2026-05-12" --end "2026-05-13" --cookie-path "M:\Users\xupai\Downloads"
```

也可以直接指定客户 ID：

```bash
node sangfor_downloader.js --customer "客户中文名" --customer-id 26912728 --start "2026-05-12" --end "2026-05-13" --cookie-path "M:\Users\xupai\Downloads"
```

默认不传 `--type`，等同于 `--type all`。可选值：

```text
all, asset, event, alarm, vuln
```

## 实际执行流程
1. 解析命令行参数。未传 `--type` 时，默认值是 `all`。
2. 读取 Cookie 和 X-Csrftoken。
3. 校验 `--start`、`--end` 日期格式。
4. 如果没有传 `--customer-id`，调用 `POST /order/v1/user/company_simple_info` 分页查询客户列表，用 `--customer` 的中文名精确匹配出 `customer_id`。
5. 构造公共请求参数：`customerId`、`customerName`、`startTime`、`endTime`、`pageSize`、`maxPages`。
6. 如果是 `--response-only`：
   - `all` 会依次保存资产表导出接口 response、事件表首页 response、告警表首页 response、资产漏洞表首页 response。
   - 不会下载资产表文件，也不会生成 Excel。
7. 非 `--response-only` 时，先下载资产表。触发条件是 `--type all`、`asset`、`event`、`vuln`。
   - 先 GET 访问：`https://soar.sangfor.com.cn/index.html#/customer/{customer_id}/business?company_name={客户中文名}`，模拟进入客户业务资产页面。
   - 再 POST：`/gateway/asset-mgr-service/order/v1/asset/download`。
   - 从响应 `data.url` 继续 GET 下载导出文件。
   - 资产表不再单独保存为 xlsx；下载后的工作簿会作为 `{customer或customer_id}_report.xlsx` 里的 `资产表` sheet。
8. 下载事件表。触发条件是 `--type all` 或 `event`。
   - 接口：`/gateway/event-mgr/external/event_table`。
   - 使用分页获取全部 `data.list`。
   - 转换后写入总报告的 `事件表` sheet。
   - “内网外网资产”优先读取本次刚下载的资产表；如果资产表下载失败，则回退读取 `--asset-map-file` 或默认 `./data/asset.xlsx`。
9. 下载告警表。触发条件是 `--type all` 或 `alarm`。
   - 接口：`/gateway/event-mgr/external/alarm_table`。
   - 使用分页获取全部 `data.list`。
   - 转换后写入总报告的 `告警表` sheet。
   - 当前告警表转换不依赖资产表。
10. 下载资产漏洞表。触发条件是 `--type all` 或 `vuln`。
    - 接口：`/gateway/vuln-manager/vm/order/v1/vulnmgr/vuln_list_port_split`。
    - 使用分页获取全部 `data.list`。
    - 从每条漏洞取 `name`、`manager`、`asset`。
    - `受影响主机/位置` 保留原始 `asset`。
    - 从 `asset` 提取 IPv4 写入 `IP`。
    - 遍历 `second_level`，取每个字典的 `vuln_status`；多个 `second_level` 拆成多行。
    - `vuln_status` 映射为中文跟进状态。
    - `内/外网` 优先按本次刚下载资产表里的 `*IP/URL`、`IP/URL`、`资产IP/URL` 等列匹配 IP，取对应 `安全域`；资产表下载失败时回退 `--asset-map-file` 或默认 `./data/asset.xlsx`。
    - 转换后写入总报告的 `资产漏洞表` sheet。
11. 生成总报告 `{customer或customer_id}_report.xlsx`。sheet 顺序固定为：
    - `数据统计`
    - `资产表`
    - `暴露面`
    - `资产漏洞表`
    - `告警表`
    - `事件表`
    当前尚未开发或本次未下载的数据表保留空 sheet。

## 默认输出
默认命令成功时，当前目录只会出现一份 Excel 报告：

```text
{customer或customer_id}_report.xlsx
```

如果某个接口失败，脚本会记录该表失败信息，并继续处理后续表。

## 资产表用途
资产表是后续表格的映射来源，不只是输出文件：

- 事件表用它补 `内网外网资产`。
- 资产漏洞表用它补 `内/外网`。
- 匹配列候选包括：`*IP/URL`、`IP/URL`、`资产IP/URL`、`asset`、`ip`、`host_ip`、`资产`、`资产IP`、`hostIp`。
- 安全域列候选包括：`security_domain`、`securityDomain`、`domain`、`内网外网资产`、`安全域`。

## 文件说明
- `sangfor_downloader.js` - 主下载脚本和流程编排
- `cookie_reader.js` - Cookie 读取模块
- `api_client.js` - 接口请求、资产表导出和文件下载
- `soar_transformer.js` - 事件表、告警表、资产漏洞表字段转换
- `data_formatter.js` - 单一总报告 Excel 生成
- `data/manage_sub_type_map.json` - 事件子类型映射
- `data/asset.xlsx` - 资产表下载失败时的备用资产映射文件

## 参数与环境变量
- 常用参数：`--customer`、`--customer-id`、`--start`、`--end`、`--cookie-path`、`--type`
- 映射参数：`--asset-map-file`、`--manage-sub-type-map-file`
- 环境变量：
  - `SANGFOR_COOKIE_PATH`
  - `SANGFOR_ASSET_MAP_FILE`
  - `SANGFOR_MANAGE_SUB_TYPE_MAP_FILE`
