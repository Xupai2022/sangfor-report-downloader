# 数据统计单元格取数请求条件（当前实现）

本文以 `data_formatter.js` 的 `buildStatisticsCells()` 和 `sangfor_downloader.js` 的实际调用为准，说明写入“数据统计”Sheet 的单元格所依赖的请求及请求体筛选条件。

## 统一参数

- `customer_id`：解析客户名称得到的客户 ID（MSS 请求中为字符串；事件/告警请求中为单元素数组）。
- `startDate/endDate`：报告起止日期。
- MSS 事件/告警/漏洞/弱口令请求的时间均转换为毫秒时间戳；开始为起始日 00:00:00，结束通常为结束日 23:59:59（弱口令结束值为结束日次日 00:00:00）。
- XDR 请求使用秒级绝对时间字符串/时间戳（由 `buildReportSecondRange` 生成），并在请求体 `globalCondition.time` 中传入。

## 请求 Profile

下面的编号在单元格映射中使用。

### MSS-E：事件表

接口：`POST /gateway/event-mgr/external/event_table`，分页抓取全部页。

请求体核心条件：`company_id=[customer_id]`、`create_time=[startTimestamp,endOfDay]`；其余筛选均为空/默认值（`asset_type=-1`、`hw_status=-1`、`protection_type=[]`、`event_status=[]`、`event_grading_tags=[]`、`my_customer=0`、`my_event=0`、`expired_customer=0`、`is_default=0` 等）。

### MSS-A：告警表

接口：`POST /gateway/alarm-mgr/v1/alarm_table/alarm_list`，分页抓取全部页。

请求体核心条件：`company_id=[customer_id]`、`create_time=[startTimestamp,endOfDay]`；`order={latest_time:'desc'}`，`asset_type=-1`、`hw_status=-1`、`attack_direction=-1`，`emergency_degree/risk_level/sip_sync_status/certainty_level/attack_state/dependability/principal_filter='all'`，其余文本/数组筛选为空。

### MSS-V：资产漏洞表

接口：`POST /gateway/vuln-manager/vm/order/v1/vulnmgr/vuln_list_port_split`，分页抓取全部页。

请求体核心条件：`company_id=customer_id`、`found_time=[startTimestamp,endOfDay]`；`vulnerability_status=[]`、`fix_level=-1`、`is_intranet=-1`、`service_status=0`、`vuln_type=-1`、`is_high_availability=-1`、`vulnerability_level=-1`，关键词、资产列表、场景等均为空。

### MSS-W0：全部弱口令汇总

接口：`POST /gateway/vuln-manager/vm/order/v1/weak_pwd/summary_list`，分页抓取。

请求体：`company_id=customer_id`、`found_time=[startTimestamp,endDate+1日00:00:00]`、`is_admin=0`、`deal_status=[]`；`order='asc'`，其余筛选数组为空，`reappear=0`。

### MSS-W：弱口令汇总（默认管理员口径）

同一 `summary_list` 接口；`company_id` 和 `found_time` 同上，调用未覆盖 `isAdmin`，因此构造器默认 `is_admin=1`；`deal_status=[]`。

### MSS-WH：已处理弱口令

同一接口；`company_id`、`found_time` 同上，`is_admin=0`、`deal_status=[2]`。

### MSS-WHA：管理员已处理弱口令

同一接口；`company_id`、`found_time` 同上，`is_admin=1`、`deal_status=[2]`。

> 弱口令汇总在拿到 IP 列表后，还会对每个 IP 调用 `POST .../weak_pwd/list`；该请求体是在对应汇总条件基础上增加 `ip=<该 IP>`。因此凡使用 `weakPwd*TotalsByIp` 的单元格，同时涉及汇总请求和逐 IP 明细请求。

### MSS-AS：资产表

接口：`POST /gateway/asset-mgr-service/order/v1/asset/download`。

请求体：`company_id=customer_id`、`is_select_all=1`、`service_status=[1]`（除非显式传入其他值）；资产类型、业务等级、时间、关键词、标签等筛选为空，`is_alive=-1`，其余 `filter_items` 使用代码中的固定开关。

### MSS-EX：暴露面

1. 目标公司选项：`POST .../exposed_surface_mss/get_target_company_option`，仅 `company_id=customer_id`。
2. 暴露面报表：`POST .../exposed_surface_mss/export_result_report`，`company_id=customer_id`、`ops=1`、`asset_tag=[2,5,3]`、`target_company=[target_company_id]`。
3. 端口/IP 统计：`POST .../exposed_surface_mss/ip_list_statistics`，`company_id=customer_id`、`target_company_id=[target_company_id]`、`attack_authorisation=0`，`keyword=''`，`related_task=[]`、`asset_tag=[]`。

### MSS-O(type)：订单分支

接口：`POST /gateway/customer-mgr-service/order/v1/branch/dev?_method=GET`。

请求体固定为 `company_id=customer_id`、`type=指定类型`、`status=0`、`platform_type='scloud'`、`belong_local_xdr=0`、`order='asc'`、`offset=0`、`limit=10`。调用类型：3、25、999、12。

### MSS-T：TopN

先调用 `load_condition` 和 `device_list`（请求体均仅 `company_id=customer_id`），再分别调用 TopN 接口。三次 TopN 请求共同条件：`company_id=customer_id`、`attack_type=[{id:'security_log'}]`、`attack_direction=[{id:'1'}]`、`device_infos=<device_list 返回的规范化设备组>`、`latest_time=[startTimestamp,endOfDay]`、`date_scope=1`、`resource_group_ids=[]`、`only_asset=0`；威胁类型和攻击源地域 `ip_type='src_ip', topn=100`，目的 IP 为 `ip_type='dst_ip', topn=5`。

### XDR-L：日志 count

接口：`POST /api/apex/logsearch/v1/log/search/count?...`；请求体为 `time:{start,end}`、`tables:['NetworkSecurityLog','EndpointSecurityLog']`。

### XDR-DIR(in2out/out2in)：访问方向日志 count

同一日志 count 接口；除 `time`、`tables` 外增加 `filter.logicalOp='and'`，条件 `field='accessDirection'`、`conditionalOp='IN'`、`value=['in2out']` 或 `['out2in']`。

### XDR-R：外到内且拒绝

接口：`POST /ngsoc/INCIDENT/api/v1/table/count/analysisTableQueryHandler?...`。`globalCondition.time.timeField='recordTimestamp'`、报告时间范围；SPL 条件：`srcIpTag=0` 且 `dstIpTag=1`，并筛选 `action` 为“拒绝”（前端值 `action=[2]`）；视图 `NetworkSecurityLogView+EndpointSecurityLogView`。

### XDR-IT：事件表 count

接口：`.../table/count/incidentTableQueryHandler?...`；时间字段 `endTime`，报告范围；视图 `IncidentView`；无额外 SPL（但表扩展条件固定 `xthConfirm=true`）。

### XDR-AT：告警表 count

接口：`.../table/count/alertTableQueryHandler?...`；时间字段 `lastTime`，报告范围；视图 `AlertView`；无额外 SPL。

### XDR-TD-A / XDR-TD-I：威胁定义分布

分别在 XDR-AT/XDR-IT 请求体上增加 `fieldName='threatDefine'`、`fieldValue={}`、`headerType='alertThreatDefine'` 或 `'eventThreatDefine'`。

### XDR-H：已处置事件

接口：`.../table/query/incidentTableQueryHandler?...` 的 count；时间字段 `endTime`；SPL/前端条件 `dealStatus=[3]`（处置完成）；扩展条件 `xthConfirm=true`。

### XDR-M：人工决策已处置且非白名单

事件表时间范围 `endTime`；SPL：`manualType != 1`、`dealStatus` 为处置完成、`whiteStatus != 已加白`。

### XDR-NW：已处置（非白名单）

事件表时间范围 `endTime`；SPL：`manualType != 1`、`dealStatus in [6,3,5]`（已管制/处置完成/已忽略）、`whiteStatus='未加白'`。

### XDR-P：处置中/挂起

事件表时间范围 `endTime`；`dealStatus in [4,2]`（处置中、挂起）。

### XDR-W：待处置

事件表时间范围 `endTime`；`dealStatus=[1]`（待处置）。

## 单元格映射

以下只列出涉及请求取数（或请求结果再计算）的单元格；纯模板值、空白值及仅由已算单元格相加/拼接得到的单元格不列出。

### 事件/告警/漏洞/资产/弱口令（MSS）

- `D6`：`D10 + D11 + D12`；同时涉及 MSS-V（高危且可利用漏洞）、MSS-W（管理员口径弱口令）/MSS-E（账号安全事件）以及 MSS-E（未公开威胁）的请求结果。
- `D7:D9`：MSS-AS、MSS-V、MSS-E、MSS-A 的结果按业务系统 IP 关联后的统计；四类请求均按各自 profile 的客户和报告时间条件执行。
- `D10`：MSS-V（高危且可利用）记录数。
- `D11`：MSS-W + MSS-E（弱口令/账号安全管理子类型事件）。
- `D12`：MSS-E（报告期内 `type` 包含“未公开威胁”），并在返回的 `affected_assets` 中统计唯一 IPv4。
- `D14:D15`：MSS-V 全量结果按跟进状态“已防护/已修复”统计。
- `D16`：MSS-WHA 的 `data.total` + MSS-E 报告期内账号安全事件中 `event_status=已闭环`。
- `D34`：MSS-V 返回的高危漏洞中“已防护”数量/高危总数。
- `D17`、`D67:D68`、`G110`、`G123`：复用 D12 的 MSS-E 请求结果。
- `D18`：MSS-A 全量结果中 `type` 包含“漏洞利用攻击”。
- `D27`：XDR-L（重保时间段逐段请求后求和）。
- `D28`：XDR-L（每个节假日重保区间逐段请求后按节假日归属汇总）。
- `D29`、`D91`、`F91:F97`、`G91:G97`：XDR-L 重保/节假日区间请求结果及其按天换算。
- `D35:D36`、`G51:G55`、`E49`：MSS-E 返回事件中事件分类属于 `isEventCategoryType` 且 `push_status=已通知`，分别平均识别/响应/遏制/处置/闭环时长字段。
- `D37`、`D83:D85`、`E83:E85`、`F83:F85`、`D110:D111`：MSS-V 返回结果按漏洞跟进状态、漏洞等级、高危 IP 等字段统计。
- `D62`：MSS-A 全量返回结果中 `type` 包含“威胁”的告警数。
- `D112`：MSS-W（`is_admin=1, deal_status=[]`）弱口令总数 + MSS-E 账号安全管理子类型事件数；`D113`：MSS-W 汇总返回的 IP 与 MSS-E 账号安全事件 `affected_assets` 的唯一 IP 数；`D114`：MSS-E 返回的未公开威胁事件数；`D115`：复用 D12；`D116:D117`：MSS-V 已防护/已修复漏洞数；`D118`：MSS-WHA 管理员已处理弱口令总数 + MSS-E 已闭环账号安全事件数。
- `D38`、`C49`、`G8`、`G9`、`F49`、`G7`、`G12`、`D63:D64`、`G10`、`G14`：MSS-E 返回事件按报告时间、事件等级分类、推送状态、事件状态、IP 关联等字段统计。
- `C80`、`D80`、`F80`、`I82`、`K81:K83`：MSS-V 返回结果按高危/可利用、内外网、跟进状态等字段统计。
- `I79:I80`、`D128`：MSS-EX 暴露面导出报表（I79/I80 读取导出 Sheet 的 E4/E5；D128 汇总 E2:E13）。
- `I81`：MSS-EX 端口/IP 统计接口返回值。
- `K79:K80`、`G16:G18`：MSS-AS 资产表按资产类型、网络域、服务类型统计。
- `G4`、`D22`：事件统计对象中的策略优化数量（由事件数据处理得到，底层为 MSS-E）。
- `G5`：MSS-A 报告期内 `type` 包含“外部威胁”。
- `G6`、`H6`、`I6`：MSS-A 返回告警按三个业务系统 IP 关联统计。
- `I122`、`I123`、`G125`：MSS-WH/MSS-W0 汇总及逐 IP `weak_pwd/list` 请求结果。
- `I124`、`G126`：MSS-E 报告期内指定管理子类型事件；`G126` 由 MSS-V 返回记录名称匹配“弱口令”。
- `D126`：MSS-E 报告期内管理子类型为“网站篡改”或“黑链”的事件数。
- `D130`：在 `protectStartDate~protectEndDate` 重新调用 MSS-E，将请求体时间条件覆盖为该重保区间，统计返回事件行数。
- `G111:G119`：MSS-AS 业务系统 IP 集合 + MSS-V/MSS-E，以及 MSS-W0/MSS-WH 逐 IP明细请求；按系统计算总数、已处理数和比例。
- `G122`、`G124`：分别复用 MSS-V 高危漏洞数、MSS-V 已修复漏洞数。

### XDR、订单、TopN

- `D20:D24`、`G21:G24`：`D23` 使用 XDR-R；`D24` 使用 MSS-A 全量结果中 `attack_direction=内->外`；其余为这些结果的本地组合。
- `D61`、`C71:N72`：XDR-DIR(out2in)，D61 为报告期总数，月度行按每月独立时间范围请求。
- `C71:N73`：XDR-DIR(out2in/in2out) 月度请求；第 72 行 out2in，第 73 行 in2out。
- `G99`：XDR-L 报告期日志总数。
- `G100`：XDR-AT 报告期告警总数。
- `G101`：XDR-IT 报告期事件总数。
- `G102`：XDR-H 报告期已处置事件总数。
- `G103`：XDR-M 报告期人工决策已处置非白名单事件总数。
- `G104`：`G102-G103`，不新增请求。
- `G105`：XDR-NW 报告期已处置非白名单事件总数。
- `G106`：XDR-P 报告期处置中/挂起事件总数。
- `G107`：XDR-W 报告期待处置事件总数。
- `G133:M134`：XDR-TD-A；每个威胁定义项的 count/占比来自告警威胁定义请求。
- `G135:P137`：XDR-TD-I；每个威胁定义项的 count/占比来自事件威胁定义请求。
- `D100`：MSS-O(type=3) + MSS-O(type=25) + MSS-O(type=999)。
- `D101`：MSS-O(type=3)。
- `D102`：MSS-O(type=25) + MSS-O(type=999)。
- `D105`：MSS-O(type=12)。
- `C41:E45`、`G41:I45`、`K41:M45`：MSS-T 的 TopN 三个接口返回的威胁类型、攻击源地理位置、目的 IP 列表（条件见 MSS-T）。

### 月度事件/告警/漏洞趋势

- `C58:N59`：复用 MSS-E 全量返回记录，按每月 `create_time` 过滤并平均响应时长。
- `C75:N75`：复用 MSS-A 全量返回记录，按每月 `create_time` 计数。
- `C86:N88`：复用 MSS-V 全量返回记录，按每月漏洞更新时间及跟进状态统计。

## 明确跳过的单元格

以下单元格当前为模板常量、空白、日期/名称展示、命令行直接传入值，或仅由其他单元格本地运算得到，没有独立请求体，因此按要求跳过：`J1`、`L1`、`M1`、`M3`、`D3:D5`、`D13`、`D25`、`D30`、`D32:D33`、`D39`、`M61`、`D65:D66`、`G11`、`G20`、`G22:G24`、`D90`、`D92`、`D121:D122`、`D129`、`N79`，以及各趋势区超出报告月份后被清空的单元格。
