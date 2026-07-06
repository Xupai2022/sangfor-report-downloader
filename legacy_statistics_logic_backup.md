# 数据统计旧逻辑备份

目的：备份 `数据统计` sheet 中 `D3:D18`、`D21:D25` 与 `G4:G12` 的旧/补充计算逻辑，后续重写时可直接对照。

来源文件：`data_formatter.js`

快照提交：`15c83fe4385b5c69e24b5d73482cc99e02130742`

说明：

- 这份备份以“表达式、代码片段、字段口径”为准。
- 行号在后续改动后会漂移，因此不作为长期引用依据。
- 如果需要回看原始上下文，以本提交版本的 `data_formatter.js` 为准。

## 写入位置

- `D3:D9` 由 `businessSystemCells` 写入。
- `D10:D18` 由 `cells` 对象直接写入。
- `D21:D25` 由 `cells` 对象直接写入。
- `G4:G12` 中，当前代码只写入 `G4/G5/G6/G7/G8/G9/G11/G12`，没有对 `G10` 赋值。

## D3:D18 旧逻辑

### D3、D4、D5

业务系统名称，来自 `businessSystems` 参数。

- 标准化入口：`normalizeBusinessSystems(value)`
- 拆分规则：按 `，、,；;` 和换行拆分
- 数量限制：最多 3 个
- 写入规则：
  - `D3 = 第 1 个业务系统名称`
  - `D4 = 第 2 个业务系统名称`
  - `D5 = 第 3 个业务系统名称`

### D6

统计周期内漏洞数：

```js
const d6 = vulnRows.filter(row => isDateInRange(row['更新时间'], range)).length;
```

含义：`vulnRows` 中 `更新时间` 落在 `startDate ~ endDate` 内的记录数。

### D7、D8、D9

分别对应 `D3/D4/D5` 三个业务系统的关联总数。

#### 1. 业务系统 IP 收集

从资产表收集每个业务系统的 IP 集：

- 从 `assetWorksheet` 第 3 行开始读取：`rows.slice(2)`
- 业务系统列：`B` 列，对应 `row[1]`
- IP 列：`C` 列，对应 `row[2]`
- 若资产表中“业务系统”文本 `includes(系统名)`，则把该行里提取出的全部 IPv4 加入该系统 IP 集

代码：

```js
function collectBusinessSystemIpSets(assetWorksheet, businessSystems) {
  const systems = normalizeBusinessSystems(businessSystems);
  const result = systems.map(name => ({ name, ips: new Set() }));
  if (result.length === 0 || !assetWorksheet) return result;

  const rows = getWorksheetRows(assetWorksheet);
  rows.slice(2).forEach((row) => {
    const businessSystemText = toText(row && row[1]);
    if (!businessSystemText) return;
    const rowIps = extractIPv4Texts(row && row[2]);
    if (rowIps.length === 0) return;

    result.forEach((item) => {
      if (!businessSystemText.includes(item.name)) return;
      rowIps.forEach(ip => item.ips.add(ip));
    });
  });

  return result;
}
```

#### 2. 关联总数计算

每个业务系统的总数为：

```js
total = vulnCount + eventCount + alarmCount
```

各项定义：

- `vulnCount`：`vulnRows` 中 `IP` 或 `受影响主机/位置` 包含该系统任一 IP 的记录数
- `eventCount`：`eventRows` 中 `host_ip` 或 `affected_assets` 包含该系统任一 IP 的记录数
- `alarmCount`：`alarmRows` 中 `host_ip` 包含该系统任一 IP 的记录数

代码：

```js
function buildBusinessSystemStatistics(statsContext) {
  const systems = collectBusinessSystemIpSets(
    statsContext && statsContext.assetWorksheet,
    statsContext && statsContext.businessSystems
  );
  const eventRows = statsContext && Array.isArray(statsContext.eventRows) ? statsContext.eventRows : [];
  const alarmRows = statsContext && Array.isArray(statsContext.alarmRows) ? statsContext.alarmRows : [];
  const vulnRows = statsContext && Array.isArray(statsContext.vulnRows) ? statsContext.vulnRows : [];

  return systems.map((system) => {
    const vulnCount = countRowsContainingIps(vulnRows, ['IP', '受影响主机/位置'], system.ips);
    const eventCount = countRowsContainingIps(eventRows, ['host_ip', 'affected_assets'], system.ips);
    const alarmCount = countRowsContainingIps(alarmRows, ['host_ip'], system.ips);

    return {
      name: system.name,
      ips: Array.from(system.ips),
      total: vulnCount + eventCount + alarmCount,
      vulnCount,
      eventCount,
      alarmCount
    };
  });
}
```

写入规则：

- `D7 = 第 1 个业务系统 total`
- `D8 = 第 2 个业务系统 total`
- `D9 = 第 3 个业务系统 total`

测试样例：

- `tests/protection_atomic.test.js` 中 `testBusinessSystemStatisticsByAssetIps`

### D10

```js
const e80WeakPwdTotal = toNumericOrNull(weakPwdSummaryTotal) || 0;
const e80AccountSecurityEventCount = toNumericOrNull(
  eventStats && eventStats.accountSecurityEventCountForE80
) || 0;
const e80 = e80WeakPwdTotal + e80AccountSecurityEventCount;
const d10 = e80;
```

含义：弱口令汇总数 + 指定账号安全事件数。

### D11

```js
const d11 = eventRows.filter((row) => (
  toText(row.event_grading_tag) === '最新威胁'
  && isDateInRange(row.create_time, range)
)).length;
```

含义：统计周期内，`event_grading_tag = 最新威胁` 的事件数。

### D12

```js
const d12 = eventRows
  .filter(row => (
    isDateInRange(row.create_time, range)
    && toText(row.type).includes('未公开威胁')
  ))
  .reduce((total, row) => total + countAffectedAssetIps(row.affected_assets), 0);
```

含义：统计周期内，`type` 包含 `未公开威胁` 的事件，对每条事件统计 `affected_assets` 中唯一 IPv4 数量，再累加。

`countAffectedAssetIps`：

```js
function countAffectedAssetIps(value) {
  if (!Array.isArray(value)) return 0;
  const uniqueIps = new Set();
  value.forEach((item) => {
    const text = String(item || '').trim();
    if (isIPv4Text(text)) uniqueIps.add(text);
  });
  return uniqueIps.size;
}
```

### D13

当前代码未赋值，项目中无 `D13` 写入逻辑。

说明：最终值取决于模板 `data.xlsx` 中是否已有内容或公式。

### D14

```js
const d14 = (toNumericOrNull(d15) || 0) + d16 + d17;
```

含义：`D15 + D16 + D17`

### D15

```js
const d15 = d12;
```

含义：直接等于 `D12`。

### D16

```js
const d16 = vulnRows.filter(row => toText(row['跟进状态']) === '已防护').length;
```

含义：全部漏洞中 `跟进状态 = 已防护` 的数量，不按时间过滤。

### D17

```js
const d17 = vulnRows.filter(row => toText(row['跟进状态']) === '已修复').length;
```

含义：全部漏洞中 `跟进状态 = 已修复` 的数量，不按时间过滤。

### D18

```js
const d18 = alarmRows.filter(row => toText(row.type).includes('漏洞利用攻击')).length;
```

含义：全部告警中 `type` 包含 `漏洞利用攻击` 的数量，不按时间过滤。

## D20:D25 补充逻辑

说明：这一段不是最早快照主体的一部分，这里补充当前仓库中 `D20:D25` 的写入规则，方便后续重取数时参考。

### D20

```js
D20: d21 + g22
```

当前代码中：

```js
const d21 = (toNumericOrNull(xdrRejectedExternalToInternalCount) || 0) + d24;
const g22 = d14 + d15;
const d20 = d21 + g22;
```

即：

```text
D20 = D21 + G22
```

### D21

```js
D21: d21
```

当前代码中：

```js
const d21 = (toNumericOrNull(xdrRejectedExternalToInternalCount) || 0) + d24;
```

即：

```text
D21 = D23 + D24
```

### D22

```js
D22: g4
```

即：

```text
D22 = G4
```

### D23

```js
D23: xdrRejectedExternalToInternalCount
```

含义：严格按指定 XDR INCIDENT count 接口返回的 `data.total` 取值。

接口：

```text
/ngsoc/INCIDENT/api/v1/table/count/analysisTableQueryHandler?viewRegionId=ffffffffffffffffffffffff&onlySelfPlatform=false
```

核心筛选条件：

```js
spl.mappedSpl = '(srcIpTag = 0 and dstIpTag = 1) | filter 动作  in { "拒绝" }'
spl.originalSpl = '(srcIpTag = 0 and dstIpTag = 1) | filter 动作  in { "拒绝" }'
```

并带：

```js
globalCondition: {
  branchIds: [],
  time: {
    timeField: 'recordTimestamp',
    end: { type: 'absolute', value: params.end },
    begin: { type: 'absolute', value: params.start }
  }
}
```

以及前端筛选展示参数：

```js
frontRender: [
  {
    displayField: '动作',
    field: 'action',
    value: [2],
    valueText: '拒绝',
    headerType: 'metaType',
    searchType: 'selector',
    type: 'number',
    filterSelect: 'renderValue',
    isValueNegate: false
  }
]
```

视图固定为：

```js
viewName: 'NetworkSecurityLogView+EndpointSecurityLogView'
table.viewName: 'NetworkSecurityLogView+EndpointSecurityLogView'
```

取值规则：

- 时间范围：报表 `startDate ~ endDate` 转成秒级绝对时间
- 条件 1：`srcIpTag = 0`
- 条件 2：`dstIpTag = 1`
- 条件 3：`动作 = 拒绝`，对应 `action = 2`
- 返回值：直接取响应中的 `data.total`

示例响应：

```json
{"strCode":null,"message":"成功","data":{"total":26354},"code":0}
```

最终：

```text
D23 = response.data.total
```

### D24

```js
D24: alarmRows.filter((row) => (
  toText(row.attack_direction) === '内-外'
)).length
```

含义：直接在告警表统计 `访问方向 = 内-外` 的记录数。

说明：

- 这里不再复用 `D16`
- 不依赖 Excel 公式
- 直接对 `alarmRows` 做计数
- `AE` 列对应转换后的字段名是 `attack_direction`

即：

```text
D24 = COUNTIFS(告警表!AE:AE, "内-外")
```

### D25

```js
D25: ''
```

含义：当前代码固定写空字符串，没有计算逻辑。

说明：程序不会主动赋值，最终显示效果取决于模板中原值是否被保留；按当前写法会跳过写入。

## G21:G24 补充逻辑

### G21

```js
G21: ''
```

含义：固定空字符串。

### G22

```js
G22: g22
```

当前代码中：

```js
const g22 = d14 + d15;
```

即：

```text
G22 = G23 + G24
```

### G23

```text
G23 = D14
```

### G24

```text
G24 = D15
```

## G4:G12 旧逻辑

### G4

```js
const g4 = eventStats && eventStats.strategyOptimizeCount !== undefined
  ? eventStats.strategyOptimizeCount
  : '';
```

含义：直接取 `eventStats.strategyOptimizeCount`，未提供则为空字符串。

### G5

```js
const g5 = alarmRows.filter((row) => (
  isDateInRange(row.create_time, range)
  && toText(row.type).includes('外部威胁')
)).length;
```

含义：统计周期内，`type` 包含 `外部威胁` 的告警数。

### G6

```js
const g6Numbers = eventRows
  .filter(row => isThreatEventType(row.event_grading_tag))
  .map(row => toNumericOrNull(row['响应时长']))
  .filter(value => value !== null);
const g6 = averageNumbers(g6Numbers);
```

含义：威胁类事件的平均响应时长。

说明：这里不按时间范围过滤，只要求 `isThreatEventType(event_grading_tag)`。

### G7

```js
const g7 = eventRows.filter((row) => (
  isDateInRange(row.create_time, range)
  && isEventCategoryType(row.event_grading_tag)
)).length;
```

含义：统计周期内，事件分类属于 `isEventCategoryType` 的事件数。

### G8

```js
const g8Numbers = eventRows
  .filter((row) => (
    isEventCategoryType(row.event_grading_tag)
    && toText(row.push_status) === '已通告'
  ))
  .map(row => toNumericOrNull(row['响应时长']))
  .filter(value => value !== null);
const g8 = averageNumbers(g8Numbers);
```

含义：事件分类属于 `isEventCategoryType` 且 `push_status = 已通告` 的事件，平均响应时长。

说明：这里也不按时间范围过滤。

### G9

```js
const g9HandledCount = eventRows.filter((row) => (
  isDateInRange(row.create_time, range)
  && isEventCategoryType(row.event_grading_tag)
  && toText(row.event_status).includes('已')
)).length;
const g9IgnoredCount = eventRows.filter((row) => (
  isDateInRange(row.create_time, range)
  && isEventCategoryType(row.event_grading_tag)
  && toText(row.event_status) === '不处置'
)).length;
const g9 = g7 > 0 ? formatRatioAsPercentage((g9HandledCount + g9IgnoredCount) / g7, 2) : '0%';
```

含义：统计周期内，事件分类属于 `isEventCategoryType` 的事件中，状态包含 `已` 或等于 `不处置` 的占比。

### G10

当前代码未赋值，项目中无 `G10` 写入逻辑。

说明：最终值取决于模板 `data.xlsx` 中是否已有内容或公式。

### G11

```js
const g11 = d6 + d10 + d12;
```

含义：`D6 + D10 + D12`

### G12

```js
const g12 = eventRows.filter((row) => (
  isDateInRange(row.create_time, range)
  && isStatisticsEventCountType(row.event_grading_tag)
)).length;
```

含义：统计周期内，事件分类满足 `isStatisticsEventCountType` 的事件数。

## 备注

- 本备份仅记录旧逻辑，不改动现有实现。
- 后续如果重写这部分逻辑，优先保留这个文件，便于核对差异和回退口径。
