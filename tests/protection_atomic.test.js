const assert = require('assert');
const XLSX = require('xlsx');
const dataFormatter = require('../data_formatter');
const apiClient = require('../api_client');
const soarTransformer = require('../soar_transformer');

function ownerLabels(range) {
  return (range.owners || []).map(owner => `${owner.source}:${owner.name}`).sort();
}

function compactRanges(ranges) {
  return ranges.map(range => ({
    startDate: range.startDate,
    endDate: range.endDate,
    owners: ownerLabels(range)
  }));
}

function findHoliday(holidayPeriods, name) {
  return holidayPeriods.find(period => period.name === name);
}

function assertRange(actual, expected) {
  assert.deepStrictEqual(compactRanges(actual), expected);
}

function testInclusiveDayCounts() {
  assert.strictEqual(dataFormatter.countInclusiveNaturalDays('2026-02-01', '2026-02-01'), 1);
  assert.strictEqual(dataFormatter.countInclusiveNaturalDays('2026-02-01', '2026-02-28'), 28);
  assert.strictEqual(dataFormatter.countInclusiveNaturalDays('2024-02-01', '2024-02-29'), 29);
  assert.strictEqual(dataFormatter.countInclusiveNaturalDays('2026-01-27', '2026-03-03'), 36);
  assert.strictEqual(dataFormatter.countInclusiveNaturalDays('2026-01-27', '2026-01-31'), 5);
  assert.strictEqual(dataFormatter.countInclusiveNaturalDays('2026-03-01', '2026-03-03'), 3);
  assert.strictEqual(dataFormatter.countInclusiveNaturalDays('2025-12-31', '2026-01-02'), 3);
}

function testHolidayClipping() {
  let query = dataFormatter.buildProtectionAtomicSecondRanges({
    startDate: '2026-01-01',
    endDate: '2026-12-31'
  });
  let spring = findHoliday(query.holidayPeriods, '春节');
  assert.strictEqual(spring.startDate, '2026-02-15');
  assert.strictEqual(spring.endDate, '2026-02-23');
  assert.strictEqual(spring.dayCount, 9);

  query = dataFormatter.buildProtectionAtomicSecondRanges({
    startDate: '2026-02-18',
    endDate: '2026-03-03'
  });
  spring = findHoliday(query.holidayPeriods, '春节');
  assert.strictEqual(spring.startDate, '2026-02-18');
  assert.strictEqual(spring.endDate, '2026-02-23');
  assert.strictEqual(spring.dayCount, 6);

  query = dataFormatter.buildProtectionAtomicSecondRanges({
    startDate: '2026-01-27',
    endDate: '2026-02-18'
  });
  spring = findHoliday(query.holidayPeriods, '春节');
  assert.strictEqual(spring.startDate, '2026-02-15');
  assert.strictEqual(spring.endDate, '2026-02-18');
  assert.strictEqual(spring.dayCount, 4);

  query = dataFormatter.buildProtectionAtomicSecondRanges({
    startDate: '2026-03-01',
    endDate: '2026-03-31'
  });
  assert.strictEqual(findHoliday(query.holidayPeriods, '春节'), undefined);
  assert.strictEqual(query.ranges.length, 0);
}

function testNonOverlappingManualAndHoliday() {
  const query = dataFormatter.buildProtectionAtomicSecondRanges({
    startDate: '2026-02-01',
    endDate: '2026-02-28',
    protectStartDate: '2026-02-01',
    protectEndDate: '2026-02-03'
  });
  assertRange(query.ranges, [
    { startDate: '2026-02-01', endDate: '2026-02-03', owners: ['manual:护网时间'] },
    { startDate: '2026-02-15', endDate: '2026-02-23', owners: ['holiday:春节'] }
  ]);
}

function testLeftOverlap() {
  const query = dataFormatter.buildProtectionAtomicSecondRanges({
    startDate: '2026-02-14',
    endDate: '2026-02-23',
    protectStartDate: '2026-02-14',
    protectEndDate: '2026-02-18'
  });
  assertRange(query.ranges, [
    { startDate: '2026-02-14', endDate: '2026-02-14', owners: ['manual:护网时间'] },
    { startDate: '2026-02-15', endDate: '2026-02-18', owners: ['holiday:春节', 'manual:护网时间'] },
    { startDate: '2026-02-19', endDate: '2026-02-23', owners: ['holiday:春节'] }
  ]);
}

function testRightOverlap() {
  const query = dataFormatter.buildProtectionAtomicSecondRanges({
    startDate: '2026-02-15',
    endDate: '2026-02-25',
    protectStartDate: '2026-02-20',
    protectEndDate: '2026-02-25'
  });
  assertRange(query.ranges, [
    { startDate: '2026-02-15', endDate: '2026-02-19', owners: ['holiday:春节'] },
    { startDate: '2026-02-20', endDate: '2026-02-23', owners: ['holiday:春节', 'manual:护网时间'] },
    { startDate: '2026-02-24', endDate: '2026-02-25', owners: ['manual:护网时间'] }
  ]);
}

function testManualContainsHoliday() {
  const query = dataFormatter.buildProtectionAtomicSecondRanges({
    startDate: '2026-02-01',
    endDate: '2026-02-28',
    protectStartDate: '2026-02-01',
    protectEndDate: '2026-02-28'
  });
  assertRange(query.ranges, [
    { startDate: '2026-02-01', endDate: '2026-02-14', owners: ['manual:护网时间'] },
    { startDate: '2026-02-15', endDate: '2026-02-23', owners: ['holiday:春节', 'manual:护网时间'] },
    { startDate: '2026-02-24', endDate: '2026-02-28', owners: ['manual:护网时间'] }
  ]);
}

function testHolidayContainsManual() {
  const query = dataFormatter.buildProtectionAtomicSecondRanges({
    startDate: '2026-02-15',
    endDate: '2026-02-23',
    protectStartDate: '2026-02-17',
    protectEndDate: '2026-02-19'
  });
  assertRange(query.ranges, [
    { startDate: '2026-02-15', endDate: '2026-02-16', owners: ['holiday:春节'] },
    { startDate: '2026-02-17', endDate: '2026-02-19', owners: ['holiday:春节', 'manual:护网时间'] },
    { startDate: '2026-02-20', endDate: '2026-02-23', owners: ['holiday:春节'] }
  ]);
}

function testOverlappingHolidaysAndThreeWayOverlap() {
  const query = dataFormatter.buildProtectionAtomicSecondRanges({
    startDate: '2025-10-01',
    endDate: '2025-10-08',
    protectStartDate: '2025-10-05',
    protectEndDate: '2025-10-06'
  });
  assertRange(query.ranges, [
    { startDate: '2025-10-01', endDate: '2025-10-04', owners: ['holiday:国庆'] },
    { startDate: '2025-10-05', endDate: '2025-10-05', owners: ['holiday:国庆', 'manual:护网时间'] },
    { startDate: '2025-10-06', endDate: '2025-10-06', owners: ['holiday:中秋', 'holiday:国庆', 'manual:护网时间'] },
    { startDate: '2025-10-07', endDate: '2025-10-07', owners: ['holiday:中秋', 'holiday:国庆'] },
    { startDate: '2025-10-08', endDate: '2025-10-08', owners: ['holiday:中秋'] }
  ]);
}

function testAdjacentRangesAreNotOverlapped() {
  const query = dataFormatter.buildProtectionAtomicSecondRanges({
    startDate: '2026-02-12',
    endDate: '2026-02-23',
    protectStartDate: '2026-02-12',
    protectEndDate: '2026-02-14'
  });
  assertRange(query.ranges, [
    { startDate: '2026-02-12', endDate: '2026-02-14', owners: ['manual:护网时间'] },
    { startDate: '2026-02-15', endDate: '2026-02-23', owners: ['holiday:春节'] }
  ]);
}

function testNoProtectionRanges() {
  const query = dataFormatter.buildProtectionAtomicSecondRanges({
    startDate: '2026-03-01',
    endDate: '2026-03-10'
  });
  assert.deepStrictEqual(query.ranges, []);
  assert.deepStrictEqual(query.holidayPeriods, []);
}

function testAtomicRangesMatchMergedRangesWhenNotOverlapped() {
  const context = {
    startDate: '2026-02-01',
    endDate: '2026-02-28',
    protectStartDate: '2026-02-01',
    protectEndDate: '2026-02-03'
  };
  const atomic = dataFormatter.buildProtectionAtomicSecondRanges(context).ranges
    .map(range => ({ start: range.start, end: range.end }));
  const merged = dataFormatter.buildProtectionSecondRanges(context);
  assert.deepStrictEqual(atomic, merged);
}

function testCountAggregation() {
  const query = dataFormatter.buildProtectionAtomicSecondRanges({
    startDate: '2026-02-14',
    endDate: '2026-02-23',
    protectStartDate: '2026-02-14',
    protectEndDate: '2026-02-18'
  });
  const queriedRanges = query.ranges.map((range, index) => ({
    ...range,
    count: [5, 40, 50][index]
  }));
  const summary = apiClient.aggregateXdrLogSearchCountDetails(queriedRanges);
  const spring = findHoliday(query.holidayPeriods, '春节');
  assert.strictEqual(summary.total, 95);
  assert.strictEqual(summary.holidayCounts[spring.key], 90);
}

function testOverlappingHolidayAggregation() {
  const query = dataFormatter.buildProtectionAtomicSecondRanges({
    startDate: '2025-10-01',
    endDate: '2025-10-08'
  });
  const queriedRanges = query.ranges.map((range, index) => ({
    ...range,
    count: [50, 20, 10][index]
  }));
  const summary = apiClient.aggregateXdrLogSearchCountDetails(queriedRanges);
  const national = findHoliday(query.holidayPeriods, '国庆');
  const midAutumn = findHoliday(query.holidayPeriods, '中秋');
  assert.strictEqual(summary.total, 80);
  assert.strictEqual(summary.holidayCounts[national.key], 70);
  assert.strictEqual(summary.holidayCounts[midAutumn.key], 30);
}

function testAssetDownloadRequestUsesServiceStatusOneByDefault() {
  const body = apiClient.buildAssetDownloadRequestBody({
    customerId: '99368607'
  });

  assert.deepStrictEqual(body.service_status, [1]);
}

function testStatisticsCellsHolidayDailyAverage() {
  const query = dataFormatter.buildProtectionAtomicSecondRanges({
    startDate: '2026-02-18',
    endDate: '2026-03-03'
  });
  const spring = findHoliday(query.holidayPeriods, '春节');
  const cells = dataFormatter.buildStatisticsCells({
    startDate: '2026-02-18',
    endDate: '2026-03-03',
    xdrLogSearchCount: 60,
    xdrHolidayLogSearchCounts: {
      [spring.key]: 10
    }
  });
  assert.strictEqual(cells.F91, '春节（2026/02/18-2026/02/23）');
  assert.strictEqual(cells.G91, 1.67);
  assert.strictEqual(cells.G92, '');
  assert.strictEqual(cells.H91, '100%');
  assert.strictEqual(cells.H92, '');
  assert.strictEqual(cells.D27, 60);
}

function testStatisticsCellsIntegerHolidayDailyAverage() {
  const query = dataFormatter.buildProtectionAtomicSecondRanges({
    startDate: '2026-02-01',
    endDate: '2026-02-28'
  });
  const spring = findHoliday(query.holidayPeriods, '春节');
  const cells = dataFormatter.buildStatisticsCells({
    startDate: '2026-02-01',
    endDate: '2026-02-28',
    xdrHolidayLogSearchCounts: {
      [spring.key]: 900
    }
  });
  assert.strictEqual(cells.G91, 100);
}

function testStatisticsCellsMissingHolidayCountLeavesAverageBlank() {
  const cells = dataFormatter.buildStatisticsCells({
    startDate: '2026-02-01',
    endDate: '2026-02-28',
    xdrHolidayLogSearchCounts: {}
  });
  assert.strictEqual(cells.F91, '春节（2026/02/15-2026/02/23）');
  assert.strictEqual(cells.G91, '');
}

function testStatisticsCellsManualOnlyLeavesHolidayAverageBlank() {
  const cells = dataFormatter.buildStatisticsCells({
    startDate: '2026-03-01',
    endDate: '2026-03-10',
    protectStartDate: '2026-03-02',
    protectEndDate: '2026-03-03',
    xdrLogSearchCount: 30,
    xdrHolidayLogSearchCounts: {}
  });
  assert.strictEqual(cells.D27, 30);
  assert.strictEqual(cells.F91, '');
  assert.strictEqual(cells.G91, '');
  assert.strictEqual(cells.H91, '');
}

function testStatisticsCellsUseUpdatedD6D10D11() {
  const cells = dataFormatter.buildStatisticsCells({
    startDate: '2026-03-01',
    endDate: '2026-03-31',
    vulnRows: [
      { '更新时间': '2026-03-01', '漏洞等级': '高危', '是否可利用': '是' },
      { '更新时间': '2026-03-31 23:59:59' },
      { '更新时间': '2026-04-01' }
    ],
    eventRows: [
      {
        create_time: '2026-03-10',
        type: '未公开威胁',
        affected_assets: ['10.0.0.1', '10.0.0.2', '10.0.0.2']
      },
      {
        create_time: '2026-04-01',
        type: '未公开威胁',
        affected_assets: ['10.0.0.3']
      },
      {
        create_time: '2026-03-15',
        __manageSubTypeDisplay: '弱口令检测'
      },
      {
        create_time: '2026-03-16',
        manage_sub_type_cn: '账号安全异常'
      },
      {
        create_time: '2026-03-17',
        manage_sub_type_name: '弱密码爆破'
      }
    ],
    weakPwdSummaryTotal: 3
  });

  assert.strictEqual(cells.D6, 9);
  assert.strictEqual(cells.D10, 1);
  assert.strictEqual(cells.D11, 6);
  assert.strictEqual(cells.D12, 2);
  assert.strictEqual(cells.D6, cells.D10 + cells.D11 + cells.D12);
  assert.strictEqual(cells.E80, 3);
  assert.strictEqual(cells.G11, '');
}

function testWeakPwdSummaryRequestMatchesCapturedTimeRange() {
  const body = apiClient.buildWeakPwdSummaryRequestBody({
    startTime: '2026-01-01',
    endTime: '2026-04-29',
    customerId: '26912728'
  });

  assert.deepStrictEqual(body.found_time, [1767196800000, 1777478400000]);
  assert.strictEqual(body.is_admin, 1);
  assert.deepStrictEqual(body.service_status, []);
  assert.strictEqual(body.company_id, '26912728');
}

function testWeakPwdSummaryRequestSupportsDealStatusFilter() {
  const body = apiClient.buildWeakPwdSummaryRequestBody({
    startTime: '2026-01-01',
    endTime: '2026-04-01',
    customerId: '26912728',
    dealStatus: [2]
  });

  assert.deepStrictEqual(body.found_time, [1767196800000, 1775059200000]);
  assert.deepStrictEqual(body.deal_status, [2]);
  assert.deepStrictEqual(body.service_status, []);
}

function testWeakPwdSummaryRequestSupportsPagination() {
  const body = apiClient.buildWeakPwdSummaryRequestBody({
    startTime: '2026-01-01',
    endTime: '2026-04-01',
    customerId: '26912728',
    page: 3,
    pageSize: 200
  });

  assert.strictEqual(body.offset, 400);
  assert.strictEqual(body.limit, 200);
  assert.deepStrictEqual(body.found_time, [1767196800000, 1775059200000]);
}

function testPopulateStatisticsSheetWritesAndClearsHolidayAverages() {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([]);
  ws.G92 = { t: 'n', v: 999 };
  ws.H92 = { t: 's', v: '100%' };
  ws.F92 = { t: 's', v: '旧节假日' };
  ws['!ref'] = 'A1:H97';
  XLSX.utils.book_append_sheet(wb, ws, '数据统计');

  const query = dataFormatter.buildProtectionAtomicSecondRanges({
    startDate: '2026-02-01',
    endDate: '2026-02-28'
  });
  const spring = findHoliday(query.holidayPeriods, '春节');
  dataFormatter.populateStatisticsSheet(wb, {
    startDate: '2026-02-01',
    endDate: '2026-02-28',
    xdrHolidayLogSearchCounts: {
      [spring.key]: 900
    }
  });

  assert.strictEqual(ws.F91.v, '春节（2026/02/15-2026/02/23）');
  assert.strictEqual(ws.G91.v, 100);
  assert.strictEqual(ws.H91.v, '100%');
  assert.strictEqual(ws.F92.v, '');
  assert.strictEqual(ws.G92.v, '');
  assert.strictEqual(ws.H92.v, '');
}

function testBusinessSystemStatisticsByAssetIps() {
  const assetWorksheet = XLSX.utils.aoa_to_sheet([
    ['', '', ''],
    ['序号', '业务系统', 'IP'],
    [1, '核心系统A', '10.0.0.1'],
    [2, '系统B', '10.0.0.2, 10.0.0.3'],
    [3, '核心系统A和系统C', '10.0.0.4'],
    [4, '其他系统', '10.0.0.5']
  ]);

  const cells = dataFormatter.buildStatisticsCells({
    startDate: '2026-03-01',
    endDate: '2026-03-31',
    businessSystems: ['核心系统A', '系统B', '系统C'],
    assetWorksheet,
    vulnRows: [
      { IP: '10.0.0.1', '受影响主机/位置': '', '漏洞等级': '高危', '是否可利用': '是', '更新时间': '2026-03-05' },
      { IP: '', '受影响主机/位置': 'http://10.0.0.2/path', '漏洞等级': '高危', '是否可利用': '是', '更新时间': '2026-03-06' },
      { IP: '10.0.0.4', '受影响主机/位置': '10.0.0.3', '漏洞等级': '中危', '是否可利用': '是', '更新时间': '2026-03-07' }
    ],
    eventRows: [
      {
        create_time: '2026-03-10',
        host_ip: '10.0.0.4',
        affected_assets: ['10.0.0.9'],
        manage_sub_type_cn: '账号安全异常'
      },
      {
        create_time: '2026-03-11',
        host_ip: '',
        affected_assets: ['10.0.0.3', '10.0.0.2'],
        type: '未公开威胁'
      }
    ],
    weakPwdSummaryList: [
      { ip: '10.0.0.1' },
      { ip: '10.0.0.3' },
      { ip: '10.0.0.4' }
    ]
  });

  assert.strictEqual(cells.D3, '核心系统A');
  assert.strictEqual(cells.D4, '系统B');
  assert.strictEqual(cells.D5, '系统C');
  assert.strictEqual(cells.D7, 4);
  assert.strictEqual(cells.D8, 4);
  assert.strictEqual(cells.D9, 2);
}

function testStatisticsCellsUseUpdatedD13ToD17() {
  const cells = dataFormatter.buildStatisticsCells({
    startDate: '2026-03-01',
    endDate: '2026-03-31',
    vulnRows: [
      { '跟进状态': '已防护', '更新时间': '2026-03-05' },
      { '跟进状态': '已防护', '更新时间': '2026-04-01' },
      { '跟进状态': '已修复', '更新时间': '2026-03-06' },
      { '跟进状态': '已修复', '更新时间': '2026-02-28' }
    ],
    eventRows: [
      {
        create_time: '2026-03-10',
        type: '未公开威胁',
        affected_assets: ['10.0.0.1', '10.0.0.2']
      },
      {
        create_time: '2026-03-12',
        event_status: '已闭环',
        manage_sub_type_cn: '账号安全异常'
      },
      {
        create_time: '2026-03-13',
        event_status: '已闭环',
        manage_sub_type_cn: '弱密码风险'
      },
      {
        create_time: '2026-03-14',
        event_status: '处置中',
        manage_sub_type_cn: '弱口令检测'
      }
    ],
    weakPwdHandledTotal: 6
  });

  assert.strictEqual(cells.D12, 2);
  assert.strictEqual(cells.D13, 4);
  assert.strictEqual(cells.D14, 1);
  assert.strictEqual(cells.D15, 1);
  assert.strictEqual(cells.D16, 8);
  assert.strictEqual(cells.D17, 2);
}

function testStatisticsCellsUseXdrRejectedExternalToInternalForD23() {
  const cells = dataFormatter.buildStatisticsCells({
    startDate: '2026-03-01',
    endDate: '2026-03-31',
    xdrRejectedExternalToInternalCount: 26354,
    eventRows: [
      {
        create_time: '2026-03-10',
        type: '未公开威胁',
        affected_assets: ['10.0.0.1', '10.0.0.2']
      }
    ]
  });

  assert.strictEqual(cells.D17, 2);
  assert.strictEqual(cells.D23, 26354);
}

function testStatisticsCellsUseAlarmAttackDirectionForD24AndMirrorD14D15ToG23G24() {
  const cells = dataFormatter.buildStatisticsCells({
    startDate: '2026-03-01',
    endDate: '2026-03-31',
    eventStats: {
      strategyOptimizeCount: 7
    },
    xdrRejectedExternalToInternalCount: 20,
    vulnRows: [
      { '跟进状态': '已防护', '更新时间': '2026-03-05' },
      { '跟进状态': '已修复', '更新时间': '2026-03-06' },
      { '跟进状态': '已防护', '更新时间': '2026-04-01' }
    ],
    alarmRows: [
      { attack_direction: '内-外' },
      { attack_direction: '外-内' },
      { attack_direction: '内-外' },
      { attack_direction: '未知' }
    ]
  });

  assert.strictEqual(cells.D14, 1);
  assert.strictEqual(cells.D15, 1);
  assert.strictEqual(cells.D24, 2);
  assert.strictEqual(cells.D21, 22);
  assert.strictEqual(cells.D22, 7);
  assert.strictEqual(cells.D20, 24);
  assert.strictEqual(cells.G21, '');
  assert.strictEqual(cells.G22, 2);
  assert.strictEqual(cells.G23, 1);
  assert.strictEqual(cells.G24, 1);
}

function testXdrRejectedExternalToInternalRequestMatchesCapturedFilters() {
  const body = apiClient.buildXdrRejectedExternalToInternalCountRequestBody({
    start: 1782907517,
    end: 1782993917
  });

  assert.strictEqual(body.spl.mappedSpl, '(srcIpTag = 0 and dstIpTag = 1) | filter 动作  in { "拒绝" }');
  assert.strictEqual(body.spl.originalSpl, '(srcIpTag = 0 and dstIpTag = 1) | filter 动作  in { "拒绝" }');
  assert.deepStrictEqual(body.spl.extensionParams.frontRender[0].value, [2]);
  assert.strictEqual(body.spl.extensionParams.frontRender[0].valueText, '拒绝');
  assert.strictEqual(body.globalCondition.time.timeField, 'recordTimestamp');
  assert.strictEqual(body.globalCondition.time.begin.value, 1782907517);
  assert.strictEqual(body.globalCondition.time.end.value, 1782993917);
  assert.strictEqual(body.viewName, 'NetworkSecurityLogView+EndpointSecurityLogView');
  assert.strictEqual(body.table.viewName, 'NetworkSecurityLogView+EndpointSecurityLogView');
}

function testStatisticsCellsUseUpdatedG7ToG10AndG12() {
  const assetWorksheet = XLSX.utils.aoa_to_sheet([
    ['', '', ''],
    ['序号', '业务系统', 'IP'],
    [1, '核心系统A', '10.0.0.1'],
    [2, '系统B', '10.0.0.2,10.0.0.3'],
    [3, '系统C', '10.0.0.4']
  ]);

  const cells = dataFormatter.buildStatisticsCells({
    startDate: '2026-03-01',
    endDate: '2026-03-31',
    businessSystems: ['核心系统A', '系统B', '系统C'],
    assetWorksheet,
    vulnRows: [
      { IP: '10.0.0.1', '受影响主机/位置': '', '漏洞等级': '高危', '是否可利用': '是', '更新时间': '2026-03-05', '跟进状态': '已闭环' }
    ],
    weakPwdSummaryList: [
      { ip: '10.0.0.2' }
    ],
    weakPwdHandledList: [
      { ip: '10.0.0.2' }
    ],
    alarmRows: [
      { host_ip: '10.0.0.1' },
      { host_ip: '10.0.0.4' },
      { host_ip: '10.0.0.99' }
    ],
    eventRows: [
      {
        create_time: '2026-03-10',
        event_grading_tag: '一般威胁',
        push_status: '已通告',
        '响应时长': 30,
        host_ip: '10.0.0.1',
        affected_assets: []
      },
      {
        create_time: '2026-03-11',
        event_grading_tag: '重大事件',
        push_status: '已通告',
        '响应时长': 20,
        host_ip: '',
        affected_assets: ['10.0.0.3'],
        type: '未公开威胁',
        event_status: '已闭环'
      },
      {
        create_time: '2026-03-12',
        event_grading_tag: '一般事件',
        push_status: '未通告',
        '响应时长': 40,
        host_ip: '10.0.0.9',
        affected_assets: ['10.0.0.4'],
        event_status: '已闭环',
        manage_sub_type_cn: '账号安全异常'
      },
      {
        create_time: '2026-03-13',
        event_grading_tag: '其他',
        push_status: '未通告',
        '响应时长': 50,
        host_ip: '10.0.0.99',
        affected_assets: []
      }
    ]
  });

  assert.strictEqual(cells.G6, 1);
  assert.strictEqual(cells.H6, 0);
  assert.strictEqual(cells.I6, 1);
  assert.strictEqual(cells.G7, 30);
  assert.strictEqual(cells.G8, 2);
  assert.strictEqual(cells.G9, 20);
  assert.strictEqual(cells.G10, 2);
  assert.strictEqual(cells.G12, cells.D7 + cells.D8 + cells.D9);
  assert.strictEqual(cells.G13, 4);
  assert.strictEqual(cells.G14, 3);
  assert.strictEqual(cells.C18, 1);
  assert.strictEqual(cells.D18, 2);
  assert.strictEqual(cells.E18, 1);
}

function testStatisticsCellsUseUpdatedG16ToG18() {
  const assetWorksheet = XLSX.utils.aoa_to_sheet([
    ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['序号', '业务系统', 'IP', '', '服务类型', '资产类型', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '安全域'],
    [1, '系统A', '10.0.0.1', '', '服务内（7*24H）', '服务器', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '内网'],
    [2, '系统B', '10.0.0.2', '', '服务内（5*8H）', '终端', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '内网'],
    [3, '系统A', '10.0.0.3', '', '服务外', '服务器', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '外网'],
    [4, '系统C', '10.0.0.4', '', '服务内（7*24H）', '终端', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '内网']
  ]);

  const cells = dataFormatter.buildStatisticsCells({
    startDate: '2026-03-01',
    endDate: '2026-03-31',
    assetWorksheet
  });

  assert.strictEqual(cells.G16, 3);
  assert.strictEqual(cells.G17, 1);
  assert.strictEqual(cells.G18, 2);
}

function testBusinessSystemStatisticsStartsAtB3C3() {
  const assetWorksheet = XLSX.utils.aoa_to_sheet([
    ['', '核心系统A', '10.0.0.99'],
    ['', '核心系统A', '10.0.0.98'],
    [1, '核心系统A', '10.0.0.1']
  ]);

  const stats = dataFormatter.buildBusinessSystemStatistics({
    businessSystems: ['核心系统A'],
    assetWorksheet,
    range: {
      start: new Date('2026-03-01T00:00:00'),
      end: new Date('2026-03-31T23:59:59.999')
    },
    vulnRows: [{ IP: '10.0.0.1' }, { IP: '10.0.0.99' }],
    eventRows: [],
    weakPwdSummaryList: []
  });

  assert.deepStrictEqual(stats[0].ips, ['10.0.0.1']);
  assert.strictEqual(stats[0].total, 0);
}

function testBusinessSystemStatisticsRejectsMoreThanThree() {
  assert.throws(
    () => dataFormatter.normalizeBusinessSystems(['a', 'b', 'c', 'd']),
    /最多支持 3 个业务系统/
  );
}

function testPopulateStatisticsSheetWritesAndClearsBusinessSystemCells() {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([]);
  ['D3', 'D4', 'D5', 'D7', 'D8', 'D9'].forEach((address) => {
    ws[address] = { t: 's', v: '旧值' };
  });
  ws['!ref'] = 'A1:D9';
  XLSX.utils.book_append_sheet(wb, ws, '数据统计');
  const assetWorksheet = XLSX.utils.aoa_to_sheet([
    ['', '', ''],
    ['序号', '业务系统', 'IP'],
    [1, '核心系统A', '10.0.0.1']
  ]);

  dataFormatter.populateStatisticsSheet(wb, {
    startDate: '2026-03-01',
    endDate: '2026-03-31',
    businessSystems: ['核心系统A'],
    assetWorksheet,
    vulnRows: [],
    eventRows: [],
    weakPwdSummaryList: []
  });

  assert.strictEqual(ws.D3.v, '核心系统A');
  assert.strictEqual(ws.D7.v, 0);
  assert.strictEqual(ws.D4.v, '');
  assert.strictEqual(ws.D5.v, '');
  assert.strictEqual(ws.D8.v, '');
  assert.strictEqual(ws.D9.v, '');
}

function testStrategyOptimizeDeviceOfflineAlwaysMapsToBusinessContinuityRisk() {
  const result = soarTransformer.transformEventDocsWithStats([
    {
      event_grading_tag: 0,
      manage_type: 'STRATEGY_OPTIMIZE',
      manage_type_cn: '策略调优',
      manage_sub_type: 'DEVICE_OFFLINE',
      manage_sub_type_cn: '设备离线',
      event_name: '设备离线',
      create_time: '2026-06-01 10:00:00',
      event_status: 'finished',
      push_status: 1
    }
  ]);

  assert.strictEqual(result.rows.length, 1);
  assert.strictEqual(result.rows[0].event_grading_tag, '业务连续性风险');
}

function testStrategyOptimizeLogCheckExceptionAlwaysIgnored() {
  const result = soarTransformer.transformEventDocsWithStats([
    {
      event_grading_tag: 2,
      manage_type: 'STRATEGY_OPTIMIZE',
      manage_type_cn: '策略调优',
      manage_sub_type: 'LOG_CHECK_EXCEPTION',
      manage_sub_type_cn: '日志检测异常',
      event_name: '日志检测异常',
      create_time: '2026-06-01 10:00:00',
      event_status: 'finished',
      push_status: 1
    }
  ]);

  assert.strictEqual(result.rows.length, 0);
}

const tests = [
  testInclusiveDayCounts,
  testHolidayClipping,
  testNonOverlappingManualAndHoliday,
  testLeftOverlap,
  testRightOverlap,
  testManualContainsHoliday,
  testHolidayContainsManual,
  testOverlappingHolidaysAndThreeWayOverlap,
  testAdjacentRangesAreNotOverlapped,
  testNoProtectionRanges,
  testAtomicRangesMatchMergedRangesWhenNotOverlapped,
  testCountAggregation,
  testOverlappingHolidayAggregation,
  testAssetDownloadRequestUsesServiceStatusOneByDefault,
  testStatisticsCellsHolidayDailyAverage,
  testStatisticsCellsIntegerHolidayDailyAverage,
  testStatisticsCellsMissingHolidayCountLeavesAverageBlank,
  testStatisticsCellsManualOnlyLeavesHolidayAverageBlank,
  testStatisticsCellsUseUpdatedD6D10D11,
  testWeakPwdSummaryRequestMatchesCapturedTimeRange,
  testWeakPwdSummaryRequestSupportsDealStatusFilter,
  testWeakPwdSummaryRequestSupportsPagination,
  testPopulateStatisticsSheetWritesAndClearsHolidayAverages,
  testBusinessSystemStatisticsByAssetIps,
  testStatisticsCellsUseUpdatedD13ToD17,
  testStatisticsCellsUseXdrRejectedExternalToInternalForD23,
  testStatisticsCellsUseAlarmAttackDirectionForD24AndMirrorD14D15ToG23G24,
  testStatisticsCellsUseUpdatedG7ToG10AndG12,
  testStatisticsCellsUseUpdatedG16ToG18,
  testXdrRejectedExternalToInternalRequestMatchesCapturedFilters,
  testBusinessSystemStatisticsStartsAtB3C3,
  testBusinessSystemStatisticsRejectsMoreThanThree,
  testPopulateStatisticsSheetWritesAndClearsBusinessSystemCells,
  testStrategyOptimizeDeviceOfflineAlwaysMapsToBusinessContinuityRisk,
  testStrategyOptimizeLogCheckExceptionAlwaysIgnored
];

tests.forEach((test) => {
  test();
  console.log(`ok - ${test.name}`);
});
