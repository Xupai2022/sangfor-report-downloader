const assert = require('assert');
const XLSX = require('xlsx');
const dataFormatter = require('../data_formatter');
const apiClient = require('../api_client');

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

function testStatisticsCellsG11UsesD6D10D12() {
  const cells = dataFormatter.buildStatisticsCells({
    startDate: '2026-03-01',
    endDate: '2026-03-31',
    vulnRows: [
      { '更新时间': '2026-03-01' },
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
      }
    ],
    weakPwdSummaryTotal: 3,
    eventStats: {
      accountSecurityEventCountForE80: 4
    }
  });

  assert.strictEqual(cells.D6, 2);
  assert.strictEqual(cells.D10, 7);
  assert.strictEqual(cells.D12, 2);
  assert.strictEqual(cells.E80, 7);
  assert.strictEqual(cells.G11, 11);
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
  testStatisticsCellsHolidayDailyAverage,
  testStatisticsCellsIntegerHolidayDailyAverage,
  testStatisticsCellsMissingHolidayCountLeavesAverageBlank,
  testStatisticsCellsManualOnlyLeavesHolidayAverageBlank,
  testStatisticsCellsG11UsesD6D10D12,
  testPopulateStatisticsSheetWritesAndClearsHolidayAverages
];

tests.forEach((test) => {
  test();
  console.log(`ok - ${test.name}`);
});
