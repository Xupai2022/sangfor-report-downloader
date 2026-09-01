/**
 * 深信服报告下载 - 数据格式化与Excel生成模块
 * 将API返回的数据格式化为标准xlsx表格
 */

const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

/**
 * 推断字段类型
 * @param {*} value - 值
 * @returns {string} 类型: string, number, boolean, date, array, object
 */
function inferType(value) {
  if (value === null || value === undefined) return 'string';
  if (typeof value === 'boolean') return 'boolean';
  if (typeof value === 'number') return 'number';
  if (value instanceof Date) return 'date';
  if (Array.isArray(value)) return 'array';
  if (typeof value === 'object') return 'object';
  return 'string';
}

/**
 * 展平嵌套对象
 * @param {Object} obj - 原始对象
 * @param {string} prefix - 键名前缀
 * @returns {Object} 展平后的对象
 */
function flattenObject(obj, prefix = '') {
  const result = {};
  
  for (const key in obj) {
    if (!Object.prototype.hasOwnProperty.call(obj, key)) continue;
    
    const newKey = prefix ? `${prefix}.${key}` : key;
    const value = obj[key];
    
    if (value && typeof value === 'object' && !Array.isArray(value) && !(value instanceof Date)) {
      Object.assign(result, flattenObject(value, newKey));
    } else {
      result[newKey] = value;
    }
  }
  
  return result;
}

/**
 * 将数组中所有对象的字段统一为相同的列
 * @param {Array} dataArray - 数据数组
 * @returns {Array} 展平且字段统一的数据数组
 */
function normalizeDataFields(dataArray) {
  if (!Array.isArray(dataArray) || dataArray.length === 0) {
    return [];
  }
  
  // 收集所有字段
  const allFields = new Set();
  dataArray.forEach(item => {
    const flat = flattenObject(item);
    Object.keys(flat).forEach(key => allFields.add(key));
  });
  
  // 转换为有序数组
  const fieldOrder = Array.from(allFields).sort((a, b) => {
    // 常用字段排在前面
    const priorityFields = ['id', 'name', 'time', 'type', 'level', 'status', 'source', 'ip'];
    const aPriority = priorityFields.findIndex(f => a.toLowerCase().includes(f));
    const bPriority = priorityFields.findIndex(f => b.toLowerCase().includes(f));
    
    if (aPriority !== -1 && bPriority !== -1) return aPriority - bPriority;
    if (aPriority !== -1) return -1;
    if (bPriority !== -1) return 1;
    return a.localeCompare(b);
  });

// 标准化每一条数据
  return dataArray.map(item => {
    const flat = flattenObject(item);
    const normalized = {};
    fieldOrder.forEach(field => {
      normalized[field] = flat[field] !== undefined ? flat[field] : '';
    });
    return normalized;
  });
}

/**
 * 格式化单元格值用于显示
 * @param {*} value - 原始值
 * @param {string} type - 类型
 * @returns {*} 格式化后的值
 */
function formatCellValue(value, type) {
  if (value === null || value === undefined || value === '') {
    return '';
  }
  
  switch (type) {
    case 'boolean':
      return value ? '是' : '否';
    case 'date':
      if (value instanceof Date) {
        return value.toLocaleString('zh-CN');
      }
      return value;
    case 'array':
      if (!Array.isArray(value)) return String(value);
      return value.some(item => item && typeof item === 'object')
        ? JSON.stringify(value)
        : value.join(', ');
    case 'object':
      return JSON.stringify(value);
    default:
      return String(value);
  }
}

function isIPv4Text(value) {
  return /^(?:(?:25[0-5]|2[0-4]\d|1?\d?\d)\.){3}(?:25[0-5]|2[0-4]\d|1?\d?\d)$/.test(String(value || '').trim());
}

function countAffectedAssetIps(value) {
  if (!Array.isArray(value)) return 0;
  const uniqueIps = new Set();
  value.forEach((item) => {
    const text = String(item || '').trim();
    if (isIPv4Text(text)) uniqueIps.add(text);
  });
  return uniqueIps.size;
}

function shouldKeepStringValue(key, value) {
  if (typeof value !== 'string') return false;

  const lowerKey = key.toLowerCase();
  if (/(^|[._-])(id|uuid|no|code|ip|phone|mobile|tel)([._-]|$)/.test(lowerKey)) {
    return true;
  }

  if (/^\d{12,}$/.test(value.trim())) {
    return true;
  }

  return false;
}

/**
 * 智能格式化数据 - 尝试解析字符串形式的日期和数字
 * @param {Array} dataArray - 数据数组
 * @returns {Array} 格式化后的数据数组
 */
function smartFormatData(dataArray) {
  return dataArray.map(row => {
    const formatted = {};
    
    for (const key in row) {
      let value = row[key];
      
      // 尝试解析为数字
      if (typeof value === 'string' && !shouldKeepStringValue(key, value)) {
        const num = Number(value);
        if (!isNaN(num) && value.trim() !== '') {
          value = num;
        }
      }
      
      // 尝试识别日期字段
      const dateFields = ['time', 'date', 'timestamp', 'created', 'updated', 'modified', 'occurred'];
      const isDateField = dateFields.some(f => key.toLowerCase().includes(f));
      
      if (isDateField && value) {
        if (typeof value === 'number') {
          // 时间戳（毫秒）
          if (value > 1e12) {
            value = new Date(value).toLocaleString('zh-CN');
          } else {
            value = new Date(value * 1000).toLocaleString('zh-CN');
          }
        } else if (typeof value === 'string') {
          // ISO 日期字符串
          const date = new Date(value);
          if (!isNaN(date.getTime())) {
            value = date.toLocaleString('zh-CN');
          }
        }
      }

      formatted[key] = value;
    }
    
    return formatted;
  });
}
/**
 * 创建 Excel 工作表
 * @param {Array} data - 数据数组
 * @returns {Object} XLSX.WorkSheet
 */
function createWorksheet(data, preferredHeaders = null, headerDisplayMap = null) {
  // 规范化数据字段
  let normalizedData = normalizeDataFields(data);
  
  // 智能格式化
  normalizedData = smartFormatData(normalizedData);
  
  // 如果数据为空，返回空工作表
  if (normalizedData.length === 0) {
    return XLSX.utils.aoa_to_sheet([]);
  }
  
  // 提取表头
  let headers = Object.keys(normalizedData[0]);
  if (Array.isArray(preferredHeaders) && preferredHeaders.length > 0) {
    const dynamic = headers.filter(h => !preferredHeaders.includes(h));
    headers = [...preferredHeaders, ...dynamic];
  }
  
  // 构建工作表数据（第一行可使用展示名称）
  const displayHeaders = headers.map((header) => {
    if (headerDisplayMap && typeof headerDisplayMap === 'object' && headerDisplayMap[header]) {
      return headerDisplayMap[header];
    }
    return header;
  });
  const aoa = [displayHeaders];
  
  // 添加数据行
  normalizedData.forEach(row => {
    const rowData = headers.map(header => {
      const value = row[header];
      const type = inferType(value);
      return formatCellValue(value, type);
    });
    aoa.push(rowData);
  });
  
  // 创建工作表
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  
  // 设置列宽
  const colWidths = headers.map((header, i) => {
    // 计算该列最大宽度
    let maxLen = header.length;
    normalizedData.forEach(row => {
      const val = String(row[header] || '');
      if (val.length > maxLen) maxLen = val.length;
    });
    return { wch: Math.min(Math.max(maxLen, 10), 50) }; // 最小10，最大50
  });
  ws['!cols'] = colWidths;
  
  return ws;
}

/**
 * 创建 Excel 工作簿
 * @param {Array} data - 数据数组
 * @param {string} sheetName - 工作表名称
 * @returns {Object} XLSX.WorkBook
 */
function createWorkbook(data, sheetName = 'Sheet1', preferredHeaders = null, headerDisplayMap = null) {
  const ws = createWorksheet(data, preferredHeaders, headerDisplayMap);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  return wb;
}

/**
 * 生成报告文件名
 * @param {string} customer - 客户名称
 * @param {string} startDate - 开始日期
 * @param {string} endDate - 结束日期
 * @param {string} type - 报告类型 (event/alarm)
 * @returns {string} 文件名
 */
function generateFileName(customer, startDate, endDate, type, customerId = '') {
  const safeCustomer = customer || customerId || 'unknown';
  if (type === 'report') {
    return `${safeCustomer}_report.xlsx`;
  }
  return `${safeCustomer}_${type}.xlsx`;
}

/**
 * 保存数据为 xlsx 文件
 * @param {Array} data - 数据数组
 * @param {string} filePath - 文件路径
 * @param {string} sheetName - 工作表名称
 */
function saveToExcel(data, filePath, sheetName = 'Sheet1', preferredHeaders = null, headerDisplayMap = null) {
  const wb = createWorkbook(data, sheetName, preferredHeaders, headerDisplayMap);
  
  // 确保目录存在
  const dir = path.dirname(filePath);
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
  
  // 写入文件
  XLSX.writeFile(wb, filePath);
  console.log(`[DataFormatter] 文件已保存: ${filePath}`);
}

function ensureOutputDir(filePath) {
  const dir = path.dirname(filePath);
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
}

function emptyWorksheet() {
  return XLSX.utils.aoa_to_sheet([]);
}

function readWorkbookFromBuffer(buffer) {
  if (!buffer) return null;
  return XLSX.read(buffer, { type: Buffer.isBuffer(buffer) ? 'buffer' : 'array' });
}

function readWorkbookFromFile(filePath) {
  if (!filePath) return null;
  if (!fs.existsSync(filePath)) {
    throw new Error(`模板文件不存在: ${filePath}`);
  }
  return XLSX.readFile(filePath);
}

function getFirstWorksheet(wb) {
  if (!wb || !Array.isArray(wb.SheetNames) || wb.SheetNames.length === 0) {
    return null;
  }
  return wb.Sheets[wb.SheetNames[0]] || null;
}

function getWorksheetByNames(wb, preferredNames = []) {
  if (!wb || !Array.isArray(wb.SheetNames) || wb.SheetNames.length === 0) {
    return null;
  }

  for (const name of preferredNames) {
    if (wb.Sheets[name]) {
      return wb.Sheets[name];
    }
  }

  const statsSheetName = wb.SheetNames.find(name => String(name).includes('统计'));
  if (statsSheetName && wb.Sheets[statsSheetName]) {
    return wb.Sheets[statsSheetName];
  }

  return null;
}

function countWorksheetDataRows(ws) {
  if (!ws || !ws['!ref']) return 0;
  return XLSX.utils.sheet_to_json(ws, { defval: '' })
    .filter(row => Object.values(row).some(value => String(value).trim() !== ''))
    .length;
}

function appendReportSheet(wb, name, ws) {
  XLSX.utils.book_append_sheet(wb, ws || emptyWorksheet(), name);
}

function replaceWorksheet(wb, name, ws) {
  if (!wb || !wb.Sheets || !Array.isArray(wb.SheetNames)) {
    throw new Error('无效的模板工作簿');
  }

  if (!wb.SheetNames.includes(name)) {
    throw new Error(`模板缺少 Sheet: ${name}`);
  }

  wb.Sheets[name] = ws || emptyWorksheet();
}

function ensureWorksheetsExist(wb, names) {
  if (!wb || !Array.isArray(wb.SheetNames)) {
    throw new Error('无效的模板工作簿');
  }

  names.forEach((name) => {
    if (!wb.SheetNames.includes(name)) {
      throw new Error(`模板缺少 Sheet: ${name}`);
    }
  });
}

function formatReportDate(value) {
  if (!value) return '';
  const match = /^(\d{4})[-/](\d{2})[-/](\d{2})$/.exec(String(value).trim());
  if (match) {
    return `${match[1]}/${match[2]}/${match[3]}`;
  }

  const date = parseReportDateValue(value);
  if (!date) return String(value);

  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}/${month}/${day}`;
}

const HOLIDAY_PROTECTION_PERIODS = [
  { name: '元旦', startDate: '2023-12-30', endDate: '2024-01-01' },
  { name: '春节', startDate: '2024-02-10', endDate: '2024-02-17' },
  { name: '清明', startDate: '2024-04-04', endDate: '2024-04-06' },
  { name: '五一', startDate: '2024-05-01', endDate: '2024-05-05' },
  { name: '端午', startDate: '2024-06-08', endDate: '2024-06-10' },
  { name: '中秋', startDate: '2024-09-15', endDate: '2024-09-17' },
  { name: '国庆', startDate: '2024-10-01', endDate: '2024-10-07' },
  { name: '元旦', startDate: '2025-01-01', endDate: '2025-01-01' },
  { name: '春节', startDate: '2025-01-28', endDate: '2025-02-04' },
  { name: '清明', startDate: '2025-04-04', endDate: '2025-04-06' },
  { name: '五一', startDate: '2025-05-01', endDate: '2025-05-05' },
  { name: '端午', startDate: '2025-05-31', endDate: '2025-06-02' },
  { name: '国庆', startDate: '2025-10-01', endDate: '2025-10-07' },
  { name: '中秋', startDate: '2025-10-06', endDate: '2025-10-08' },
  { name: '元旦', startDate: '2026-01-01', endDate: '2026-01-03' },
  { name: '春节', startDate: '2026-02-15', endDate: '2026-02-23' },
  { name: '清明', startDate: '2026-04-04', endDate: '2026-04-06' },
  { name: '五一', startDate: '2026-05-01', endDate: '2026-05-05' },
  { name: '端午', startDate: '2026-06-19', endDate: '2026-06-21' },
  { name: '中秋', startDate: '2026-09-25', endDate: '2026-09-27' },
  { name: '国庆', startDate: '2026-10-01', endDate: '2026-10-07' }
];

function parseReportDateValue(value) {
  if (value === null || value === undefined || value === '') return null;

  if (value instanceof Date) {
    return Number.isNaN(value.getTime()) ? null : value;
  }

  if (typeof value === 'number') {
    if (value > 1e12) {
      const date = new Date(value);
      return Number.isNaN(date.getTime()) ? null : date;
    }

    if (value > 1e9) {
      const date = new Date(value * 1000);
      return Number.isNaN(date.getTime()) ? null : date;
    }

    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      return new Date(parsed.y, parsed.m - 1, parsed.d, parsed.H, parsed.M, parsed.S || 0);
    }

    return null;
  }

  if (typeof value !== 'string') return null;

  const text = value.trim();
  if (!text) return null;

  const directDate = new Date(text);
  if (!Number.isNaN(directDate.getTime())) return directDate;

  const normalizedText = text.replace(/\//g, '-');
  const normalizedDate = new Date(normalizedText);
  if (!Number.isNaN(normalizedDate.getTime())) return normalizedDate;

  return null;
}

function buildInclusiveDateRange(startDate, endDate) {
  const start = parseReportDateValue(startDate);
  const end = parseReportDateValue(endDate);
  if (!start || !end) {
    throw new Error(`统计时间范围无效: ${startDate} ~ ${endDate}`);
  }

  const rangeStart = new Date(start.getFullYear(), start.getMonth(), start.getDate(), 0, 0, 0, 0);
  const rangeEnd = new Date(end.getFullYear(), end.getMonth(), end.getDate(), 23, 59, 59, 999);
  if (rangeStart > rangeEnd) {
    throw new Error(`统计时间范围开始日期不能晚于结束日期: ${startDate} ~ ${endDate}`);
  }
  return { start: rangeStart, end: rangeEnd };
}

function addMonths(date, monthOffset) {
  return new Date(date.getFullYear(), date.getMonth() + monthOffset, 1);
}

function formatYearMonth(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  return `${year}-${month}`;
}

function buildMonthRange(date) {
  const start = new Date(date.getFullYear(), date.getMonth(), 1, 0, 0, 0, 0);
  const end = new Date(date.getFullYear(), date.getMonth() + 1, 0, 23, 59, 59, 999);
  return { start, end };
}

function isDateInRange(value, range) {
  const date = parseReportDateValue(value);
  if (!date || !range) return false;
  return date >= range.start && date <= range.end;
}

function rangesOverlap(left, right) {
  return left.start <= right.end && right.start <= left.end;
}

function rangeContains(outer, inner) {
  return outer.start <= inner.start && outer.end >= inner.end;
}

function formatReportDateRange(startDate, endDate) {
  if (!startDate || !endDate) return '';
  return `${formatReportDate(startDate)}-${formatReportDate(endDate)}`;
}

function formatIsoDate(date) {
  const parsed = parseReportDateValue(date);
  if (!parsed) return '';
  const year = parsed.getFullYear();
  const month = String(parsed.getMonth() + 1).padStart(2, '0');
  const day = String(parsed.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function buildProtectionPeriods(statsContext, reportRange) {
  const periods = [];
  const protectStartDate = statsContext.protectStartDate
    || (statsContext.protectionPeriod && statsContext.protectionPeriod.startDate)
    || '';
  const protectEndDate = statsContext.protectEndDate
    || (statsContext.protectionPeriod && statsContext.protectionPeriod.endDate)
    || '';

  if (protectStartDate && protectEndDate) {
    periods.push({
      source: 'manual',
      name: '护网时间',
      startDate: protectStartDate,
      endDate: protectEndDate,
      range: buildInclusiveDateRange(protectStartDate, protectEndDate)
    });
  }

  HOLIDAY_PROTECTION_PERIODS.forEach((holiday) => {
    const holidayRange = buildInclusiveDateRange(holiday.startDate, holiday.endDate);
    if (!rangesOverlap(holidayRange, reportRange)) return;
    periods.push(Object.assign({}, holiday, {
      source: 'holiday',
      range: holidayRange
    }));
  });

  return periods;
}

function clipProtectionPeriodsToReportRange(statsContext, reportRange) {
  return buildProtectionPeriods(statsContext, reportRange)
    .map((period) => {
      if (!period.range) return null;
      const startMs = Math.max(period.range.start.getTime(), reportRange.start.getTime());
      const endMs = Math.min(period.range.end.getTime(), reportRange.end.getTime());
      if (startMs > endMs) return null;
      const range = {
        start: new Date(startMs),
        end: new Date(endMs)
      };
      const originalStartDate = period.startDate;
      const originalEndDate = period.endDate;
      const startDate = formatIsoDate(range.start);
      const endDate = formatIsoDate(range.end);
      return {
        ...period,
        originalStartDate,
        originalEndDate,
        startDate,
        endDate,
        key: buildProtectionPeriodKey(period, originalStartDate, originalEndDate),
        dayCount: countInclusiveNaturalDays(range.start, range.end),
        range
      };
    })
    .filter(Boolean);
}

function buildProtectionPeriodKey(period, startDate, endDate) {
  const source = period && period.source ? period.source : 'unknown';
  const name = period && period.name ? period.name : '';
  return `${source}:${name}:${formatIsoDate(startDate)}~${formatIsoDate(endDate)}`;
}

function startOfNaturalDay(date) {
  const parsed = parseReportDateValue(date);
  if (!parsed) return null;
  return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate(), 0, 0, 0, 0);
}

function countInclusiveNaturalDays(startDate, endDate) {
  const startDay = startOfNaturalDay(startDate);
  const endDay = startOfNaturalDay(endDate);
  if (!startDay || !endDay || startDay > endDay) return 0;
  return Math.round((endDay.getTime() - startDay.getTime()) / 86400000) + 1;
}

function mergeInclusiveRanges(periods) {
  const ranges = periods
    .map(period => period.range)
    .filter(Boolean)
    .map(range => ({ startMs: range.start.getTime(), endMs: range.end.getTime() }))
    .sort((a, b) => a.startMs - b.startMs || a.endMs - b.endMs);

  const merged = [];
  ranges.forEach((range) => {
    const last = merged[merged.length - 1];
    if (!last || range.startMs > last.endMs + 1) {
      merged.push(Object.assign({}, range));
      return;
    }
    if (range.endMs > last.endMs) {
      last.endMs = range.endMs;
    }
  });
  return merged;
}

function buildProtectionSecondRanges(statsContext) {
  const reportRange = buildInclusiveDateRange(statsContext.startDate, statsContext.endDate);
  const clippedPeriods = clipProtectionPeriodsToReportRange(statsContext, reportRange);

  return mergeInclusiveRanges(clippedPeriods).map(range => ({
    start: Math.floor(range.startMs / 1000),
    end: Math.floor(range.endMs / 1000)
  }));
}

function buildProtectionAtomicSecondRanges(statsContext) {
  const reportRange = buildInclusiveDateRange(statsContext.startDate, statsContext.endDate);
  const periods = clipProtectionPeriodsToReportRange(statsContext, reportRange);
  const holidayPeriods = periods.filter(period => period.source === 'holiday');

  if (periods.length === 0) {
    return {
      ranges: [],
      holidayPeriods
    };
  }

  const boundaries = new Set();
  periods.forEach((period) => {
    boundaries.add(period.range.start.getTime());
    boundaries.add(period.range.end.getTime() + 1);
  });

  const sortedBoundaries = Array.from(boundaries).sort((a, b) => a - b);
  const ranges = [];
  for (let index = 0; index < sortedBoundaries.length - 1; index += 1) {
    const startMs = sortedBoundaries[index];
    const endMs = sortedBoundaries[index + 1] - 1;
    if (startMs > endMs) continue;

    const owners = periods
      .filter(period => period.range.start.getTime() <= startMs && period.range.end.getTime() >= endMs)
      .map(period => ({
        source: period.source,
        name: period.name,
        key: period.key
      }));

    if (owners.length === 0) continue;

    ranges.push({
      start: Math.floor(startMs / 1000),
      end: Math.floor(endMs / 1000),
      startDate: formatIsoDate(new Date(startMs)),
      endDate: formatIsoDate(new Date(endMs)),
      owners
    });
  }

  return {
    ranges,
    holidayPeriods
  };
}

function buildReportSecondRange(statsContext) {
  const reportRange = buildInclusiveDateRange(statsContext.startDate, statsContext.endDate);
  return {
    start: Math.floor(reportRange.start.getTime() / 1000),
    end: Math.floor(reportRange.end.getTime() / 1000)
  };
}

function buildReportMonthSecondRanges(statsContext, maxMonths = 12) {
  const reportRange = buildInclusiveDateRange(statsContext.startDate, statsContext.endDate);
  const ranges = [];
  let cursor = new Date(reportRange.start.getFullYear(), reportRange.start.getMonth(), 1, 0, 0, 0, 0);

  while (
    cursor <= reportRange.end
    && ranges.length < maxMonths
  ) {
    const year = cursor.getFullYear();
    const month = cursor.getMonth();
    const monthStart = new Date(year, month, 1, 0, 0, 0, 0);
    const monthEnd = new Date(year, month + 1, 0, 23, 59, 59, 999);
    const startMs = Math.max(monthStart.getTime(), reportRange.start.getTime());
    const endMs = Math.min(monthEnd.getTime(), reportRange.end.getTime());

    if (startMs <= endMs) {
      ranges.push({
        monthLabel: formatYearMonth(monthStart),
        start: Math.floor(startMs / 1000),
        end: Math.floor(endMs / 1000)
      });
    }

    cursor = new Date(year, month + 1, 1, 0, 0, 0, 0);
  }

  return ranges;
}

function isTimestampInMergedRanges(timestamp, ranges) {
  if (!Number.isFinite(timestamp) || !Array.isArray(ranges) || ranges.length === 0) return false;

  let low = 0;
  let high = ranges.length - 1;
  while (low <= high) {
    const mid = Math.floor((low + high) / 2);
    const range = ranges[mid];
    if (timestamp < range.startMs) {
      high = mid - 1;
    } else if (timestamp > range.endMs) {
      low = mid + 1;
    } else {
      return true;
    }
  }
  return false;
}

function getFirstFieldValue(row, fieldNames) {
  if (!row || !Array.isArray(fieldNames)) return '';
  for (const fieldName of fieldNames) {
    if (row[fieldName] !== undefined && row[fieldName] !== null && row[fieldName] !== '') {
      return row[fieldName];
    }
  }
  return '';
}

function countRowsInProtectionRanges(rows, fieldNames, mergedRanges) {
  if (!Array.isArray(rows) || rows.length === 0 || mergedRanges.length === 0) return 0;

  let count = 0;
  rows.forEach((row) => {
    const date = parseReportDateValue(getFirstFieldValue(row, fieldNames));
    if (date && isTimestampInMergedRanges(date.getTime(), mergedRanges)) {
      count += 1;
    }
  });
  return count;
}

function toText(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function getEventManageSubTypeName(row) {
  if (!row || typeof row !== 'object') return '';
  return toText(
    row.__manageSubTypeDisplay
    || row.manage_sub_type_cn
    || row.manage_sub_type_name
    || row['事件二级类型']
    || row['二级类型']
  );
}

function countEventRowsByManageSubTypeKeywords(eventRows, keywords) {
  if (!Array.isArray(eventRows) || !Array.isArray(keywords) || keywords.length === 0) return 0;
  const normalizedKeywords = keywords.map(keyword => toText(keyword)).filter(Boolean);
  if (normalizedKeywords.length === 0) return 0;

  return eventRows.filter((row) => {
    const name = getEventManageSubTypeName(row);
    return name && normalizedKeywords.some(keyword => name.includes(keyword));
  }).length;
}

function countEventRowsByManageSubTypeKeywordsAndIps(eventRows, keywords, ipSet) {
  if (!Array.isArray(eventRows) || !ipSet || ipSet.size === 0) return 0;
  const normalizedKeywords = Array.isArray(keywords)
    ? keywords.map(keyword => toText(keyword)).filter(Boolean)
    : [];
  if (normalizedKeywords.length === 0) return 0;

  return eventRows.filter((row) => {
    const name = getEventManageSubTypeName(row);
    if (!name || !normalizedKeywords.some(keyword => name.includes(keyword))) return false;
    return rowContainsAnyIp(row, ['host_ip', 'affected_assets'], ipSet);
  }).length;
}

function countMatchedAffectedAssetIps(value, ipSet) {
  if (!ipSet || ipSet.size === 0) return 0;
  return extractIPv4Texts(value).filter(ip => ipSet.has(ip)).length;
}

function countWeakPwdRowsContainingIps(weakPwdSummaryList, ipSet) {
  if (!Array.isArray(weakPwdSummaryList) || !ipSet || ipSet.size === 0) return 0;
  return weakPwdSummaryList.filter((row) => {
    const ips = extractIPv4Texts(row && row.ip);
    return ips.some(ip => ipSet.has(ip));
  }).length;
}

function buildTopEventManageSubTypeStats(eventRows, limit = 5) {
  const counts = new Map();
  (Array.isArray(eventRows) ? eventRows : []).forEach((row) => {
    const name = getEventManageSubTypeName(row);
    if (!name) return;
    const current = counts.get(name);
    if (current) {
      current.count += 1;
    } else {
      counts.set(name, { name, count: 1, firstIndex: counts.size });
    }
  });

  return Array.from(counts.values())
    .sort((a, b) => {
      if (b.count !== a.count) return b.count - a.count;
      return a.firstIndex - b.firstIndex;
    })
    .map(({ name, count }) => ({ name, count }))
    .slice(0, limit);
}

function buildTopnReportCells(topnReportStats) {
  const cells = {};
  const srcIpGeos = Array.isArray((topnReportStats || {}).srcIpGeos)
    ? topnReportStats.srcIpGeos
    : [];
  const threatTypes = Array.isArray((topnReportStats || {}).threatTypes)
    ? topnReportStats.threatTypes
    : [];
  const dstIps = Array.isArray((topnReportStats || {}).dstIps)
    ? topnReportStats.dstIps
    : [];

  srcIpGeos.slice(0, 5).forEach((item, index) => {
    const row = 62 + index;
    const name = toText((item || {}).src_ip_geo);
    if (name) cells[`F${row}`] = name;
    if ((item || {}).attack_count !== undefined && (item || {}).attack_count !== null) {
      const count = toNumericOrNull(item.attack_count);
      cells[`G${row}`] = count === null ? item.attack_count : count;
    }
  });

  threatTypes.slice(0, 5).forEach((item, index) => {
    const row = 62 + index;
    const name = toText((item || {}).threat_type);
    if (name) cells[`I${row}`] = name;
    if ((item || {}).attack_count !== undefined && (item || {}).attack_count !== null) {
      const count = toNumericOrNull(item.attack_count);
      cells[`J${row}`] = count === null ? item.attack_count : count;
    }
  });

  dstIps.slice(0, 5).forEach((item, index) => {
    const row = 62 + index;
    const name = toText((item || {}).dst_ip);
    if (name) cells[`L${row}`] = name;
    if ((item || {}).attack_count !== undefined && (item || {}).attack_count !== null) {
      const count = toNumericOrNull(item.attack_count);
      cells[`M${row}`] = count === null ? item.attack_count : count;
    }
  });

  return cells;
}

function toNumericOrNull(value) {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : null;
  }

  const normalized = String(value).trim();
  if (!normalized) return null;

  const number = Number(normalized);
  return Number.isFinite(number) ? number : null;
}

function roundTo(value, digits = 2) {
  const factor = 10 ** digits;
  return Math.round(value * factor) / factor;
}

function formatRatioAsPercentage(value, digits = 2) {
  if (value === null || value === undefined || !Number.isFinite(value)) return '';
  const percentage = roundTo(value * 100, digits);
  return `${Number(percentage.toFixed(digits))}%`;
}

function averageNumbers(values) {
  if (!Array.isArray(values) || values.length === 0) return 0;
  const total = values.reduce((sum, item) => sum + item, 0);
  return roundTo(total / values.length, 2);
}

function isThreatEventType(eventType) {
  const text = toText(eventType);
  if (!text) return false;
  return ['重大威胁', '一般威胁'].includes(text) || text.includes('威胁');
}

function isEventCategoryType(eventType) {
  const text = toText(eventType);
  if (!text) return false;
  return ['重大事件', '重要事件', '一般事件'].includes(text) || text.includes('事件');
}

function isStatisticsEventCountType(eventType) {
  const text = toText(eventType);
  if (!text) return false;
  return text.includes('事件') || ['重大威胁', '一般威胁'].includes(text);
}

function getWorksheetCell(ws, address) {
  if (!ws[address]) {
    ws[address] = { t: 's', v: '' };
  }
  return ws[address];
}

function ensureWorksheetRefIncludesCell(ws, address) {
  if (!ws || !address) return;
  const cellRef = XLSX.utils.decode_cell(address);
  let range;
  if (ws['!ref']) {
    range = XLSX.utils.decode_range(ws['!ref']);
  } else {
    range = { s: Object.assign({}, cellRef), e: Object.assign({}, cellRef) };
  }

  range.s.r = Math.min(range.s.r, cellRef.r);
  range.s.c = Math.min(range.s.c, cellRef.c);
  range.e.r = Math.max(range.e.r, cellRef.r);
  range.e.c = Math.max(range.e.c, cellRef.c);
  ws['!ref'] = XLSX.utils.encode_range(range);
}

function getWorksheetCellValue(ws, address) {
  if (!ws || !ws[address]) return '';
  const cell = ws[address];
  return cell.v === null || cell.v === undefined ? '' : cell.v;
}

function getWorksheetRows(ws) {
  if (!ws || !ws['!ref']) return [];
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
}

function findHeaderRowAndIndexes(rows, headerNames) {
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] || [];
    const indexes = {};
    let matchedCount = 0;
    for (const headerName of headerNames) {
      const colIndex = row.findIndex(value => toText(value) === headerName);
      indexes[headerName] = colIndex;
      if (colIndex >= 0) matchedCount += 1;
    }
    if (matchedCount > 0) {
      return { rowIndex, indexes };
    }
  }
  return { rowIndex: -1, indexes: {} };
}

function countAssetRowsByCriteria(assetWorksheet, criteria) {
  const rows = getWorksheetRows(assetWorksheet);
  if (rows.length === 0) return 0;

  const headerNames = ['服务类型', '资产类型', '安全域'];
  const { rowIndex, indexes } = findHeaderRowAndIndexes(rows, headerNames);
  const startRow = rowIndex >= 0 ? rowIndex + 1 : 0;
  const serviceTypeIndex = indexes['服务类型'] >= 0 ? indexes['服务类型'] : 4;
  const assetTypeIndex = indexes['资产类型'] >= 0 ? indexes['资产类型'] : 5;
  const securityDomainIndex = indexes['安全域'] >= 0 ? indexes['安全域'] : 32;

  return rows.slice(startRow).filter((row) => {
    if (!row || row.every(value => toText(value) === '')) return false;
    if (criteria.assetType !== undefined && toText(row[assetTypeIndex]) !== criteria.assetType) return false;
    if (criteria.securityDomain !== undefined && toText(row[securityDomainIndex]) !== criteria.securityDomain) return false;
    if (criteria.serviceType !== undefined && !toText(row[serviceTypeIndex]).includes(criteria.serviceType)) return false;
    return true;
  }).length;
}

function countDistinctBusinessSystems(assetWorksheet) {
  const rows = getWorksheetRows(assetWorksheet);
  if (rows.length === 0) return 0;

  const { rowIndex, indexes } = findHeaderRowAndIndexes(rows, ['业务系统']);
  const startRow = rowIndex >= 0 ? rowIndex + 1 : 2;
  const businessSystemIndex = indexes['业务系统'] >= 0 ? indexes['业务系统'] : 1;
  const names = new Set();

  rows.slice(startRow).forEach((row) => {
    const value = toText(row && row[businessSystemIndex]);
    if (!value) return;
    names.add(value);
  });

  return names.size;
}

function normalizeBusinessSystems(value) {
  if (value === undefined || value === null || value === '') return [];

  const rawItems = Array.isArray(value) ? value : [value];
  const normalized = [];
  for (const rawItem of rawItems) {
    String(rawItem || '')
      .split(/[，、,；;\r\n]+/)
      .map(item => toText(item))
      .filter(Boolean)
      .forEach(item => normalized.push(item));
  }

  if (normalized.length > 3) {
    throw new Error('最多支持 3 个业务系统。请精简后重试。');
  }

  return normalized;
}

function extractIPv4Texts(value) {
  const ips = new Set();
  const visit = (item) => {
    if (item === null || item === undefined) return;
    if (Array.isArray(item)) {
      item.forEach(visit);
      return;
    }
    if (typeof item === 'object') {
      Object.values(item).forEach(visit);
      return;
    }

    const matches = String(item).match(/\b(?:(?:25[0-5]|2[0-4]\d|1?\d?\d)\.){3}(?:25[0-5]|2[0-4]\d|1?\d?\d)\b/g);
    if (matches) {
      matches.forEach(ip => ips.add(ip));
    }
  };

  visit(value);
  return Array.from(ips);
}

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

function rowContainsAnyIp(row, fieldNames, ipSet) {
  if (!row || !ipSet || ipSet.size === 0) return false;
  for (const fieldName of fieldNames) {
    const ips = extractIPv4Texts(row[fieldName]);
    if (ips.some(ip => ipSet.has(ip))) return true;
  }
  return false;
}

function countRowsContainingIps(rows, fieldNames, ipSet) {
  if (!Array.isArray(rows) || !ipSet || ipSet.size === 0) return 0;
  return rows.filter(row => rowContainsAnyIp(row, fieldNames, ipSet)).length;
}

function buildBusinessSystemStatistics(statsContext) {
  const systems = collectBusinessSystemIpSets(
    statsContext && statsContext.assetWorksheet,
    statsContext && statsContext.businessSystems
  );
  const eventRows = statsContext && Array.isArray(statsContext.eventRows) ? statsContext.eventRows : [];
  const vulnRows = statsContext && Array.isArray(statsContext.vulnRows) ? statsContext.vulnRows : [];
  const weakPwdSummaryList = statsContext && Array.isArray(statsContext.weakPwdSummaryList)
    ? statsContext.weakPwdSummaryList
    : [];
  const weakPwdHandledList = statsContext && Array.isArray(statsContext.weakPwdHandledList)
    ? statsContext.weakPwdHandledList
    : [];
  const range = statsContext && statsContext.range;
  const keywordList = ['弱口令', '弱密码', '账号安全'];

  return systems.map((system) => {
    const d10Count = vulnRows.filter((row) => (
      toText(row['漏洞等级']) === '高危'
      && toText(row['是否可利用']) === '是'
      && rowContainsAnyIp(row, ['IP', '受影响主机/位置'], system.ips)
    )).length;
    const d11WeakPwdCount = countWeakPwdRowsContainingIps(weakPwdSummaryList, system.ips);
    const d11EventCount = countEventRowsByManageSubTypeKeywordsAndIps(eventRows, keywordList, system.ips);
    const d12Count = eventRows
      .filter((row) => (
        isDateInRange(row.create_time, range)
        && toText(row.type).includes('未公开威胁')
      ))
      .reduce((total, row) => total + countMatchedAffectedAssetIps(row.affected_assets, system.ips), 0);
    const closedLoopD10Count = vulnRows.filter((row) => (
      toText(row['漏洞等级']) === '高危'
      && toText(row['是否可利用']) === '是'
      && toText(row['跟进状态']) === '已闭环'
      && rowContainsAnyIp(row, ['IP', '受影响主机/位置'], system.ips)
    )).length;
    const closedLoopD11WeakPwdCount = countWeakPwdRowsContainingIps(weakPwdHandledList, system.ips);
    const closedLoopD11EventCount = eventRows.filter((row) => (
      isDateInRange(row.create_time, range)
      && toText(row.event_status) === '已闭环'
      && keywordList.some(keyword => getEventManageSubTypeName(row).includes(keyword))
      && rowContainsAnyIp(row, ['host_ip', 'affected_assets'], system.ips)
    )).length;
    const closedLoopD12Count = eventRows
      .filter((row) => (
        isDateInRange(row.create_time, range)
        && toText(row.event_status) === '已闭环'
        && toText(row.type).includes('未公开威胁')
      ))
      .reduce((total, row) => total + countMatchedAffectedAssetIps(row.affected_assets, system.ips), 0);

    return {
      name: system.name,
      ips: Array.from(system.ips),
      total: d10Count + d11WeakPwdCount + d11EventCount + d12Count,
      closedLoopTotal: closedLoopD10Count + closedLoopD11WeakPwdCount + closedLoopD11EventCount + closedLoopD12Count,
      d10Count,
      d11WeakPwdCount,
      d11EventCount,
      d12Count,
      closedLoopD10Count,
      closedLoopD11WeakPwdCount,
      closedLoopD11EventCount,
      closedLoopD12Count
    };
  });
}

function sumWeakPwdTicketTotalsForIps(totalsByIp, ipSet) {
  if (!totalsByIp || typeof totalsByIp !== 'object' || !ipSet || ipSet.size === 0) return 0;
  return Object.entries(totalsByIp).reduce((total, [ip, count]) => (
    ipSet.has(ip) ? total + (toNumericOrNull(count) || 0) : total
  ), 0);
}

function buildCoreSystemTicketStatistics({
  ipSet,
  vulnRows,
  eventRows,
  weakPwdAllTotalsByIp,
  weakPwdHandledTotalsByIp
}) {
  const totalVulnCount = vulnRows.filter(row => (
    rowContainsAnyIp(row, ['IP', '受影响主机/位置'], ipSet)
  )).length;
  const handledVulnCount = vulnRows.filter(row => (
    ['已防护', '已修复'].includes(toText(row['跟进状态']))
    && rowContainsAnyIp(row, ['IP', '受影响主机/位置'], ipSet)
  )).length;
  const totalLatestThreatCount = eventRows.filter(row => (
    toText(row.event_grading_tag) === '最新威胁'
    && rowContainsAnyIp(row, ['host_ip', 'affected_assets'], ipSet)
  )).length;
  const handledLatestThreatCount = eventRows.filter(row => (
    toText(row.event_grading_tag) === '最新威胁'
    && toText(row.event_status) === '已闭环'
    && rowContainsAnyIp(row, ['host_ip', 'affected_assets'], ipSet)
  )).length;
  const totalWeakPwdCount = sumWeakPwdTicketTotalsForIps(weakPwdAllTotalsByIp, ipSet);
  const handledWeakPwdCount = sumWeakPwdTicketTotalsForIps(weakPwdHandledTotalsByIp, ipSet);
  const total = totalVulnCount
    + totalWeakPwdCount
    + totalLatestThreatCount;
  const handledTotal = handledVulnCount
    + handledWeakPwdCount
    + handledLatestThreatCount;

  return {
    total,
    handledTotal,
    totalVulnCount,
    totalWeakPwdCount,
    totalLatestThreatCount,
    ratio: total > 0 ? handledTotal / total : 0
  };
}

function setWorksheetCellValue(ws, address, value) {
  const cell = getWorksheetCell(ws, address);
  ensureWorksheetRefIncludesCell(ws, address);

  if (value === '' || value === null || value === undefined) {
    cell.t = 's';
    cell.v = '';
    delete cell.w;
    delete cell.f;
    return;
  }

  if (typeof value === 'number') {
    cell.t = 'n';
    cell.v = value;
  } else if (typeof value === 'boolean') {
    cell.t = 'b';
    cell.v = value;
  } else {
    cell.t = 's';
    cell.v = String(value);
  }

  delete cell.w;
  delete cell.f;
}

function setWorksheetBlankPercentCell(ws, address) {
  const cell = getWorksheetCell(ws, address);
  cell.t = 's';
  cell.v = '';
  cell.z = '0%';
  delete cell.w;
  delete cell.f;
}

function setWorksheetPercentCellValue(ws, address, value) {
  setWorksheetCellValue(ws, address, value);
  ws[address].z = '0.00%';
}

function buildThreatDefineStatisticsCells(items, {
  labelRow,
  countRow = null,
  ratioRow,
  startColumn = 2,
  maxItems = Number.POSITIVE_INFINITY
}) {
  const rows = Array.isArray(items) ? items : [];
  const total = rows.reduce((sum, item) => {
    const count = Number(item && item.count);
    return Number.isFinite(count) && count >= 0 ? sum + count : sum;
  }, 0);
  const cells = {};
  const percentCells = [];

  rows.slice(0, maxItems).forEach((item, index) => {
    const count = Number(item && item.count);
    if (!Number.isFinite(count) || count < 0) return;
    const col = XLSX.utils.encode_col(startColumn + index);
    cells[`${col}${labelRow}`] = item && item.label ? String(item.label) : '';
    if (countRow !== null) {
      cells[`${col}${countRow}`] = count;
    }
    cells[`${col}${ratioRow}`] = total > 0 ? count / total : 0;
    percentCells.push(`${col}${ratioRow}`);
  });

  return { cells, percentCells };
}

function buildStatisticsCells(statsContext) {
  const {
    customer,
    startDate,
    endDate,
    protectStartDate,
    protectEndDate,
    eventRows = [],
    eventStats = {},
    alarmRows = [],
    vulnRows = [],
    businessSystems = [],
    assetWorksheet = null,
    exposedSurfaceWorksheet = null,
    exposedSurfacePortCount = '',
    weakPwdSummaryTotal = 0,
    weakPwdSummaryList = [],
    weakPwdAllSummaryList = [],
    weakPwdAllTotalsByIp = {},
    weakPwdAllTotal = 0,
    weakPwdHandledTotal = 0,
    weakPwdHandledAdminTotal = 0,
    weakPwdHandledList = [],
    weakPwdHandledTotalsByIp = {},
    g126 = 0,
    d129 = '',
    d130 = '',
    topnReportStats = null,
    xdrIn2outLogSearchCount = '',
    xdrOut2inLogSearchCount = '',
    xdrRejectedExternalToInternalCount = '',
    xdrG99LogSearchCount = '',
    xdrG100AlertCount = '',
    xdrG101IncidentCount = '',
    xdrG102IncidentHandledCount = '',
    xdrG103IncidentCount = '',
    xdrG105IncidentCount = '',
    xdrG106IncidentCount = '',
    xdrG107IncidentCount = '',
    xdrAlertThreatDefineCounts = [],
    xdrIncidentThreatDefineCounts = [],
    xdrMonthlyIn2outLogSearchCounts = [],
    xdrMonthlyOut2inLogSearchCounts = [],
    xdrLogSearchCount = '',
    xdrHolidayLogSearchCounts = {},
    orderBranchCounts = {}
  } = statsContext || {};

  const range = buildInclusiveDateRange(startDate, endDate);
  const protectionPeriods = clipProtectionPeriodsToReportRange(
    Object.assign({}, statsContext, { protectStartDate, protectEndDate }),
    range
  );
  const holidayPeriods = protectionPeriods.filter(period => period.source === 'holiday');
  // D90/D91 and their source metrics only cover the built-in holiday protection periods.
  const mergedHolidayProtectionRanges = mergeInclusiveRanges(holidayPeriods);
  const importantProtectionCount = holidayPeriods.length;
  const startMonth = addMonths(range.start, 0);
  const displayMonthCount = Math.min(
    12,
    ((range.end.getFullYear() - startMonth.getFullYear()) * 12)
      + (range.end.getMonth() - startMonth.getMonth())
      + 1
  );

  const legacyD6 = vulnRows.length;
  const d11 = eventRows.filter((row) => (
    toText(row.event_grading_tag) === '最新威胁'
    && isDateInRange(row.create_time, range)
  )).length;
  const d12 = eventRows
    .filter(row => (
      isDateInRange(row.create_time, range)
      && toText(row.type).includes('未公开威胁')
    ))
    .reduce((total, row) => total + countAffectedAssetIps(row.affected_assets), 0);
  const accountSecurityKeywords = ['弱口令', '弱密码', '账号安全'];
  const accountSecurityKeywordEventCount = countEventRowsByManageSubTypeKeywords(
    eventRows,
    accountSecurityKeywords
  );
  const d113AssetIps = new Set();
  weakPwdSummaryList.forEach((row) => {
    extractIPv4Texts(row && row.ip).forEach(ip => d113AssetIps.add(ip));
  });
  eventRows.forEach((row) => {
    const manageSubType = getEventManageSubTypeName(row);
    if (!manageSubType || !accountSecurityKeywords.some(keyword => manageSubType.includes(keyword))) return;
    extractIPv4Texts(row && row.affected_assets).forEach(ip => d113AssetIps.add(ip));
  });
  const d113 = d113AssetIps.size;
  const d114 = eventRows.filter(row => toText(row.type).includes('未公开威胁')).length;
  const d115 = d12;
  const g110 = d115;
  const accountSecurityClosedEventCount = eventRows.filter((row) => (
    isDateInRange(row.create_time, range)
    && toText(row.event_status) === '已闭环'
    && ['弱口令', '弱密码', '账号安全'].some(keyword => getEventManageSubTypeName(row).includes(keyword))
  )).length;
  const d14 = vulnRows.filter((row) => (
    toText(row['跟进状态']) === '已防护'
  )).length;
  const d15 = vulnRows.filter((row) => (
    toText(row['跟进状态']) === '已修复'
  )).length;
  const d116 = d14;
  const d117 = d15;
  const d16 = (toNumericOrNull(weakPwdHandledAdminTotal) || 0) + accountSecurityClosedEventCount;
  const d118 = d16;
  const d17 = d12;
  const d13 = d17 + d14 + d15;
  const d24 = alarmRows.filter((row) => (
    toText(row.attack_direction) === '内-外'
  )).length;
  const g22 = d14 + d15;
  const d21 = (toNumericOrNull(xdrRejectedExternalToInternalCount) || 0) + d24;
  const d20 = d21 + g22;
  const highRiskVulnRows = vulnRows.filter(row => toText(row['漏洞等级']) === '高危');
  const highRiskProtectedCount = highRiskVulnRows.filter(row => toText(row['跟进状态']) === '已防护').length;
  const d34 = highRiskVulnRows.length > 0
    ? formatRatioAsPercentage(highRiskProtectedCount / highRiskVulnRows.length, 2)
    : '0%';
  const d35Numbers = eventRows
    .filter((row) => (
      isEventCategoryType(row.event_grading_tag)
      && toText(row.push_status) === '已通告'
    ))
    .map(row => toNumericOrNull(row['识别时长']))
    .filter(value => value !== null);
  const d35 = averageNumbers(d35Numbers);
  const d36Numbers = eventRows
    .filter((row) => (
      isEventCategoryType(row.event_grading_tag)
      && toText(row.push_status) === '已通告'
    ))
    .map(row => toNumericOrNull(row['响应时长']))
    .filter(value => value !== null);
  const d36 = averageNumbers(d36Numbers);
  const d37 = vulnRows.filter(row => toText(row['跟进状态']).includes('已')).length;
  const d38 = eventRows.filter(row => isDateInRange(row.create_time, range)).length;
  const c80 = vulnRows.filter((row) => (
    toText(row['漏洞等级']) === '高危'
    && toText(row['是否可利用']) === '是'
  )).length;
  const d80 = vulnRows.filter((row) => (
    toText(row['内/外网']) === '外网'
    && toText(row['跟进状态']).includes('已')
  )).length;
  const highRiskExploitableRows = vulnRows.filter((row) => (
    toText(row['漏洞等级']) === '高危'
    && toText(row['是否可利用']) === '是'
  ));
  const highRiskExploitableHandledCount = highRiskExploitableRows
    .filter(row => toText(row['跟进状态']).includes('已'))
    .length;
  const f80 = highRiskExploitableRows.length > 0
    ? formatRatioAsPercentage(highRiskExploitableHandledCount / highRiskExploitableRows.length, 2)
    : '0%';
  const countVulnByLevel = (level) => vulnRows
    .filter(row => toText(row['漏洞等级']) === level)
    .length;
  const countHandledVulnByLevel = (level) => vulnRows
    .filter((row) => (
      toText(row['漏洞等级']) === level
      && toText(row['跟进状态']).includes('已')
    ))
    .length;
  const d83 = countVulnByLevel('高危');
  const d84 = countVulnByLevel('中危');
  const d85 = countVulnByLevel('低危');
  // D110/D111 对应资产漏洞表中高危漏洞总数及其唯一 IP 数量。
  const highRiskIpSet = new Set(
    highRiskVulnRows
      .map(row => toText(row && row['IP']).trim())
      .filter(Boolean)
  );
  const d110 = highRiskVulnRows.length;
  const d111 = highRiskIpSet.size;
  const d122 = d110 + d115;
  const d126 = eventRows.filter((row) => (
    isDateInRange(row.create_time, range)
    && ['\u7f51\u7ad9\u7be1\u6539', '\u9ed1\u94fe'].some(keyword => (
      getEventManageSubTypeName(row).includes(keyword)
    ))
  )).length;
  const d121 = `${d122 + d126}+15-3/15-4/15-5`;
  const e83 = countHandledVulnByLevel('高危');
  const e84 = countHandledVulnByLevel('中危');
  const e85 = countHandledVulnByLevel('低危');
  const f83 = d83 > 0 ? formatRatioAsPercentage(e83 / d83, 2) : '0%';
  const f84 = d84 > 0 ? formatRatioAsPercentage(e84 / d84, 2) : '0%';
  const f85 = d85 > 0 ? formatRatioAsPercentage(e85 / d85, 2) : '0%';
  const i79 = getWorksheetCellValue(exposedSurfaceWorksheet, 'E4');
  const i80 = getWorksheetCellValue(exposedSurfaceWorksheet, 'E5');
  const i81 = exposedSurfacePortCount;
  const d128 = Array.from({ length: 12 }, (_, index) => index + 2)
    .reduce((total, rowNumber) => {
      const value = toNumericOrNull(getWorksheetCellValue(
        exposedSurfaceWorksheet,
        `E${rowNumber}`
      ));
      return total + (value === null ? 0 : value);
    }, 0);
  const i82 = vulnRows.filter(row => toText(row['内/外网']) === '外网').length;
  const k79 = countAssetRowsByCriteria(assetWorksheet, {
    assetType: '服务器',
    securityDomain: '内网'
  });
  const k80 = countAssetRowsByCriteria(assetWorksheet, {
    assetType: '服务器',
    securityDomain: '内网',
    serviceType: '服务内'
  });
  const k81 = vulnRows.filter(row => toText(row['内/外网']) === '内网').length;
  const k82 = vulnRows.filter((row) => (
    toText(row['内/外网']) === '内网'
    && toText(row['跟进状态']) === '已防护'
  )).length;
  const k83 = vulnRows.filter((row) => (
    toText(row['内/外网']) === '内网'
    && toText(row['跟进状态']) === '已修复'
  )).length;

  const g5 = alarmRows.filter((row) => (
    isDateInRange(row.create_time, range)
    && toText(row.type).includes('外部威胁')
  )).length;
  const g16 = countDistinctBusinessSystems(assetWorksheet);
  const g17 = countAssetRowsByCriteria(assetWorksheet, {
    assetType: '服务器',
    serviceType: '服务内'
  });
  const g18 = countAssetRowsByCriteria(assetWorksheet, {
    assetType: '终端',
    serviceType: '服务内'
  });
  const d62 = alarmRows.filter(row => toText(row.type).includes('威胁')).length;
  const d63 = eventRows.filter((row) => (
    isDateInRange(row.create_time, range)
    && ['重大威胁', '一般威胁'].includes(toText(row.event_grading_tag))
  )).length;
  const d64Numbers = eventRows
    .filter((row) => (
      isThreatEventType(row.event_grading_tag)
      && toText(row.push_status) === '已通告'
    ))
    .map(row => toNumericOrNull(row['响应时长']))
    .filter(value => value !== null);
  const d64 = averageNumbers(d64Numbers);

  const g6Numbers = eventRows
    .filter(row => isThreatEventType(row.event_grading_tag))
    .map(row => toNumericOrNull(row['响应时长']))
    .filter(value => value !== null);
  const legacyG6 = averageNumbers(g6Numbers);

  const g7 = eventRows.filter((row) => (
    isDateInRange(row.create_time, range)
    && isEventCategoryType(row.event_grading_tag)
  )).length;

  const g8Numbers = eventRows
    .filter((row) => (
      isEventCategoryType(row.event_grading_tag)
      && toText(row.push_status) === '已通告'
    ))
    .map(row => toNumericOrNull(row['响应时长']))
    .filter(value => value !== null);
  const g8 = averageNumbers(g8Numbers);
  const e49Numbers = eventRows
    .filter((row) => (
      isEventCategoryType(row.event_grading_tag)
      && toText(row.push_status) === '已通告'
    ))
    .map(row => toNumericOrNull(row['处置时长']))
    .filter(value => value !== null);
  const e49 = averageNumbers(e49Numbers);
  const averageAnnouncedEventDuration = (fieldName) => {
    const numbers = eventRows
      .filter((row) => (
        isEventCategoryType(row.event_grading_tag)
        && toText(row.push_status) === '已通告'
      ))
      .map(row => toNumericOrNull(row[fieldName]))
      .filter(value => value !== null);
    return averageNumbers(numbers);
  };

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
  const legacyG12 = eventRows.filter((row) => (
    isDateInRange(row.create_time, range)
    && isStatisticsEventCountType(row.event_grading_tag)
  )).length;
  const d28 = countRowsInProtectionRanges(
    alarmRows,
    ['create_time', '创建时间', '告警创建时间'],
    mergedHolidayProtectionRanges
  );
  const d29 = countRowsInProtectionRanges(
    eventRows,
    ['create_time', '事件创建时间', '创建时间'],
    mergedHolidayProtectionRanges
  );
  const g4 = eventStats && eventStats.strategyOptimizeCount !== undefined
    ? eventStats.strategyOptimizeCount
    : '';
  const e80WeakPwdTotal = toNumericOrNull(weakPwdSummaryTotal) || 0;
  const e80AccountSecurityEventCount = toNumericOrNull(
    eventStats && eventStats.accountSecurityEventCountForE80
  ) || 0;
  const e80 = e80WeakPwdTotal + e80AccountSecurityEventCount;
  const d10 = c80;
  const d11WeakPwdTotal = e80WeakPwdTotal + accountSecurityKeywordEventCount;
  const d112 = d11WeakPwdTotal;
  const d6 = d10 + d11WeakPwdTotal + d12;
  const businessSystemStats = buildBusinessSystemStatistics({
    businessSystems,
    assetWorksheet,
    range,
    eventRows,
    vulnRows,
    weakPwdSummaryList,
    weakPwdHandledList
  });
  const coreSystemTicketStats = businessSystemStats.map((system) => buildCoreSystemTicketStatistics({
    ipSet: new Set((system && system.ips) || []),
    vulnRows,
    eventRows,
    weakPwdAllTotalsByIp,
    weakPwdHandledTotalsByIp
  }));
  const [coreSystem1 = {}, coreSystem2 = {}, coreSystem3 = {}] = coreSystemTicketStats;
  if (businessSystemStats[0]) {
    console.log(`[G111] System: ${businessSystemStats[0].name}; IP count: ${(businessSystemStats[0].ips || []).length}; Vulnerabilities: ${coreSystem1.totalVulnCount || 0}; Weak passwords: ${coreSystem1.totalWeakPwdCount || 0}; Latest threats: ${coreSystem1.totalLatestThreatCount || 0}; Total: ${coreSystem1.total || 0}`);
  } else {
    console.log('[G111] No first business system configured; G111=0');
  }
  const g111 = coreSystem1.total || 0;
  const g112 = coreSystem1.handledTotal || 0;
  const g113 = coreSystem1.ratio || 0;
  const g114 = coreSystem2.total || 0;
  const g115 = coreSystem2.handledTotal || 0;
  const g116 = coreSystem2.ratio || 0;
  const g117 = coreSystem3.total || 0;
  const g118 = coreSystem3.handledTotal || 0;
  const g119 = coreSystem3.ratio || 0;
  const aggregatedBusinessSystemIps = new Set();
  businessSystemStats.forEach((item) => {
    (item && Array.isArray(item.ips) ? item.ips : []).forEach(ip => aggregatedBusinessSystemIps.add(ip));
  });
  const i123 = sumWeakPwdTicketTotalsForIps(
    weakPwdHandledTotalsByIp,
    aggregatedBusinessSystemIps
  );
  const i124 = eventRows.filter((row) => (
    isDateInRange(row.create_time, range)
    && toText(row.event_status) === '已闭环'
    && ['\u6728\u9a6c', '\u75c5\u6bd2'].some(keyword => (
      getEventManageSubTypeName(row).includes(keyword)
    ))
  )).length;
  const businessSystemAlarmCounts = businessSystemStats.map((item) => {
    const ipSet = new Set(item && Array.isArray(item.ips) ? item.ips : []);
    return countRowsContainingIps(alarmRows, ['host_ip'], ipSet);
  });
  const g6 = businessSystemAlarmCounts[0] || 0;
  const h6 = businessSystemAlarmCounts[1] || 0;
  const i6 = businessSystemAlarmCounts[2] || 0;
  const g10 = eventRows.filter((row) => (
    isEventCategoryType(row.event_grading_tag)
    && rowContainsAnyIp(row, ['host_ip', 'affected_assets'], aggregatedBusinessSystemIps)
  )).length;
  const g12 = businessSystemStats.reduce((total, item) => total + (Number(item && item.total) || 0), 0);
  const g13 = businessSystemStats.reduce((total, item) => total + (Number(item && item.closedLoopTotal) || 0), 0);
  const g14 = eventRows.filter((row) => (
    isStatisticsEventCountType(row.event_grading_tag)
    && rowContainsAnyIp(row, ['host_ip', 'affected_assets'], aggregatedBusinessSystemIps)
  )).length;
  const businessSystemCells = {};
  for (let index = 0; index < 3; index += 1) {
    const item = businessSystemStats[index];
    businessSystemCells[`D${3 + index}`] = item ? item.name : '';
    businessSystemCells[`D${7 + index}`] = item ? item.total : '';
    businessSystemCells[`${XLSX.utils.encode_col(2 + index)}18`] = item ? item.closedLoopTotal : '';
  }
  const topEventManageSubTypeStats = buildTopEventManageSubTypeStats(eventRows, 5);
  const topEventManageSubTypeCells = {};
  for (let index = 0; index < 5; index += 1) {
    const row = 51 + index;
    const item = topEventManageSubTypeStats[index];
    if (item) {
      topEventManageSubTypeCells[`C${row}`] = item.name;
      topEventManageSubTypeCells[`D${row}`] = item.count;
    }
  }
  const topnReportCells = buildTopnReportCells(topnReportStats);
  const monthTrendCells = {};
  const blankTrendCells = [];
  for (let row = 51; row <= 55; row += 1) {
    blankTrendCells.push(`C${row}`, `D${row}`);
  }
  for (let row = 62; row <= 66; row += 1) {
    blankTrendCells.push(`F${row}`, `G${row}`, `I${row}`, `J${row}`, `L${row}`, `M${row}`);
  }
  for (let i = 0; i < 12; i += 1) {
    const col = XLSX.utils.encode_col(2 + i);
    const month = addMonths(startMonth, i);
    const monthRange = buildMonthRange(month);
    const shouldDisplayMonth = i < displayMonthCount;
    if (!shouldDisplayMonth) {
      ['58', '59', '71', '72', '73', '74', '75', '86', '87', '88']
        .forEach(row => blankTrendCells.push(`${col}${row}`));
      continue;
    }
    const responseNumbers = eventRows
      .filter(row => isDateInRange(row.create_time, monthRange))
      .map(row => toNumericOrNull(row['响应时长']))
      .filter(value => value !== null);
    monthTrendCells[`${col}58`] = formatYearMonth(month);
    monthTrendCells[`${col}59`] = averageNumbers(responseNumbers);
    monthTrendCells[`${col}71`] = formatYearMonth(month);
    monthTrendCells[`${col}72`] = '';
    monthTrendCells[`${col}73`] = '';
    monthTrendCells[`${col}74`] = '';
    monthTrendCells[`${col}75`] = alarmRows
      .filter(row => isDateInRange(row.create_time, monthRange))
      .length;
    monthTrendCells[`${col}86`] = formatYearMonth(month);
    monthTrendCells[`${col}87`] = vulnRows
      .filter(row => isDateInRange(row['更新时间'], monthRange))
      .length;
    monthTrendCells[`${col}88`] = vulnRows
      .filter((row) => (
        toText(row['跟进状态']).includes('已')
        && isDateInRange(row['更新时间'], monthRange)
      ))
      .length;
  }
  const monthlyOut2inCells = {};
  for (let i = 0; i < 12; i += 1) {
    const col = XLSX.utils.encode_col(2 + i);
    monthlyOut2inCells[`${col}71`] = '';
    monthlyOut2inCells[`${col}72`] = '';
    monthlyOut2inCells[`${col}73`] = '';
  }
  (Array.isArray(xdrMonthlyOut2inLogSearchCounts) ? xdrMonthlyOut2inLogSearchCounts : [])
    .slice(0, 12)
    .forEach((item, index) => {
      if (index >= displayMonthCount) return;
      const col = XLSX.utils.encode_col(2 + index);
      monthlyOut2inCells[`${col}71`] = item && item.monthLabel ? item.monthLabel : '';
      monthlyOut2inCells[`${col}72`] = item && item.count !== undefined ? item.count : '';
    });
  (Array.isArray(xdrMonthlyIn2outLogSearchCounts) ? xdrMonthlyIn2outLogSearchCounts : [])
    .slice(0, 12)
    .forEach((item, index) => {
      if (index >= displayMonthCount) return;
      const col = XLSX.utils.encode_col(2 + index);
      monthlyOut2inCells[`${col}73`] = item && item.count !== undefined ? item.count : '';
    });

  const holidayCells = {};
  for (let row = 4; row <= 30; row += 1) {
    holidayCells[`L${row}`] = '';
    holidayCells[`M${row}`] = '';
  }
  holidayPeriods.forEach((period, index) => {
    const row = 4 + index;
    holidayCells[`L${row}`] = period.name;
    holidayCells[`M${row}`] = formatReportDateRange(period.startDate, period.endDate);
  });
  const holidaySummaryCells = {};
  const holidaySummaryBlankCells = [];
  for (let row = 91; row <= 97; row += 1) {
    holidaySummaryCells[`F${row}`] = '';
    holidaySummaryCells[`G${row}`] = '';
    holidaySummaryCells[`H${row}`] = '';
    holidaySummaryBlankCells.push(`F${row}`, `G${row}`, `H${row}`);
  }
  holidayPeriods.slice(0, 7).forEach((period, index) => {
    const row = 91 + index;
    holidaySummaryCells[`F${row}`] = `${period.name}（${formatReportDateRange(period.startDate, period.endDate)}）`;
    const count = xdrHolidayLogSearchCounts && Object.prototype.hasOwnProperty.call(xdrHolidayLogSearchCounts, period.key)
      ? toNumericOrNull(xdrHolidayLogSearchCounts[period.key])
      : null;
    holidaySummaryCells[`G${row}`] = count !== null && period.dayCount > 0
      ? roundTo(count / period.dayCount, 2)
      : '';
    holidaySummaryCells[`H${row}`] = '100%';
  });

  const d101 = toNumericOrNull(orderBranchCounts && orderBranchCounts.d101) || 0;
  const d102 = toNumericOrNull(orderBranchCounts && orderBranchCounts.d102) || 0;
  const d100 = d101 + d102;
  const d105 = toNumericOrNull(orderBranchCounts && orderBranchCounts.d105) || 0;
  const g99 = toNumericOrNull(xdrG99LogSearchCount) || 0;
  const g100 = toNumericOrNull(xdrG100AlertCount) || 0;
  const g101 = toNumericOrNull(xdrG101IncidentCount) || 0;
  const g102 = toNumericOrNull(xdrG102IncidentHandledCount) || 0;
  const g103 = toNumericOrNull(xdrG103IncidentCount) || 0;
  const g104 = g102 - g103;
  const g105 = toNumericOrNull(xdrG105IncidentCount) || 0;
  const g106 = toNumericOrNull(xdrG106IncidentCount) || 0;
  const g107 = toNumericOrNull(xdrG107IncidentCount) || 0;
  const alertThreatDefineStatistics = buildThreatDefineStatisticsCells(xdrAlertThreatDefineCounts, {
    labelRow: 133,
    ratioRow: 134,
    maxItems: 7
  });
  const incidentThreatDefineStatistics = buildThreatDefineStatisticsCells(xdrIncidentThreatDefineCounts, {
    labelRow: 135,
    countRow: 136,
    ratioRow: 137,
    maxItems: 10
  });

  const cells = {
    J1: customer || '',
    L1: formatReportDate(startDate),
    M1: formatReportDate(endDate),
    M3: protectStartDate && protectEndDate ? formatReportDateRange(protectStartDate, protectEndDate) : '',
    ...holidayCells,
    ...businessSystemCells,
    D6: d6,
    D10: d10,
    D11: d11WeakPwdTotal,
    D12: d12,
    D13: d13,
    D14: d14,
    D15: d15,
    D16: d16,
    D17: d17,
    D20: d20,
    D21: d21,
    D22: g4,
    D23: xdrRejectedExternalToInternalCount,
    D24: d24,
    D25: '',
    D27: xdrLogSearchCount,
    D28: d28,
    D29: d29,
    D30: '手写',
    D32: '100%',
    D33: '',
    D34: d34,
    D35: d35,
    D36: d36,
    D37: d37,
    D38: d38,
    D39: '',
    C49: g7,
    D49: g8,
    E49: e49,
    F49: g9,
    ...topEventManageSubTypeCells,
    G51: averageAnnouncedEventDuration('识别时长'),
    G52: averageAnnouncedEventDuration('响应时长'),
    G53: averageAnnouncedEventDuration('遏制时长'),
    G54: averageAnnouncedEventDuration('处置时长'),
    G55: averageAnnouncedEventDuration('闭环时长'),
    ...monthTrendCells,
    ...monthlyOut2inCells,
    D61: xdrOut2inLogSearchCount,
    D62: d62,
    D63: d63,
    D64: d64,
    D65: '',
    D66: '',
    D67: d11,
    D68: d12,
    ...topnReportCells,
    M61: '',
    C80: c80,
    D80: d80,
    E80: e80,
    F80: f80,
    G21: '',
    G22: g22,
    G23: d14,
    G24: d15,
    D83: d83,
    D84: d84,
    D85: d85,
    D110: d110,
    D111: d111,
    D112: d112,
    D113: d113,
    D114: d114,
    D115: d115,
    D116: d116,
    D117: d117,
    D118: d118,
    D121: d121,
    D122: d122,
    D126: d126,
    D129: d129,
    D130: d130,
    D128: d128,
    E83: e83,
    E84: e84,
    E85: e85,
    F83: f83,
    F84: f84,
    F85: f85,
    I79: i79,
    I80: i80,
    I81: i81,
    I82: i82,
    I122: toNumericOrNull(weakPwdHandledTotal) || 0,
    I123: i123,
    I124: i124,
    K79: k79,
    K80: k80,
    K81: k81,
    K82: k82,
    K83: k83,
    N79: '',
    D90: importantProtectionCount,
    D91: d29,
    D92: '100%',
    ...holidaySummaryCells,
    G4: g4,
    G5: g5,
    G6: g6,
    G7: legacyG6,
    G8: g7,
    G9: g8,
    H6: h6,
    I6: i6,
    G10: g10,
    G11: '',
    G12: g12,
    G13: g13,
    G14: g14,
    G16: g16,
    G17: g17,
    G18: g18,
    G99: g99,
    G100: g100,
    G101: g101,
    G102: g102,
    G103: g103,
    G104: g104,
    G105: g105,
    G106: g106,
    G107: g107,
    G110: g110,
    G111: g111,
    G112: g112,
    G113: g113,
    G114: g114,
    G115: g115,
    G116: g116,
    G117: g117,
    G118: g118,
    G119: g119,
    G122: d110,
    G123: d12,
    G124: d15,
    G125: toNumericOrNull(weakPwdAllTotal) || 0,
    G126: toNumericOrNull(g126) || 0,
    ...alertThreatDefineStatistics.cells,
    ...incidentThreatDefineStatistics.cells,
    D100: d100,
    D101: d101,
    D102: d102,
    D105: d105
  };
  cells.__blankCells = blankTrendCells.concat(holidaySummaryBlankCells, ['D3', 'D4', 'D5', 'D7', 'D8', 'D9']);
  cells.__percentCells = alertThreatDefineStatistics.percentCells
    .concat(incidentThreatDefineStatistics.percentCells, ['G113', 'G116', 'G119']);
  return cells;
}

function populateStatisticsSheet(wb, statsContext) {
  if (!wb || !wb.Sheets || !wb.Sheets['数据统计']) {
    throw new Error('模板缺少数据统计 Sheet');
  }

  const ws = wb.Sheets['数据统计'];
  const cells = buildStatisticsCells(statsContext);
  const blankCells = Array.isArray(cells.__blankCells) ? cells.__blankCells : [];
  const percentCells = Array.isArray(cells.__percentCells) ? cells.__percentCells : [];
  delete cells.__blankCells;
  delete cells.__percentCells;

  blankCells.forEach(address => {
    setWorksheetCellValue(ws, address, '');
  });

  Object.entries(cells).forEach(([address, value]) => {
    if (value === '' || value === null || value === undefined) return;
    setWorksheetCellValue(ws, address, value);
  });
  percentCells.forEach(address => {
    if (Object.prototype.hasOwnProperty.call(cells, address)) {
      setWorksheetPercentCellValue(ws, address, cells[address]);
    }
  });

  return cells;
}

/**
 * 生成完整报告
 * @param {Object} options - 选项 { customer, startDate, endDate, eventData, alarmData, outputDir }
 * @returns {Object} 生成的文件信息
 */
function generateReport(options) {
  const {
    customer,
    customerId,
    startDate,
    endDate,
    protectStartDate,
    protectEndDate,
    reportTemplatePath,
    assetWorkbookBuffer,
    exposedSurfaceWorkbookBuffer,
    exposedSurfacePortCount,
    weakPwdSummaryTotal,
    weakPwdSummaryList,
    weakPwdAllSummaryList,
    weakPwdAllTotalsByIp,
    weakPwdAllTotal,
    weakPwdHandledTotal,
    weakPwdHandledAdminTotal,
    weakPwdHandledList,
    weakPwdHandledTotalsByIp,
    g126,
    d129,
    d130,
    topnReportStats,
    eventData,
    eventStats,
    alarmData,
    vulnData,
    businessSystems,
    vulnRawRowCount,
    xdrIn2outLogSearchCount,
    xdrOut2inLogSearchCount,
    xdrRejectedExternalToInternalCount,
    xdrG99LogSearchCount,
    xdrG100AlertCount,
    xdrG101IncidentCount,
    xdrG102IncidentHandledCount,
    xdrG103IncidentCount,
    xdrG105IncidentCount,
    xdrG106IncidentCount,
    xdrG107IncidentCount,
    xdrAlertThreatDefineCounts,
    xdrIncidentThreatDefineCounts,
    xdrMonthlyIn2outLogSearchCounts,
    xdrMonthlyOut2inLogSearchCounts,
    xdrLogSearchCount,
    xdrHolidayLogSearchCounts,
    orderBranchCounts,
    outputDir,
    eventHeaders,
    alarmHeaders,
    vulnHeaders,
    eventHeaderDisplayMap,
    alarmHeaderDisplayMap,
    vulnHeaderDisplayMap
  } = options;
  
  const result = {
    success: true,
    files: [],
    errors: []
  };
  
  const outputPath = outputDir || __dirname;
  
  try {
    const reportFileName = generateFileName(customer, startDate, endDate, 'report', customerId);
    const reportFilePath = path.join(outputPath, reportFileName);
    const wb = readWorkbookFromFile(reportTemplatePath);
    ensureWorksheetsExist(wb, ['数据统计', '资产表', '暴露面', '资产漏洞表', '告警表', '事件表']);
    const assetWorkbook = readWorkbookFromBuffer(assetWorkbookBuffer);
    const assetWorksheet = getFirstWorksheet(assetWorkbook);
    const exposedSurfaceWorkbook = readWorkbookFromBuffer(exposedSurfaceWorkbookBuffer);
    const exposedSurfaceWorksheet = getWorksheetByNames(exposedSurfaceWorkbook, ['统计', '统计sheet', '统计Sheet']);
    if (exposedSurfaceWorkbookBuffer && !exposedSurfaceWorksheet) {
      throw new Error('暴露面工作簿中未找到统计 Sheet');
    }
    const assetRowCount = countWorksheetDataRows(assetWorksheet);
    const exposedSurfaceRowCount = countWorksheetDataRows(exposedSurfaceWorksheet);
    const vulnRowCount = Array.isArray(vulnData) ? vulnData.length : 0;
    const alarmRowCount = Array.isArray(alarmData) ? alarmData.length : 0;
    const eventRowCount = Array.isArray(eventData) ? eventData.length : 0;

    replaceWorksheet(wb, '资产表', assetWorksheet);
    replaceWorksheet(wb, '暴露面', exposedSurfaceWorksheet);
    replaceWorksheet(wb, '资产漏洞表', createWorksheet(vulnData || [], vulnHeaders || null, vulnHeaderDisplayMap || null));
    replaceWorksheet(wb, '告警表', createWorksheet(alarmData || [], alarmHeaders || null, alarmHeaderDisplayMap || null));
    replaceWorksheet(wb, '事件表', createWorksheet(eventData || [], eventHeaders || null, eventHeaderDisplayMap || null));
    populateStatisticsSheet(wb, {
      customer,
      startDate,
      endDate,
      protectStartDate,
      protectEndDate,
      eventRows: eventData || [],
      eventStats: eventStats || {},
      alarmRows: alarmData || [],
      vulnRows: vulnData || [],
      businessSystems,
      assetWorksheet,
      exposedSurfaceWorksheet,
      exposedSurfacePortCount,
      weakPwdSummaryTotal,
      weakPwdSummaryList,
      weakPwdAllSummaryList,
      weakPwdAllTotalsByIp,
      weakPwdAllTotal,
      weakPwdHandledTotal,
      weakPwdHandledAdminTotal,
      weakPwdHandledList,
      weakPwdHandledTotalsByIp,
      g126,
      d129,
      d130,
      topnReportStats,
      xdrIn2outLogSearchCount,
      xdrOut2inLogSearchCount,
      xdrRejectedExternalToInternalCount,
      xdrG99LogSearchCount,
      xdrG100AlertCount,
      xdrG101IncidentCount,
      xdrG102IncidentHandledCount,
      xdrG103IncidentCount,
      xdrG105IncidentCount,
      xdrG106IncidentCount,
      xdrG107IncidentCount,
      xdrAlertThreatDefineCounts,
      xdrIncidentThreatDefineCounts,
      xdrMonthlyIn2outLogSearchCounts,
      xdrMonthlyOut2inLogSearchCounts,
      xdrLogSearchCount,
      xdrHolidayLogSearchCounts,
      orderBranchCounts
    });

    ensureOutputDir(reportFilePath);
    XLSX.writeFile(wb, reportFilePath);
    result.files.push({
      type: 'report',
      path: reportFilePath,
      name: reportFileName,
      rowCount: assetRowCount + exposedSurfaceRowCount + vulnRowCount + alarmRowCount + eventRowCount,
      sheets: [
        { name: '数据统计', rowCount: 0 },
        { name: '资产表', rowCount: assetRowCount },
        { name: '暴露面', rowCount: exposedSurfaceRowCount },
        {
          name: '资产漏洞表',
          rowCount: vulnRowCount,
          ...(Number.isInteger(vulnRawRowCount) ? { rawRowCount: vulnRawRowCount } : {})
        },
        { name: '告警表', rowCount: alarmRowCount },
        { name: '事件表', rowCount: eventRowCount }
      ]
    });

    console.log(`[DataFormatter] 报告已保存: ${reportFilePath}`);
    console.log(`[DataFormatter] 使用模板: ${reportTemplatePath}`);
    console.log(`[DataFormatter] 已写入 Sheet: 资产表、暴露面、资产漏洞表、告警表、事件表`);
    console.log(`[DataFormatter] 已更新 Sheet: 数据统计`);
    
  } catch (error) {
    result.success = false;
    result.errors.push(error.message);
    console.error(`[DataFormatter] 生成报告失败: ${error.message}`);
  }
  
  return result;
}

/**
 * 分析数据结构（用于调试）
 * @param {Array} data - 数据数组
 * @returns {Object} 结构分析结果
 */
function analyzeDataStructure(data) {
  if (!Array.isArray(data) || data.length === 0) {
    return { fields: [], sample: null };
  }
  
  const fields = new Set();
  const sample = flattenObject(data[0]);
  
  data.forEach(item => {
    Object.keys(flattenObject(item)).forEach(key => fields.add(key));
  });
  
  return {
    fieldCount: fields.size,
    fields: Array.from(fields).sort(),
    sample
  };
}

module.exports = {
  inferType,
  flattenObject,
  normalizeDataFields,
  formatCellValue,
  smartFormatData,
  createWorksheet,
  createWorkbook,
  generateFileName,
  saveToExcel,
  formatReportDate,
  parseReportDateValue,
  countInclusiveNaturalDays,
  buildReportSecondRange,
  buildReportMonthSecondRanges,
  buildProtectionSecondRanges,
  buildProtectionAtomicSecondRanges,
  normalizeBusinessSystems,
  collectBusinessSystemIpSets,
  buildBusinessSystemStatistics,
  buildStatisticsCells,
  populateStatisticsSheet,
  generateReport,
  analyzeDataStructure
};
