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
  return { start: rangeStart, end: rangeEnd };
}

function isDateInRange(value, range) {
  const date = parseReportDateValue(value);
  if (!date || !range) return false;
  return date >= range.start && date <= range.end;
}

function toText(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
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
  return ['重要威胁', '一般威胁'].includes(text) || text.includes('威胁');
}

function isEventCategoryType(eventType) {
  const text = toText(eventType);
  if (!text) return false;
  return ['重大事件', '重要事件', '一般事件'].includes(text) || text.includes('事件');
}

function getWorksheetCell(ws, address) {
  if (!ws[address]) {
    ws[address] = { t: 's', v: '' };
  }
  return ws[address];
}

function setWorksheetCellValue(ws, address, value) {
  const cell = getWorksheetCell(ws, address);

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

function buildStatisticsCells(statsContext) {
  const {
    customer,
    startDate,
    endDate,
    eventRows = [],
    eventStats = {},
    alarmRows = [],
    vulnRows = []
  } = statsContext || {};

  const range = buildInclusiveDateRange(startDate, endDate);

  const d6 = vulnRows.filter(row => isDateInRange(row['更新时间'], range)).length;
  const d11 = eventRows.filter((row) => (
    toText(row.event_grading_tag) === '最新威胁'
    && isDateInRange(row.create_time, range)
  )).length;
  const d16 = vulnRows.filter(row => toText(row['跟进状态']) === '已防护').length;
  const d17 = vulnRows.filter(row => toText(row['跟进状态']) === '已修复').length;
  const d15 = '';
  const d14 = (toNumericOrNull(d15) || 0) + d16 + d17;
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

  const g5 = alarmRows.filter((row) => (
    isDateInRange(row.create_time, range)
    && toText(row.type).includes('外部威胁')
  )).length;

  const g6Numbers = eventRows
    .filter(row => isThreatEventType(row.event_grading_tag))
    .map(row => toNumericOrNull(row['响应时长']))
    .filter(value => value !== null);
  const g6 = averageNumbers(g6Numbers);

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
  const g4 = eventStats && eventStats.strategyOptimizeCount !== undefined
    ? eventStats.strategyOptimizeCount
    : '';

  return {
    J1: customer || '',
    L1: formatReportDate(startDate),
    M1: formatReportDate(endDate),
    D3: '',
    D4: '',
    D5: '',
    D6: d6,
    D7: '',
    D8: '',
    D9: '',
    D10: '',
    D11: d11,
    D12: '',
    D14: d14,
    D15: d15,
    D16: d16,
    D17: d17,
    D18: '',
    D20: '',
    D21: '',
    D22: '',
    D23: d17,
    D24: d16,
    D25: '',
    D27: '',
    D28: '',
    D29: '',
    D30: '',
    D32: '',
    D33: '',
    D34: d34,
    D35: d35,
    D36: d36,
    D37: d37,
    D38: d38,
    D39: '',
    G4: g4,
    G5: g5,
    G6: g6,
    G7: g7,
    G8: g8,
    G9: g9
  };
}

function populateStatisticsSheet(wb, statsContext) {
  if (!wb || !wb.Sheets || !wb.Sheets['数据统计']) {
    throw new Error('模板缺少数据统计 Sheet');
  }

  const ws = wb.Sheets['数据统计'];
  const cells = buildStatisticsCells(statsContext);

  Object.entries(cells).forEach(([address, value]) => {
    setWorksheetCellValue(ws, address, value);
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
    reportTemplatePath,
    assetWorkbookBuffer,
    exposedSurfaceWorkbookBuffer,
    eventData,
    eventStats,
    alarmData,
    vulnData,
    vulnRawRowCount,
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
      eventRows: eventData || [],
      eventStats: eventStats || {},
      alarmRows: alarmData || [],
      vulnRows: vulnData || []
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
  buildStatisticsCells,
  populateStatisticsSheet,
  generateReport,
  analyzeDataStructure
};
