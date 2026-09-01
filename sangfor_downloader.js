/**
 * 深信服报告下载器 - 主脚本
 * 自动下载指定客户、指定时间段的【资产表】、【暴露面】、【事件表】、【告警表】和【资产漏洞表】Excel数据
 * 
 * 使用方法:
 *   node sangfor_downloader.js                                    # 交互式
 *   node sangfor_downloader.js --customer "客户名" --start "2024-01-01" --end "2024-01-31"
 *   node sangfor_downloader.js --type asset                       # 只下载资产表
 *   node sangfor_downloader.js --type event                       # 只下载事件表
 *   node sangfor_downloader.js --type alarm                       # 只下载告警表
 *   node sangfor_downloader.js --type vuln                        # 只下载资产漏洞表
 *   node sangfor_downloader.js --type exposed                     # 只下载暴露面
 */
const path = require('path');
const fs = require('fs');

// 引入模块
const cookieReader = require('./cookie_reader');
const apiClient = require('./api_client');

const DEFAULT_REPORT_TEMPLATE_PATH = path.join(__dirname, 'data.xlsx');
const DEFAULT_COOKIE_DOWNLOAD_DIR = process.env.USERNAME
  ? path.join('M:\\Users', process.env.USERNAME, 'Downloads')
  : __dirname;

// 配置
const CONFIG = {
  // Cookie 文件路径 (由浏览器插件生成)
  cookiePath: process.env.SANGFOR_COOKIE_PATH || path.join(DEFAULT_COOKIE_DOWNLOAD_DIR, 'cookies.txt'),
  xdrCookiePath: process.env.SANGFOR_XDR_COOKIE_PATH || path.join(DEFAULT_COOKIE_DOWNLOAD_DIR, 'xdr_cookies.txt'),
  xdrBaseUrl: process.env.SANGFOR_XDR_BASE_URL || '',

  // 报告模板路径
  reportTemplatePath: process.env.SANGFOR_REPORT_TEMPLATE_PATH || DEFAULT_REPORT_TEMPLATE_PATH,
  
  // 输出目录
  outputDir: __dirname,
  
  // 默认请求参数
  defaultPageSize: 100,
  defaultVulnPageSize: 50,
  defaultPageDelayMs: 10,
  maxRetries: 3,
  retryDelay: 2000
};

const EVENT_HEADER_DISPLAY_MAP = {
  event_grading_tag: '事件类型',
  event_name: '事件名称',
  create_time: '事件创建时间',
  type: '类型',
  host_ip: '主机 IP',
  event_status: '状态',
  service_status: '服务类型',
  latest_time: '最近更新时间',
  checkout_time: '组件检出时间',
  dispose_time: '专家研判时间',
  contain_time: '事件遏制时间',
  finished_time: '事件闭环时间',
  incidence: '影响范围',
  update_protected_time: '防护时间',
  affected_assets: '受影响资产',
  update_announced_time: '通告时间',
  update_accept_risk_time: '接受风险时间',
  push_status: '通告状态',
  wechat_push_time: '实际通告时间'
};

const ALARM_HEADER_DISPLAY_MAP = {
  alarm_name: '告警名称',
  host_ip: '主机IP',
  type: '威胁类型',
  event_status: '告警状态',
  first_time: '首次触发时间',
  latest_time: '最近触发时间',
  service_status: '服务类型',
  create_time: '创建时间',
  attack_state: '攻击结果',
  attack_direction: '访问方向',
  current_operate_time: '操作时间',
  rejected_event_id: '驳回事件ID',
  current_operator: '驳回人',
  reject_reason: '驳回原因'
};

function normalizeReportType(type) {
  const normalized = String(type || 'all').trim().toLowerCase();
  const aliases = {
    exposure: 'exposed',
    exposed_surface: 'exposed',
    exposed_surface_mss: 'exposed',
    surface: 'exposed',
    expose: 'exposed',
    '暴露面': 'exposed',
    '漏洞': 'vuln'
  };
  return aliases[normalized] || normalized || 'all';
}

function shouldFetchType(selectedType, targetType) {
  return selectedType === 'all' || selectedType === targetType;
}

function resolvePageSizeForTarget(targetType, options) {
  if (targetType === 'vuln' && !options.pageSizeProvided) {
    return CONFIG.defaultVulnPageSize;
  }
  return options.pageSize;
}

function validateReportType(type) {
  const allowed = ['all', 'asset', 'event', 'alarm', 'vuln', 'exposed'];
  if (!allowed.includes(type)) {
    throw new Error(`报告类型错误: ${type}，可选值: ${allowed.join(', ')}`);
  }
}

function normalizeBusinessSystems(value) {
  if (value === undefined || value === null || value === '') return [];

  const rawItems = Array.isArray(value) ? value : [value];
  const normalized = [];
  for (const rawItem of rawItems) {
    String(rawItem || '')
      .split(/[，、,；;\r\n]+/)
      .map(item => item.replace(/\s+/g, ' ').trim())
      .filter(Boolean)
      .forEach(item => normalized.push(item));
  }

  if (normalized.length > 3) {
    throw new Error('最多支持 3 个业务系统。请精简后重试。');
  }

  return normalized;
}

/**
 * 解析命令行参数
 */
function parseArgs() {
  const args = process.argv.slice(2);
  const options = {
    customer: '',
    startDate: '',
    endDate: '',
    type: 'all',  // all, asset, event, alarm, vuln
    customerId: '',
    protectStartDate: '',
    protectEndDate: '',
    pageSize: CONFIG.defaultPageSize,
    pageSizeProvided: false,
    pageDelayMs: CONFIG.defaultPageDelayMs,
    responseOnly: false,
    cookiePath: CONFIG.cookiePath,
    xdrCookiePath: CONFIG.xdrCookiePath,
    xdrBaseUrl: CONFIG.xdrBaseUrl,
    outputDir: CONFIG.outputDir,
    reportTemplatePath: CONFIG.reportTemplatePath,
    manageSubTypeMapFile: process.env.SANGFOR_MANAGE_SUB_TYPE_MAP_FILE || path.join(__dirname, 'data', 'manage_sub_type_map.json'),
    assetMapFile: process.env.SANGFOR_ASSET_MAP_FILE || path.join(__dirname, 'data', 'asset.xlsx'),
    businessSystems: []
  };
  
  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    
    switch (arg) {
      case '--customer':
      case '-c':
        options.customer = args[++i] || '';
        break;
      case '--start':
      case '-s':
        options.startDate = args[++i] || '';
        break;
      case '--end':
      case '-e':
        options.endDate = args[++i] || '';
        break;
      case '--type':
      case '-t':
        options.type = normalizeReportType(args[++i] || 'all');
        break;
      case '--customer-id':
        options.customerId = args[++i] || '';
        break;
      case '--protect-start':
        options.protectStartDate = args[++i] || '';
        break;
      case '--protect-end':
        options.protectEndDate = args[++i] || '';
        break;
      case '--cookie-path':
        options.cookiePath = args[++i] || CONFIG.cookiePath;
        break;
      case '--xdr-cookie-path':
        options.xdrCookiePath = args[++i] || CONFIG.xdrCookiePath;
        break;
      case '--xdr-base-url':
        options.xdrBaseUrl = args[++i] || CONFIG.xdrBaseUrl;
        break;
      case '--output-dir':
        options.outputDir = args[++i] || CONFIG.outputDir;
        break;
      case '--report-template':
        options.reportTemplatePath = args[++i] || CONFIG.reportTemplatePath;
        break;
      case '--page-size':
        options.pageSize = parseInt(args[++i]) || CONFIG.defaultPageSize;
        options.pageSizeProvided = true;
        break;
      case '--page-delay-ms':
        options.pageDelayMs = parseInt(args[++i]);
        if (!Number.isFinite(options.pageDelayMs) || options.pageDelayMs < 0) {
          options.pageDelayMs = CONFIG.defaultPageDelayMs;
        }
        break;
      case '--manage-sub-type-map-file':
        options.manageSubTypeMapFile = args[++i] || '';
        break;
      case '--asset-map-file':
        options.assetMapFile = args[++i] || '';
        break;
      case '--business-systems':
        options.businessSystems = args[++i] || '';
        break;
      case '--response-only':
      case '--debug-response':
        options.responseOnly = true;
        break;
      case '--help':
      case '-h':
        showHelp();
        process.exit(0);
    }
  }
  
  options.type = normalizeReportType(options.type);
  options.businessSystems = normalizeBusinessSystems(options.businessSystems);
  return options;
}

/**
 * 显示帮助信息
 */
function showHelp() {
  console.log(`
深信服报告下载器 - 帮助信息
============================

用法:
  node sangfor_downloader.js [选项]

选项:
  --customer, -c <名称>    客户名称
  --start, -s <日期>       开始日期 (格式: YYYY-MM-DD)
  --end, -e <日期>         结束日期 (格式: YYYY-MM-DD)
  --protect-start <日期>   护网开始日期 (可选，格式: YYYY-MM-DD)
  --protect-end <日期>     护网结束日期 (可选，格式: YYYY-MM-DD)
  --type, -t <类型>        报告类型: all(默认), asset, event, alarm, vuln, exposed
  --customer-id <ID>       客户ID (可选)
  --cookie-path <路径>     Cookie 文件路径 (默认: M:\\Users\\%USERNAME%\\Downloads\\cookies.txt，也可用 SANGFOR_COOKIE_PATH)
  --xdr-cookie-path <路径> XDR Cookie 文件路径 (默认: M:\\Users\\%USERNAME%\\Downloads\\xdr_cookies.txt，也可用 SANGFOR_XDR_COOKIE_PATH)
  --xdr-base-url <域名>    XDR 域名覆盖，例如 xdrsz.sangfor.com.cn（也可用 SANGFOR_XDR_BASE_URL）
  --output-dir <路径>      输出目录 (默认: 脚本所在目录)
  --page-size <大小>       每页数量 (默认: 100，资产漏洞表默认: 50)
  --page-delay-ms <毫秒>   每页请求后的等待时间 (默认: 10)
  --manage-sub-type-map-file <路径> manage_sub_type 映射文件（默认 ./data/manage_sub_type_map.json）
  --asset-map-file <路径>  资产IP与安全域映射文件（默认 ./data/asset.xlsx，支持 xlsx/json）
  --business-systems <名称> 业务系统列表，逗号分隔，最多 3 个
  --response-only          只请求第一页并保存原始 response JSON，不生成 Excel
  --help, -h               显示帮助信息

示例:
  node sangfor_downloader.js -c "测试客户" -s "2024-01-01" -e "2024-01-31"
  node sangfor_downloader.js --customer-id 26912728 --start "2026-05-12" --end "2026-05-13"
  node sangfor_downloader.js --response-only --customer-id 26912728 --start "2026-05-12" --end "2026-05-13"
  `);
}

/**
 * 交互式获取参数
 */
async function interactiveInput() {
  const readline = require('readline');
  
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });
  
  const question = (prompt) => new Promise((resolve) => {
    rl.question(prompt, resolve);
  });
  
  console.log('\n===== 深信服报告下载器 =====\n');
  
  try {
    const customer = await question('客户名称: ');
    const startDate = await question('开始日期 (YYYY-MM-DD): ');
    const endDate = await question('结束日期 (YYYY-MM-DD): ');
    const type = await question('报告类型 (all/asset/event/alarm/vuln/exposed, 默认all): ');
    
    return {
      customer: customer.trim(),
      startDate: startDate.trim(),
      endDate: endDate.trim(),
      type: normalizeReportType(type.trim() || 'all'),
      customerId: '',
      pageSize: CONFIG.defaultPageSize,
      pageSizeProvided: false,
      pageDelayMs: CONFIG.defaultPageDelayMs,
      responseOnly: false,
      cookiePath: CONFIG.cookiePath,
      xdrCookiePath: CONFIG.xdrCookiePath,
      outputDir: CONFIG.outputDir,
      manageSubTypeMapFile: process.env.SANGFOR_MANAGE_SUB_TYPE_MAP_FILE || path.join(__dirname, 'data', 'manage_sub_type_map.json'),
      assetMapFile: process.env.SANGFOR_ASSET_MAP_FILE || path.join(__dirname, 'data', 'asset.xlsx')
    };
  } finally {
    rl.close();
  }
}

function sanitizeFilePart(value) {
  return String(value || 'unknown').replace(/[\\/:*?"<>|]/g, '_');
}

function saveJsonResponse(type, response, options) {
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
  const fileName = `${sanitizeFilePart(options.customer || options.customerId)}_${type}_response_${timestamp}.json`;
  const filePath = path.join(CONFIG.outputDir, fileName);

  fs.writeFileSync(filePath, JSON.stringify(response, null, 2), 'utf8');
  console.log(`[调试] ${type} 原始 response 已保存: ${filePath}`);

  return filePath;
}

async function fetchAndSaveFirstResponse(cookieInfo, requestParams, options) {
  const files = [];
  const targets = options.type === 'all' ? ['asset', 'exposed', 'event', 'alarm', 'vuln'] : [options.type];

  for (const target of targets) {
    if (target === 'asset') {
      console.log(`\n--- 获取 资产表导出 response ---`);
      const response = await requestWithRetry(
        (params) => apiClient.fetchAssetDownloadInfo(cookieInfo, params),
        requestParams
      );
      files.push(saveJsonResponse(target, response, options));
      continue;
    }

    if (target === 'exposed') {
      console.log(`\n--- 获取 暴露面目标公司与导出 response ---`);
      const optionResponse = await requestWithRetry(
        (params) => apiClient.fetchExposedTargetCompanyOptions(cookieInfo, params),
        requestParams
      );
      files.push(saveJsonResponse('exposed_target_company_option', optionResponse, options));

      const exportResponse = await requestWithRetry(
        (params) => apiClient.fetchExposedSurfaceExportInfo(cookieInfo, params),
        requestParams
      );
      files.push(saveJsonResponse('exposed_export', exportResponse, options));

      const ipListStatisticsResponse = await requestWithRetry(
        (params) => apiClient.fetchExposedSurfaceIpListStatistics(cookieInfo, params),
        requestParams
      );
      files.push(saveJsonResponse('exposed_ip_list_statistics_port', { port: ipListStatisticsResponse }, options));
      continue;
    }

    const fetchFunc = target === 'alarm'
      ? apiClient.fetchAlarmTable
      : target === 'vuln'
        ? apiClient.fetchVulnTable
        : apiClient.fetchEventTable;
    const displayName = target === 'alarm' ? '告警表' : target === 'vuln' ? '资产漏洞表' : '事件表';
    console.log(`\n--- 获取 ${displayName} 首页 response ---`);
    const response = await requestWithRetry(
      (params) => fetchFunc(cookieInfo, {
        ...params,
        page: 1,
        pageSize: resolvePageSizeForTarget(target, options)
      }),
      requestParams
    );
    files.push(saveJsonResponse(target, response, options));
  }

  return files;
}

/**
 * 验证日期格式
 */
function validateDate(dateStr) {
  if (!dateStr) return true; // 空日期可能允许
  const regex = /^\d{4}-\d{2}-\d{2}$/;
  if (!regex.test(dateStr)) {
    throw new Error(`日期格式错误: ${dateStr}，正确格式: YYYY-MM-DD`);
  }
  return true;
}

function subtractDays(dateStr, days) {
  if (!dateStr) return dateStr;
  const date = new Date(`${dateStr}T00:00:00`);
  if (Number.isNaN(date.getTime())) return dateStr;
  date.setDate(date.getDate() - days);
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/**
 * 等待用户按回车继续
 */
function waitForContinue() {
  return new Promise((resolve) => {
    const rl = readline.createInterface({ input: process.stdin });
    rl.question('\n按回车键继续...', () => {
      rl.close();
      resolve();
    });
  });
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * 带重试的请求
 */
async function requestWithRetry(requestFunc, options, maxRetries = CONFIG.maxRetries) {
  let lastError;
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      console.log(`[主程序] 第 ${attempt} 次尝试...`);
      return await requestFunc(options);
    } catch (error) {
      lastError = error;
      console.warn(`[主程序] 第 ${attempt} 次尝试失败: ${error.message}`);
      
      if (attempt < maxRetries) {
        console.log(`[主程序] ${CONFIG.retryDelay/1000} 秒后重试...`);
        await new Promise(r => setTimeout(r, CONFIG.retryDelay));
      }
    }
  }
  
  throw lastError;
}

/**
 * 主下载流程
 */
async function downloadReports(options) {
  console.log('\n========================================');
  console.log('   深信服报告下载器');
  console.log('========================================\n');
  
  const startTime = Date.now();
  
  try {
    // 1. 读取 Cookie
    console.log('[步骤 1/5] 读取 Cookie...');
    const cookieInfo = cookieReader.getCookieInfo(options.cookiePath || CONFIG.cookiePath);
    
    if (!cookieInfo.validation.valid) {
      console.warn(`[警告] Cookie 可能不完整，缺失字段: ${cookieInfo.validation.missing.join(', ')}`);
      console.warn('[警告] 继续执行，但如果请求失败请重新登录获取 Cookie');
    }
    
    // 2. 验证日期参数
    console.log('\n[步骤 2/5] 验证参数...');
    validateDate(options.startDate);
    validateDate(options.endDate);
    options.type = normalizeReportType(options.type);
    validateReportType(options.type);
    options.businessSystems = normalizeBusinessSystems(options.businessSystems);
    console.log(`[参数] 客户: ${options.customer || '未指定'}`);
    console.log(`[参数] 客户ID(输入): ${options.customerId || '未指定'}`);
    console.log(`[参数] 时间: ${options.startDate} ~ ${options.endDate}`);
    console.log(`[参数] 类型: ${options.type}`);
    if (options.businessSystems.length > 0) {
      console.log(`[参数] 业务系统: ${options.businessSystems.join(', ')}`);
    }

    let resolvedCustomerId = options.customerId || '';
    if (!resolvedCustomerId) {
      if (!options.customer) {
        throw new Error('缺少客户信息。请至少提供 --customer 中文公司名称，或直接提供 --customer-id。');
      }
      console.log('[步骤 2.5/5] 解析客户ID...');
      resolvedCustomerId = await apiClient.resolveCompanyIdByName(
        cookieInfo,
        options.customer,
        { pageSize: 100, keyword: '', pageDelayMs: options.pageDelayMs }
      );
      console.log(`[参数] 客户ID(解析): ${resolvedCustomerId}`);
    }
    
    // 3. 准备请求参数
    const requestParams = {
      customerId: resolvedCustomerId,
      customerName: options.customer,
      startTime: options.startDate,
      endTime: options.endDate,
      pageSize: options.pageSize,
      pageDelayMs: options.pageDelayMs
    };

    if (options.responseOnly) {
      console.log('\n[调试模式] 只获取第一页 response，不生成 Excel。');
      const files = await fetchAndSaveFirstResponse(cookieInfo, requestParams, options);
      return {
        success: true,
        files: files.map(filePath => ({ type: 'response', path: filePath, name: path.basename(filePath), rowCount: 0 })),
        errors: []
      };
    }
    
    const results = {
      assetWorkbookBuffer: null,
      exposedSurfaceWorkbookBuffer: null,
      eventData: null,
      alarmData: null,
      vulnData: null,
      exposedSurfacePortCount: null,
      weakPwdSummaryTotal: null,
      weakPwdSummaryList: [],
      weakPwdAllSummaryList: [],
      weakPwdAllTotalsByIp: {},
      weakPwdAllTotal: null,
      weakPwdHandledTotal: null,
      weakPwdHandledList: [],
      weakPwdHandledTotalsByIp: {},
      g126: null,
      d129: null,
      d130: null,
      topnReportStats: null,
      xdrLogSearchCount: null,
      xdrHolidayLogSearchCounts: {},
      xdrProtectionQueriedRanges: [],
      xdrIn2outLogSearchCount: null,
      xdrOut2inLogSearchCount: null,
      xdrRejectedExternalToInternalCount: null,
      xdrG99LogSearchCount: null,
      xdrG100AlertCount: null,
      xdrG101IncidentCount: null,
      xdrG102IncidentHandledCount: null,
      xdrG103IncidentCount: null,
      xdrG105IncidentCount: null,
      xdrG106IncidentCount: null,
      xdrG107IncidentCount: null,
      xdrAlertThreatDefineCounts: [],
      xdrIncidentThreatDefineCounts: [],
      xdrMonthlyIn2outLogSearchCounts: [],
      xdrMonthlyOut2inLogSearchCounts: [],
      orderBranchCounts: {},
      assetError: null,
      exposedSurfaceError: null,
      eventError: null,
      alarmError: null,
      vulnError: null,
      weakPwdError: null,
      topnError: null,
      xdrError: null
    };
    
    // 4. 下载数据
    console.log('\n[步骤 3/5] 下载数据...');

    if (shouldFetchType(options.type, 'asset') || shouldFetchType(options.type, 'event') || shouldFetchType(options.type, 'vuln')) {
      console.log('\n--- 下载资产表 ---');
      try {
        const assetFile = await requestWithRetry(
          (params) => apiClient.fetchAssetTableFile(cookieInfo, params),
          requestParams
        );
        results.assetWorkbookBuffer = assetFile.buffer;
        console.log(`[成功] 资产表: 已下载，稍后写入总报告`);
      } catch (error) {
        results.assetError = error.message;
        console.error(`[失败] 资产表: ${error.message}`);
      }
    }

    const runtimeAssetMapFile = results.assetWorkbookBuffer ? null : options.assetMapFile;

    if (shouldFetchType(options.type, 'exposed')) {
      console.log('\n--- 下载暴露面 ---');
      try {
        const exposedSurfaceFile = await requestWithRetry(
          (params) => apiClient.fetchExposedSurfaceFile(cookieInfo, params),
          requestParams
        );
        results.exposedSurfaceWorkbookBuffer = exposedSurfaceFile.buffer;
        console.log(`[成功] 暴露面: 已下载，稍后写入总报告`);

        results.exposedSurfacePortCount = await requestWithRetry(
          (params) => apiClient.fetchExposedSurfaceIpListStatistics(cookieInfo, params),
          requestParams
        );
        console.log(`[成功] 暴露面端口统计: ${results.exposedSurfacePortCount}`);
      } catch (error) {
        results.exposedSurfaceError = error.message;
        console.error(`[失败] 暴露面: ${error.message}`);
      }
    }
    
    if (shouldFetchType(options.type, 'event')) {
      console.log('\n--- 下载事件表 ---');
      try {
        results.eventData = await requestWithRetry(
          (params) => apiClient.fetchAllPages(apiClient.fetchEventTable, cookieInfo, params, 'data'),
          requestParams
        );
        console.log(`[成功] 事件表: ${results.eventData.length} 条记录`);
      } catch (error) {
        results.eventError = error.message;
        console.error(`[失败] 事件表: ${error.message}`);
      }
    }

    
    if (shouldFetchType(options.type, 'alarm')) {
      console.log('\n--- 下载告警表 ---');
      try {
        results.alarmData = await requestWithRetry(
          (params) => apiClient.fetchAllPages(apiClient.fetchAlarmTable, cookieInfo, params, 'data'),
          requestParams
        );
        console.log(`[成功] 告警表: ${results.alarmData.length} 条记录`);
      } catch (error) {
        results.alarmError = error.message;
        console.error(`[失败] 告警表: ${error.message}`);
      }
    }

    // D129 is calculated from second_level entries, so the vulnerability rows
    // must be loaded whenever the protection-period statistic is requested.
    if (shouldFetchType(options.type, 'vuln') || (options.protectStartDate && options.protectEndDate)) {
      console.log('\n--- 下载资产漏洞表 ---');
      try {
        results.vulnData = await requestWithRetry(
          (params) => apiClient.fetchAllPages(apiClient.fetchVulnTable, cookieInfo, {
            ...params,
            pageSize: resolvePageSizeForTarget('vuln', options)
          }, 'data'),
          requestParams
        );
        results.g126 = results.vulnData.filter((row) => (
          String(row && row.name || '').includes('弱口令')
        )).length;
        console.log(`[成功] 资产漏洞表: ${results.vulnData.length} 条记录`);
      } catch (error) {
        results.vulnError = error.message;
        console.error(`[失败] 资产漏洞表: ${error.message}`);
      }
    }
    
    // 5. 生成报告
    console.log('\n[步骤 4/5] 生成 Excel 报告...');
    let dataFormatter;
    let soarTransformer;
    try {
      soarTransformer = require('./soar_transformer');
      dataFormatter = require('./data_formatter');
    } catch (error) {
      if (error.code === 'MODULE_NOT_FOUND' && error.message.includes('xlsx')) {
        throw new Error('缺少 xlsx 依赖。请在当前目录执行 npm install 后重试。');
      }
      throw error;
    }

    console.log('\n--- 查询全部弱口令统计 (G111) ---');
    try {
      const weakPwdAllSummary = await requestWithRetry(
        (params) => apiClient.fetchWeakPwdSummary(cookieInfo, {
          ...params,
          isAdmin: 0,
          pageSize: 100
        }),
        requestParams
      );
      results.weakPwdAllSummaryList = weakPwdAllSummary.list;
      results.weakPwdAllTotalsByIp = weakPwdAllSummary.totalsByIp;
      results.weakPwdAllTotal = weakPwdAllSummary.total;
      console.log(`[成功] 全部弱口令统计: ${results.weakPwdAllTotal}`);
    } catch (error) {
      results.weakPwdError = error.message;
      throw new Error(`全部弱口令统计失败: ${error.message}`);
    }

    console.log('\n--- 查询弱口令统计 ---');
    try {
      const weakPwdSummary = await requestWithRetry(
        (params) => apiClient.fetchWeakPwdSummary(cookieInfo, params),
        requestParams
      );
      results.weakPwdSummaryTotal = weakPwdSummary.total;
      results.weakPwdSummaryList = weakPwdSummary.list;
      console.log(`[成功] 弱口令统计: ${results.weakPwdSummaryTotal}`);
    } catch (error) {
      results.weakPwdError = error.message;
      throw new Error(`弱口令统计失败: ${error.message}`);
    }

    console.log('\n--- 查询已处理弱口令统计 ---');
    try {
      const handledWeakPwdSummary = await requestWithRetry(
        (params) => apiClient.fetchWeakPwdSummary(cookieInfo, {
          ...params,
          isAdmin: 0,
          dealStatus: [2]
        }),
        requestParams
      );
      results.weakPwdHandledTotal = handledWeakPwdSummary.total;
      results.weakPwdHandledList = handledWeakPwdSummary.list;
      results.weakPwdHandledTotalsByIp = handledWeakPwdSummary.totalsByIp;
      console.log(`[成功] 已处理弱口令统计: ${results.weakPwdHandledTotal}`);
    } catch (error) {
      results.weakPwdError = error.message;
      throw new Error(`已处理弱口令统计失败: ${error.message}`);
    }

    console.log('\n--- 查询漏洞端口拆分统计 (D129) ---');
    if (!options.protectStartDate || !options.protectEndDate) {
      results.d129 = '无hw时间无数据';
      console.log(`[跳过] 未传入完整 hw 时间，D129=${results.d129}`);
    } else {
      try {
        const d129RequestParams = {
          ...requestParams,
          startTime: subtractDays(options.protectStartDate, 7),
          endTime: options.protectEndDate
        };
        results.d129 = soarTransformer.countVulnSecondLevelEntries(results.vulnData, {
          startTime: d129RequestParams.startTime,
          endTime: d129RequestParams.endTime,
          statuses: [2, 4]
        });
        console.log(`[成功] 漏洞端口拆分统计: D129=${results.d129}`);
      } catch (error) {
        results.vulnPortSplitError = error.message;
        throw new Error(`漏洞端口拆分统计失败: ${error.message}`);
      }
    }

    console.log('\n--- 查询 hw 时间事件总数 (D130) ---');
    if (!options.protectStartDate || !options.protectEndDate) {
      results.d130 = '无hw时间无数据';
      console.log(`[跳过] 未传入完整 hw 时间，D130=${results.d130}`);
    } else {
      try {
        const hwEventRequestParams = {
          ...requestParams,
          startTime: options.protectStartDate,
          endTime: options.protectEndDate
        };
        const hwEventRows = await requestWithRetry(
          (params) => apiClient.fetchAllPages(apiClient.fetchEventTable, cookieInfo, params, 'data'),
          hwEventRequestParams
        );
        results.d130 = Array.isArray(hwEventRows) ? hwEventRows.length : 0;
        console.log(`[成功] hw 时间事件总数: D130=${results.d130}`);
      } catch (error) {
        results.hwEventError = error.message;
        throw new Error(`hw 时间事件统计失败: ${error.message}`);
      }
    }

    console.log('\n--- 查询订单列表统计 (D100/D101/D102/D105) ---');
    try {
      const d101 = await requestWithRetry(
        (params) => apiClient.fetchOrderBranchTotal(cookieInfo, params, 3),
        requestParams
      );
      const d102Type25 = await requestWithRetry(
        (params) => apiClient.fetchOrderBranchTotal(cookieInfo, params, 25),
        requestParams
      );
      const d102Type999 = await requestWithRetry(
        (params) => apiClient.fetchOrderBranchTotal(cookieInfo, params, 999),
        requestParams
      );
      const d105 = await requestWithRetry(
        (params) => apiClient.fetchOrderBranchTotal(cookieInfo, params, 12),
        requestParams
      );
      results.orderBranchCounts = {
        d101,
        d102Type25,
        d102Type999,
        d102: d102Type25 + d102Type999,
        d100: d101 + d102Type25 + d102Type999,
        d105
      };
      console.log(`[成功] 订单列表统计: D101=${results.orderBranchCounts.d101} (type=3)`);
      console.log(`[成功] 订单列表统计: D102=${results.orderBranchCounts.d102} (type=25=${d102Type25} + type=999=${d102Type999})`);
      console.log(`[成功] 订单列表统计: D100=${results.orderBranchCounts.d100}`);
      console.log(`[成功] 订单列表统计: D105=${results.orderBranchCounts.d105} (type=12)`);
    } catch (error) {
      results.orderBranchError = error.message;
      throw new Error(`订单列表统计失败: ${error.message}`);
    }

    console.log('\n--- 查询 SOAR TopN 统计 ---');
    try {
      results.topnReportStats = await requestWithRetry(
        (params) => apiClient.fetchTopnReportStats(cookieInfo, params),
        requestParams
      );
      const threatCount = Array.isArray(results.topnReportStats.threatTypes)
        ? results.topnReportStats.threatTypes.length
        : 0;
      const geoCount = Array.isArray(results.topnReportStats.srcIpGeos)
        ? results.topnReportStats.srcIpGeos.length
        : 0;
      const dstIpCount = Array.isArray(results.topnReportStats.dstIps)
        ? results.topnReportStats.dstIps.length
        : 0;
      console.log(`[成功] SOAR TopN 统计: 威胁类型 ${threatCount} 条，攻击源地理位置 ${geoCount} 条，目的 IP ${dstIpCount} 条`);
    } catch (error) {
      results.topnError = error.message;
      throw new Error(`SOAR TopN 统计失败: ${error.message}`);
    }

    console.log('\n--- 查询 XDR 重保日志统计 ---');
    try {
      const reportRange = dataFormatter.buildReportSecondRange({
        startDate: options.startDate,
        endDate: options.endDate
      });
      const reportMonthRanges = dataFormatter.buildReportMonthSecondRanges({
        startDate: options.startDate,
        endDate: options.endDate
      });
      const protectionQuery = dataFormatter.buildProtectionAtomicSecondRanges({
        startDate: options.startDate,
        endDate: options.endDate
      });
      const protectionRanges = protectionQuery.ranges;
      console.log(`[XDR] in2out/out2in 查询时间段: ${reportRange.start} ~ ${reportRange.end}`);
      console.log(`[XDR] 月度 in2out/out2in 查询时间段: ${reportMonthRanges.length} 段`);
      console.log(`[XDR] 重保原子查询时间段: ${protectionRanges.length} 段`);
      const xdrCookieInfo = cookieReader.getCookieInfo(options.xdrCookiePath || CONFIG.xdrCookiePath);
      if (options.xdrBaseUrl) {
        xdrCookieInfo.xdrBaseUrl = options.xdrBaseUrl;
      }
      results.xdrIn2outLogSearchCount = await requestWithRetry(
        () => apiClient.fetchXdrIn2outLogSearchCount(xdrCookieInfo, reportRange),
        requestParams
      );
      results.xdrOut2inLogSearchCount = await requestWithRetry(
        () => apiClient.fetchXdrOut2inLogSearchCount(xdrCookieInfo, reportRange),
        requestParams
      );
      results.xdrRejectedExternalToInternalCount = await requestWithRetry(
        () => apiClient.fetchXdrRejectedExternalToInternalCount(xdrCookieInfo, reportRange),
        requestParams
      );
      results.xdrG99LogSearchCount = await requestWithRetry(
        () => apiClient.fetchXdrLogSearchCount(xdrCookieInfo, reportRange),
        requestParams
      );
      results.xdrG101IncidentCount = await requestWithRetry(
        () => apiClient.fetchXdrIncidentTableCount(xdrCookieInfo, reportRange),
        requestParams
      );
      results.xdrG100AlertCount = await requestWithRetry(
        () => apiClient.fetchXdrAlertTableCount(xdrCookieInfo, reportRange),
        requestParams
      );
      results.xdrAlertThreatDefineCounts = await requestWithRetry(
        () => apiClient.fetchXdrAlertThreatDefineCounts(xdrCookieInfo, reportRange),
        requestParams
      );
      results.xdrIncidentThreatDefineCounts = await requestWithRetry(
        () => apiClient.fetchXdrIncidentThreatDefineCounts(xdrCookieInfo, reportRange),
        requestParams
      );
      results.xdrG102IncidentHandledCount = await requestWithRetry(
        () => apiClient.fetchXdrIncidentTableQueryCount(xdrCookieInfo, reportRange),
        requestParams
      );
      results.xdrG103IncidentCount = await requestWithRetry(
        () => apiClient.fetchXdrG103IncidentCount(xdrCookieInfo, reportRange),
        requestParams
      );
      results.xdrG105IncidentCount = await requestWithRetry(
        () => apiClient.fetchXdrG105IncidentCount(xdrCookieInfo, reportRange),
        requestParams
      );
      results.xdrG106IncidentCount = await requestWithRetry(
        () => apiClient.fetchXdrG106IncidentCount(xdrCookieInfo, reportRange),
        requestParams
      );
      results.xdrG107IncidentCount = await requestWithRetry(
        () => apiClient.fetchXdrG107IncidentCount(xdrCookieInfo, reportRange),
        requestParams
      );
      results.xdrMonthlyIn2outLogSearchCounts = [];
      results.xdrMonthlyOut2inLogSearchCounts = [];
      for (let index = 0; index < reportMonthRanges.length; index += 1) {
        const monthRange = reportMonthRanges[index];
        console.log(`[XDR] 月度 in2out/out2in count 查询 ${index + 1}/${reportMonthRanges.length}: ${monthRange.monthLabel} ${monthRange.start} ~ ${monthRange.end}`);
        const in2outCount = await requestWithRetry(
          () => apiClient.fetchXdrIn2outLogSearchCount(xdrCookieInfo, monthRange),
          requestParams
        );
        const out2inCount = await requestWithRetry(
          () => apiClient.fetchXdrOut2inLogSearchCount(xdrCookieInfo, monthRange),
          requestParams
        );
        results.xdrMonthlyIn2outLogSearchCounts.push({
          monthLabel: monthRange.monthLabel,
          count: in2outCount
        });
        results.xdrMonthlyOut2inLogSearchCounts.push({
          monthLabel: monthRange.monthLabel,
          count: out2inCount
        });
        const delayMs = Number.isFinite(options.pageDelayMs) ? options.pageDelayMs : 10;
        if (delayMs > 0 && index < reportMonthRanges.length - 1) {
          await sleep(delayMs);
        }
      }
      if (protectionRanges.length === 0) {
        results.xdrLogSearchCount = 0;
        results.xdrHolidayLogSearchCounts = {};
        results.xdrProtectionQueriedRanges = [];
      } else {
        const xdrProtectionDetails = await requestWithRetry(
          () => apiClient.fetchXdrLogSearchCountDetailsForRanges(xdrCookieInfo, protectionRanges, {
            pageDelayMs: options.pageDelayMs
          }),
          requestParams
        );
        results.xdrLogSearchCount = xdrProtectionDetails.total;
        results.xdrHolidayLogSearchCounts = xdrProtectionDetails.holidayCounts;
        results.xdrProtectionQueriedRanges = xdrProtectionDetails.queriedRanges;
      }
      console.log(`[成功] XDR in2out 日志统计: ${results.xdrIn2outLogSearchCount}`);
      console.log(`[成功] XDR out2in 日志统计: ${results.xdrOut2inLogSearchCount}`);
      console.log(`[成功] XDR 外到内拒绝日志统计: ${results.xdrRejectedExternalToInternalCount}`);
      console.log(`[成功] XDR 报告期日志统计 (G99): ${results.xdrG99LogSearchCount}`);
      console.log(`[成功] XDR 报告期事件表统计 (G101): ${results.xdrG101IncidentCount}`);
      console.log(`[成功] XDR 报告期告警表统计 (G100): ${results.xdrG100AlertCount}`);
      console.log(`[成功] XDR 告警威胁定义统计: ${results.xdrAlertThreatDefineCounts.map(item => `${item.label}:${item.count}`).join(', ')}`);
      console.log(`[成功] XDR 事件威胁定义统计: ${results.xdrIncidentThreatDefineCounts.map(item => `${item.label}:${item.count}`).join(', ')}`);
      console.log(`[成功] XDR 报告期已处置事件表统计 (G102): ${results.xdrG102IncidentHandledCount}`);
      console.log(`[成功] XDR 报告期人工决策已处置非白名单事件表统计 (G103): ${results.xdrG103IncidentCount}`);
      console.log(`[成功] XDR 报告期已处置(非白名单)事件表统计 (G105): ${results.xdrG105IncidentCount}`);
      console.log(`[成功] XDR 报告期处置中/挂起事件表统计 (G106): ${results.xdrG106IncidentCount}`);
      console.log(`[成功] XDR 报告期待处置事件表统计 (G107): ${results.xdrG107IncidentCount}`);
      console.log(`[成功] XDR 月度 in2out 日志统计: ${results.xdrMonthlyIn2outLogSearchCounts.map(item => `${item.monthLabel}:${item.count}`).join(', ')}`);
      console.log(`[成功] XDR 月度 out2in 日志统计: ${results.xdrMonthlyOut2inLogSearchCounts.map(item => `${item.monthLabel}:${item.count}`).join(', ')}`);
      console.log(`[成功] XDR 重保日志统计: ${results.xdrLogSearchCount}`);
      console.log(`[成功] XDR 节假日重保日志统计: ${Object.entries(results.xdrHolidayLogSearchCounts).map(([key, count]) => `${key}:${count}`).join(', ')}`);
    } catch (error) {
      results.xdrError = error.message;
      throw new Error(`XDR 重保日志统计失败: ${error.message}`);
    }

    const transformedEventResult = results.eventData
      ? soarTransformer.transformEventDocsWithStats(results.eventData, {
        manageSubTypeMapFile: options.manageSubTypeMapFile,
        assetMapFile: runtimeAssetMapFile,
        assetWorkbookBuffer: results.assetWorkbookBuffer
      })
      : null;
    const transformedEventData = transformedEventResult ? transformedEventResult.rows : null;
    const transformedAlarmData = results.alarmData
      ? soarTransformer.transformAlarmDocs(results.alarmData, {
        manageSubTypeMapFile: options.manageSubTypeMapFile
      })
      : null;
    const transformedVulnData = results.vulnData
      ? soarTransformer.transformVulnDocs(results.vulnData, {
        assetMapFile: runtimeAssetMapFile,
        assetWorkbookBuffer: results.assetWorkbookBuffer
      })
      : null;

    const reportResult = dataFormatter.generateReport({
      customer: options.customer,
      customerId: resolvedCustomerId,
      startDate: options.startDate,
      endDate: options.endDate,
      protectStartDate: options.protectStartDate,
      protectEndDate: options.protectEndDate,
      assetWorkbookBuffer: results.assetWorkbookBuffer,
      exposedSurfaceWorkbookBuffer: results.exposedSurfaceWorkbookBuffer,
      exposedSurfacePortCount: results.exposedSurfacePortCount,
      weakPwdSummaryTotal: results.weakPwdSummaryTotal,
      weakPwdSummaryList: results.weakPwdSummaryList,
      weakPwdAllSummaryList: results.weakPwdAllSummaryList,
      weakPwdAllTotalsByIp: results.weakPwdAllTotalsByIp,
      weakPwdAllTotal: results.weakPwdAllTotal,
      weakPwdHandledTotal: results.weakPwdHandledTotal,
      weakPwdHandledList: results.weakPwdHandledList,
      weakPwdHandledTotalsByIp: results.weakPwdHandledTotalsByIp,
      g126: results.g126,
      d129: results.d129,
      d130: results.d130,
      topnReportStats: results.topnReportStats,
      reportTemplatePath: options.reportTemplatePath,
      eventData: transformedEventData,
      eventStats: transformedEventResult ? transformedEventResult.stats : null,
      alarmData: transformedAlarmData,
      vulnData: transformedVulnData,
      businessSystems: options.businessSystems,
      vulnRawRowCount: Array.isArray(results.vulnData) ? results.vulnData.length : null,
      xdrLogSearchCount: results.xdrLogSearchCount,
      xdrHolidayLogSearchCounts: results.xdrHolidayLogSearchCounts,
      xdrIn2outLogSearchCount: results.xdrIn2outLogSearchCount,
      xdrOut2inLogSearchCount: results.xdrOut2inLogSearchCount,
      xdrRejectedExternalToInternalCount: results.xdrRejectedExternalToInternalCount,
      xdrG99LogSearchCount: results.xdrG99LogSearchCount,
      xdrG100AlertCount: results.xdrG100AlertCount,
      xdrG101IncidentCount: results.xdrG101IncidentCount,
      xdrG102IncidentHandledCount: results.xdrG102IncidentHandledCount,
      xdrG103IncidentCount: results.xdrG103IncidentCount,
      xdrG105IncidentCount: results.xdrG105IncidentCount,
      xdrG106IncidentCount: results.xdrG106IncidentCount,
      xdrG107IncidentCount: results.xdrG107IncidentCount,
      xdrAlertThreatDefineCounts: results.xdrAlertThreatDefineCounts,
      xdrIncidentThreatDefineCounts: results.xdrIncidentThreatDefineCounts,
      xdrMonthlyIn2outLogSearchCounts: results.xdrMonthlyIn2outLogSearchCounts,
      xdrMonthlyOut2inLogSearchCounts: results.xdrMonthlyOut2inLogSearchCounts,
      orderBranchCounts: results.orderBranchCounts,
      outputDir: options.outputDir || CONFIG.outputDir,
      eventHeaders: soarTransformer.EVENT_OUTPUT_FIELDS,
      alarmHeaders: soarTransformer.ALARM_OUTPUT_FIELDS,
      vulnHeaders: soarTransformer.VULN_OUTPUT_FIELDS,
      eventHeaderDisplayMap: EVENT_HEADER_DISPLAY_MAP,
      alarmHeaderDisplayMap: ALARM_HEADER_DISPLAY_MAP
    });

    // 6. 完成
    console.log('\n[步骤 5/5] 完成！');
    
    const elapsed = ((Date.now() - startTime) / 1000).toFixed(2);
    
    console.log('\n========================================');
    console.log('   执行结果');
    console.log('========================================');
    console.log(`耗时: ${elapsed} 秒`);
    console.log(`成功: ${reportResult.success ? '是' : '否'}`);
    
    if (reportResult.files.length > 0) {
      console.log('\n生成的文件:');
      reportResult.files.forEach((f, i) => {
        console.log(`  ${i + 1}. ${f.name}`);
        console.log(`     路径: ${f.path}`);
        console.log(`     行数: ${f.rowCount}`);
        if (Array.isArray(f.sheets)) {
          f.sheets.forEach(sheet => {
            if (
              sheet.name === '资产漏洞表' &&
              Number.isInteger(sheet.rawRowCount)
            ) {
              console.log(`     - ${sheet.name}: ${sheet.rawRowCount} 条原始记录，${sheet.rowCount} 条 Excel 明细行`);
            } else {
              console.log(`     - ${sheet.name}: ${sheet.rowCount} 行`);
            }
          });
        }
      });
    }
    
    if (reportResult.errors.length > 0) {
      console.log('\n错误:');
      reportResult.errors.forEach((e, i) => {
        console.log(`  ${i + 1}. ${e}`);
      });
    }
    
    // 输出未下载部分的错误
    if (results.assetError || results.exposedSurfaceError || results.eventError || results.alarmError || results.vulnError) {
      console.log('\n下载详情:');
      console.log(`  资产表: ${results.assetWorkbookBuffer ? '成功' : '失败 - ' + results.assetError}`);
      console.log(`  暴露面: ${results.exposedSurfaceWorkbookBuffer ? '成功' : '失败 - ' + results.exposedSurfaceError}`);
      console.log(`  事件表: ${results.eventData ? '成功 (' + results.eventData.length + '条)' : '失败 - ' + results.eventError}`);
      console.log(`  告警表: ${results.alarmData ? '成功 (' + results.alarmData.length + '条)' : '失败 - ' + results.alarmError}`);
      console.log(`  资产漏洞表: ${results.vulnData ? '成功 (' + results.vulnData.length + '条)' : '失败 - ' + results.vulnError}`);
    }
    
    console.log('\n========================================\n');
    
    return reportResult;
    
  } catch (error) {
    console.error('\n[错误] 执行失败:', error.message);
    console.error('[错误] 详细信息:', error.stack);
    throw error;
  }
}

/**
 * 入口函数
 */
async function main() {
  try {
    // 解析参数
    const args = parseArgs();
    
    // 如果没有参数，进入交互模式
    const hasCliArgs = process.argv.slice(2).length > 0;
    const options = hasCliArgs
      ? args
      : await interactiveInput();
    
    // 执行下载
    await downloadReports(options);
    
  } catch (error) {
    console.error('\n[致命错误]', error.message);
    process.exit(1);
  }
}

// 如果是直接运行此脚本
if (require.main === module) {
  main();
}

// 导出供外部调用
module.exports = {
  downloadReports,
  CONFIG,
  parseArgs,
  normalizeBusinessSystems
};
