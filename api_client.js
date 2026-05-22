/**
 * 深信服报告下载 - API 请求封装
 * 处理事件表、告警表、资产漏洞表、资产表和暴露面的接口请求
 */

const https = require('https');
const http = require('http');
const zlib = require('zlib');

// API 配置
const API_CONFIG = {
  baseUrl: 'soar.sangfor.com.cn',
  eventEndpoint: '/gateway/event-mgr/external/event_table',
  alarmEndpoint: '/gateway/alarm-mgr/v1/alarm_table/alarm_list',
  vulnEndpoint: '/gateway/vuln-manager/vm/order/v1/vulnmgr/vuln_list_port_split',
  assetDownloadEndpoint: '/gateway/asset-mgr-service/order/v1/asset/download',
  exposedTargetCompanyOptionEndpoint: '/gateway/vuln-manager/vm/order/v1/vulnmgr/exposed_surface_mss/get_target_company_option',
  exposedExportEndpoint: '/gateway/vuln-manager/vm/order/v1/vulnmgr/exposed_surface_mss/export_result_report',
  companyEndpoint: '/order/v1/user/company_simple_info'
};

/**
 * 将日期字符串转换为时间戳（毫秒）
 * @param {string} dateStr - 日期字符串，格式 YYYY-MM-DD
 * @returns {number} 时间戳
 */
function dateToTimestamp(dateStr) {
  if (!dateStr) return 0;
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(dateStr);
  if (!match) {
    return new Date(dateStr).getTime();
  }

  const [, year, month, day] = match.map(Number);
  return new Date(year, month - 1, day, 0, 0, 0, 0).getTime();
}

/**
 * 构建事件表/告警表的请求体（符合接口规范）
 * @param {Object} params - 请求参数
 * @returns {Object} 请求体对象
 */
function buildRequestBody(params) {
  // 转换日期为时间戳
  const startTimestamp = dateToTimestamp(params.startTime);
  const endTimestamp = dateToTimestamp(params.endTime);
  // 对齐前端抓包：结束时间设为当天 23:59:59.000
  const endOfDay = endTimestamp + 24 * 60 * 60 * 1000 - 1000;

  return {
    order: params.order || {},
    offset: (params.page - 1) * (params.pageSize || 100),
    limit: params.pageSize || 100,
    asset_type: params.asset_type !== undefined ? params.asset_type : -1,
    hw_status: params.hw_status !== undefined ? params.hw_status : -1,
    protection_type: params.protection_type || [],
    company_id: params.customerId ? [params.customerId] : [],
    create_time: [startTimestamp, endOfDay],
    event_name: params.event_name || '',
    event_id: params.event_id || '',
    handler_name: params.handler_name || '',
    task_name: params.task_name || '',
    fuzzy_search: params.fuzzy_search || '',
    event_status: params.event_status || [],
    host_ip: params.host_ip || '',
    hostname: params.hostname || '',
    judgment: params.judgment || [],
    latest_time: params.latest_time || [],
    manage_type: params.manage_type || [],
    more_search: params.more_search || '',
    my_customer: params.my_customer !== undefined ? params.my_customer : 0,
    priority: params.priority || 'all',
    risk_level: params.risk_level || [],
    screen_tag: params.screen_tag || '',
    service_status: params.service_status || [],
    dst_ip: params.dst_ip || '',
    src_ip: params.src_ip || '',
    tag: params.tag || '',
    flow_service_level: params.flow_service_level || [],
    service_group_id_list: params.service_group_id_list || [],
    my_event: params.my_event !== undefined ? params.my_event : 0,
    expired_customer: params.expired_customer !== undefined ? params.expired_customer : 0,
    is_default: params.is_default !== undefined ? params.is_default : 0,
    event_grading_tags: params.event_grading_tags || []
  };
}

/**
 * 生成固定的请求头
 * @param {string} cookieString - Cookie 字符串
 * @param {string} csrfToken - X-Csrftoken
 * @returns {Object} 完整的请求头对象
 */
function generateHeaders(cookieString, csrfToken, overrides = {}) {
  const traceId = generateUUID();
  
  return {
    'host': API_CONFIG.baseUrl,
    'accept': 'application/json, text/javascript, */*; q=0.01',
    'accept-encoding': 'gzip, deflate, br',
    'accept-language': 'zh-CN,zh;q=0.9',
    'cache-control': 'no-cache',
    'content-type': 'application/json',
    'cookie': cookieString,
    'origin': 'https://soar.sangfor.com.cn',
    'pragma': 'no-cache',
    'priority': 'u=1, i',
    'referer': 'https://soar.sangfor.com.cn/index.html',
    'sec-ch-ua': '"Chromium";v="148", "Google Chrome";v="148", "Not/A)Brand";v="99"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-origin',
    'timezone': '+08:00',
    'traceid': traceId,
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36',
    'x-csrftoken': csrfToken,
    'x-requested-with': 'XMLHttpRequest',
    ...overrides
  };
}

function buildAlarmRequestBody(params) {
  const startTimestamp = dateToTimestamp(params.startTime);
  const endTimestamp = dateToTimestamp(params.endTime);
  const endOfDay = endTimestamp + 24 * 60 * 60 * 1000 - 1;

  return {
    order: params.order || { latest_time: 'desc' },
    offset: (params.page - 1) * (params.pageSize || 100),
    limit: params.pageSize || 100,
    event_status: [],
    hw_status: params.hw_status !== undefined ? params.hw_status : -1,
    protection_type: params.protection_type || [],
    service_status: [],
    asset_type: params.asset_type !== undefined ? params.asset_type : -1,
    emergency_degree: params.emergency_degree || 'all',
    risk_level: params.risk_level || 'all',
    sip_sync_status: params.sip_sync_status || 'all',
    certainty_level: params.certainty_level || 'all',
    attack_state: params.attack_state || 'all',
    attack_direction: params.attack_direction !== undefined ? params.attack_direction : -1,
    alarm_source: [],
    source: params.source || [],
    dependability: params.dependability || 'all',
    company_id: params.customerId ? [params.customerId] : [],
    service_group_id_list: params.service_group_id_list || [],
    create_time: [startTimestamp, endOfDay],
    current_operator: params.current_operator || '',
    alarm_name: params.alarm_name || '',
    usecase_id: params.usecase_id || '',
    alarm_id: params.alarm_id || '',
    rule_id: params.rule_id || '',
    event_id: params.event_id || '',
    task_name: params.task_name || '',
    examine_uid: params.examine_uid || '',
    host_ip: params.host_ip || '',
    dst_ip: params.dst_ip || '',
    ioc_value: params.ioc_value || '',
    latest_time: params.latest_time || [],
    manage_type: params.manage_type || [],
    more_search: params.more_search || '',
    my_customer: params.my_customer !== undefined ? params.my_customer : 0,
    principal_filter: params.principal_filter || 'all',
    screen_tag: params.screen_tag || '',
    src_ip: params.src_ip || '',
    tag: params.tag || '',
    hostname: params.hostname || '',
    is_default: params.is_default !== undefined ? params.is_default : 0
  };
}

/**
 * 生成 UUID
 */
function generateUUID() {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
    const r = Math.random() * 16 | 0;
    const v = c === 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
}

/**
 * 发送 HTTP POST 请求
 * @param {string} url - 请求 URL
 * @param {Object} headers - 请求头
 * @param {string} body - 请求体
 * @returns {Promise<Object>} 响应数据
 */
function httpPost(url, headers, body) {
  return new Promise((resolve, reject) => {
    const urlObj = new URL(url);
    const isHttps = urlObj.protocol === 'https:';
    const lib = isHttps ? https : http;
    
    const options = {
      hostname: urlObj.hostname,
      port: urlObj.port || (isHttps ? 443 : 80),
      path: urlObj.pathname + urlObj.search,
      method: 'POST',
      headers: {
        ...headers,
        'content-length': Buffer.byteLength(body)
      },
      timeout: 30000
    };
    
    console.log(`[ApiClient] 发送请求到: ${url}`);
    console.log(`[ApiClient] 请求路径: ${options.path}`);
    
    const req = lib.request(options, (res) => {
      const chunks = [];
      
      res.on('data', (chunk) => {
        chunks.push(chunk);
      });
      
      res.on('end', () => {
        const raw = Buffer.concat(chunks);
        let decoded;

        try {
          decoded = decodeResponseBody(raw, res.headers['content-encoding']);
        } catch (e) {
          reject(new Error(`响应解压失败: ${e.message}`));
          return;
        }

        const data = decoded.toString('utf8');
        console.log(`[ApiClient] 响应状态: ${res.statusCode}`);
        console.log(`[ApiClient] 响应数据长度: ${data.length} 字节`);
        
        if (res.statusCode !== 200) {
          reject(new Error(`HTTP ${res.statusCode}: ${data.substring(0, 500)}`));
          return;
        }
        
        try {
          const jsonData = JSON.parse(data);
          resolve(jsonData);
        } catch (e) {
          reject(new Error(`JSON 解析失败: ${e.message}\n原始数据: ${data.substring(0, 500)}`));
        }
      });
    });
    
 req.on('error', (e) => {
      console.error(`[ApiClient] 请求错误: ${e.message}`);
      reject(e);
    });
    
    req.on('timeout', () => {
      req.destroy();
      reject(new Error('请求超时 (30秒)'));
    });
    
    req.write(body);
    req.end();
  });
}

function httpGetBuffer(url, headers = {}, redirectCount = 0) {
  return new Promise((resolve, reject) => {
    const urlObj = new URL(url);
    const isHttps = urlObj.protocol === 'https:';
    const lib = isHttps ? https : http;

    const options = {
      hostname: urlObj.hostname,
      port: urlObj.port || (isHttps ? 443 : 80),
      path: urlObj.pathname + urlObj.search,
      method: 'GET',
      headers,
      timeout: 30000
    };

    console.log(`[ApiClient] 发送 GET 请求到: ${url}`);
    console.log(`[ApiClient] 请求路径: ${options.path}`);

    const req = lib.request(options, (res) => {
      const location = res.headers.location;
      if (res.statusCode >= 300 && res.statusCode < 400 && location) {
        if (redirectCount >= 5) {
          reject(new Error(`GET 重定向次数过多: ${url}`));
          return;
        }
        const nextUrl = new URL(location, url).toString();
        resolve(httpGetBuffer(nextUrl, headers, redirectCount + 1));
        return;
      }

      const chunks = [];
      res.on('data', chunk => chunks.push(chunk));
      res.on('end', () => {
        let decoded;
        try {
          decoded = decodeResponseBody(Buffer.concat(chunks), res.headers['content-encoding']);
        } catch (e) {
          reject(new Error(`响应解压失败: ${e.message}`));
          return;
        }

        console.log(`[ApiClient] GET 响应状态: ${res.statusCode}`);
        console.log(`[ApiClient] GET 响应数据长度: ${decoded.length} 字节`);

        if (res.statusCode !== 200) {
          reject(new Error(`HTTP ${res.statusCode}: ${decoded.toString('utf8').substring(0, 500)}`));
          return;
        }

        resolve({
          buffer: decoded,
          headers: res.headers,
          statusCode: res.statusCode
        });
      });
    });

    req.on('error', (e) => {
      console.error(`[ApiClient] GET 请求错误: ${e.message}`);
      reject(e);
    });

    req.on('timeout', () => {
      req.destroy();
      reject(new Error('GET 请求超时 (30秒)'));
    });

    req.end();
  });
}

function readUInt16LE(buffer, offset) {
  if (offset + 2 > buffer.length) {
    throw new Error('ZIP 文件结构不完整');
  }
  return buffer.readUInt16LE(offset);
}

function readUInt32LE(buffer, offset) {
  if (offset + 4 > buffer.length) {
    throw new Error('ZIP 文件结构不完整');
  }
  return buffer.readUInt32LE(offset);
}

function findEndOfCentralDirectory(buffer) {
  const minOffset = Math.max(0, buffer.length - 0xffff - 22);
  for (let offset = buffer.length - 22; offset >= minOffset; offset--) {
    if (readUInt32LE(buffer, offset) === 0x06054b50) {
      return offset;
    }
  }
  return -1;
}

function listZipEntries(buffer) {
  if (!Buffer.isBuffer(buffer) || buffer.length < 22) {
    return [];
  }

  const eocdOffset = findEndOfCentralDirectory(buffer);
  if (eocdOffset < 0) {
    return [];
  }

  const entryCount = readUInt16LE(buffer, eocdOffset + 10);
  const centralDirOffset = readUInt32LE(buffer, eocdOffset + 16);
  const entries = [];
  let offset = centralDirOffset;

  for (let i = 0; i < entryCount; i++) {
    if (offset + 46 > buffer.length || readUInt32LE(buffer, offset) !== 0x02014b50) {
      throw new Error('ZIP 中央目录结构异常');
    }

    const flags = readUInt16LE(buffer, offset + 8);
    const compressionMethod = readUInt16LE(buffer, offset + 10);
    const compressedSize = readUInt32LE(buffer, offset + 20);
    const uncompressedSize = readUInt32LE(buffer, offset + 24);
    const fileNameLength = readUInt16LE(buffer, offset + 28);
    const extraLength = readUInt16LE(buffer, offset + 30);
    const commentLength = readUInt16LE(buffer, offset + 32);
    const localHeaderOffset = readUInt32LE(buffer, offset + 42);
    const nameStart = offset + 46;
    const nameEnd = nameStart + fileNameLength;
    const encoding = (flags & 0x0800) ? 'utf8' : 'utf8';
    const fileName = buffer.slice(nameStart, nameEnd).toString(encoding);

    entries.push({
      fileName,
      compressionMethod,
      compressedSize,
      uncompressedSize,
      localHeaderOffset
    });

    offset = nameEnd + extraLength + commentLength;
  }

  return entries;
}

function extractZipEntry(buffer, entry) {
  const offset = entry.localHeaderOffset;
  if (offset + 30 > buffer.length || readUInt32LE(buffer, offset) !== 0x04034b50) {
    throw new Error(`ZIP 本地文件头异常: ${entry.fileName}`);
  }

  const fileNameLength = readUInt16LE(buffer, offset + 26);
  const extraLength = readUInt16LE(buffer, offset + 28);
  const dataStart = offset + 30 + fileNameLength + extraLength;
  const dataEnd = dataStart + entry.compressedSize;
  if (dataEnd > buffer.length) {
    throw new Error(`ZIP 文件数据不完整: ${entry.fileName}`);
  }

  const compressedData = buffer.slice(dataStart, dataEnd);
  if (entry.compressionMethod === 0) {
    return compressedData;
  }
  if (entry.compressionMethod === 8) {
    return zlib.inflateRawSync(compressedData);
  }

  throw new Error(`不支持的 ZIP 压缩方式 ${entry.compressionMethod}: ${entry.fileName}`);
}

function extractXlsxFromZipOrSelf(buffer) {
  const entries = listZipEntries(buffer);
  if (!entries.length) {
    throw new Error('下载文件不是有效的 zip/xlsx 文件');
  }

  if (entries.some(entry => entry.fileName === '[Content_Types].xml')) {
    return buffer;
  }

  const xlsxEntry = entries.find(entry => /\.xlsx$/i.test(entry.fileName) && !entry.fileName.endsWith('/'));
  if (!xlsxEntry) {
    throw new Error(`zip 中未找到 xlsx 文件，包含文件: ${entries.map(entry => entry.fileName).join(', ')}`);
  }

  console.log(`[ApiClient] 从暴露面 zip 提取 xlsx: ${xlsxEntry.fileName}`);
  const xlsxBuffer = extractZipEntry(buffer, xlsxEntry);
  const nestedEntries = listZipEntries(xlsxBuffer);
  if (!nestedEntries.some(entry => entry.fileName === '[Content_Types].xml')) {
    throw new Error(`提取出的文件不是有效 xlsx: ${xlsxEntry.fileName}`);
  }
  return xlsxBuffer;
}

function buildVulnRequestBody(params) {
  const pageSize = params.pageSize || 100;
  const page = params.page || 1;

  return {
    order: {},
    offset: (page - 1) * pageSize,
    limit: pageSize,
    keyword: '',
    vulnerability_status: [],
    dev_id: [],
    fix_level: -1,
    src_storage: [],
    src_type: ['tss'],
    is_intranet: -1,
    service_status: 0,
    vuln_type: -1,
    is_high_availability: -1,
    protection_rule: [],
    asset_list: [],
    scene_tag: [],
    scan_method: -1,
    found_time: [],
    last_time: [],
    vulnerability_level: -1,
    company_id: String(params.customerId || '')
  };
}

function buildAssetDownloadRequestBody(params) {
  return {
    asset_list: [],
    exclude_asset_list: [],
    is_select_all: 1,
    order: 'asc',
    service_status: [],
    is_alive: -1,
    ip_url_keyword: '',
    asset_type: [],
    business_level: [],
    first_time: [],
    update_time: [],
    keyword: '',
    asset_tag: [],
    agent_status: [],
    af_defend_status: '',
    database: [],
    middleware: [],
    os: [],
    developing_languages: [],
    development_framework: [],
    server_port: [],
    company_id: String(params.customerId || ''),
    asset_group_id: 'all',
    filter_items: {
      business_name: 1,
      hostname: 1,
      business_level: 1,
      asset_group_name: 1,
      asset_type: 1,
      asset_tag: 1,
      is_alive: 1,
      adapter: 1,
      authorization_type: 1,
      is_service: 1,
      branch: 1,
      is_auto_scan: 0,
      os: 1,
      af_defend_status: 1,
      agent_status: 1,
      first_time: 1,
      update_time: 1,
      owner_name: 0,
      app_type: 0,
      app_brand: 0,
      server_port: 0,
      public_ip: 0,
      public_port: 0,
      security_domain: 1,
      database: 0,
      middleware: 0,
      developing_languages: 0,
      development_framework: 0,
      owner_department: 0,
      asset_owner_name: 0,
      agent_id: 0,
      proof_status: 0,
      task_name: 0
    }
  };
}

function buildExposedTargetCompanyOptionRequestBody(params) {
  return {
    company_id: String(params.customerId || '')
  };
}

function buildExposedExportRequestBody(params) {
  return {
    ops: 1,
    asset_tag: [2, 5, 3],
    target_company: Array.isArray(params.targetCompanyIds) ? params.targetCompanyIds : [],
    company_id: String(params.customerId || '')
  };
}

function buildCustomerBusinessUrl(params) {
  const companyName = encodeURIComponent(params.customerName || '');
  return `https://${API_CONFIG.baseUrl}/index.html#/customer/${params.customerId}/business?company_name=${companyName}`;
}

function buildHomeUrl() {
  return `https://${API_CONFIG.baseUrl}/index.html`;
}

function decodeResponseBody(buffer, encoding) {
  const normalized = String(encoding || '').toLowerCase();

  if (normalized.includes('gzip')) {
    return zlib.gunzipSync(buffer);
  }

  if (normalized.includes('deflate')) {
    return zlib.inflateSync(buffer);
  }

  if (normalized.includes('br')) {
    return zlib.brotliDecompressSync(buffer);
  }

  return buffer;
}

function extractRowsFromResponse(response, dataKey = 'data') {
  if (Array.isArray(response)) {
    return response;
  }

  if (!response || typeof response !== 'object') {
    return [];
  }

  if (Array.isArray(response[dataKey])) {
    return response[dataKey];
  }

  if (response[dataKey] && Array.isArray(response[dataKey].list)) {
    return response[dataKey].list;
  }

  if (response[dataKey] && Array.isArray(response[dataKey].rows)) {
    return response[dataKey].rows;
  }

  if (response[dataKey] && Array.isArray(response[dataKey].records)) {
    return response[dataKey].records;
  }

  if (Array.isArray(response.list)) {
    return response.list;
  }

  if (Array.isArray(response.rows)) {
    return response.rows;
  }

  if (Array.isArray(response.records)) {
    return response.records;
  }

  return [];
}

function getTotalFromResponse(response, dataKey = 'data') {
  if (!response || typeof response !== 'object') {
    return null;
  }

  const data = response[dataKey] && typeof response[dataKey] === 'object' ? response[dataKey] : {};
  const candidates = [
    response.total,
    response.count,
    response.total_count,
    response.totalCount,
    data.total,
    data.count,
    data.total_count,
    data.totalCount
  ];

  for (const value of candidates) {
    const numberValue = Number(value);
    if (Number.isFinite(numberValue) && numberValue >= 0) {
      return numberValue;
    }
  }

  return null;
}

function getHasMoreFromResponse(response, rows, allDataLength, pageSize, dataKey = 'data') {
  if (response && typeof response === 'object') {
    const data = response[dataKey] && typeof response[dataKey] === 'object' ? response[dataKey] : {};

    if (typeof response.hasMore === 'boolean') return response.hasMore;
    if (typeof response.has_more === 'boolean') return response.has_more;
    if (typeof data.hasMore === 'boolean') return data.hasMore;
    if (typeof data.has_more === 'boolean') return data.has_more;

    // Some SOAR table APIs cap or otherwise misreport total, so do not use it
    // as a pagination boundary. Continue until the server returns a short page.
  }

  return rows.length >= pageSize;
}

function buildPageFingerprint(rows) {
  if (!Array.isArray(rows) || rows.length === 0) {
    return 'empty';
  }

  return JSON.stringify({
    length: rows.length,
    first: rows[0],
    last: rows[rows.length - 1]
  });
}

/**
 * 请求事件表数据
 * @param {Object} cookieInfo - Cookie 信息 { cookieString, csrfToken }
 * @param {Object} params - 请求参数 { customerId, startTime, endTime, page, pageSize }
 */
async function fetchEventTable(cookieInfo, params) {
  const { cookieString, csrfToken } = cookieInfo;
  
  const requestBody = buildRequestBody(params);
  
  const headers = generateHeaders(cookieString, csrfToken);
  
  const url = `https://${API_CONFIG.baseUrl}${API_CONFIG.eventEndpoint}`;
  
  console.log(`[ApiClient] 请求事件表参数:`, JSON.stringify(requestBody, null, 2));
  
  return httpPost(url, headers, JSON.stringify(requestBody));
}

/**
 * 请求告警表数据
 * @param {Object} cookieInfo - Cookie 信息
 * @param {Object} params - 请求参数
 */
async function fetchAlarmTable(cookieInfo, params) {
  const { cookieString, csrfToken } = cookieInfo;
  
  const requestBody = buildAlarmRequestBody(params);
  
  const headers = generateHeaders(cookieString, csrfToken);
  
  const url = `https://${API_CONFIG.baseUrl}${API_CONFIG.alarmEndpoint}`;
  
  console.log(`[ApiClient] 请求告警表参数:`, JSON.stringify(requestBody, null, 2));
  
  return httpPost(url, headers, JSON.stringify(requestBody));
}

/**
 * 请求资产漏洞表数据
 * @param {Object} cookieInfo - Cookie 信息
 * @param {Object} params - 请求参数
 */
async function fetchVulnTable(cookieInfo, params) {
  const { cookieString, csrfToken } = cookieInfo;

  const requestBody = buildVulnRequestBody(params);

  const headers = generateHeaders(cookieString, csrfToken);

  const url = `https://${API_CONFIG.baseUrl}${API_CONFIG.vulnEndpoint}`;

  console.log(`[ApiClient] 请求资产漏洞表参数:`, JSON.stringify(requestBody, null, 2));

  return httpPost(url, headers, JSON.stringify(requestBody));
}

async function visitHomePage(cookieInfo) {
  const { cookieString, csrfToken } = cookieInfo;
  const headers = generateHeaders(cookieString, csrfToken, {
    accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    referer: buildHomeUrl(),
    'sec-fetch-dest': 'document',
    'sec-fetch-mode': 'navigate',
    'sec-fetch-site': 'same-origin'
  });
  delete headers['content-type'];
  delete headers['x-requested-with'];

  return httpGetBuffer(buildHomeUrl(), headers);
}

function extractTargetCompanyList(response) {
  const data = response && typeof response === 'object' ? response.data : null;
  if (data && Array.isArray(data.target_company_list)) return data.target_company_list;
  if (Array.isArray(response && response.target_company_list)) return response.target_company_list;
  return [];
}

function normalizeCompanyText(value) {
  return String(value || '').trim();
}

function resolveTargetCompanyId(optionResponse, targetName = '') {
  const rows = extractTargetCompanyList(optionResponse);
  if (!rows.length) {
    throw new Error(`暴露面目标公司列表为空: ${JSON.stringify(optionResponse).substring(0, 500)}`);
  }

  const normalizedTarget = normalizeCompanyText(targetName);
  if (normalizedTarget) {
    const matched = rows.find(row => normalizeCompanyText(row.target_company) === normalizedTarget);
    if (matched && matched.target_company_id) {
      return String(matched.target_company_id);
    }
  }

  if (rows.length === 1 && rows[0].target_company_id) {
    return String(rows[0].target_company_id);
  }

  const available = rows
    .map(row => `${row.target_company || ''}(${row.target_company_id || ''})`)
    .join(', ');
  throw new Error(`无法唯一确定暴露面 target_company_id，目标客户: ${targetName || '未指定'}，可选项: ${available}`);
}

async function fetchExposedTargetCompanyOptions(cookieInfo, params) {
  const { cookieString, csrfToken } = cookieInfo;
  await visitHomePage(cookieInfo);

  const requestBody = buildExposedTargetCompanyOptionRequestBody(params);
  const headers = generateHeaders(cookieString, csrfToken, {
    referer: buildHomeUrl()
  });
  const url = `https://${API_CONFIG.baseUrl}${API_CONFIG.exposedTargetCompanyOptionEndpoint}`;

  console.log(`[ApiClient] 请求暴露面目标公司参数:`, JSON.stringify(requestBody, null, 2));

  return httpPost(url, headers, JSON.stringify(requestBody));
}

async function fetchExposedSurfaceExportInfo(cookieInfo, params) {
  const { cookieString, csrfToken } = cookieInfo;
  const optionResponse = params.targetCompanyIds && params.targetCompanyIds.length
    ? null
    : await fetchExposedTargetCompanyOptions(cookieInfo, params);
  const targetCompanyIds = params.targetCompanyIds && params.targetCompanyIds.length
    ? params.targetCompanyIds
    : [resolveTargetCompanyId(optionResponse, params.customerName)];

  const requestBody = buildExposedExportRequestBody({
    ...params,
    targetCompanyIds
  });
  const headers = generateHeaders(cookieString, csrfToken, {
    referer: buildHomeUrl()
  });
  const url = `https://${API_CONFIG.baseUrl}${API_CONFIG.exposedExportEndpoint}`;

  console.log(`[ApiClient] 请求暴露面导出参数:`, JSON.stringify(requestBody, null, 2));

  return httpPost(url, headers, JSON.stringify(requestBody));
}

async function downloadExposedSurfaceFile(cookieInfo, exportResponse) {
  const { cookieString, csrfToken } = cookieInfo;
  const relativeUrl = (((exportResponse || {}).data || {}).url) || exportResponse.url;
  if (!relativeUrl) {
    throw new Error(`暴露面导出响应缺少下载地址: ${JSON.stringify(exportResponse).substring(0, 500)}`);
  }

  const downloadUrl = new URL(relativeUrl, `https://${API_CONFIG.baseUrl}`).toString();
  const headers = generateHeaders(cookieString, csrfToken, {
    accept: 'application/zip,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*',
    referer: buildHomeUrl()
  });
  delete headers['content-type'];

  const downloaded = await httpGetBuffer(downloadUrl, headers);
  return {
    ...downloaded,
    buffer: extractXlsxFromZipOrSelf(downloaded.buffer)
  };
}

async function fetchExposedSurfaceFile(cookieInfo, params) {
  const exportInfo = await fetchExposedSurfaceExportInfo(cookieInfo, params);
  return downloadExposedSurfaceFile(cookieInfo, exportInfo);
}

async function visitCustomerBusinessPage(cookieInfo, params) {
  const { cookieString, csrfToken } = cookieInfo;
  const businessUrl = buildCustomerBusinessUrl(params);
  const headers = generateHeaders(cookieString, csrfToken, {
    accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    referer: 'https://soar.sangfor.com.cn/index.html',
    'sec-fetch-dest': 'document',
    'sec-fetch-mode': 'navigate',
    'sec-fetch-site': 'same-origin'
  });
  delete headers['content-type'];
  delete headers['x-requested-with'];

  return httpGetBuffer(businessUrl, headers);
}

async function fetchAssetDownloadInfo(cookieInfo, params) {
  const { cookieString, csrfToken } = cookieInfo;
  const businessUrl = buildCustomerBusinessUrl(params);

  await visitCustomerBusinessPage(cookieInfo, params);

  const requestBody = buildAssetDownloadRequestBody(params);
  const headers = generateHeaders(cookieString, csrfToken, {
    referer: businessUrl
  });
  const url = `https://${API_CONFIG.baseUrl}${API_CONFIG.assetDownloadEndpoint}`;

  console.log(`[ApiClient] 请求资产表导出参数:`, JSON.stringify(requestBody, null, 2));

  return httpPost(url, headers, JSON.stringify(requestBody));
}

async function downloadAssetFile(cookieInfo, assetDownloadResponse, params) {
  const { cookieString, csrfToken } = cookieInfo;
  const relativeUrl = (((assetDownloadResponse || {}).data || {}).url) || assetDownloadResponse.url;
  if (!relativeUrl) {
    throw new Error(`资产表导出响应缺少下载地址: ${JSON.stringify(assetDownloadResponse).substring(0, 500)}`);
  }

  const downloadUrl = new URL(relativeUrl, `https://${API_CONFIG.baseUrl}`).toString();
  const businessUrl = buildCustomerBusinessUrl(params);
  const headers = generateHeaders(cookieString, csrfToken, {
    accept: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,application/octet-stream,*/*',
    referer: businessUrl
  });
  delete headers['content-type'];

  return httpGetBuffer(downloadUrl, headers);
}

async function fetchAssetTableFile(cookieInfo, params) {
  const downloadInfo = await fetchAssetDownloadInfo(cookieInfo, params);
  return downloadAssetFile(cookieInfo, downloadInfo, params);
}

function buildCompanyRequestBody(params) {
  return {
    id_ignore: true,
    service_status: 1,
    offset: params.offset || 0,
    limit: params.limit || 100,
    my_customer: 0,
    my_customer_first_handler: 0,
    keyword: params.keyword || ''
  };
}

function readValueByCandidates(obj, keys) {
  for (const key of keys) {
    if (obj[key] !== undefined && obj[key] !== null && String(obj[key]).trim() !== '') {
      return obj[key];
    }
  }
  return null;
}

function extractCompanyRows(result) {
  if (!result || typeof result !== 'object') return [];
  if (Array.isArray(result.data)) return result.data;
  if (result.data && Array.isArray(result.data.list)) return result.data.list;
  if (result.data && Array.isArray(result.data.rows)) return result.data.rows;
  if (Array.isArray(result.list)) return result.list;
  return [];
}

function matchCompanyName(row, targetName) {
  const target = String(targetName || '').trim();
  if (!target) return false;
  const nameCandidates = [
    'company_name',
    'name',
    'customer_name',
    'companyName',
    'company_simple_name',
    'simple_name',
    'label'
  ];
  for (const key of nameCandidates) {
    const value = row[key];
    if (value === undefined || value === null) continue;
    const text = String(value).trim();
    if (!text) continue;
    if (text === target) return true;
  }
  return false;
}

function extractCompanyId(row) {
  const id = readValueByCandidates(row, ['company_id', 'id', 'companyId', 'value']);
  if (id === null || id === undefined) return '';
  return String(id).trim();
}

async function fetchCompanyPage(cookieInfo, params = {}) {
  const { cookieString, csrfToken } = cookieInfo;
  const headers = generateHeaders(cookieString, csrfToken);
  const url = `https://${API_CONFIG.baseUrl}${API_CONFIG.companyEndpoint}`;
  const body = buildCompanyRequestBody(params);
  return httpPost(url, headers, JSON.stringify(body));
}

async function resolveCompanyIdByName(cookieInfo, companyName, options = {}) {
  const pageSize = options.pageSize || 100;
  const keyword = options.keyword || '';
  const pageDelayMs = Number.isFinite(options.pageDelayMs) ? options.pageDelayMs : 10;
  let offset = 0;
  let page = 1;
  let previousFingerprint = null;

  while (true) {
    console.log(`[ApiClient] 查询客户ID，第 ${page} 页，offset=${offset}`);
    const result = await fetchCompanyPage(cookieInfo, { offset, limit: pageSize, keyword });
    const rows = extractCompanyRows(result);
    if (!rows.length) break;

    const currentFingerprint = buildPageFingerprint(rows);
    if (currentFingerprint === previousFingerprint) {
      throw new Error(`客户列表分页疑似未生效，offset=${offset} 返回了重复数据，已中止以避免死循环。`);
    }
    previousFingerprint = currentFingerprint;

    for (const row of rows) {
      if (matchCompanyName(row, companyName)) {
        const companyId = extractCompanyId(row);
        if (companyId) {
          return companyId;
        }
      }
    }

    if (rows.length < pageSize) break;
    offset += pageSize;
    page++;
    if (pageDelayMs > 0) {
      await sleep(pageDelayMs);
    }
  }

  throw new Error(`未找到客户名称对应的 customer_id: ${companyName}`);
}

/**
 * 获取所有分页数据
 * @param {Function} fetchFunc - 单页请求函数
 * @param {Object} cookieInfo - Cookie 信息
 * @param {Object} params - 请求参数
 * @param {string} dataKey - 返回数据中的数据键名
 */
async function fetchAllPages(fetchFunc, cookieInfo, params, dataKey = 'data') {
  const allData = [];
  let page = 1;
  let hasMore = true;
  let previousFingerprint = null;
  
  console.log(`[ApiClient] 开始分页获取数据...`);
  
  while (hasMore) {
    console.log(`[ApiClient] 获取第 ${page} 页...`);
    
    try {
      const result = await fetchFunc(cookieInfo, { ...params, page });
      
      const rows = extractRowsFromResponse(result, dataKey);

      if (rows.length > 0) {
        const currentFingerprint = buildPageFingerprint(rows);
        if (currentFingerprint === previousFingerprint) {
          throw new Error(`分页疑似未生效，第 ${page} 页返回了与上一页相同的数据，已中止以避免死循环。`);
        }
        previousFingerprint = currentFingerprint;

        allData.push(...rows);
        console.log(`[ApiClient] 第 ${page} 页获取到 ${rows.length} 条数据，累计 ${allData.length} 条`);
        hasMore = getHasMoreFromResponse(result, rows, allData.length, params.pageSize || 100, dataKey);
      } else {
        console.log(`[ApiClient] 第 ${page} 页没有可导出的列表数据。响应摘要:`, JSON.stringify(result).substring(0, 300));
        hasMore = false;
      }
      
      page++;
      
      const pageDelayMs = Number.isFinite(params.pageDelayMs) ? params.pageDelayMs : 10;
      if (pageDelayMs > 0) {
        await sleep(pageDelayMs);
      }
    } catch (error) {
      console.error(`[ApiClient] 第 ${page} 页请求失败: ${error.message}`);
      throw error;
    }
  }
  
  console.log(`[ApiClient] 分页获取完成，共 ${allData.length} 条数据`);
  return allData;
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

module.exports = {
  API_CONFIG,
  buildRequestBody,
  buildAlarmRequestBody,
  buildVulnRequestBody,
  buildAssetDownloadRequestBody,
  buildExposedTargetCompanyOptionRequestBody,
  buildExposedExportRequestBody,
  buildCustomerBusinessUrl,
  buildHomeUrl,
  httpPost,
  httpGetBuffer,
  extractXlsxFromZipOrSelf,
  extractRowsFromResponse,
  getTotalFromResponse,
  generateHeaders,
  fetchEventTable,
  fetchAlarmTable,
  fetchVulnTable,
  visitHomePage,
  fetchExposedTargetCompanyOptions,
  fetchExposedSurfaceExportInfo,
  downloadExposedSurfaceFile,
  fetchExposedSurfaceFile,
  visitCustomerBusinessPage,
  fetchAssetDownloadInfo,
  downloadAssetFile,
  fetchAssetTableFile,
  fetchAllPages,
  fetchCompanyPage,
  resolveCompanyIdByName
};
