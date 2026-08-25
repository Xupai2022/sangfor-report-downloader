/**
 * 深信服报告下载 - API 请求封装
 * 处理事件表、告警表、资产漏洞表、资产表和暴露面的接口请求
 */

const https = require('https');
const http = require('http');
const zlib = require('zlib');

const REQUEST_TIMEOUT_MS = 75000;
const REQUEST_TIMEOUT_SECONDS = REQUEST_TIMEOUT_MS / 1000;

function normalizeHost(value) {
  const raw = String(value || '').trim();
  if (!raw) {
    return '';
  }

  const withScheme = raw.includes('://') ? raw : `https://${raw}`;
  try {
    return new URL(withScheme).hostname;
  } catch (error) {
    return raw.replace(/^https?:\/\//i, '').replace(/\/.*$/, '');
  }
}

// API 配置
const API_CONFIG = {
  baseUrl: 'soar.sangfor.com.cn',
  xdrBaseUrl: normalizeHost(process.env.SANGFOR_XDR_BASE_URL || 'xdr.sangfor.com.cn'),
  eventEndpoint: '/gateway/event-mgr/external/event_table',
  alarmEndpoint: '/gateway/alarm-mgr/v1/alarm_table/alarm_list',
  vulnEndpoint: '/gateway/vuln-manager/vm/order/v1/vulnmgr/vuln_list_port_split',
  assetDownloadEndpoint: '/gateway/asset-mgr-service/order/v1/asset/download',
  exposedTargetCompanyOptionEndpoint: '/gateway/vuln-manager/vm/order/v1/vulnmgr/exposed_surface_mss/get_target_company_option',
  exposedExportEndpoint: '/gateway/vuln-manager/vm/order/v1/vulnmgr/exposed_surface_mss/export_result_report',
  exposedIpListStatisticsEndpoint: '/gateway/vuln-manager/vm/order/v1/vulnmgr/exposed_surface_mss/ip_list_statistics',
  weakPwdSummaryEndpoint: '/gateway/vuln-manager/vm/order/v1/weak_pwd/summary_list',
  weakPwdListEndpoint: '/gateway/vuln-manager/vm/order/v1/weak_pwd/list',
  topnLoadConditionEndpoint: '/order/v1/tool_box/topn/load_condition',
  topnDeviceListEndpoint: '/order/v1/tool_box/topn/device_list',
  topnThreatTypeEndpoint: '/order/v1/tool_box/topn/threat_type',
  topnSrcIpGeoEndpoint: '/order/v1/tool_box/topn/src_ip_geo',
  topnDstIpEndpoint: '/order/v1/tool_box/topn/dst_ip',
  companyEndpoint: '/order/v1/user/company_simple_info',
  xdrLogSearchCountEndpoint: '/api/apex/logsearch/v1/log/search/count?enableCache=true&viewRegionId=ffffffffffffffffffffffff&onlySelfPlatform=false',
  xdrIncidentAnalysisCountEndpoint: '/ngsoc/INCIDENT/api/v1/table/count/analysisTableQueryHandler?viewRegionId=ffffffffffffffffffffffff&onlySelfPlatform=false',
  xdrIncidentTableCountEndpoint: '/ngsoc/INCIDENT/api/v1/table/count/incidentTableQueryHandler?viewRegionId=ffffffffffffffffffffffff&onlySelfPlatform=false',
  xdrIncidentTableQueryEndpoint: '/ngsoc/INCIDENT/api/v1/table/query/incidentTableQueryHandler?viewRegionId=ffffffffffffffffffffffff&onlySelfPlatform=false',
  xdrAlertTableQueryEndpoint: '/ngsoc/INCIDENT/api/v1/table/query/alertTableQueryHandler?viewRegionId=ffffffffffffffffffffffff&onlySelfPlatform=false',
  orderBranchEndpoint: '/gateway/customer-mgr-service/order/v1/branch/dev?_method=GET'
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

function getXdrBaseUrl(cookieInfo = {}) {
  return normalizeHost(cookieInfo.xdrBaseUrl || API_CONFIG.xdrBaseUrl);
}

function generateXdrHeaders(cookieString, csrfToken, overrides = {}, baseUrl = API_CONFIG.xdrBaseUrl) {
  const xdrBaseUrl = normalizeHost(baseUrl || API_CONFIG.xdrBaseUrl);

  return {
    'host': xdrBaseUrl,
    'accept': 'application/json, text/plain, */*',
    'accept-encoding': 'gzip, deflate, br',
    'accept-language': 'zh-CN,zh;q=0.9',
    'connection': 'keep-alive',
    'content-type': 'application/json',
    'cookie': cookieString,
    'origin': `https://${xdrBaseUrl}`,
    'referer': `https://${xdrBaseUrl}/`,
    'sec-ch-ua': '"Chromium";v="148", "Google Chrome";v="148", "Not/A)Brand";v="99"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-origin',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36',
    'x-csrf-token': csrfToken,
    'x-requested-with': 'XMLHttpRequest',
    ...overrides
  };
}

function buildXdrLogSearchCountRequestBody(params) {
  return {
    time: {
      start: params.start,
      end: params.end
    },
    tables: ['NetworkSecurityLog', 'EndpointSecurityLog']
  };
}

function buildXdrAccessDirectionLogSearchCountRequestBody(params, accessDirection) {
  return {
    filter: {
      logicalOp: 'and',
      filters: [
        {
          field: 'accessDirection',
          conditionalOp: 'IN',
          value: [accessDirection]
        }
      ]
    },
    time: {
      start: params.start,
      end: params.end
    },
    tables: ['NetworkSecurityLog', 'EndpointSecurityLog']
  };
}

function buildXdrIn2outLogSearchCountRequestBody(params) {
  return buildXdrAccessDirectionLogSearchCountRequestBody(params, 'in2out');
}

function buildXdrOut2inLogSearchCountRequestBody(params) {
  return buildXdrAccessDirectionLogSearchCountRequestBody(params, 'out2in');
}

function buildXdrRejectedExternalToInternalCountRequestBody(params) {
  return {
    extensionParams: {},
    spl: {
      mappedSpl: '(srcIpTag = 0 and dstIpTag = 1) | filter 动作  in { "拒绝" }',
      originalSpl: '(srcIpTag = 0 and dstIpTag = 1) | filter 动作  in { "拒绝" }',
      extensionParams: {
        frontRender: [
          {
            displayField: '动作',
            field: 'action',
            value: [2],
            headerType: 'metaType',
            searchType: 'selector',
            valueText: '拒绝',
            isValueNegate: false,
            type: 'number',
            filterSelect: 'renderValue'
          }
        ],
        mappedInputSpl: '',
        originalInputSpl: ''
      }
    },
    serviceInfo: {
      appName: 'incident',
      servletContextPath: '/',
      serviceType: 'table',
      handler: 'analysisTableQueryHandler'
    },
    globalCondition: {
      branchIds: [],
      time: {
        timeField: 'recordTimestamp',
        end: { type: 'absolute', value: params.end },
        begin: { type: 'absolute', value: params.start }
      }
    },
    table: {
      enable: true,
      viewName: 'NetworkSecurityLogView+EndpointSecurityLogView',
      aggregationStrategies: null,
      tableFields: [
        { field: 'recordTimestamp', show: true, selected: true, sort: 'desc', columnWidth: 100, fixed: null, dataType: 'value' },
        { field: 'productType', show: true, selected: true, sort: 'disable', columnWidth: 80, fixed: null, dataType: 'value' },
        { field: 'ruleName', show: true, selected: true, sort: 'disable', columnWidth: 120, fixed: null, dataType: 'value' },
        { field: 'deviceId', show: true, selected: true, sort: 'disable', columnWidth: 120, fixed: null, dataType: 'value' },
        { field: 'deviceIp', show: true, selected: true, sort: 'disable', columnWidth: 120, fixed: null, dataType: 'value' },
        { field: 'srcIp', show: true, selected: true, sort: 'disable', columnWidth: 130, fixed: null, dataType: 'value' },
        { field: 'srcPort', show: true, selected: true, sort: 'disable', columnWidth: 80, fixed: null, dataType: 'value' },
        { field: 'dstIp', show: true, selected: true, sort: 'disable', columnWidth: 130, fixed: null, dataType: 'value' },
        { field: 'dstPort', show: true, selected: true, sort: 'disable', columnWidth: 80, fixed: null, dataType: 'value' },
        { field: 'threatSubTypeProxy', show: true, selected: true, sort: 'disable', columnWidth: 110, fixed: null, dataType: 'value' },
        { field: 'attackState', show: true, selected: true, sort: 'disable', columnWidth: 90, fixed: null, dataType: 'value' },
        { field: 'severity', show: true, selected: true, sort: 'disable', columnWidth: 95, fixed: null, dataType: 'value' },
        { field: 'requestMethod', show: true, selected: true, sort: 'disable', columnWidth: 150, fixed: null, dataType: 'value' },
        { field: 'url', show: true, selected: true, sort: 'disable', columnWidth: 110, fixed: null, dataType: 'value' },
        { field: 'xForwardedFor', show: true, selected: true, sort: 'disable', columnWidth: 140, fixed: null, dataType: 'value' },
        { field: 'respStatus', show: true, selected: true, sort: 'disable', columnWidth: 110, fixed: null, dataType: 'value' },
        { field: 'virusName', show: true, selected: true, sort: 'disable', columnWidth: 120, fixed: null, dataType: 'value' },
        { field: 'username', show: true, selected: true, sort: 'disable', columnWidth: 150, fixed: null, dataType: 'value' },
        { field: 'hostIp', show: true, selected: true, sort: 'disable', columnWidth: 110, fixed: null, dataType: 'value' },
        { field: 'fileMd5', show: true, selected: true, sort: 'disable', columnWidth: 130, fixed: null, dataType: 'value' }
      ],
      pageNum: 1,
      pageSize: 20,
      serviceInfo: {
        appName: 'incident',
        servletContextPath: '/',
        serviceType: 'table',
        handler: 'analysisTableQueryHandler'
      },
      subTable: null,
      rightClicked: false,
      selectAllPage: false,
      routers: [],
      rightActions: [
        { name: 'addFilter', type: 'filter', params: null, actionParams: null, applicableCols: null },
        { name: 'removeFilter', type: 'filter', params: null, actionParams: null, applicableCols: null },
        { name: 'copyCellText', type: 'copy', params: null, actionParams: null, applicableCols: null }
      ],
      extensionParams: {},
      tag: null
    },
    viewName: 'NetworkSecurityLogView+EndpointSecurityLogView',
    model: 'simple',
    autoRefresh: false,
    enableHistory: true
  };
}

function buildXdrIncidentTableCountRequestBody(params) {
  return {
    extensionParams: null,
    spl: {
      mappedSpl: '',
      originalSpl: '',
      extensionParams: {
        frontRender: [],
        mappedInputSpl: '',
        originalInputSpl: ''
      }
    },
    serviceInfo: {
      appName: 'incident',
      servletContextPath: '/',
      serviceType: 'table',
      handler: 'incidentTableQueryHandler'
    },
    globalCondition: {
      branchIds: [],
      time: {
        timeField: 'endTime',
        begin: { type: 'absolute', unit: null, value: params.start },
        end: { type: 'absolute', unit: null, value: params.end }
      }
    },
    table: {
      enable: true,
      viewName: 'IncidentView',
      aggregationStrategies: null,
      tableFields: [],
      pageNum: 1,
      pageSize: 1000,
      serviceInfo: {
        appName: 'incident',
        servletContextPath: '/',
        serviceType: 'table',
        handler: 'incidentTableQueryHandler'
      },
      subTable: null,
      rightClicked: true,
      selectAllPage: true,
      routers: [],
      rightActions: [],
      extensionParams: {
        spl: 'filter xthConfirm= true'
      },
      tag: null
    },
    viewName: 'IncidentView',
    model: 'simple',
    viewInstanceId: '',
    enableHistory: true
  };
}

function buildXdrAlertTableCountRequestBody(params) {
  return {
    extensionParams: null,
    spl: {
      mappedSpl: '',
      originalSpl: '',
      extensionParams: {
        frontRender: [],
        mappedInputSpl: '',
        originalInputSpl: ''
      }
    },
    serviceInfo: {
      appName: 'incident',
      servletContextPath: '/',
      serviceType: 'table',
      handler: 'alertTableQueryHandler'
    },
    globalCondition: {
      branchIds: [],
      time: {
        timeField: 'lastTime',
        begin: { type: 'absolute', value: params.start },
        end: { type: 'absolute', value: params.end }
      }
    },
    table: {
      enable: true,
      viewName: 'AlertView',
      aggregationStrategies: null,
      tableFields: [],
      pageNum: 1,
      pageSize: 50,
      serviceInfo: {
        appName: 'incident',
        servletContextPath: '/',
        serviceType: 'table',
        handler: 'alertTableQueryHandler'
      },
      subTable: null,
      rightClicked: false,
      selectAllPage: true,
      routers: [],
      rightActions: [],
      extensionParams: {},
      tag: null
    },
    viewName: 'AlertView',
    model: 'expert',
    autoRefresh: false,
    viewInstanceId: '',
    enableHistory: true
  };
}

function buildXdrIncidentTableQueryCountRequestBody(params) {
  return {
    extensionParams: null,
    spl: {
      mappedSpl: 'filter 处置状态  in { "处置完成" }',
      originalSpl: 'filter 处置状态  in { "处置完成" }',
      extensionParams: {
        frontRender: [
          {
            displayField: '处置状态',
            field: 'dealStatus',
            value: [3],
            headerType: 'metaType',
            searchType: 'selector',
            valueText: '处置完成',
            isValueNegate: false,
            type: 'string',
            filterSelect: 'renderValue'
          }
        ],
        mappedInputSpl: '',
        originalInputSpl: ''
      }
    },
    serviceInfo: {
      appName: 'incident',
      servletContextPath: '/',
      serviceType: 'table',
      handler: 'incidentTableQueryHandler'
    },
    globalCondition: {
      branchIds: [],
      time: {
        timeField: 'endTime',
        begin: { type: 'absolute', unit: null, value: params.start },
        end: { type: 'absolute', unit: null, value: params.end }
      }
    },
    table: {
      enable: true,
      viewName: 'IncidentView',
      aggregationStrategies: null,
      tableFields: [
        { field: 'dealStatus', show: true, selected: true, sort: 'disable', columnWidth: 120, fixed: false, dataType: 'value' },
        { field: 'endTime', show: true, selected: true, sort: 'desc', columnWidth: 160, fixed: false, dataType: 'value' }
      ],
      pageNum: 1,
      pageSize: 1000,
      serviceInfo: {
        appName: 'incident',
        servletContextPath: '/',
        serviceType: 'table',
        handler: 'incidentTableQueryHandler'
      },
      subTable: null,
      rightClicked: false,
      selectAllPage: true,
      routers: [],
      rightActions: [],
      extensionParams: {
        spl: 'filter xthConfirm= true'
      },
      tag: null
    },
    viewName: 'IncidentView',
    model: 'simple',
    autoRefresh: false,
    viewInstanceId: '6904b17d8a099b2dc81a73a7',
    enableHistory: true
  };
}

function buildXdrG105IncidentCountRequestBody(params) {
  return {
    extensionParams: null,
    spl: {
      mappedSpl: 'manualType != 1 | filter 处置状态  in { "已遏制", "处置完成", "已忽略" } | filter 加白状态  in { "未加白" }',
      originalSpl: 'manualType != 1 | filter 处置状态  in { "已遏制", "处置完成", "已忽略" } | filter 加白状态  in { "未加白" }',
      extensionParams: {
        frontRender: [
          {
            displayField: '处置状态',
            field: 'dealStatus',
            value: [6, 3, 5],
            headerType: 'metaType',
            searchType: 'selector',
            valueText: '已遏制, 处置完成, 已忽略',
            isValueNegate: false,
            type: 'string',
            filterSelect: 'renderValue'
          },
          {
            displayField: '加白状态',
            field: 'whiteStatus',
            value: ['未加白'],
            headerType: 'alertWhiteStatus',
            searchType: 'selector',
            valueText: '未加白',
            isValueNegate: false,
            type: 'string',
            filterSelect: 'renderValue'
          }
        ],
        mappedInputSpl: 'manualType != 1',
        originalInputSpl: 'manualType != 1'
      }
    },
    serviceInfo: {
      appName: 'incident',
      servletContextPath: '/',
      serviceType: 'table',
      handler: 'incidentTableQueryHandler'
    },
    globalCondition: {
      branchIds: [],
      time: {
        timeField: 'endTime',
        begin: { type: 'absolute', value: params.start },
        end: { type: 'absolute', value: params.end }
      }
    },
    table: {
      enable: true,
      viewName: 'IncidentView',
      aggregationStrategies: null,
      tableFields: [],
      pageNum: 1,
      pageSize: 50,
      serviceInfo: {
        appName: 'incident',
        servletContextPath: '/',
        serviceType: 'table',
        handler: 'incidentTableQueryHandler'
      },
      subTable: null,
      rightClicked: true,
      selectAllPage: true,
      routers: [],
      rightActions: [],
      extensionParams: {
        spl: 'filter xthConfirm= true'
      },
      tag: null
    },
    viewName: 'IncidentView',
    model: 'expert',
    autoRefresh: false,
    viewInstanceId: '6719cafbd619f06373baa6ba',
    enableHistory: true
  };
}

function buildXdrG103IncidentCountRequestBody(params) {
  return {
    extensionParams: null,
    spl: {
      mappedSpl: 'filter 处置执行方式 in { "人工决策处置" } and  加白状态 != "已加白"  | filter 处置状态  in { "处置完成" }',
      originalSpl: 'filter 处置执行方式 in { "人工决策处置" } and  加白状态 != "已加白"  | filter 处置状态  in { "处置完成" }',
      extensionParams: {
        frontRender: [
          {
            displayField: '处置状态',
            field: 'dealStatus',
            value: [3],
            headerType: 'metaType',
            searchType: 'selector',
            valueText: '处置完成',
            isValueNegate: false,
            type: 'string',
            filterSelect: 'renderValue'
          }
        ],
        mappedInputSpl: 'filter 处置执行方式 in { "人工决策处置" } and  加白状态 != "已加白" ',
        originalInputSpl: 'filter 处置执行方式 in { "人工决策处置" } and  加白状态 != "已加白" '
      }
    },
    serviceInfo: {
      appName: 'incident',
      servletContextPath: '/',
      serviceType: 'table',
      handler: 'incidentTableQueryHandler'
    },
    globalCondition: {
      branchIds: [],
      time: {
        timeField: 'endTime',
        begin: { type: 'absolute', value: params.start },
        end: { type: 'absolute', value: params.end }
      }
    },
    table: {
      enable: true,
      viewName: 'IncidentView',
      aggregationStrategies: null,
      tableFields: [],
      pageNum: 1,
      pageSize: 50,
      serviceInfo: {
        appName: 'incident',
        servletContextPath: '/',
        serviceType: 'table',
        handler: 'incidentTableQueryHandler'
      },
      subTable: null,
      rightClicked: true,
      selectAllPage: true,
      routers: [],
      rightActions: [],
      extensionParams: {
        spl: 'filter xthConfirm= true'
      },
      tag: null
    },
    viewName: 'IncidentView',
    model: 'expert',
    autoRefresh: false,
    viewInstanceId: '6719cafbd619f06373baa6ba',
    enableHistory: true
  };
}

function buildXdrG106IncidentCountRequestBody(params) {
  return {
    extensionParams: null,
    spl: {
      mappedSpl: 'filter 处置状态  in { "处置中", "挂起" }',
      originalSpl: 'filter 处置状态  in { "处置中", "挂起" }',
      extensionParams: {
        frontRender: [
          {
            displayField: '处置状态',
            field: 'dealStatus',
            value: [4, 2],
            headerType: 'metaType',
            searchType: 'selector',
            valueText: '处置中, 挂起',
            isValueNegate: false,
            type: 'string',
            filterSelect: 'renderValue'
          }
        ],
        mappedInputSpl: '',
        originalInputSpl: ''
      }
    },
    serviceInfo: {
      appName: 'incident',
      servletContextPath: '/',
      serviceType: 'table',
      handler: 'incidentTableQueryHandler'
    },
    globalCondition: {
      branchIds: [],
      time: {
        timeField: 'endTime',
        begin: { type: 'absolute', value: params.start },
        end: { type: 'absolute', value: params.end }
      }
    },
    table: {
      enable: true,
      viewName: 'IncidentView',
      aggregationStrategies: null,
      tableFields: [],
      pageNum: 1,
      pageSize: 50,
      serviceInfo: {
        appName: 'incident',
        servletContextPath: '/',
        serviceType: 'table',
        handler: 'incidentTableQueryHandler'
      },
      subTable: null,
      rightClicked: true,
      selectAllPage: true,
      routers: [],
      rightActions: [],
      extensionParams: {
        spl: 'filter xthConfirm= true'
      },
      tag: null
    },
    viewName: 'IncidentView',
    model: 'expert',
    autoRefresh: false,
    viewInstanceId: '',
    enableHistory: true
  };
}

function buildXdrG107IncidentCountRequestBody(params) {
  return {
    extensionParams: null,
    spl: {
      mappedSpl: 'filter 处置状态  in { "待处置" }',
      originalSpl: 'filter 处置状态  in { "待处置" }',
      extensionParams: {
        frontRender: [
          {
            displayField: '处置状态',
            field: 'dealStatus',
            value: [1],
            headerType: 'metaType',
            searchType: 'selector',
            valueText: '待处置',
            isValueNegate: false,
            type: 'string',
            filterSelect: 'renderValue'
          }
        ],
        mappedInputSpl: '',
        originalInputSpl: ''
      }
    },
    serviceInfo: {
      appName: 'incident',
      servletContextPath: '/',
      serviceType: 'table',
      handler: 'incidentTableQueryHandler'
    },
    globalCondition: {
      branchIds: [],
      time: {
        timeField: 'endTime',
        begin: { type: 'absolute', value: params.start },
        end: { type: 'absolute', value: params.end }
      }
    },
    table: {
      enable: true,
      viewName: 'IncidentView',
      aggregationStrategies: null,
      tableFields: [],
      pageNum: 1,
      pageSize: 50,
      serviceInfo: {
        appName: 'incident',
        servletContextPath: '/',
        serviceType: 'table',
        handler: 'incidentTableQueryHandler'
      },
      subTable: null,
      rightClicked: true,
      selectAllPage: true,
      routers: [],
      rightActions: [],
      extensionParams: {
        spl: 'filter xthConfirm= true'
      },
      tag: null
    },
    viewName: 'IncidentView',
    model: 'expert',
    autoRefresh: false,
    viewInstanceId: '',
    enableHistory: true
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
      timeout: REQUEST_TIMEOUT_MS
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
      reject(new Error(`请求超时 (${REQUEST_TIMEOUT_SECONDS}秒)`));
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
      timeout: REQUEST_TIMEOUT_MS
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
      reject(new Error(`GET 请求超时 (${REQUEST_TIMEOUT_SECONDS}秒)`));
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

function buildWeakPwdSummaryRequestBody(params) {
  const startTimestamp = dateToTimestamp(params.startTime);
  const endTimestamp = dateToTimestamp(params.endTime);
  const endOfDay = endTimestamp + 24 * 60 * 60 * 1000;
  const dealStatus = Array.isArray(params.dealStatus) ? params.dealStatus : [];
  const pageSize = params.pageSize || 100;
  const page = params.page || 1;

  return {
    order: 'asc',
    offset: (page - 1) * pageSize,
    limit: pageSize,
    keyword: '',
    login_status: [],
    is_intranet: [],
    business_level: [],
    src_storage: [],
    src_type: [],
    service_status: [],
    scene_tag: [],
    login_time: [],
    found_time: [startTimestamp, endOfDay],
    reappear: 0,
    is_admin: 1,
    risk_level: [],
    deal_status: dealStatus,
    company_id: String(params.customerId || '')
  };
}

function buildWeakPwdListRequestBody(params) {
  return {
    ...buildWeakPwdSummaryRequestBody(params),
    ip: String(params.ip || '')
  };
}

function buildTopnLoadConditionRequestBody(params) {
  return {
    company_id: String(params.customerId || '')
  };
}

function buildTopnDeviceListRequestBody(params) {
  return {
    company_id: String(params.customerId || '')
  };
}

function normalizeTopnConditionIds(ids) {
  if (!Array.isArray(ids)) return [];
  return ids
    .map(id => String(id || '').trim())
    .filter(id => id)
    .map(id => ({ id }));
}

function normalizeTopnDeviceInfos(deviceInfos) {
  if (!Array.isArray(deviceInfos)) return [];
  return deviceInfos
    .map((deviceInfo) => {
      const deviceType = String((deviceInfo || {}).device_type || '').trim();
      const deviceIdList = Array.isArray((deviceInfo || {}).device_id_list)
        ? deviceInfo.device_id_list.map(id => String(id || '').trim()).filter(id => id)
        : [];
      if (!deviceType || deviceIdList.length === 0) return null;
      return {
        device_type: deviceType,
        device_id_list: deviceIdList
      };
    })
    .filter(Boolean);
}

function normalizeTopnDeviceInfosFromDeviceList(deviceList) {
  if (!Array.isArray(deviceList)) return [];
  return deviceList
    .map((deviceGroup) => {
      const deviceType = String((deviceGroup || {}).device_type || '').trim();
      const deviceIdList = Array.isArray((deviceGroup || {}).device_info_list)
        ? deviceGroup.device_info_list
          .map(deviceInfo => String((deviceInfo || {}).device_id || '').trim())
          .filter(id => id)
        : [];
      if (!deviceType || deviceIdList.length === 0) return null;
      return {
        device_type: deviceType,
        device_id_list: deviceIdList
      };
    })
    .filter(Boolean);
}

function buildTopnRequestBody(params, topnCondition, options = {}) {
  const startTimestamp = dateToTimestamp(params.startTime);
  const endTimestamp = dateToTimestamp(params.endTime);
  const endOfDay = endTimestamp + 24 * 60 * 60 * 1000 - 1;
  const conditionData = (topnCondition || {}).data || topnCondition || {};

  return {
    company_id: String(params.customerId || ''),
    attack_type: conditionData.attack_type || normalizeTopnConditionIds(conditionData.attack_type_ids),
    attack_direction: conditionData.attack_direction || normalizeTopnConditionIds(conditionData.attack_direction_ids),
    ip_type: options.ipType || 'src_ip',
    topn: options.topn || 100,
    device_infos: normalizeTopnDeviceInfos(conditionData.device_infos),
    latest_time: [startTimestamp, endOfDay],
    date_scope: 1,
    resource_group_ids: [],
    only_asset: 0
  };
}

function buildAssetDownloadRequestBody(params) {
  const serviceStatus = Array.isArray(params.service_status)
    ? params.service_status
    : [1];

  return {
    asset_list: [],
    exclude_asset_list: [],
    is_select_all: 1,
    order: 'asc',
    service_status: serviceStatus,
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

function buildExposedIpListStatisticsRequestBody(params) {
  return {
    keyword: '',
    target_company_id: Array.isArray(params.targetCompanyIds) ? params.targetCompanyIds : [],
    related_task: [],
    asset_tag: [],
    company_id: String(params.customerId || ''),
    attack_authorisation: 0
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

async function fetchWeakPwdSummaryTotal(cookieInfo, params) {
  const { cookieString, csrfToken } = cookieInfo;

  const requestBody = buildWeakPwdSummaryRequestBody(params);
  const headers = generateHeaders(cookieString, csrfToken);
  const url = `https://${API_CONFIG.baseUrl}${API_CONFIG.weakPwdSummaryEndpoint}`;

  console.log(`[ApiClient] 请求弱口令统计参数:`, JSON.stringify(requestBody, null, 2));

  const result = await httpPost(url, headers, JSON.stringify(requestBody));
  if (!result || result.code !== 0) {
    throw new Error(`弱口令统计接口返回异常: ${JSON.stringify(result).substring(0, 500)}`);
  }

  const total = Number((((result || {}).data || {}).total) || 0);
  if (!Number.isFinite(total)) {
    throw new Error(`弱口令统计 total 不是有效数字: ${JSON.stringify(result).substring(0, 500)}`);
  }

  return total;
}

function buildOrderBranchRequestBody(params, orderType) {
  return {
    order: 'asc',
    offset: 0,
    limit: 10,
    keyword: '',
    type: orderType,
    status: 0,
    company_id: String(params.customerId || ''),
    platform_type: 'scloud',
    belong_local_xdr: 0
  };
}

async function fetchOrderBranchTotal(cookieInfo, params, orderType) {
  const { cookieString, csrfToken } = cookieInfo;

  const requestBody = buildOrderBranchRequestBody(params, orderType);
  const headers = generateHeaders(cookieString, csrfToken);
  const url = `https://${API_CONFIG.baseUrl}${API_CONFIG.orderBranchEndpoint}`;

  console.log(`[ApiClient] 请求订单列表统计(type=${orderType})参数:`, JSON.stringify(requestBody, null, 2));

  const result = await httpPost(url, headers, JSON.stringify(requestBody));
  if (!result || result.code !== 0 || !result.data) {
    throw new Error(`订单列表统计(type=${orderType})接口返回异常: ${JSON.stringify(result).substring(0, 500)}`);
  }

  const total = Number(result.data.total);
  if (!Number.isFinite(total)) {
    throw new Error(`订单列表统计(type=${orderType})响应 data.total 不是有效数字: ${JSON.stringify(result).substring(0, 500)}`);
  }

  return total;
}

async function fetchWeakPwdSummary(cookieInfo, params) {
  console.log(`[ApiClient] 开始分页获取弱口令统计...`);

  const allRows = [];
  const pageSize = params.pageSize || 100;
  let page = 1;
  let hasMore = true;
  let previousFingerprint = null;

  while (hasMore) {
    console.log(`[ApiClient] 获取弱口令统计第 ${page} 页...`);
    const { cookieString, csrfToken } = cookieInfo;
    const requestBody = buildWeakPwdSummaryRequestBody({
      ...params,
      page,
      pageSize
    });
    const headers = generateHeaders(cookieString, csrfToken);
    const url = `https://${API_CONFIG.baseUrl}${API_CONFIG.weakPwdSummaryEndpoint}`;

    console.log(`[ApiClient] 请求弱口令统计参数:`, JSON.stringify(requestBody, null, 2));

    const result = await httpPost(url, headers, JSON.stringify(requestBody));
    if (!result || result.code !== 0) {
      throw new Error(`弱口令统计接口返回异常: ${JSON.stringify(result).substring(0, 500)}`);
    }

    const rows = extractRowsFromResponse(result, 'data');
    if (rows.length === 0) {
      console.log(`[ApiClient] 弱口令统计第 ${page} 页没有数据。`);
      hasMore = false;
      continue;
    }

    const currentFingerprint = buildPageFingerprint(rows);
    if (currentFingerprint === previousFingerprint) {
      throw new Error(`弱口令统计分页疑似未生效，第 ${page} 页返回了与上一页相同的数据，已中止以避免死循环。`);
    }
    previousFingerprint = currentFingerprint;

    allRows.push(...rows);
    console.log(`[ApiClient] 弱口令统计第 ${page} 页获取到 ${rows.length} 条数据，累计 ${allRows.length} 条`);

    hasMore = getHasMoreFromResponse(result, rows, allRows.length, pageSize, 'data');
    page++;

    const pageDelayMs = Number.isFinite(params.pageDelayMs) ? params.pageDelayMs : 10;
    if (pageDelayMs > 0 && hasMore) {
      await sleep(pageDelayMs);
    }
  }

  console.log(`[ApiClient] 弱口令统计分页获取完成，共 ${allRows.length} 条数据`);

  const ips = [...new Set(allRows
    .map(row => String(row && row.ip || '').trim())
    .filter(Boolean))];
  let total = 0;
  for (let index = 0; index < ips.length; index += 1) {
    const ip = ips[index];
    const { cookieString, csrfToken } = cookieInfo;
    const requestBody = buildWeakPwdListRequestBody({ ...params, ip });
    const headers = generateHeaders(cookieString, csrfToken);
    const url = `https://${API_CONFIG.baseUrl}${API_CONFIG.weakPwdListEndpoint}`;

    console.log(`[ApiClient] 请求弱口令明细 ${index + 1}/${ips.length}，IP: ${ip}`);
    const result = await httpPost(url, headers, JSON.stringify(requestBody));
    if (!result || result.code !== 0) {
      throw new Error(`IP ${ip} 弱口令明细接口返回异常: ${JSON.stringify(result).substring(0, 500)}`);
    }

    const ipTotal = Number((((result || {}).data || {}).total) || 0);
    if (!Number.isFinite(ipTotal)) {
      throw new Error(`IP ${ip} 弱口令明细 total 不是有效数字: ${JSON.stringify(result).substring(0, 500)}`);
    }
    total += ipTotal;
    console.log(`[ApiClient] IP ${ip} 弱口令数: ${ipTotal}，累计: ${total}`);

    const pageDelayMs = Number.isFinite(params.pageDelayMs) ? params.pageDelayMs : 10;
    if (pageDelayMs > 0 && index < ips.length - 1) {
      await sleep(pageDelayMs);
    }
  }

  return {
    total,
    list: allRows
  };
}

async function fetchTopnLoadCondition(cookieInfo, params) {
  const { cookieString, csrfToken } = cookieInfo;

  const requestBody = buildTopnLoadConditionRequestBody(params);
  const headers = generateHeaders(cookieString, csrfToken);
  const url = `https://${API_CONFIG.baseUrl}${API_CONFIG.topnLoadConditionEndpoint}`;

  console.log(`[ApiClient] 请求 TopN 查询条件参数:`, JSON.stringify(requestBody, null, 2));

  const result = await httpPost(url, headers, JSON.stringify(requestBody));
  if (!result || result.code !== 0 || !result.data) {
    throw new Error(`TopN 查询条件接口返回异常: ${JSON.stringify(result).substring(0, 500)}`);
  }

  return result.data;
}

async function fetchTopnDeviceList(cookieInfo, params) {
  const { cookieString, csrfToken } = cookieInfo;

  const requestBody = buildTopnDeviceListRequestBody(params);
  const headers = generateHeaders(cookieString, csrfToken);
  const url = `https://${API_CONFIG.baseUrl}${API_CONFIG.topnDeviceListEndpoint}`;

  console.log(`[ApiClient] 请求 TopN 设备列表参数:`, JSON.stringify(requestBody, null, 2));

  const result = await httpPost(url, headers, JSON.stringify(requestBody));
  if (!result || result.code !== 0 || !result.data || !Array.isArray(result.data.list)) {
    throw new Error(`TopN 设备列表接口返回异常: ${JSON.stringify(result).substring(0, 500)}`);
  }

  return normalizeTopnDeviceInfosFromDeviceList(result.data.list);
}

function extractTopnList(result, label) {
  if (!result || result.code !== 0 || !result.data || !Array.isArray(result.data.list)) {
    throw new Error(`${label} 接口返回异常: ${JSON.stringify(result).substring(0, 500)}`);
  }
  return result.data.list;
}

async function fetchTopnThreatTypes(cookieInfo, params, loadCondition) {
  const { cookieString, csrfToken } = cookieInfo;

  const requestBody = buildTopnRequestBody(params, loadCondition);
  const headers = generateHeaders(cookieString, csrfToken);
  const url = `https://${API_CONFIG.baseUrl}${API_CONFIG.topnThreatTypeEndpoint}`;

  console.log(`[ApiClient] 请求威胁类型 TopN 参数:`, JSON.stringify(requestBody, null, 2));

  const result = await httpPost(url, headers, JSON.stringify(requestBody));
  return extractTopnList(result, '威胁类型 TopN');
}

async function fetchTopnSrcIpGeos(cookieInfo, params, loadCondition) {
  const { cookieString, csrfToken } = cookieInfo;

  const requestBody = buildTopnRequestBody(params, loadCondition);
  const headers = generateHeaders(cookieString, csrfToken);
  const url = `https://${API_CONFIG.baseUrl}${API_CONFIG.topnSrcIpGeoEndpoint}`;

  console.log(`[ApiClient] 请求攻击源地理位置 TopN 参数:`, JSON.stringify(requestBody, null, 2));

  const result = await httpPost(url, headers, JSON.stringify(requestBody));
  return extractTopnList(result, '攻击源地理位置 TopN');
}

async function fetchTopnDstIps(cookieInfo, params, loadCondition) {
  const { cookieString, csrfToken } = cookieInfo;

  const requestBody = buildTopnRequestBody(params, loadCondition, {
    ipType: 'dst_ip',
    topn: 5
  });
  const headers = generateHeaders(cookieString, csrfToken);
  const url = `https://${API_CONFIG.baseUrl}${API_CONFIG.topnDstIpEndpoint}`;

  console.log(`[ApiClient] 请求目的 IP TopN 参数:`, JSON.stringify(requestBody, null, 2));

  const result = await httpPost(url, headers, JSON.stringify(requestBody));
  return extractTopnList(result, '目的 IP TopN');
}

async function fetchTopnReportStats(cookieInfo, params) {
  const deviceInfos = await fetchTopnDeviceList(cookieInfo, params);
  const topnCondition = {
    attack_type: [{ id: 'security_log' }],
    attack_direction: [{ id: '1' }],
    device_infos: deviceInfos
  };
  const [threatTypes, srcIpGeos, dstIps] = await Promise.all([
    fetchTopnThreatTypes(cookieInfo, params, topnCondition),
    fetchTopnSrcIpGeos(cookieInfo, params, topnCondition),
    fetchTopnDstIps(cookieInfo, params, topnCondition)
  ]);

  return {
    threatTypes,
    srcIpGeos,
    dstIps
  };
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

async function fetchExposedSurfaceIpListStatistics(cookieInfo, params) {
  const { cookieString, csrfToken } = cookieInfo;
  await visitHomePage(cookieInfo);

  const requestBody = buildExposedIpListStatisticsRequestBody(params);
  const headers = generateHeaders(cookieString, csrfToken, {
    referer: buildHomeUrl()
  });
  const url = `https://${API_CONFIG.baseUrl}${API_CONFIG.exposedIpListStatisticsEndpoint}`;

  console.log(`[ApiClient] 请求暴露面 IP 列表统计参数:`, JSON.stringify(requestBody, null, 2));

  const result = await httpPost(url, headers, JSON.stringify(requestBody));
  if (!result || result.code !== 0 || !result.data || result.data.port === undefined) {
    throw new Error(`暴露面 IP 列表统计接口返回异常: ${JSON.stringify(result).substring(0, 500)}`);
  }

  const port = Number(result.data.port);
  if (!Number.isFinite(port)) {
    throw new Error(`暴露面 IP 列表统计 port 不是有效数字: ${JSON.stringify(result).substring(0, 500)}`);
  }
  return port;
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

async function fetchXdrLogSearchCount(cookieInfo, params) {
  const { cookieString, csrfToken } = cookieInfo;
  if (!Number.isFinite(params.start) || !Number.isFinite(params.end)) {
    throw new Error(`XDR count 时间范围无效: ${params.start} ~ ${params.end}`);
  }
  if (params.start > params.end) {
    throw new Error(`XDR count 开始时间不能晚于结束时间: ${params.start} ~ ${params.end}`);
  }

  const xdrBaseUrl = getXdrBaseUrl(cookieInfo);
  const url = `https://${xdrBaseUrl}${API_CONFIG.xdrLogSearchCountEndpoint}`;
  const headers = generateXdrHeaders(cookieString, csrfToken, {}, xdrBaseUrl);
  const body = JSON.stringify(buildXdrLogSearchCountRequestBody(params));
  const result = await httpPost(url, headers, body);

  if (!result || result.code !== 'Success') {
    throw new Error(`XDR count 接口返回异常: ${JSON.stringify(result).substring(0, 500)}`);
  }

  const count = Number(result.data);
  if (!Number.isFinite(count)) {
    throw new Error(`XDR count 响应 data 不是有效数字: ${JSON.stringify(result).substring(0, 500)}`);
  }
  return count;
}

async function fetchXdrIn2outLogSearchCount(cookieInfo, params) {
  return fetchXdrAccessDirectionLogSearchCount(cookieInfo, params, 'in2out');
}

async function fetchXdrOut2inLogSearchCount(cookieInfo, params) {
  return fetchXdrAccessDirectionLogSearchCount(cookieInfo, params, 'out2in');
}

async function fetchXdrAccessDirectionLogSearchCount(cookieInfo, params, accessDirection) {
  const { cookieString, csrfToken } = cookieInfo;
  if (!Number.isFinite(params.start) || !Number.isFinite(params.end)) {
    throw new Error(`XDR ${accessDirection} count 时间范围无效: ${params.start} ~ ${params.end}`);
  }
  if (params.start > params.end) {
    throw new Error(`XDR ${accessDirection} count 开始时间不能晚于结束时间: ${params.start} ~ ${params.end}`);
  }

  const xdrBaseUrl = getXdrBaseUrl(cookieInfo);
  const url = `https://${xdrBaseUrl}${API_CONFIG.xdrLogSearchCountEndpoint}`;
  const headers = generateXdrHeaders(cookieString, csrfToken, {}, xdrBaseUrl);
  const body = JSON.stringify(buildXdrAccessDirectionLogSearchCountRequestBody(params, accessDirection));
  const result = await httpPost(url, headers, body);

  if (!result || result.code !== 'Success') {
    throw new Error(`XDR ${accessDirection} count 接口返回异常: ${JSON.stringify(result).substring(0, 500)}`);
  }

  const count = Number(result.data);
  if (!Number.isFinite(count)) {
    throw new Error(`XDR ${accessDirection} count 响应 data 不是有效数字: ${JSON.stringify(result).substring(0, 500)}`);
  }
  return count;
}

async function fetchXdrRejectedExternalToInternalCount(cookieInfo, params) {
  const { cookieString, csrfToken } = cookieInfo;
  if (!Number.isFinite(params.start) || !Number.isFinite(params.end)) {
    throw new Error(`XDR rejected external->internal count 时间范围无效: ${params.start} ~ ${params.end}`);
  }
  if (params.start > params.end) {
    throw new Error(`XDR rejected external->internal count 开始时间不能晚于结束时间: ${params.start} ~ ${params.end}`);
  }

  const xdrBaseUrl = getXdrBaseUrl(cookieInfo);
  const url = `https://${xdrBaseUrl}${API_CONFIG.xdrIncidentAnalysisCountEndpoint}`;
  const headers = generateXdrHeaders(cookieString, csrfToken, {}, xdrBaseUrl);
  const body = JSON.stringify(buildXdrRejectedExternalToInternalCountRequestBody(params));
  const result = await httpPost(url, headers, body);

  if (!result || result.code !== 0 || !result.data) {
    throw new Error(`XDR rejected external->internal count 接口返回异常: ${JSON.stringify(result).substring(0, 500)}`);
  }

  const count = Number(result.data.total);
  if (!Number.isFinite(count)) {
    throw new Error(`XDR rejected external->internal count 响应 data.total 不是有效数字: ${JSON.stringify(result).substring(0, 500)}`);
  }
  return count;
}

async function fetchXdrIncidentTableCount(cookieInfo, params) {
  const { cookieString, csrfToken } = cookieInfo;
  if (!Number.isFinite(params.start) || !Number.isFinite(params.end)) {
    throw new Error(`XDR incidentTable count 时间范围无效: ${params.start} ~ ${params.end}`);
  }
  if (params.start > params.end) {
    throw new Error(`XDR incidentTable count 开始时间不能晚于结束时间: ${params.start} ~ ${params.end}`);
  }

  const xdrBaseUrl = getXdrBaseUrl(cookieInfo);
  const url = `https://${xdrBaseUrl}${API_CONFIG.xdrIncidentTableCountEndpoint}`;
  const headers = generateXdrHeaders(cookieString, csrfToken, {}, xdrBaseUrl);
  const body = JSON.stringify(buildXdrIncidentTableCountRequestBody(params));
  const result = await httpPost(url, headers, body);

  if (!result || result.code !== 0 || !result.data) {
    throw new Error(`XDR incidentTable count 接口返回异常: ${JSON.stringify(result).substring(0, 500)}`);
  }

  const count = Number(result.data.total);
  if (!Number.isFinite(count)) {
    throw new Error(`XDR incidentTable count 响应 data.total 不是有效数字: ${JSON.stringify(result).substring(0, 500)}`);
  }
  return count;
}

async function fetchXdrAlertTableCount(cookieInfo, params) {
  const { cookieString, csrfToken } = cookieInfo;
  if (!Number.isFinite(params.start) || !Number.isFinite(params.end)) {
    throw new Error(`XDR alertTable count 时间范围无效: ${params.start} ~ ${params.end}`);
  }
  if (params.start > params.end) {
    throw new Error(`XDR alertTable count 开始时间不能晚于结束时间: ${params.start} ~ ${params.end}`);
  }

  const xdrBaseUrl = getXdrBaseUrl(cookieInfo);
  const url = `https://${xdrBaseUrl}${API_CONFIG.xdrAlertTableQueryEndpoint}`;
  const headers = generateXdrHeaders(cookieString, csrfToken, {}, xdrBaseUrl);
  const body = JSON.stringify(buildXdrAlertTableCountRequestBody(params));
  const result = await httpPost(url, headers, body);

  if (!result || result.code !== 0 || !result.data) {
    throw new Error(`XDR alertTable count 接口返回异常: ${JSON.stringify(result).substring(0, 500)}`);
  }

  const count = Number(result.data.total);
  if (!Number.isFinite(count)) {
    throw new Error(`XDR alertTable count 响应 data.total 不是有效数字: ${JSON.stringify(result).substring(0, 500)}`);
  }
  return count;
}

async function fetchXdrIncidentTableQueryCount(cookieInfo, params) {
  const { cookieString, csrfToken } = cookieInfo;
  if (!Number.isFinite(params.start) || !Number.isFinite(params.end)) {
    throw new Error(`XDR incidentTable query count 时间范围无效: ${params.start} ~ ${params.end}`);
  }
  if (params.start > params.end) {
    throw new Error(`XDR incidentTable query count 开始时间不能晚于结束时间: ${params.start} ~ ${params.end}`);
  }

  const xdrBaseUrl = getXdrBaseUrl(cookieInfo);
  const url = `https://${xdrBaseUrl}${API_CONFIG.xdrIncidentTableQueryEndpoint}`;
  const headers = generateXdrHeaders(cookieString, csrfToken, {}, xdrBaseUrl);
  const body = JSON.stringify(buildXdrIncidentTableQueryCountRequestBody(params));
  console.log('[ApiClient] G102 incidentTable query 请求参数:', body);
  const result = await httpPost(url, headers, body);

  if (!result || result.code !== 0 || !result.data) {
    console.error('[ApiClient] G102 incidentTable query 响应异常:', JSON.stringify(result).substring(0, 1000));
    throw new Error(`XDR incidentTable query count 接口返回异常: ${JSON.stringify(result).substring(0, 500)}`);
  }

  const count = Number(result.data.total);
  if (!Number.isFinite(count)) {
    throw new Error(`XDR incidentTable query count 响应 data.total 不是有效数字: ${JSON.stringify(result).substring(0, 500)}`);
  }
  console.log(`[G102] XDR incidentTable query 返回 total: ${count}`);
  return count;
}

async function fetchXdrG103IncidentCount(cookieInfo, params) {
  const { cookieString, csrfToken } = cookieInfo;
  if (!Number.isFinite(params.start) || !Number.isFinite(params.end)) {
    throw new Error(`XDR G103 incident count 时间范围无效: ${params.start} ~ ${params.end}`);
  }
  if (params.start > params.end) {
    throw new Error(`XDR G103 incident count 开始时间不能晚于结束时间: ${params.start} ~ ${params.end}`);
  }

  const xdrBaseUrl = getXdrBaseUrl(cookieInfo);
  const url = `https://${xdrBaseUrl}${API_CONFIG.xdrIncidentTableCountEndpoint}`;
  const headers = generateXdrHeaders(cookieString, csrfToken, {}, xdrBaseUrl);
  const body = JSON.stringify(buildXdrG103IncidentCountRequestBody(params));
  const result = await httpPost(url, headers, body);

  if (!result || result.code !== 0 || !result.data) {
    throw new Error(`XDR G103 incident count 接口返回异常: ${JSON.stringify(result).substring(0, 500)}`);
  }

  const count = Number(result.data.total);
  if (!Number.isFinite(count)) {
    throw new Error(`XDR G103 incident count 响应 data.total 不是有效数字: ${JSON.stringify(result).substring(0, 500)}`);
  }
  return count;
}

async function fetchXdrG105IncidentCount(cookieInfo, params) {
  const { cookieString, csrfToken } = cookieInfo;
  if (!Number.isFinite(params.start) || !Number.isFinite(params.end)) {
    throw new Error(`XDR G105 incident count 时间范围无效: ${params.start} ~ ${params.end}`);
  }
  if (params.start > params.end) {
    throw new Error(`XDR G105 incident count 开始时间不能晚于结束时间: ${params.start} ~ ${params.end}`);
  }

  const xdrBaseUrl = getXdrBaseUrl(cookieInfo);
  const url = `https://${xdrBaseUrl}${API_CONFIG.xdrIncidentTableCountEndpoint}`;
  const headers = generateXdrHeaders(cookieString, csrfToken, {}, xdrBaseUrl);
  const body = JSON.stringify(buildXdrG105IncidentCountRequestBody(params));
  const result = await httpPost(url, headers, body);

  if (!result || result.code !== 0 || !result.data) {
    throw new Error(`XDR G105 incident count 接口返回异常: ${JSON.stringify(result).substring(0, 500)}`);
  }

  const count = Number(result.data.total);
  if (!Number.isFinite(count)) {
    throw new Error(`XDR G105 incident count 响应 data.total 不是有效数字: ${JSON.stringify(result).substring(0, 500)}`);
  }
  return count;
}

async function fetchXdrG106IncidentCount(cookieInfo, params) {
  const { cookieString, csrfToken } = cookieInfo;
  if (!Number.isFinite(params.start) || !Number.isFinite(params.end)) {
    throw new Error(`XDR G106 incident count 时间范围无效: ${params.start} ~ ${params.end}`);
  }
  if (params.start > params.end) {
    throw new Error(`XDR G106 incident count 开始时间不能晚于结束时间: ${params.start} ~ ${params.end}`);
  }

  const xdrBaseUrl = getXdrBaseUrl(cookieInfo);
  const url = `https://${xdrBaseUrl}${API_CONFIG.xdrIncidentTableQueryEndpoint}`;
  const headers = generateXdrHeaders(cookieString, csrfToken, {}, xdrBaseUrl);
  const body = JSON.stringify(buildXdrG106IncidentCountRequestBody(params));
  const result = await httpPost(url, headers, body);

  if (!result || result.code !== 0 || !result.data) {
    throw new Error(`XDR G106 incident count 接口返回异常: ${JSON.stringify(result).substring(0, 500)}`);
  }

  const count = Number(result.data.total);
  if (!Number.isFinite(count)) {
    throw new Error(`XDR G106 incident count 响应 data.total 不是有效数字: ${JSON.stringify(result).substring(0, 500)}`);
  }
  return count;
}

async function fetchXdrG107IncidentCount(cookieInfo, params) {
  const { cookieString, csrfToken } = cookieInfo;
  if (!Number.isFinite(params.start) || !Number.isFinite(params.end)) {
    throw new Error(`XDR G107 incident count 时间范围无效: ${params.start} ~ ${params.end}`);
  }
  if (params.start > params.end) {
    throw new Error(`XDR G107 incident count 开始时间不能晚于结束时间: ${params.start} ~ ${params.end}`);
  }

  const xdrBaseUrl = getXdrBaseUrl(cookieInfo);
  const url = `https://${xdrBaseUrl}${API_CONFIG.xdrIncidentTableQueryEndpoint}`;
  const headers = generateXdrHeaders(cookieString, csrfToken, {}, xdrBaseUrl);
  const body = JSON.stringify(buildXdrG107IncidentCountRequestBody(params));
  const result = await httpPost(url, headers, body);

  if (!result || result.code !== 0 || !result.data) {
    throw new Error(`XDR G107 incident count 接口返回异常: ${JSON.stringify(result).substring(0, 500)}`);
  }

  const count = Number(result.data.total);
  if (!Number.isFinite(count)) {
    throw new Error(`XDR G107 incident count 响应 data.total 不是有效数字: ${JSON.stringify(result).substring(0, 500)}`);
  }
  return count;
}

async function fetchXdrLogSearchCountForRanges(cookieInfo, ranges, params = {}) {
  if (!Array.isArray(ranges) || ranges.length === 0) {
    console.log('[ApiClient] XDR 重保时间段为空，count 记为 0');
    return 0;
  }

  let total = 0;
  for (let index = 0; index < ranges.length; index += 1) {
    const range = ranges[index];
    console.log(`[ApiClient] XDR count 查询 ${index + 1}/${ranges.length}: ${range.start} ~ ${range.end}`);
    const count = await fetchXdrLogSearchCount(cookieInfo, range);
    total += count;
    console.log(`[ApiClient] XDR count 本段 ${count}，累计 ${total}`);

    const delayMs = Number.isFinite(params.pageDelayMs) ? params.pageDelayMs : 10;
    if (delayMs > 0 && index < ranges.length - 1) {
      await sleep(delayMs);
    }
  }
  return total;
}

function aggregateXdrLogSearchCountDetails(queriedRanges) {
  const summary = {
    total: 0,
    holidayCounts: {}
  };

  (Array.isArray(queriedRanges) ? queriedRanges : []).forEach((range) => {
    const count = Number(range && range.count);
    if (!Number.isFinite(count)) return;
    summary.total += count;

    (Array.isArray(range.owners) ? range.owners : [])
      .filter(owner => owner && owner.source === 'holiday' && owner.key)
      .forEach((owner) => {
        summary.holidayCounts[owner.key] = (summary.holidayCounts[owner.key] || 0) + count;
      });
  });

  return summary;
}

async function fetchXdrLogSearchCountDetailsForRanges(cookieInfo, ranges, params = {}) {
  if (!Array.isArray(ranges) || ranges.length === 0) {
    console.log('[ApiClient] XDR 重保时间段为空，count 记为 0');
    return {
      total: 0,
      holidayCounts: {},
      queriedRanges: []
    };
  }

  const queriedRanges = [];
  for (let index = 0; index < ranges.length; index += 1) {
    const range = ranges[index];
    const ownerText = Array.isArray(range.owners) && range.owners.length > 0
      ? range.owners.map(owner => `${owner.source}:${owner.name}`).join(',')
      : 'unknown';
    console.log(`[ApiClient] XDR count 查询 ${index + 1}/${ranges.length}: ${range.start} ~ ${range.end} owners=${ownerText}`);
    const count = await fetchXdrLogSearchCount(cookieInfo, range);
    queriedRanges.push({
      start: range.start,
      end: range.end,
      startDate: range.startDate,
      endDate: range.endDate,
      owners: Array.isArray(range.owners) ? range.owners : [],
      count
    });
    const partialSummary = aggregateXdrLogSearchCountDetails(queriedRanges);
    console.log(`[ApiClient] XDR count 本段 ${count}，累计 ${partialSummary.total}`);

    const delayMs = Number.isFinite(params.pageDelayMs) ? params.pageDelayMs : 10;
    if (delayMs > 0 && index < ranges.length - 1) {
      await sleep(delayMs);
    }
  }

  const summary = aggregateXdrLogSearchCountDetails(queriedRanges);
  return {
    total: summary.total,
    holidayCounts: summary.holidayCounts,
    queriedRanges
  };
}

async function fetchXdrIn2outLogSearchCountForRanges(cookieInfo, ranges, params = {}) {
  if (!Array.isArray(ranges) || ranges.length === 0) {
    console.log('[ApiClient] XDR in2out 重保时间段为空，count 记为 0');
    return 0;
  }

  let total = 0;
  for (let index = 0; index < ranges.length; index += 1) {
    const range = ranges[index];
    console.log(`[ApiClient] XDR in2out count 查询 ${index + 1}/${ranges.length}: ${range.start} ~ ${range.end}`);
    const count = await fetchXdrIn2outLogSearchCount(cookieInfo, range);
    total += count;
    console.log(`[ApiClient] XDR in2out count 本段 ${count}，累计 ${total}`);

    const delayMs = Number.isFinite(params.pageDelayMs) ? params.pageDelayMs : 10;
    if (delayMs > 0 && index < ranges.length - 1) {
      await sleep(delayMs);
    }
  }
  return total;
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

module.exports = {
  API_CONFIG,
  buildRequestBody,
  buildAlarmRequestBody,
  buildVulnRequestBody,
  buildWeakPwdSummaryRequestBody,
  buildWeakPwdListRequestBody,
  buildTopnLoadConditionRequestBody,
  buildTopnDeviceListRequestBody,
  buildTopnRequestBody,
  buildAssetDownloadRequestBody,
  buildExposedTargetCompanyOptionRequestBody,
  buildExposedExportRequestBody,
  buildExposedIpListStatisticsRequestBody,
  buildXdrLogSearchCountRequestBody,
  buildXdrAccessDirectionLogSearchCountRequestBody,
  buildXdrIn2outLogSearchCountRequestBody,
  buildXdrOut2inLogSearchCountRequestBody,
  buildXdrRejectedExternalToInternalCountRequestBody,
  buildXdrIncidentTableCountRequestBody,
  buildXdrIncidentTableQueryCountRequestBody,
  buildXdrG103IncidentCountRequestBody,
  buildXdrG105IncidentCountRequestBody,
  buildXdrG106IncidentCountRequestBody,
  buildXdrG107IncidentCountRequestBody,
  buildXdrAlertTableCountRequestBody,
  buildCustomerBusinessUrl,
  buildHomeUrl,
  httpPost,
  httpGetBuffer,
  extractXlsxFromZipOrSelf,
  extractRowsFromResponse,
  getTotalFromResponse,
  normalizeHost,
  generateHeaders,
  getXdrBaseUrl,
  generateXdrHeaders,
  fetchEventTable,
  fetchAlarmTable,
  fetchVulnTable,
  fetchWeakPwdSummary,
  fetchWeakPwdSummaryTotal,
  buildOrderBranchRequestBody,
  fetchOrderBranchTotal,
  fetchTopnLoadCondition,
  fetchTopnDeviceList,
  fetchTopnThreatTypes,
  fetchTopnSrcIpGeos,
  fetchTopnDstIps,
  fetchTopnReportStats,
  visitHomePage,
  fetchExposedTargetCompanyOptions,
  fetchExposedSurfaceExportInfo,
  downloadExposedSurfaceFile,
  fetchExposedSurfaceFile,
  fetchExposedSurfaceIpListStatistics,
  visitCustomerBusinessPage,
  fetchAssetDownloadInfo,
  downloadAssetFile,
  fetchAssetTableFile,
  fetchAllPages,
  fetchXdrLogSearchCount,
  fetchXdrLogSearchCountForRanges,
  aggregateXdrLogSearchCountDetails,
  fetchXdrLogSearchCountDetailsForRanges,
  fetchXdrAccessDirectionLogSearchCount,
  fetchXdrIn2outLogSearchCount,
  fetchXdrOut2inLogSearchCount,
  fetchXdrRejectedExternalToInternalCount,
  fetchXdrIncidentTableCount,
  fetchXdrIncidentTableQueryCount,
  fetchXdrG103IncidentCount,
  fetchXdrG105IncidentCount,
  fetchXdrG106IncidentCount,
  fetchXdrG107IncidentCount,
  fetchXdrAlertTableCount,
  fetchXdrIn2outLogSearchCountForRanges,
  fetchCompanyPage,
  resolveCompanyIdByName
};
