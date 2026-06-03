/**
 * SOAR 事件/告警数据转换器
 * 对齐 mss_ai_ppt_sample_assets 的 mongo_collector + event_transformer 处理逻辑
 */

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const LATEST_THREAT_TYPE = '最新威胁';

const EVENT_MANAGE_TYPE_DISPLAY = {
  INTRANET_THREAT: '内部威胁',
  MANAGEMENT: '管理要求',
  ACTIVE_INVOLVEMENT: '主动投入',
  SERVICE_SPREAD: '服务蔓延',
  UNDECLARED_THREAT: '未公开威胁',
  INTERNET_THREAT: '外部威胁',
  VULNERABILITY: '脆弱性',
  CUSTOMER_FEEDBACK: '用户反馈',
  STRATEGY_OPTIMIZE: '策略调优',
  EMERGENCY_RESPONSE: '应急工单',
  STRATEGY_SERVICE: '策略工单',
  OTHER: '其他',
  CONSULTANT: '咨询问题',
  PRODUCTION: '产品问题',
  '94': '脆弱性风险',
  '214': '访问风险',
  '201': '服务探测',
  '215': '主机探测',
  '90': '网站攻击',
  '203': '后门通信',
  '204': '账号爆破',
  '205': '攻击利用',
  '96': '邮件攻击',
  '10': 'Dos攻击',
  '207': '黑链',
  '30': '漏洞攻击',
  '208': '黑客工具',
  '213': '数据库攻击利用',
  '40': '访问恶意文件',
  '209': '感染病毒',
  '212': '行为异常',
  '216': '流量异常',
  '217': '登录异常',
  '218': '终端行为异常',
  '219': '容器异常'
};

const EVENT_EVENT_STATUS_DISPLAY = {
  inited: '未处置',
  finished: '已闭环',
  disposal: '处置中',
  suspend: '定期跟进',
  accept_risk: '不处置',
  announced: '已通告',
  protected: '已防护',
  new_link_alarm: '有告警更新'
};

const EVENT_SERVICE_STATUS_DISPLAY = {
  0: '服务内（7*24H）',
  1: '服务外',
  3: '服务内（5*8H）',
  '-1': '全部服务'
};

const EVENT_GRADING_TAG_DISPLAY = {
  0: '重大事件',
  1: '重要事件',
  2: '一般事件',
  3: '重要威胁',
  4: '一般威胁',
  5: '其他'
};

const STRATEGY_OPTIMIZE_MANAGE_SUB_TYPES = new Set([
  'STRATEGY_OPTIMIZE',
  'EDR_STRATEGY_OPTIMIZE',
  'SIP_STRATEGY_OPTIMIZE',
  'AF_STRATEGY_OPTIMIZE',
  'TSS_STRATEGY_OPTIMIZE',
  'CWPP_STRATEGY_OPTIMIZE',
  'NTA_STRATEGY_OPTIMIZE'
]);

const EVENT_SCENE_STAT_NAMES = {
  penetrationFramework: '渗透框架',
  trojan: '木马',
  proxyTool: '代理工具',
  mining: '挖矿',
  accountSecurity: '账号安全'
};

const EVENT_GRADING_TAG_5_MANAGE_TYPE_RULES = {
  INTRANET_THREAT: '一般事件',
  EMERGENCY_RESPONSE: '重大事件',
  UNDECLARED_THREAT: '最新威胁',
  '209': '一般事件',
  INTERNET_THREAT: '一般威胁',
  VULNERABILITY: '重大威胁'
};

const PUSH_STATUS_DISPLAY = {
  1: '已通告',
  '-1': '未通告'
};

const ALARM_EVENT_STATUS_DISPLAY = {
  inited: '未处置',
  finished: '已完成',
  reject_ignore: '已忽略（驳回）',
  rejected: '已驳回',
  disposal: '处置中',
  link_event: '已关联告警',
  relate_event: '生成事件',
  ignore: '已忽略',
  white: '已加白',
  auto_event: '生成事件（自动）',
  auto_ignore: '已忽略（自动）'
};

const ATTACK_STATE_DISPLAY = {
  0: 'judge（待研判）',
  1: 'fail（失败）',
  2: 'succ（成功）',
  3: 'compromised（失陷）',
  4: 'anomaly（异常）',
  5: 'attempt（尝试）'
};

const ATTACK_DIRECTION_DISPLAY = {
  0: '未知',
  1: '内-外',
  2: '外-内',
  3: '内-内',
  4: '外-外'
};

const ALARM_SERVICE_STATUS_DISPLAY = {
  0: '服务内（7*24H）',
  1: '服务外',
  3: '服务内（5*8H）',
  '-1': '全部服务'
};

const REJECT_REASON_DISPLAY = {
  '1': '业务触发',
  '2': '技术误报',
  '3': '接受风险',
  '4': '误推送',
  '5': '正报延迟'
};

const EVENT_OUTPUT_FIELDS = [
  'event_grading_tag',
  'create_time',
  'type',
  'event_name',
  'host_ip',
  '内网外网资产',
  'event_status',
  'service_status',
  'latest_time',
  'checkout_time',
  'dispose_time',
  'contain_time',
  'finished_time',
  'incidence',
  'update_protected_time',
  'affected_assets',
  'update_announced_time',
  'update_accept_risk_time',
  'push_status',
  'wechat_push_time',
  '识别时长',
  '响应时长',
  '遏制时长',
  '处置时长',
  '闭环时长'
];

const ALARM_OUTPUT_FIELDS = [
  'alarm_name',
  'host_ip',
  'type',
  'event_status',
  'first_time',
  'latest_time',
  'service_status',
  'create_time',
  'attack_state',
  'attack_direction',
  'current_operate_time',
  'rejected_event_id',
  'current_operator',
  'reject_reason'
];

const VULN_OUTPUT_FIELDS = [
  '漏洞名称',
  '漏洞等级',
  '是否可利用',
  '受影响主机/位置',
  'IP',
  '内/外网',
  '跟进状态',
  '更新时间'
];

const VULN_STATUS_DISPLAY = {
  1: '处置中',
  2: '已修复',
  3: '接受风险',
  4: '已防护',
  5: '未审核',
  6: '超时未审核',
  7: '超时未修复',
  8: '已标为误判',
  9: '修复失败',
  10: '已搁置',
  11: '已加白'
};

const VULN_LEVEL_DISPLAY = {
  0: '低危',
  1: '中危',
  2: '高危'
};

function loadManageSubTypeMap(filePath) {
  if (!filePath || !fs.existsSync(filePath)) return {};
  try {
    const payload = JSON.parse(fs.readFileSync(filePath, 'utf8'));
    const items = ((((payload || {}).data || {}).manage_sub_type) || []);
    const map = {};
    for (const item of items) {
      if (item && item.value !== undefined && item.name !== undefined) {
        map[String(item.value)] = String(item.name);
      }
    }
    return map;
  } catch (error) {
    console.warn(`[Transformer] manage_sub_type 映射加载失败: ${error.message}`);
    return {};
  }
}

function loadAssetSecurityDomainMap(filePath) {
  if (!filePath || !fs.existsSync(filePath)) return {};
  const ext = path.extname(filePath).toLowerCase();
  if (ext === '.xlsx' || ext === '.xls') {
    return loadAssetSecurityDomainMapFromExcel(filePath);
  }
  try {
    const payload = JSON.parse(fs.readFileSync(filePath, 'utf8'));
    const map = {};
    if (Array.isArray(payload)) {
      for (const item of payload) {
        if (!item) continue;
        const ip = item['*IP/URL'];
        const domain = item['安全域'];
        if (ip) {
          for (const candidateIp of extractHostIpCandidates(String(ip))) {
            map[candidateIp] = domain ? String(domain).trim() : '';
          }
        }
      }
      return map;
    }
    if (payload && typeof payload === 'object') {
      for (const [ip, domain] of Object.entries(payload)) {
        for (const candidateIp of extractHostIpCandidates(String(ip))) {
          map[candidateIp] = domain ? String(domain).trim() : '';
        }
      }
    }
    return map;
  } catch (error) {
    console.warn(`[Transformer] 资产安全域映射加载失败: ${error.message}`);
    return {};
  }
}

function addAssetSecurityDomainMapping(map, ipValue, domainValue) {
  if (!ipValue) return;
  const domain = domainValue ? String(domainValue).trim() : '';
  for (const candidateIp of extractHostIpCandidates(String(ipValue))) {
    map[candidateIp] = domain;
  }
}

function loadAssetSecurityDomainMapFromSheet(ws) {
  const map = {};
  if (!ws) return map;

  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const headerRowIndex = 1;
  const headerRow = rows[headerRowIndex] || [];
  const ipIndex = headerRow.findIndex(value => value === '*IP/URL');
  const domainIndex = headerRow.findIndex(value => value === '安全域');
  if (ipIndex < 0 || domainIndex < 0) return map;

  for (const row of rows.slice(headerRowIndex + 1)) {
    addAssetSecurityDomainMapping(map, row[ipIndex], row[domainIndex]);
  }

  return map;
}

function loadAssetSecurityDomainMapFromWorkbook(wb) {
  const map = {};
  if (!wb || !Array.isArray(wb.SheetNames)) return map;
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    Object.assign(map, loadAssetSecurityDomainMapFromSheet(ws));
  }
  return map;
}

function loadAssetSecurityDomainMapFromBuffer(buffer) {
  if (!buffer) return {};
  try {
    const wb = XLSX.read(buffer, { type: Buffer.isBuffer(buffer) ? 'buffer' : 'array' });
    return loadAssetSecurityDomainMapFromWorkbook(wb);
  } catch (error) {
    console.warn(`[Transformer] 资产 Excel buffer 映射加载失败: ${error.message}`);
    return {};
  }
}

function loadAssetSecurityDomainMapFromExcel(filePath) {
  try {
    const wb = XLSX.readFile(filePath);
    return loadAssetSecurityDomainMapFromWorkbook(wb);
  } catch (error) {
    console.warn(`[Transformer] 资产 Excel 映射加载失败: ${error.message}`);
    return {};
  }
}

function stringifyEnumKey(value) {
  if (value === null || value === undefined) return null;
  return String(value);
}

function decodeUnicodeEscapes(value) {
  if (value === null || value === undefined) return '';
  const text = String(value).trim();
  if (!text) return '';
  return text.replace(/\\u([0-9a-fA-F]{4})/g, (_, hex) => String.fromCharCode(parseInt(hex, 16)));
}

function pickLastArrayValue(value) {
  if (Array.isArray(value)) {
    if (value.length === 0) return '';
    const last = value[value.length - 1];
    return last === null || last === undefined ? '' : last;
  }
  return value === null || value === undefined ? '' : value;
}

function formatAlarmEventStatus(value) {
  const key = stringifyEnumKey(value);
  if (key === null) return '';
  return ALARM_EVENT_STATUS_DISPLAY[key] ?? value;
}

function formatAlarmServiceStatus(value) {
  const key = stringifyEnumKey(value);
  if (key === null) return '';
  return ALARM_SERVICE_STATUS_DISPLAY[key] ?? value;
}

function formatAttackState(value) {
  const key = stringifyEnumKey(value);
  if (key === null) return '';

  const attackStateDisplay = {
    '0': '待研判',
    '1': '失败',
    '2': '成功',
    '3': '失陷',
    '4': '异常',
    '5': '尝试',
    judge: '待研判',
    fail: '失败',
    succ: '成功',
    compromised: '失陷',
    anomaly: '异常',
    attempt: '尝试'
  };

  return attackStateDisplay[key] ?? value;
}

function formatAttackDirection(value) {
  const key = stringifyEnumKey(value);
  if (key === null) return '';
  return ATTACK_DIRECTION_DISPLAY[key] ?? value;
}

function formatRejectReason(value) {
  const key = stringifyEnumKey(value);
  if (key === null) return '';
  return REJECT_REASON_DISPLAY[key] ?? value;
}

function buildTypeDisplay(primary, secondary, separator) {
  const primaryText = primary === null || primary === undefined || primary === '' ? null : String(primary);
  const secondaryText = secondary === null || secondary === undefined || secondary === '' ? null : String(secondary);
  if (primaryText && secondaryText) return `${primaryText}${separator}${secondaryText}`;
  return primaryText || secondaryText || '';
}

function resolveManageTypeDisplay(eventDoc) {
  const manageTypeDisplay = decodeUnicodeEscapes(eventDoc.manage_type_cn);
  if (manageTypeDisplay) return manageTypeDisplay;
  const manageType = stringifyEnumKey(eventDoc.manage_type);
  return EVENT_MANAGE_TYPE_DISPLAY[manageType] ?? (manageType || '');
}

function isUndeclaredThreatManageType(eventDoc) {
  const manageType = stringifyEnumKey(eventDoc.manage_type);
  const manageTypeDisplay = resolveManageTypeDisplay(eventDoc);
  return manageType === 'UNDECLARED_THREAT' || manageTypeDisplay.includes('未公开威胁');
}

function parseDateTime(value) {
  if (value instanceof Date) return value;
  if (typeof value === 'number') {
    const ms = value > 1e12 ? value : value * 1000;
    const d = new Date(ms);
    return Number.isNaN(d.getTime()) ? null : d;
  }
  if (typeof value !== 'string') return null;
  const text = value.trim();
  if (!text) return null;

  let parsed = new Date(text);
  if (!Number.isNaN(parsed.getTime())) return parsed;

  parsed = new Date(text.replace(/\//g, '-'));
  if (!Number.isNaN(parsed.getTime())) return parsed;

  return null;
}

function isLatestThreat(tagValue) {
  return typeof tagValue === 'string' && tagValue === LATEST_THREAT_TYPE;
}

function isEventCategoryTag(tagValue) {
  const text = String(tagValue === null || tagValue === undefined ? '' : tagValue).trim();
  if (!text) return false;
  return ['重大事件', '重要事件', '一般事件'].includes(text) || text.includes('事件');
}

function createEventSceneStatSets(manageSubTypeMap) {
  const sets = {};
  for (const key of Object.keys(EVENT_SCENE_STAT_NAMES)) {
    sets[key] = new Set();
  }

  for (const [value, name] of Object.entries(manageSubTypeMap || {})) {
    for (const [key, targetName] of Object.entries(EVENT_SCENE_STAT_NAMES)) {
      if (String(name || '').trim() === targetName) {
        sets[key].add(String(value));
      }
    }
  }

  return sets;
}

function createEmptyEventSceneStats() {
  return Object.fromEntries(Object.keys(EVENT_SCENE_STAT_NAMES).map(key => [key, 0]));
}

function pickFirstDateTime(...values) {
  for (const value of values) {
    const dt = parseDateTime(value);
    if (dt) return dt;
  }
  return null;
}

function minutesBetween(later, earlier) {
  return Math.round(((later - earlier) / 60000) * 100) / 100;
}

function calcRecognitionDurationMinutes(eventDoc) {
  if (isLatestThreat(eventDoc.event_grading_tag)) return '';
  const createTime = parseDateTime(eventDoc.create_time);
  const checkoutTime = parseDateTime(eventDoc.checkout_time);
  if (!createTime || !checkoutTime) return '';
  return minutesBetween(createTime, checkoutTime);
}

function calcAccessDurationMinutes(eventDoc) {
  if (isLatestThreat(eventDoc.event_grading_tag)) return '';
  const createTime = parseDateTime(eventDoc.create_time);
  const accessTime = pickFirstDateTime(eventDoc.wechat_push_time, eventDoc.dispose_time, eventDoc.latest_time);
  if (!createTime || !accessTime) return '';
  return minutesBetween(accessTime, createTime);
}

function calcContainmentDurationMinutes(eventDoc) {
  if (isLatestThreat(eventDoc.event_grading_tag)) return '';
  const createTime = parseDateTime(eventDoc.create_time);
  const pushTime = parseDateTime(eventDoc.wechat_push_time);
  if (!createTime || !pushTime) return '';
  return minutesBetween(pushTime, createTime);
}

function calcDisposalDurationMinutes(eventDoc) {
  if (isLatestThreat(eventDoc.event_grading_tag)) return '';
  const createTime = parseDateTime(eventDoc.create_time);
  const disposalTime = pickFirstDateTime(eventDoc.contain_time, eventDoc.latest_time);
  if (!createTime || !disposalTime) return '';
  return minutesBetween(disposalTime, createTime);
}

function calcClosedLoopDurationMinutes(eventDoc) {
  if (isLatestThreat(eventDoc.event_grading_tag)) return '';
  const createTime = parseDateTime(eventDoc.create_time);
  const closedTime = pickFirstDateTime(eventDoc.finished_time, eventDoc.latest_time);
  if (!createTime || !closedTime) return '';
  return minutesBetween(closedTime, createTime);
}

function splitHostIpText(value) {
  const ipv4Matches = String(value).match(/\b(?:(?:25[0-5]|2[0-4]\d|1?\d?\d)\.){3}(?:25[0-5]|2[0-4]\d|1?\d?\d)\b/g);
  if (ipv4Matches) {
    return Array.from(new Set(ipv4Matches));
  }

  let normalized = value;
  for (const delimiter of ['，', ';', '；', '/', '|', '\n', '\t']) {
    normalized = normalized.replaceAll(delimiter, ',');
  }
  normalized = normalized.trim().replace(/\s+/g, ',');
  const seen = new Set();
  const result = [];
  for (const token of normalized.split(',')) {
    const ip = token.trim();
    if (!ip || seen.has(ip)) continue;
    seen.add(ip);
    result.push(ip);
  }
  return result;
}

function extractHostIpCandidates(value) {
  if (value === null || value === undefined || value === '') return [];
  if (typeof value === 'string') return splitHostIpText(value);
  if (Array.isArray(value)) {
    return value.flatMap(extractHostIpCandidates);
  }
  return [String(value).trim()].filter(Boolean);
}

function resolveEventSecurityDomain(hostIpValue, securityDomainByIp) {
  const domains = [];
  const seen = new Set();
  for (const ip of extractHostIpCandidates(hostIpValue)) {
    const domain = securityDomainByIp[ip];
    if (!domain || seen.has(domain)) continue;
    seen.add(domain);
    domains.push(domain);
  }
  return domains.join('、');
}

function extractIpFromAsset(value) {
  const candidates = extractHostIpCandidates(value);
  return candidates.length > 0 ? candidates[0] : '';
}

function resolveSecurityDomainByIp(ip, securityDomainByIp) {
  if (!ip) return '';
  return securityDomainByIp[ip] || '';
}

function formatVulnStatus(value) {
  const key = stringifyEnumKey(value);
  if (key === null) return '';
  return VULN_STATUS_DISPLAY[key] ?? value;
}

function formatVulnLevel(value) {
  const key = stringifyEnumKey(value);
  if (key === null) return '';
  return VULN_LEVEL_DISPLAY[key] ?? value;
}

function formatVulnHighAvailability(value) {
  const key = stringifyEnumKey(value);
  if (key === '1') return '是';
  if (key === '0') return '否';
  return '';
}

function resolveEventGradingTag5Rule(eventDoc) {
  if (isUndeclaredThreatManageType(eventDoc)) {
    return { ignore: false, eventGradingTag: LATEST_THREAT_TYPE, strategyOptimizeCount: 0 };
  }

  if (stringifyEnumKey(eventDoc.event_grading_tag) !== '5') {
    return { ignore: false, strategyOptimizeCount: 0 };
  }

  const manageType = stringifyEnumKey(eventDoc.manage_type);
  const manageSubType = stringifyEnumKey(eventDoc.manage_sub_type);

  if (manageType === 'MANAGEMENT') {
    return { ignore: true, strategyOptimizeCount: 0 };
  }

  if (manageType === 'STRATEGY_OPTIMIZE') {
    if (manageSubType === 'LOG_CHECK_EXCEPTION') {
      return { ignore: true, strategyOptimizeCount: 0 };
    }
    if (manageSubType === 'DEVICE_OFFLINE') {
      return {
        ignore: false,
        eventGradingTag: '业务连续性风险',
        strategyOptimizeCount: 0
      };
    }
    if (STRATEGY_OPTIMIZE_MANAGE_SUB_TYPES.has(manageSubType)) {
      return { ignore: false, strategyOptimizeCount: 1 };
    }
  }

  const eventGradingTag = EVENT_GRADING_TAG_5_MANAGE_TYPE_RULES[manageType];
  if (eventGradingTag) {
    return { ignore: false, eventGradingTag, strategyOptimizeCount: 0 };
  }

  return { ignore: false, strategyOptimizeCount: 0 };
}

function transformEventDoc(eventDoc, ctx, overrides = {}) {
  const manageTypeDisplay = resolveManageTypeDisplay(eventDoc);
  const manageSubTypeDisplay = decodeUnicodeEscapes(eventDoc.manage_sub_type_cn);
  const rawTag = EVENT_GRADING_TAG_DISPLAY[eventDoc.event_grading_tag] ?? eventDoc.event_grading_tag;
  const eventGradingTag = overrides.eventGradingTag ?? rawTag;

  const transformed = {
    event_grading_tag: eventGradingTag,
    event_name: eventDoc.event_name ?? '',
    create_time: eventDoc.create_time,
    type: buildTypeDisplay(manageTypeDisplay, manageSubTypeDisplay, '->'),
    host_ip: eventDoc.host_ip,
    affected_assets: eventDoc.affected_assets,
    内网外网资产: resolveEventSecurityDomain(eventDoc.host_ip, ctx.securityDomainByIp),
    event_status: EVENT_EVENT_STATUS_DISPLAY[eventDoc.event_status] ?? eventDoc.event_status,
    service_status: EVENT_SERVICE_STATUS_DISPLAY[stringifyEnumKey(eventDoc.service_status)] ?? eventDoc.service_status,
    latest_time: eventDoc.latest_time,
    checkout_time: eventDoc.checkout_time,
    dispose_time: eventDoc.dispose_time,
    contain_time: eventDoc.contain_time,
    finished_time: eventDoc.finished_time,
    incidence: eventDoc.incidence,
    update_protected_time: eventDoc.update_protected_time,
    update_announced_time: eventDoc.update_announced_time,
    update_accept_risk_time: eventDoc.update_accept_risk_time,
    push_status: PUSH_STATUS_DISPLAY[stringifyEnumKey(eventDoc.push_status)] ?? eventDoc.push_status,
    wechat_push_time: eventDoc.wechat_push_time
  };

  Object.defineProperty(transformed, '__manageSubTypeDisplay', {
    value: manageSubTypeDisplay,
    enumerable: false,
    configurable: true
  });

  transformed['识别时长'] = calcRecognitionDurationMinutes({ ...eventDoc, event_grading_tag: transformed.event_grading_tag });
  transformed['响应时长'] = calcAccessDurationMinutes({ ...eventDoc, event_grading_tag: transformed.event_grading_tag });
  transformed['遏制时长'] = calcContainmentDurationMinutes({ ...eventDoc, event_grading_tag: transformed.event_grading_tag });
  transformed['处置时长'] = calcDisposalDurationMinutes({ ...eventDoc, event_grading_tag: transformed.event_grading_tag });
  transformed['闭环时长'] = calcClosedLoopDurationMinutes({ ...eventDoc, event_grading_tag: transformed.event_grading_tag });

  return transformed;
}

function transformAlarmDoc(alarmDoc, ctx) {
  const manageTypeDisplay = decodeUnicodeEscapes(alarmDoc.manage_type_cn);
  const manageSubTypeDisplay = decodeUnicodeEscapes(alarmDoc.manage_sub_type_cn);

  return {
    alarm_name: alarmDoc.title ?? alarmDoc.alarm_name ?? '',
    host_ip: alarmDoc.host_ip ?? '',
    type: buildTypeDisplay(manageTypeDisplay, manageSubTypeDisplay, '>'),
    event_status: formatAlarmEventStatus(alarmDoc.event_status),
    first_time: alarmDoc.first_time ?? '',
    latest_time: alarmDoc.latest_time ?? '',
    service_status: formatAlarmServiceStatus(alarmDoc.service_status),
    create_time: alarmDoc.create_time ?? '',
    attack_state: formatAttackState(alarmDoc.attack_state),
    attack_direction: formatAttackDirection(alarmDoc.attack_direction),
    current_operate_time: alarmDoc.current_operate_time ?? '',
    rejected_event_id: alarmDoc.rejected_event_id ?? '',
    current_operator: pickLastArrayValue(alarmDoc.current_operator),
    reject_reason: formatRejectReason(alarmDoc.reject_reason)
  };
}

function transformVulnDoc(vulnDoc, ctx) {
  const asset = vulnDoc.asset;
  const ip = extractIpFromAsset(asset);
  const secondLevels = Array.isArray(vulnDoc.second_level) && vulnDoc.second_level.length > 0
    ? vulnDoc.second_level
    : [null];

  return secondLevels.map((secondLevel) => ({
    漏洞名称: vulnDoc.name,
    漏洞等级: formatVulnLevel(vulnDoc.level),
    是否可利用: formatVulnHighAvailability(vulnDoc.is_high_availability),
    '受影响主机/位置': asset,
    IP: ip,
    '内/外网': resolveSecurityDomainByIp(ip, ctx.securityDomainByIp),
    跟进状态: secondLevel && typeof secondLevel === 'object' ? formatVulnStatus(secondLevel.vuln_status) : '',
    更新时间: secondLevel && typeof secondLevel === 'object' ? secondLevel.last_time ?? '' : ''
  }));
}

function transformEventDocs(eventDocs, options = {}) {
  return transformEventDocsWithStats(eventDocs, options).rows;
}

function transformEventDocsWithStats(eventDocs, options = {}) {
  const ctx = {
    securityDomainByIp: options.assetWorkbookBuffer
      ? loadAssetSecurityDomainMapFromBuffer(options.assetWorkbookBuffer)
      : loadAssetSecurityDomainMap(options.assetMapFile)
  };

  const stats = {
    strategyOptimizeCount: 0,
    sceneCounts: createEmptyEventSceneStats(),
    accountSecurityEventCountForE80: 0
  };
  const rows = [];

  for (const doc of eventDocs || []) {
    const rule = resolveEventGradingTag5Rule(doc);
    if (rule.ignore) continue;
    stats.strategyOptimizeCount += rule.strategyOptimizeCount || 0;
    const transformed = transformEventDoc(doc, ctx, {
      eventGradingTag: rule.eventGradingTag
    });
    rows.push(transformed);

    const manageSubTypeDisplay = decodeUnicodeEscapes(doc.manage_sub_type_cn);
    if (manageSubTypeDisplay.includes('账号安全')) {
      stats.accountSecurityEventCountForE80 += 1;
    }

    if (isEventCategoryTag(transformed.event_grading_tag)) {
      for (const [key, targetName] of Object.entries(EVENT_SCENE_STAT_NAMES)) {
        if (manageSubTypeDisplay === targetName) {
          stats.sceneCounts[key] += 1;
        }
      }
    }
  }

  return { rows, stats };
}

function transformAlarmDocs(alarmDocs, options = {}) {
  const ctx = {};
  return (alarmDocs || []).map(doc => transformAlarmDoc(doc, ctx));
}

function transformVulnDocs(vulnDocs, options = {}) {
  const ctx = {
    securityDomainByIp: options.assetWorkbookBuffer
      ? loadAssetSecurityDomainMapFromBuffer(options.assetWorkbookBuffer)
      : loadAssetSecurityDomainMap(options.assetMapFile)
  };
  return (vulnDocs || [])
    .flatMap(doc => transformVulnDoc(doc, ctx));
}

module.exports = {
  EVENT_OUTPUT_FIELDS,
  ALARM_OUTPUT_FIELDS,
  VULN_OUTPUT_FIELDS,
  transformEventDocs,
  transformEventDocsWithStats,
  transformAlarmDocs,
  transformVulnDocs,
  loadManageSubTypeMap,
  loadAssetSecurityDomainMap,
  loadAssetSecurityDomainMapFromBuffer,
  extractIpFromAsset
};
