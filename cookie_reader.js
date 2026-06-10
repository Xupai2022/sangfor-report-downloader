/**
 * 深信服报告下载 - Cookie 读取模块
 * 从 cookies.txt 文件中读取并解析 Cookie
 */

const fs = require('fs');
const path = require('path');

// 默认 Cookie 文件路径
const DEFAULT_COOKIE_PATH = path.join(__dirname, 'cookies.txt');

function resolveCookiePath(cookiePath = DEFAULT_COOKIE_PATH) {
  if (!fs.existsSync(cookiePath)) {
    throw new Error(`Cookie 文件不存在: ${cookiePath}\n请确保浏览器插件已运行并生成 Cookie 文件`);
  }

  const stat = fs.statSync(cookiePath);
  if (!stat.isDirectory()) {
    return cookiePath;
  }

  const candidates = fs.readdirSync(cookiePath)
    .map(name => path.join(cookiePath, name))
    .filter(filePath => {
      try {
        const fileStat = fs.statSync(filePath);
        return fileStat.isFile() && /\.(txt|json|cookie|cookies)$/i.test(filePath);
      } catch (e) {
        return false;
      }
    })
    .sort((a, b) => fs.statSync(b).mtimeMs - fs.statSync(a).mtimeMs);

  if (candidates.length === 0) {
    throw new Error(`Cookie 目录中没有找到 txt/json/cookie 文件: ${cookiePath}`);
  }

  console.log(`[CookieReader] Cookie 路径是目录，使用最新文件: ${candidates[0]}`);
  return candidates[0];
}

function cookiePairsToString(pairs) {
  return pairs
    .filter(item => item && item.name && item.value !== undefined)
    .map(item => `${item.name}=${item.value}`)
    .join('; ');
}

function normalizeBaseUrl(value) {
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

function normalizeCookieContent(rawContent) {
  const content = rawContent.trim();
  if (!content) {
    throw new Error('Cookie 文件内容为空');
  }

  if (content.startsWith('{') || content.startsWith('[')) {
    const parsed = JSON.parse(content);

    if (Array.isArray(parsed)) {
      return {
        cookieString: cookiePairsToString(parsed),
        csrfToken: null,
        xdrBaseUrl: ''
      };
    }

    const cookieString = parsed.cookie || parsed.cookieString || parsed.Cookie || parsed.cookiesText;
    const xdrBaseUrl = normalizeBaseUrl(parsed.xdrBaseUrl || parsed.baseUrl || parsed.domain);
    if (typeof cookieString === 'string' && cookieString.trim()) {
      return {
        cookieString: cookieString.trim(),
        csrfToken: parsed.csrfToken || parsed.xCsrftoken || parsed['x-csrftoken'] || null,
        xdrBaseUrl
      };
    }

    if (Array.isArray(parsed.cookies)) {
      return {
        cookieString: cookiePairsToString(parsed.cookies),
        csrfToken: parsed.csrfToken || parsed.xCsrftoken || parsed['x-csrftoken'] || null,
        xdrBaseUrl
      };
    }

    throw new Error('无法识别 JSON Cookie 格式，请提供 Cookie header 字符串或浏览器 cookie 数组');
  }

  return {
    cookieString: content,
    csrfToken: null,
    xdrBaseUrl: ''
  };
}

/**
 * 从文件中读取 Cookie 字符串
 * @param {string} cookiePath - Cookie 文件路径
 * @returns {string} Cookie 字符串
 */
function readCookieString(cookiePath = DEFAULT_COOKIE_PATH) {
  try {
    const resolvedPath = resolveCookiePath(cookiePath);
    const rawContent = fs.readFileSync(resolvedPath, 'utf-8');
    return normalizeCookieContent(rawContent).cookieString;
  } catch (error) {
    console.error(`[CookieReader] 读取 Cookie 文件失败: ${error.message}`);
    throw error;
  }
}

/**
 * 解析 Cookie 字符串为对象
 * @param {string} cookieString - 原始 Cookie 字符串
 * @returns {Object} 解析后的 Cookie 对象
 */
function parseCookieString(cookieString) {
  const cookies = {};
  
  if (!cookieString || typeof cookieString !== 'string') {
    return cookies;
  }
  
  cookieString.split(';').forEach(cookie => {
    const [name, ...valueParts] = cookie.trim().split('=');
    if (name) {
      cookies[name.trim()] = valueParts.join('=').trim();
    }
  });
  
  return cookies;
}

/**
 * 从 Cookie 中提取 X-Csrftoken
 * @param {string} cookieString - Cookie 字符串
 * @returns {string} X-Csrftoken 值
 */
function extractCsrfToken(cookieString) {
  const cookies = parseCookieString(cookieString);
  
  // 尝试多个可能的 token 字段名
  const tokenKeys = ['csrf_token', 'x-csrf-token', 'X-Csrftoken', 'csrftoken', '_csrf'];
  
  for (const key of tokenKeys) {
    if (cookies[key]) {
      console.log(`[CookieReader] 找到 ${key}: ${cookies[key].substring(0, 20)}...`);
      return cookies[key];
    }
  }
  
  throw new Error('无法从 Cookie 中找到 X-Csrftoken，请检查 Cookie 是否完整');
}

/**
 * 验证 Cookie 是否完整
 * @param {string} cookieString - Cookie 字符串
 * @returns {Object} { valid: boolean, missing: string[] }
 */
function validateCookie(cookieString) {
  // 必要的 Cookie 字段
  const requiredKeys = [
    'UEDC_LOGIN_LANGUAGE',
    'tmp-change-pwd', 
    'UEDC_LOGIN_POLICY_VALUE',
    'pre_soc_token',
    'LAST_LOGIN_TYPE',
    'USER_BLESS_META_NAME',
    'csrf_token',
    'login-source',
    'soc-token'
  ];
  
  const cookies = parseCookieString(cookieString);
  const missing = requiredKeys.filter(key => !cookies[key]);
  
  return {
    valid: missing.length === 0,
    missing
  };
}

/**
 * 获取完整的 Cookie 信息
 * @param {string} cookiePath - Cookie 文件路径
 * @returns {Object} { cookieString, csrfToken, cookies }
 */
function getCookieInfo(cookiePath = DEFAULT_COOKIE_PATH) {
  console.log(`[CookieReader] 从 ${cookiePath} 读取 Cookie`);
  
  const resolvedPath = resolveCookiePath(cookiePath);
  const normalized = normalizeCookieContent(fs.readFileSync(resolvedPath, 'utf-8'));
  const cookieString = normalized.cookieString;
  const cookies = parseCookieString(cookieString);
  const csrfToken = normalized.csrfToken || extractCsrfToken(cookieString);
  const validation = validateCookie(cookieString);
  
  console.log(`[CookieReader] Cookie 解析完成，共 ${Object.keys(cookies).length} 个字段`);
  console.log(`[CookieReader] Cookie 验证: ${validation.valid ? '通过' : '缺失 ' + validation.missing.join(', ')}`);
  
  return {
    cookieString,
    csrfToken,
    xdrBaseUrl: normalized.xdrBaseUrl || '',
    cookies,
    validation
  };
}

module.exports = {
  resolveCookiePath,
  normalizeCookieContent,
  normalizeBaseUrl,
  readCookieString,
  parseCookieString,
  extractCsrfToken,
  validateCookie,
  getCookieInfo,
  DEFAULT_COOKIE_PATH
};
