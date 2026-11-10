/*
 * 打卡无效记录筛查 —— 核心解析/判定逻辑
 * 纯函数，无 DOM 依赖，可在浏览器与 Node 中复用。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.DashboardCore = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var CUTOFF_SECONDS = 9 * 3600; // 09:00:00
  var EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30); // Excel 1900 日期系统
  var PLACEHOLDERS = [
    '-', '--', '---', '—', '——', '–', '/', '\\', '无', '空', '未打卡',
    '未签退', '未签到', '待打卡', 'n/a', 'na', 'null', 'none'
  ];

  function pad2(n) {
    return (n < 10 ? '0' : '') + n;
  }

  function trimText(value) {
    return String(value == null ? '' : value).replace(/\u00a0/g, ' ').trim();
  }

  function isPlaceholderText(text) {
    var t = trimText(text).toLowerCase();
    return PLACEHOLDERS.indexOf(t) !== -1;
  }

  /** 把 ExcelJS 的单元格值（含公式、富文本、超链接）还原为基础值 */
  function normalizeCellValue(value) {
    if (value == null) return null;
    if (value instanceof Date) return value;
    if (typeof value === 'object') {
      if (Array.isArray(value.richText)) {
        return value.richText.map(function (part) { return part && part.text ? part.text : ''; }).join('');
      }
      if (Object.prototype.hasOwnProperty.call(value, 'result')) {
        return normalizeCellValue(value.result);
      }
      if (Object.prototype.hasOwnProperty.call(value, 'error')) {
        return String(value.error);
      }
      if (typeof value.text === 'string') return value.text;
    }
    return value;
  }

  /** 判断数字格式里是否包含日期（年/日）标记 */
  function formatHasDate(numFmt) {
    if (!numFmt) return false;
    var f = String(numFmt)
      .replace(/"[^"]*"/g, '')
      .replace(/\\./g, '')
      .replace(/\[[^\]]*\]/g, '')
      .replace(/_.?/g, '')
      .replace(/\*.?/g, '');
    return /[yd]/i.test(f);
  }

  function displayFromParts(parts) {
    var time = pad2(parts.h) + ':' + pad2(parts.mi) + ':' + pad2(parts.s);
    if (!parts.hasDate) return time;
    return parts.y + '-' + pad2(parts.mo) + '-' + pad2(parts.d) + ' ' + time;
  }

  function makeTime(parts, raw) {
    parts.seconds = parts.h * 3600 + parts.mi * 60 + parts.s;
    parts.display = displayFromParts(parts);
    return {
      kind: 'time',
      seconds: parts.seconds,
      hasDate: !!parts.hasDate,
      display: parts.display,
      parts: parts,
      raw: raw == null ? parts.display : String(raw)
    };
  }

  function blankResult(raw) {
    return {
      kind: 'blank',
      seconds: null,
      hasDate: false,
      display: '',
      parts: null,
      raw: raw == null ? '' : String(raw)
    };
  }

  function invalidResult(raw, display) {
    return {
      kind: 'invalid',
      seconds: null,
      hasDate: false,
      display: display == null ? trimText(raw) : String(display),
      parts: null,
      raw: raw == null ? '' : String(raw)
    };
  }

  function safeInt(value) {
    var n = parseInt(value, 10);
    return isNaN(n) ? null : n;
  }

  /** 解析纯时间字符串，兼容 08:58、9:05:30、2026-11-09 09:15:00、上午9点05分 等 */
  function parseTimeString(input) {
    var original = trimText(input);
    if (!original) return blankResult(original);

    var s = original
      .replace(/[\uFF10-\uFF19]/g, function (ch) {
        return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0);
      })
      .replace(/\uFF1A/g, ':')
      .replace(/\u3000/g, ' ');

    // 日期部分（可选）
    var dateMatch = s.match(/(\d{4})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})\s*日?/);
    var dateParts = null;
    if (dateMatch) {
      var y = safeInt(dateMatch[1]);
      var mo = safeInt(dateMatch[2]);
      var d = safeInt(dateMatch[3]);
      if (y !== null && mo !== null && d !== null && mo >= 1 && mo <= 12 && d >= 1 && d <= 31) {
        var probe = new Date(Date.UTC(y, mo - 1, d));
        if (probe.getUTCFullYear() === y && probe.getUTCMonth() === mo - 1 && probe.getUTCDate() === d) {
          dateParts = { y: y, mo: mo, d: d };
        }
      }
    }

    // 优先匹配 hh:mm[:ss]，其次匹配中文 hh点mm分
    var tm = s.match(/(上午|下午|凌晨|早上|早晨|晚上|中午|AM|PM)?\s*(\d{1,2})\s*:\s*(\d{1,2})(?:\s*:\s*(\d{1,2}))?\s*(上午|下午|凌晨|早上|早晨|晚上|中午|AM|PM)?/i);
    if (!tm) {
      tm = s.match(/(上午|下午|凌晨|早上|早晨|晚上|中午)?\s*(\d{1,2})\s*[点时]\s*(?:(\d{1,2})\s*分?)?\s*(上午|下午|凌晨|早上|早晨|晚上|中午)?/);
    }
    if (!tm) return invalidResult(original, original);

    var h = safeInt(tm[2]);
    var mi = tm[3] === undefined ? 0 : safeInt(tm[3]);
    var sec = tm[4] === undefined ? 0 : safeInt(tm[4]);
    if (h === null || mi === null || sec === null || h > 23 || mi > 59 || sec > 59) {
      return invalidResult(original, original);
    }

    var marker = (tm[1] || tm[5] || '').toUpperCase();
    var isPm = marker === 'PM' || /下午|晚上/.test(marker);
    var isAm = marker === 'AM' || /上午|早上|早晨|凌晨/.test(marker);
    var isNoon = /中午/.test(marker);
    if (isPm && h < 12) h += 12;
    if (isAm && h === 12) h = 0;
    if (isNoon && h < 11) h += 12;

    return makeTime({
      y: dateParts ? dateParts.y : 1899,
      mo: dateParts ? dateParts.mo : 12,
      d: dateParts ? dateParts.d : 30,
      h: h,
      mi: mi,
      s: sec,
      hasDate: !!dateParts
    }, original);
  }

  /** 解析 Excel 日期/时间序列号 */
  function parseSerial(serial, numFmt) {
    if (!isFinite(serial) || serial < 0) return invalidResult(String(serial), String(serial));
    var totalSeconds = Math.round(serial * 86400);
    var days = Math.floor(totalSeconds / 86400);
    var secOfDay = totalSeconds - days * 86400;
    var date = new Date(EXCEL_EPOCH_MS + days * 86400000);
    var y = date.getUTCFullYear();
    var mo = date.getUTCMonth() + 1;
    var d = date.getUTCDate();
    var h = Math.floor(secOfDay / 3600);
    var mi = Math.floor((secOfDay % 3600) / 60);
    var s = secOfDay % 60;

    var hasDate;
    if (y <= 1900) {
      // 1899/1900 的“日期”是 Excel 存放纯时间的方式，不应带出日期
      hasDate = false;
    } else if (formatHasDate(numFmt)) {
      hasDate = true;
    } else if (numFmt && String(numFmt).toLowerCase() !== 'general') {
      hasDate = false;
    } else {
      hasDate = serial >= 1;
    }

    return makeTime({ y: y, mo: mo, d: d, h: h, mi: mi, s: s, hasDate: hasDate }, String(serial));
  }

  function parseDateValue(date, numFmt) {
    if (isNaN(date.getTime())) return invalidResult(String(date), String(date));
    var y = date.getUTCFullYear();
    var mo = date.getUTCMonth() + 1;
    var d = date.getUTCDate();
    var h = date.getUTCHours();
    var mi = date.getUTCMinutes();
    var s = date.getUTCSeconds();

    var hasDate;
    if (y <= 1900) {
      hasDate = false;
    } else if (formatHasDate(numFmt)) {
      hasDate = true;
    } else if (numFmt && String(numFmt).toLowerCase() !== 'general') {
      hasDate = false;
    } else {
      hasDate = y > 1900;
    }

    return makeTime({ y: y, mo: mo, d: d, h: h, mi: mi, s: s, hasDate: hasDate }, date.toISOString());
  }

  /**
   * 解析一个“签到/签退时间”单元格值。
   * @param {*} value 原始值（可先经 normalizeCellValue）
   * @param {string} numFmt Excel 数字格式
   * @param {{placeholderAsBlank?: boolean}} options
   */
  function parseTimeValue(value, numFmt, options) {
    options = options || {};
    var placeholderAsBlank = options.placeholderAsBlank !== false;
    var v = normalizeCellValue(value);

    if (v == null) return blankResult('');
    if (typeof v === 'string') {
      var text = trimText(v);
      if (!text) return blankResult('');
      if (placeholderAsBlank && isPlaceholderText(text)) return blankResult(text);
      return parseTimeString(text);
    }
    if (v instanceof Date) return parseDateValue(v, numFmt);
    if (typeof v === 'number') return parseSerial(v, numFmt);
    if (typeof v === 'boolean') return invalidResult(String(v), String(v));
    return parseTimeString(String(v));
  }

  var REASONS = {
    'missing-in': '未签到',
    'bad-in': '签到时间无法识别',
    'late-in': '签到晚于 9:00',
    'missing-out': '未签退',
    'bad-out': '签退时间无法识别'
  };

  /** 按规则判定一条记录 */
  function evaluateRecord(signIn, signOut, cutoffSeconds) {
    var cutoff = typeof cutoffSeconds === 'number' ? cutoffSeconds : CUTOFF_SECONDS;
    var reasons = [];
    if (signIn.kind === 'blank') {
      reasons.push('missing-in');
    } else if (signIn.kind === 'invalid') {
      reasons.push('bad-in');
    } else if (signIn.seconds > cutoff) {
      reasons.push('late-in');
    }

    if (signOut.kind === 'blank') {
      reasons.push('missing-out');
    } else if (signOut.kind === 'invalid') {
      reasons.push('bad-out');
    }

    return {
      invalid: reasons.length > 0,
      reasons: reasons,
      reasonLabels: reasons.map(function (code) { return REASONS[code]; })
    };
  }

  function computeStats(records) {
    var stats = {
      total: records.length,
      invalid: 0,
      valid: 0,
      counts: { 'late-in': 0, 'missing-in': 0, 'bad-in': 0, 'missing-out': 0, 'bad-out': 0 }
    };
    records.forEach(function (record) {
      if (record.invalid) stats.invalid += 1;
      else stats.valid += 1;
      record.reasons.forEach(function (code) {
        if (stats.counts[code] != null) stats.counts[code] += 1;
      });
    });
    return stats;
  }

  function normalizeHeader(header) {
    return trimText(header).replace(/[\s\u3000()（）:：]/g, '');
  }

  function scoreHeader(header, field) {
    var s = normalizeHeader(header).toLowerCase();
    if (!s) return 0;
    if (field === 'name') {
      if (s === '姓名') return 100;
      if (s.indexOf('姓名') !== -1) return 90;
      if (s === '名字' || s.indexOf('员工姓名') !== -1) return 80;
      return 0;
    }
    if (/地点|地址|位置|状态|类型|结果|日期/.test(s)) return 0;
    if (field === 'in') {
      if (s === '签到时间' || s === '签到') return 100;
      if (s.indexOf('签到') !== -1) return 85;
      if (s.indexOf('上班') !== -1 && s.indexOf('时间') !== -1) return 80;
      if (s.indexOf('进入') !== -1) return 60;
      return 0;
    }
    if (field === 'out') {
      if (s === '签退时间' || s === '签退') return 100;
      if (s.indexOf('签退') !== -1) return 85;
      if (s.indexOf('下班') !== -1 && s.indexOf('时间') !== -1) return 80;
      if (s.indexOf('离开') !== -1) return 60;
      return 0;
    }
    return 0;
  }

  /** 根据表头自动识别三列，返回 1 基列号或 null */
  function autoMapHeaders(headers) {
    var result = { name: null, in: null, out: null };
    var used = {};
    ['name', 'in', 'out'].forEach(function (field) {
      var bestCol = null;
      var bestScore = 0;
      for (var i = 0; i < headers.length; i++) {
        var col = i + 1;
        if (used[col]) continue;
        var score = scoreHeader(headers[i], field);
        if (score > bestScore) {
          bestScore = score;
          bestCol = col;
        }
      }
      if (bestCol !== null) {
        result[field] = bestCol;
        used[bestCol] = true;
      }
    });
    return result;
  }

  /** 导出时把解析结果转换为 ExcelJS 可写入的值 */
  function toExcelCell(parsed) {
    if (parsed.kind !== 'time') {
      if (parsed.kind === 'invalid' && parsed.display) {
        return { value: parsed.display, numFmt: '@' };
      }
      return { value: null, numFmt: null };
    }
    var p = parsed.parts;
    if (parsed.hasDate) {
      return {
        value: new Date(Date.UTC(p.y, p.mo - 1, p.d, p.h, p.mi, p.s)),
        numFmt: 'yyyy-mm-dd hh:mm:ss'
      };
    }
    return {
      value: (p.h * 3600 + p.mi * 60 + p.s) / 86400,
      numFmt: 'hh:mm:ss'
    };
  }

  /** 中文日期信息（按浏览器本地时区） */
  function todayInfo(date) {
    var now = date instanceof Date ? date : new Date();
    var weekdays = ['周日', '周一', '周二', '周三', '周四', '周五', '周六'];
    return {
      year: now.getFullYear(),
      month: now.getMonth() + 1,
      day: now.getDate(),
      weekday: weekdays[now.getDay()],
      monthDay: (now.getMonth() + 1) + '月' + now.getDate() + '日',
      full: now.getFullYear() + '年' + (now.getMonth() + 1) + '月' + now.getDate() + '日',
      header: (now.getMonth() + 1) + '月' + now.getDate() + '日打卡无效人员'
    };
  }

  function columnLetter(col) {
    var s = '';
    var n = col;
    while (n > 0) {
      var mod = (n - 1) % 26;
      s = String.fromCharCode(65 + mod) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function formatFileSize(bytes) {
    if (!bytes && bytes !== 0) return '';
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / 1024 / 1024).toFixed(1) + ' MB';
  }

  return {
    CUTOFF_SECONDS: CUTOFF_SECONDS,
    REASONS: REASONS,
    normalizeCellValue: normalizeCellValue,
    parseTimeString: parseTimeString,
    parseTimeValue: parseTimeValue,
    parseSerial: parseSerial,
    evaluateRecord: evaluateRecord,
    computeStats: computeStats,
    autoMapHeaders: autoMapHeaders,
    toExcelCell: toExcelCell,
    todayInfo: todayInfo,
    columnLetter: columnLetter,
    formatFileSize: formatFileSize,
    formatHasDate: formatHasDate,
    isPlaceholderText: isPlaceholderText
  };
});
