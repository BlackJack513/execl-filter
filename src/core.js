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

  var CUTOFF_SECONDS = 9 * 3600; // 09:00:00，到达该时刻即算迟到
  var EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30); // Excel 1900 日期系统
  // 加班卡默认规则：18:00-19:00 签到，21:00 起签退
  var OVERTIME_DEFAULTS = {
    inStartSeconds: 18 * 3600,
    inEndSeconds: 19 * 3600,
    outFromSeconds: 21 * 3600
  };
  // 加班卡兜底识别：有数据且签到晚于 16:00 的行视为加班卡
  var OVERTIME_DETECT_SECONDS = 16 * 3600;
  // 三种统计场景：
  //   full     全天考勤——签到、签退都要检查
  //   morning  早上九点后——签退时间通常还都是空的，只检查签到
  //   overtime 加班打卡——只看加班卡，按加班签到/签退时间段判定
  var MODES = { FULL: 'full', MORNING: 'morning', OVERTIME: 'overtime' };
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
    'late-in': '迟到（≥ 9:00）',
    'missing-out': '未签退',
    'bad-out': '签退时间无法识别',
    'ot-missing-in': '未签到',
    'ot-bad-in': '签到时间无法识别',
    'ot-early-in': '签到早于规定时间',
    'ot-late-in': '签到晚于规定时间',
    'ot-missing-out': '未签退',
    'ot-bad-out': '签退时间无法识别',
    'ot-early-out': '签退早于规定时间',
    'ot-only-out': '缺加班签到（仅签退）',
    'ot-blank': '加班未打卡（签到、签退均为空）'
  };

  /** 秒数格式化为 hh:mm，用于规则文案 */
  function formatClock(seconds) {
    if (typeof seconds !== 'number' || !isFinite(seconds)) return '';
    var s = Math.round(seconds) % 86400;
    return pad2(Math.floor(s / 3600)) + ':' + pad2(Math.floor((s % 3600) / 60));
  }

  /** 解析界面上 hh:mm 形式的配置，无法识别时返回 null */
  function parseClockSeconds(value) {
    var parsed = parseTimeString(value);
    return parsed.kind === 'time' ? parsed.seconds : null;
  }

  function pickSeconds(value, fallback) {
    return typeof value === 'number' && isFinite(value) ? value : fallback;
  }

  /** 签退日期是否晚于签到日期（跨天加班） */
  function isCrossDay(signIn, signOut) {
    if (!signIn || !signOut) return false;
    if (signIn.kind !== 'time' || signOut.kind !== 'time') return false;
    if (!signIn.hasDate || !signOut.hasDate) return false;
    var a = signIn.parts;
    var b = signOut.parts;
    if (!a || !b) return false;
    if (b.y !== a.y) return b.y > a.y;
    if (b.mo !== a.mo) return b.mo > a.mo;
    return b.d > a.d;
  }

  function dateKeyOf(parsed) {
    if (!parsed || parsed.kind !== 'time' || !parsed.hasDate || !parsed.parts) return '';
    var p = parsed.parts;
    return p.y + '-' + p.mo + '-' + p.d;
  }

  function recordDateKey(signIn, signOut) {
    return dateKeyOf(signIn) || dateKeyOf(signOut);
  }

  function isOvertimeRow(item, outFrom) {
    if (item.signIn.kind === 'time' && item.signIn.seconds >= OVERTIME_DETECT_SECONDS) return true;
    // 只打了加班签退、没打加班签到
    return item.signIn.kind === 'blank' &&
      item.signOut.kind === 'time' &&
      item.signOut.seconds >= outFrom;
  }

  /**
   * 标记加班卡（考勤系统导出的「打卡类型」永远是传统打卡，只能靠行结构判断）：
   *   1. 同一人同一日期出现第一次 = 正常卡，第二次及以后 = 加班卡；
   *   2. 有数据的加班行：签到晚于 16:00，或签到空白但签退不早于最早签退时间；
   *   3. 第二次出现但签到、签退都为空：加班卡未打卡，标为 blankOvertime；
   *   4. 第一次出现就是空白、且后面还有正常卡时，视为导出填充行，跳过（skip）。
   * 返回值与入参为同一数组，逐项带上 overtime / blankOvertime / skip 标记。
   */
  function classifyAttendanceRows(items, options) {
    options = options || {};
    var outFrom = pickSeconds(options.outFromSeconds, OVERTIME_DEFAULTS.outFromSeconds);
    var groups = {};
    var order = [];

    items.forEach(function (item) {
      item.overtime = false;
      item.blankOvertime = false;
      item.skip = false;
      var name = trimText(item.name);
      var dateKey = trimText(item.dateKey) || recordDateKey(item.signIn, item.signOut);
      var key = name + '\u0000' + dateKey;
      if (!groups[key]) {
        groups[key] = [];
        order.push(key);
      }
      groups[key].push(item);
    });

    order.forEach(function (key) {
      var list = groups[key];
      var hasData = list.map(function (item) {
        return !(item.signIn.kind === 'blank' && item.signOut.kind === 'blank');
      });

      list.forEach(function (item, pos) {
        if (hasData[pos]) {
          item.overtime = isOvertimeRow(item, outFrom);
          return;
        }

        // 第二次及以后出现，即使签到、签退都为空，也属于加班卡未打卡
        if (pos > 0) {
          item.overtime = true;
          item.blankOvertime = true;
          return;
        }

        // 第一次出现就是空白：后面还有正常卡时，这是导出填充行
        var laterNormal = false;
        for (var i = 1; i < list.length; i++) {
          if (!hasData[i]) continue;
          if (!isOvertimeRow(list[i], outFrom)) {
            laterNormal = true;
            break;
          }
        }
        item.skip = laterNormal;
      });
    });

    return items;
  }

  /**
   * 按规则判定一条记录
   * options.mode          'full'（默认）全天考勤 | 'morning' 仅看签到
   * options.cutoffSeconds 迟到分界，默认 09:00:00，到达即判迟到
   */
  function evaluateRecord(signIn, signOut, options) {
    var mode = MODES.FULL;
    var cutoff = CUTOFF_SECONDS;
    if (typeof options === 'number') {
      cutoff = options; // 兼容旧签名 evaluateRecord(in, out, cutoffSeconds)
    } else if (options) {
      if (typeof options.cutoffSeconds === 'number') cutoff = options.cutoffSeconds;
      if (options.mode === MODES.MORNING) mode = MODES.MORNING;
    }

    var reasons = [];
    if (signIn.kind === 'blank') {
      reasons.push('missing-in');
    } else if (signIn.kind === 'invalid') {
      reasons.push('bad-in');
    } else if (signIn.seconds >= cutoff) {
      // 09:00:00 整也判为迟到
      reasons.push('late-in');
    }

    // 早上九点后统计：签退时间此时通常都是空的，不参与判定
    if (mode !== MODES.MORNING) {
      if (signOut.kind === 'blank') {
        reasons.push('missing-out');
      } else if (signOut.kind === 'invalid') {
        reasons.push('bad-out');
      }
    }

    return {
      invalid: reasons.length > 0,
      reasons: reasons,
      reasonLabels: reasons.map(function (code) { return REASONS[code]; })
    };
  }

  /**
   * 加班卡判定：签到需落在加班签到时间段内，签退不早于最早签退时间。
   * options.inStartSeconds / inEndSeconds / outFromSeconds 均可配置。
   */
  function evaluateOvertimeRecord(signIn, signOut, options) {
    options = options || {};
    var inStart = pickSeconds(options.inStartSeconds, OVERTIME_DEFAULTS.inStartSeconds);
    var inEnd = pickSeconds(options.inEndSeconds, OVERTIME_DEFAULTS.inEndSeconds);
    var outFrom = pickSeconds(options.outFromSeconds, OVERTIME_DEFAULTS.outFromSeconds);

    var reasons = [];
    var labels = [];

    function push(code, label) {
      reasons.push(code);
      labels.push(label || REASONS[code] || code);
    }

    // 加班段签到、签退都为空：整条加班卡未打卡
    if (signIn.kind === 'blank' && signOut.kind === 'blank') {
      push('ot-blank', '加班未打卡（签到、签退均为空）');
      return {
        invalid: true,
        reasons: reasons,
        reasonLabels: labels
      };
    }

    if (signIn.kind === 'blank') {
      if (signOut.kind === 'time') {
        // 加班段必须「签到 + 签退」齐全，只有签退单独标注
        push('ot-only-out', '缺加班签到（仅签退）');
      } else {
        push('ot-missing-in', '未签到');
      }
    } else if (signIn.kind === 'invalid') {
      push('ot-bad-in', '签到时间无法识别');
    } else if (signIn.seconds < inStart) {
      push('ot-early-in', '签到早于 ' + formatClock(inStart));
    } else if (signIn.seconds > inEnd) {
      push('ot-late-in', '签到晚于 ' + formatClock(inEnd));
    }

    if (signOut.kind === 'blank') {
      push('ot-missing-out', '未签退');
    } else if (signOut.kind === 'invalid') {
      push('ot-bad-out', '签退时间无法识别');
    } else if (!isCrossDay(signIn, signOut) && signOut.seconds < outFrom) {
      // 跨天签退（次日凌晨）视为正常
      push('ot-early-out', '签退早于 ' + formatClock(outFrom));
    }

    return {
      invalid: reasons.length > 0,
      reasons: reasons,
      reasonLabels: labels
    };
  }

  function computeStats(records) {
    var stats = {
      total: records.length,
      invalid: 0,
      valid: 0,
      counts: {},
      labels: {}
    };
    records.forEach(function (record) {
      if (record.invalid) stats.invalid += 1;
      else stats.valid += 1;
      record.reasons.forEach(function (code, index) {
        stats.counts[code] = (stats.counts[code] || 0) + 1;
        if (record.reasonLabels && record.reasonLabels[index]) {
          stats.labels[code] = record.reasonLabels[index];
        }
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
    if (field === 'date') {
      if (s === '日期') return 100;
      if (s === '考勤日期' || s === '打卡日期') return 95;
      if (s.indexOf('日期') !== -1) return 80;
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
    var result = { name: null, in: null, out: null, date: null };
    var used = {};
    ['name', 'in', 'out', 'date'].forEach(function (field) {
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
    OVERTIME_DEFAULTS: OVERTIME_DEFAULTS,
    OVERTIME_DETECT_SECONDS: OVERTIME_DETECT_SECONDS,
    MODES: MODES,
    REASONS: REASONS,
    normalizeCellValue: normalizeCellValue,
    parseTimeString: parseTimeString,
    parseTimeValue: parseTimeValue,
    parseSerial: parseSerial,
    formatClock: formatClock,
    parseClockSeconds: parseClockSeconds,
    isCrossDay: isCrossDay,
    evaluateRecord: evaluateRecord,
    evaluateOvertimeRecord: evaluateOvertimeRecord,
    classifyAttendanceRows: classifyAttendanceRows,
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
