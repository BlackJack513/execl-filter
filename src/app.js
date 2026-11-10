/*
 * 打卡无效记录筛查 —— 页面交互
 * 依赖：全局 ExcelJS、window.DashboardCore
 */
(function () {
  'use strict';

  var C = window.DashboardCore;
  var MAX_PREVIEW_ROWS = 500;
  var MAX_COLS = 80;
  var MAX_ROWS = 200000;
  var MAX_FILE_SIZE = 30 * 1024 * 1024;

  var el = {};
  var toastTimer = null;
  var recalcTimer = null;

  var state = {
    file: null,
    workbook: null,
    worksheets: [],
    sheet: null,
    sheetIndex: -1,
    rowCount: 0,
    maxCol: 0,
    headerRow: 1,
    dataStart: 2,
    headers: [],
    mapping: { name: '', in: '', out: '' },
    records: [],
    stats: null,
    filter: 'all'
  };

  var ICONS = {
    initial:
      '<svg width="42" height="42" viewBox="0 0 48 48" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
      '<path d="M10 6.5h16.5L36 16v25.5H10z"></path>' +
      '<path d="M26.5 6.5V16H36"></path>' +
      '<path d="M16.5 22.5h12M16.5 29h7.5"></path>' +
      '<circle cx="31.5" cy="32.5" r="6.2"></circle>' +
      '<path d="M36 37l4.2 4.2"></path>' +
      '</svg>',
    warning:
      '<svg width="42" height="42" viewBox="0 0 48 48" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
      '<path d="M24 7.5 43 41H5z"></path>' +
      '<path d="M24 18.5v10"></path>' +
      '<circle cx="24" cy="35.2" r="1.5" fill="currentColor" stroke="none"></circle>' +
      '</svg>',
    success:
      '<svg width="42" height="42" viewBox="0 0 48 48" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
      '<circle cx="24" cy="24" r="17"></circle>' +
      '<path d="m16.5 24.5 5 5 10-10.5"></path>' +
      '</svg>',
    search:
      '<svg width="42" height="42" viewBox="0 0 48 48" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
      '<circle cx="21" cy="21" r="12"></circle>' +
      '<path d="M30 30l9.5 9.5"></path>' +
      '<path d="M16 21h10M21 16v10" opacity=".55"></path>' +
      '</svg>'
  };

  function $(id) {
    return document.getElementById(id);
  }

  function cacheDom() {
    [
      'fileInput', 'dropzone', 'fileCard', 'fileName', 'fileMeta', 'changeFileBtn',
      'settingsCard', 'headerRowInput', 'dataStartInput', 'sheetSelect',
      'nameCol', 'inCol', 'outCol', 'nameSample', 'inSample', 'outSample',
      'remapBtn', 'mapHint', 'resultCard', 'resultSub', 'filterTabs',
      'stats', 'statTotal', 'statInvalid', 'statValid', 'statInvalidFoot', 'statValidFoot',
      'reasonChips', 'tableShell', 'previewBody', 'tableFoot', 'emptyState',
      'exportName', 'exportHint', 'copyBtn', 'exportBtn', 'exportBtnText',
      'loading', 'loadingFile', 'toast', 'todayText'
    ].forEach(function (id) {
      el[id] = $(id);
    });
  }

  function esc(value) {
    return String(value == null ? '' : value).replace(/[&<>"']/g, function (ch) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch];
    });
  }

  function truncate(text, length) {
    var value = String(text == null ? '' : text);
    return value.length > length ? value.slice(0, length - 1) + '…' : value;
  }

  function clampInt(value, min, max) {
    var n = parseInt(value, 10);
    if (!isFinite(n)) n = min;
    if (n < min) n = min;
    if (n > max) n = max;
    return n;
  }

  function toCol(value) {
    var n = parseInt(value, 10);
    return isFinite(n) && n > 0 ? n : '';
  }

  function tick() {
    return new Promise(function (resolve) {
      setTimeout(resolve, 30);
    });
  }

  function debounceRecalc() {
    clearTimeout(recalcTimer);
    recalcTimer = setTimeout(function () {
      recalcTimer = null;
      recalc();
    }, 120);
  }

  function flushRecalc() {
    if (!recalcTimer) return;
    clearTimeout(recalcTimer);
    recalcTimer = null;
    recalc();
  }

  function showToast(message, type) {
    clearTimeout(toastTimer);
    el.toast.textContent = message;
    el.toast.className = 'toast show' + (type ? ' is-' + type : '');
    toastTimer = setTimeout(function () {
      el.toast.className = 'toast';
    }, 3800);
  }

  function setLoading(on, fileName) {
    if (on) {
      el.loadingFile.textContent = fileName ? '文件：' + fileName : '';
      el.loading.classList.remove('hidden');
    } else {
      el.loading.classList.add('hidden');
    }
  }

  function setFilter(filter) {
    state.filter = filter;
    Array.prototype.forEach.call(el.filterTabs.querySelectorAll('.seg-btn'), function (btn) {
      btn.classList.toggle('is-active', btn.getAttribute('data-filter') === filter);
    });
  }

  function setSample(node, key) {
    if (!state.workbook) {
      node.textContent = '上传文件后自动识别';
      return;
    }
    var col = state.mapping[key];
    if (!col) {
      node.textContent = '请选择一列';
      return;
    }
    var sample = findSample(col, key);
    node.textContent = sample ? '示例：' + sample : '该列暂无数据';
  }

  function formatWallDate(value) {
    function pad(n) {
      return (n < 10 ? '0' : '') + n;
    }
    return value.getUTCFullYear() + '-' + pad(value.getUTCMonth() + 1) + '-' + pad(value.getUTCDate()) +
      ' ' + pad(value.getUTCHours()) + ':' + pad(value.getUTCMinutes()) + ':' + pad(value.getUTCSeconds());
  }

  function readCellText(cell) {
    if (!cell) return '';
    var value = C.normalizeCellValue(cell.value);
    if (value instanceof Date) return formatWallDate(value);
    var text = '';
    try {
      text = cell.text == null ? '' : String(cell.text);
    } catch (err) {
      text = '';
    }
    if (!text && value != null) text = String(value);
    return text.replace(/\s+/g, ' ').trim();
  }

  function findSample(col, key) {
    if (!state.sheet || !state.rowCount) return '';
    var last = Math.min(state.rowCount, state.dataStart + 11);
    for (var r = state.dataStart; r <= last; r++) {
      var cell = state.sheet.getCell(r, col);
      if (key === 'in' || key === 'out') {
        var parsed = C.parseTimeValue(cell.value, cell.numFmt, { placeholderAsBlank: true });
        if (parsed.kind === 'time') return truncate(parsed.display, 38);
        var fallback = readCellText(cell);
        if (fallback) return truncate(fallback, 38);
        continue;
      }
      var text = readCellText(cell);
      if (text) return truncate(text, 38);
    }
    return '';
  }

  function updateSamples() {
    setSample(el.nameSample, 'name');
    setSample(el.inSample, 'in');
    setSample(el.outSample, 'out');
  }

  function showMapHint(message) {
    el.mapHint.textContent = message;
    el.mapHint.classList.remove('hidden');
  }

  function hideMapHint() {
    el.mapHint.classList.add('hidden');
  }

  function readHeaders() {
    state.headers = [];
    if (!state.sheet) return;
    var count = Math.min(Math.max(state.maxCol || 1, 1), MAX_COLS);
    for (var c = 1; c <= count; c++) {
      var cell = state.sheet.getCell(state.headerRow, c);
      var master = cell;
      try {
        if (cell.isMerged && cell.master) master = cell.master;
      } catch (err) {
        master = cell;
      }
      var text = readCellText(master);
      var label = C.columnLetter(c) + ' 列';
      if (text) label += ' · ' + truncate(text, 18);
      else label += ' · （空表头）';
      state.headers.push({ col: c, letter: C.columnLetter(c), text: text, label: label });
    }
    state.maxCol = count;
  }

  function renderColumnOptions() {
    var options = '<option value="">— 请选择列 —</option>';
    state.headers.forEach(function (header) {
      options += '<option value="' + header.col + '">' + esc(header.label) + '</option>';
    });

    var disabled = !state.workbook || !state.headers.length;
    var fields = [
      { node: el.nameCol, key: 'name' },
      { node: el.inCol, key: 'in' },
      { node: el.outCol, key: 'out' }
    ];

    fields.forEach(function (field) {
      field.node.innerHTML = options;
      var current = state.mapping[field.key];
      var exists = false;
      state.headers.forEach(function (header) {
        if (String(header.col) === String(current)) exists = true;
      });
      field.node.value = exists ? String(current) : '';
      field.node.disabled = disabled;
      state.mapping[field.key] = exists ? Number(current) : '';
    });
  }

  function applyAutoMapping() {
    var auto = C.autoMapHeaders(state.headers.map(function (header) { return header.text; }));
    state.mapping = {
      name: auto.name || '',
      in: auto.in || '',
      out: auto.out || ''
    };
    renderColumnOptions();
    updateSamples();
    recalc();
    return !!(state.mapping.name && state.mapping.in && state.mapping.out);
  }

  function validateMapping() {
    var missing = [];
    if (!state.mapping.name) missing.push('姓名列');
    if (!state.mapping.in) missing.push('签到时间列');
    if (!state.mapping.out) missing.push('签退时间列');
    if (missing.length) return '请选择' + missing.join('、') + '，然后即可开始筛查。';
    if (state.dataStart <= state.headerRow) return '数据起始行必须大于表头所在行。';
    if (!state.headers.length) return '当前工作表没有可读取的列，请检查表头行设置。';

    var selected = [state.mapping.name, state.mapping.in, state.mapping.out];
    if (selected[0] === selected[1] || selected[0] === selected[2] || selected[1] === selected[2]) {
      return '姓名、签到时间、签退时间不能选择同一列。';
    }
    return '';
  }

  function buildRecords() {
    var records = [];
    var mapping = state.mapping;
    var lastRow = Math.min(state.rowCount, MAX_ROWS);

    for (var r = state.dataStart; r <= lastRow; r++) {
      var nameCell = state.sheet.getCell(r, mapping.name);
      var inCell = state.sheet.getCell(r, mapping.in);
      var outCell = state.sheet.getCell(r, mapping.out);

      var name = readCellText(nameCell);
      var signIn = C.parseTimeValue(inCell.value, inCell.numFmt, { placeholderAsBlank: true });
      var signOut = C.parseTimeValue(outCell.value, outCell.numFmt, { placeholderAsBlank: true });

      if (!name && signIn.kind === 'blank' && signOut.kind === 'blank') continue;

      var evaluation = C.evaluateRecord(signIn, signOut);
      records.push({
        row: r,
        name: name,
        signIn: signIn,
        signOut: signOut,
        invalid: evaluation.invalid,
        reasons: evaluation.reasons,
        reasonLabels: evaluation.reasonLabels
      });
    }
    return records;
  }

  function updateFileMeta() {
    if (!state.file) return;
    var parts = [];
    parts.push(state.worksheets.length + ' 个工作表');
    parts.push(C.formatFileSize(state.file.size));
    if (state.sheet) parts.push('当前 ' + state.rowCount + ' 行');
    el.fileMeta.textContent = parts.join(' · ');
  }

  function populateSheets() {
    el.sheetSelect.innerHTML = state.worksheets.map(function (ws, index) {
      var suffix = ws.state && ws.state !== 'visible' ? '（隐藏）' : '';
      return '<option value="' + index + '">' + esc(ws.name) + suffix + '</option>';
    }).join('');
    el.sheetSelect.disabled = state.worksheets.length <= 1;
  }

  function pickDefaultSheet() {
    var best = 0;
    var bestRows = -1;
    state.worksheets.forEach(function (ws, index) {
      if (ws.state && ws.state !== 'visible') return;
      var rows = ws.rowCount || 0;
      if (rows > bestRows) {
        bestRows = rows;
        best = index;
      }
    });
    return best;
  }

  function selectSheet(index, options) {
    options = options || {};
    if (!state.workbook || index < 0 || index >= state.worksheets.length) return;

    state.sheetIndex = index;
    state.sheet = state.worksheets[index];
    state.rowCount = state.sheet.rowCount || 0;
    state.maxCol = Math.max(state.sheet.columnCount || 0, 1);

    var maxHeader = Math.max(state.rowCount, 1);
    var headerRow = clampInt(el.headerRowInput.value, 1, maxHeader);
    el.headerRowInput.max = maxHeader;
    el.headerRowInput.value = headerRow;
    state.headerRow = headerRow;

    var minData = headerRow + 1;
    var maxData = Math.max(state.rowCount, minData);
    var dataStart = clampInt(el.dataStartInput.value, minData, maxData);
    el.dataStartInput.min = minData;
    el.dataStartInput.max = maxData;
    el.dataStartInput.value = dataStart;
    state.dataStart = dataStart;

    readHeaders();
    renderColumnOptions();

    if (options.autoMap) {
      applyAutoMapping();
    } else {
      updateSamples();
      recalc();
    }

    updateFileMeta();
  }

  function resetResultFooter() {
    var today = C.todayInfo();
    el.exportName.textContent = today.header + '.xlsx';
    el.exportBtn.disabled = true;
    el.copyBtn.disabled = true;
    el.exportBtnText.textContent = '导出 Excel';
  }

  function showEmptyState(options) {
    var tone = options.tone ? ' is-' + options.tone : '';
    el.emptyState.className = 'empty-state';
    el.emptyState.innerHTML =
      '<div class="empty-illustration' + tone + '">' + options.icon + '</div>' +
      '<h3 class="empty-title">' + options.title + '</h3>' +
      '<p class="empty-text">' + options.text + '</p>' +
      (options.steps || '');
    el.emptyState.classList.remove('hidden');
  }

  function hideEmptyState() {
    el.emptyState.classList.add('hidden');
  }

  function hideTable() {
    el.tableShell.classList.add('hidden');
  }

  function renderInitial() {
    state.records = [];
    state.stats = null;
    el.resultSub.textContent = '上传文件后自动生成';
    el.stats.classList.add('hidden');
    el.reasonChips.classList.add('hidden');
    hideTable();
    el.filterTabs.classList.add('hidden');
    el.exportHint.textContent = '上传文件并完成列映射后即可导出';
    resetResultFooter();

    showEmptyState({
      icon: ICONS.initial,
      title: '上传考勤表，开始筛查',
      text: '系统会自动识别姓名、签到时间、签退时间三列，<br>实时预览异常记录，确认后一键导出。',
      steps:
        '<div class="empty-steps">' +
        '<span class="empty-step"><i>1</i>上传文件</span>' +
        '<span class="empty-step-arrow">→</span>' +
        '<span class="empty-step"><i>2</i>确认列映射</span>' +
        '<span class="empty-step-arrow">→</span>' +
        '<span class="empty-step"><i>3</i>导出结果</span>' +
        '</div>'
    });
  }

  function renderMappingNeeded(message) {
    state.records = [];
    state.stats = null;
    showMapHint(message);
    el.resultSub.textContent = '请完成列映射';
    el.stats.classList.add('hidden');
    el.reasonChips.classList.add('hidden');
    hideTable();
    el.filterTabs.classList.add('hidden');
    el.exportHint.textContent = '完成列映射并读取数据后即可导出';
    resetResultFooter();

    showEmptyState({
      icon: ICONS.warning,
      tone: 'amber',
      title: '还差一步：完成列映射',
      text: '请在左侧分别选择<b>姓名</b>、<b>签到时间</b>、<b>签退时间</b>所在的列。<br>选择完成后，筛查结果会实时出现在这里。',
      steps:
        '<div class="empty-steps">' +
        '<span class="empty-step"><i>←</i>在左侧选择三列</span>' +
        '</div>'
    });
  }

  function renderNoData() {
    hideMapHint();
    state.records = [];
    state.stats = null;
    el.resultSub.textContent = (state.sheet ? state.sheet.name + ' · ' : '') + '没有可筛查的数据';
    el.stats.classList.add('hidden');
    el.reasonChips.classList.add('hidden');
    hideTable();
    el.filterTabs.classList.add('hidden');
    el.exportHint.textContent = '没有可导出的数据';
    resetResultFooter();

    showEmptyState({
      icon: ICONS.search,
      tone: 'amber',
      title: '没有读取到数据',
      text: '已从第 ' + state.dataStart + ' 行开始查找，但没有读取到包含姓名或打卡时间的记录。<br>请检查「表头所在行」「数据起始行」，或切换工作表。'
    });
  }

  function renderTimes(parsed, bad) {
    if (parsed.kind === 'blank') {
      return '<span class="cell-empty">空白</span>';
    }
    var display = parsed.display || parsed.raw || '（无法识别）';
    return '<span class="time-value' + (bad ? ' time-warning' : '') + '">' + esc(display) + '</span>';
  }

  function renderTable(rows) {
    var shown = rows.slice(0, MAX_PREVIEW_ROWS);
    var html = shown.map(function (record) {
      var badIn = record.reasons.indexOf('late-in') !== -1 ||
        record.reasons.indexOf('missing-in') !== -1 ||
        record.reasons.indexOf('bad-in') !== -1;
      var badOut = record.reasons.indexOf('missing-out') !== -1 ||
        record.reasons.indexOf('bad-out') !== -1;

      var nameHtml = record.name
        ? '<span>' + esc(record.name) + '</span>'
        : '<span class="name-empty">（空姓名）</span>';

      var statusHtml = record.invalid
        ? '<span class="badge badge-invalid">无效</span>' +
          '<div class="status-reasons">' + esc(record.reasonLabels.join(' · ')) + '</div>'
        : '<span class="badge badge-valid">有效</span>';

      return '<tr class="' + (record.invalid ? 'is-invalid' : '') + '">' +
        '<td class="cell-no">' + record.row + '</td>' +
        '<td class="cell-name">' + nameHtml + '</td>' +
        '<td class="cell-time' + (badIn ? ' cell-bad' : '') + '">' + renderTimes(record.signIn, badIn) + '</td>' +
        '<td class="cell-time' + (badOut ? ' cell-bad' : '') + '">' + renderTimes(record.signOut, badOut) + '</td>' +
        '<td>' + statusHtml + '</td>' +
        '</tr>';
    }).join('');

    el.previewBody.innerHTML = html;
    var rightNote = rows.length > MAX_PREVIEW_ROWS
      ? '仅预览前 ' + MAX_PREVIEW_ROWS + ' 条，导出包含全部'
      : (state.filter === 'invalid' ? '仅显示无效记录' : '异常单元格已标红');
    el.tableFoot.innerHTML =
      '<span>显示 <b>' + shown.length + '</b> / ' + rows.length + ' 条</span>' +
      '<span>' + rightNote + '</span>';
    el.tableShell.classList.remove('hidden');
    hideEmptyState();
  }

  function renderFooter() {
    var today = C.todayInfo();
    el.exportName.textContent = today.header + '.xlsx';

    var invalid = state.stats ? state.stats.invalid : 0;
    if (!state.workbook) {
      el.exportHint.textContent = '上传文件并完成列映射后即可导出';
    } else if (!state.stats) {
      el.exportHint.textContent = '完成列映射并读取数据后即可导出';
    } else if (!invalid) {
      el.exportHint.textContent = '没有无效记录，无需导出';
    } else {
      el.exportHint.textContent = '将导出 ' + invalid + ' 条无效记录 · 异常单元格标红';
    }

    el.exportBtnText.textContent = invalid > 0 ? '导出 Excel（' + invalid + '）' : '导出 Excel';
    el.exportBtn.disabled = invalid === 0;
    el.copyBtn.disabled = invalid === 0;
  }

  function renderResults() {
    hideMapHint();
    var stats = state.stats;
    var total = stats.total;
    var invalid = stats.invalid;

    el.stats.classList.remove('hidden');
    el.filterTabs.classList.remove('hidden');
    setFilter(state.filter);

    el.statTotal.textContent = String(total);
    el.statInvalid.textContent = String(invalid);
    el.statValid.textContent = String(stats.valid);
    el.statInvalidFoot.textContent = invalid
      ? '异常率 ' + ((invalid / total) * 100).toFixed(1).replace(/\.0$/, '') + '%'
      : '全部通过';
    el.statValidFoot.textContent = invalid ? '签到与签退均正常' : '全部记录均正常';

    if (invalid > 0) {
      var definitions = [
        ['late-in', '签到晚于 9:00'],
        ['missing-out', '未签退'],
        ['missing-in', '未签到'],
        ['bad-in', '签到无法识别'],
        ['bad-out', '签退无法识别']
      ];
      var chips = definitions.filter(function (item) {
        return stats.counts[item[0]] > 0;
      }).map(function (item) {
        return '<span class="reason-chip"><span class="chip-dot"></span>' +
          esc(item[1]) + ' <b>' + stats.counts[item[0]] + '</b></span>';
      }).join('');
      el.reasonChips.innerHTML = chips +
        '<span class="chips-note">同一人可能同时命中多条规则</span>';
      el.reasonChips.classList.remove('hidden');
    } else {
      el.reasonChips.classList.add('hidden');
    }

    var sheetName = state.sheet ? state.sheet.name : '';
    el.resultSub.textContent = sheetName + ' · 共 ' + total + ' 条记录' +
      (invalid ? ' · 发现 ' + invalid + ' 条无效' : ' · 全部有效');

    var rows = state.filter === 'invalid'
      ? state.records.filter(function (record) { return record.invalid; })
      : state.records;

    if (rows.length) {
      renderTable(rows);
    } else {
      hideTable();
      showEmptyState({
        icon: ICONS.success,
        tone: 'green',
        title: '本次没有发现无效记录',
        text: '所有记录均满足签到时间不晚于 09:00:00，且签退时间不为空白。',
        steps:
          '<div class="empty-steps">' +
          '<span class="empty-step"><i>✓</i>' + total + ' 条记录全部通过</span>' +
          '</div>'
      });
    }

    renderFooter();
  }

  function recalc() {
    if (!state.workbook) {
      renderInitial();
      return;
    }

    var error = validateMapping();
    if (error) {
      renderMappingNeeded(error);
      return;
    }

    hideMapHint();
    var records;
    try {
      records = buildRecords();
    } catch (err) {
      console.error(err);
      showToast('读取数据失败：' + (err && err.message ? err.message : '未知错误'), 'error');
      records = [];
    }

    state.records = records;
    state.stats = C.computeStats(records);

    if (!records.length) {
      renderNoData();
      return;
    }
    renderResults();
  }

  function onHeaderRowChange() {
    if (!state.workbook) return;
    var maxHeader = Math.max(state.rowCount, 1);
    var headerRow = clampInt(el.headerRowInput.value, 1, maxHeader);
    el.headerRowInput.value = headerRow;
    state.headerRow = headerRow;

    var minData = headerRow + 1;
    var maxData = Math.max(state.rowCount, minData);
    var dataStart = clampInt(el.dataStartInput.value, minData, maxData);
    el.dataStartInput.min = minData;
    el.dataStartInput.max = maxData;
    el.dataStartInput.value = dataStart;
    state.dataStart = dataStart;

    readHeaders();
    renderColumnOptions();
    if (!state.mapping.name || !state.mapping.in || !state.mapping.out) {
      applyAutoMapping();
      return;
    }
    updateSamples();
    recalc();
  }

  function onDataStartChange() {
    if (!state.workbook) return;
    var minData = state.headerRow + 1;
    var maxData = Math.max(state.rowCount, minData);
    var dataStart = clampInt(el.dataStartInput.value, minData, maxData);
    el.dataStartInput.value = dataStart;
    state.dataStart = dataStart;
    updateSamples();
    debounceRecalc();
  }

  function onMappingChange() {
    state.mapping = {
      name: toCol(el.nameCol.value),
      in: toCol(el.inCol.value),
      out: toCol(el.outCol.value)
    };
    updateSamples();
    debounceRecalc();
  }

  function formatLoadError(err) {
    var message = String((err && err.message) || err || '');
    if (/password|encrypt/i.test(message)) {
      return '文件受密码保护，请先在 Excel 中取消密码后再上传。';
    }
    if (message === 'NO_SHEETS') {
      return '这个工作簿里没有工作表。';
    }
    if (/zip|central directory|end of data|invalid signature|corrupt/i.test(message)) {
      return '无法读取该文件：可能不是有效的 .xlsx 文件，或文件已损坏。';
    }
    return '解析失败：' + (message || '未知错误');
  }

  function enableSettings(on) {
    el.settingsCard.classList.toggle('is-locked', !on);
    el.headerRowInput.disabled = !on;
    el.dataStartInput.disabled = !on;
    el.sheetSelect.disabled = !on || state.worksheets.length <= 1;
    el.remapBtn.disabled = !on;
    if (!on) {
      el.nameCol.disabled = true;
      el.inCol.disabled = true;
      el.outCol.disabled = true;
    }
  }

  async function loadFile(file) {
    var name = file.name || '未命名文件';
    var ext = String(name.split('.').pop() || '').toLowerCase();

    if (ext === 'xls') {
      showToast('暂不支持旧版 .xls，请在 Excel 中另存为 .xlsx 后再上传。', 'error');
      return;
    }
    if (ext !== 'xlsx' && ext !== 'xlsm') {
      showToast('请选择 .xlsx 或 .xlsm 格式的 Excel 文件。', 'error');
      return;
    }
    if (file.size > MAX_FILE_SIZE) {
      showToast('文件超过 30 MB，请拆分后再上传。', 'error');
      return;
    }
    if (!window.ExcelJS) {
      showToast('Excel 解析组件未加载，请刷新页面后重试。', 'error');
      return;
    }

    setLoading(true, name);
    await tick();

    try {
      var buffer = await file.arrayBuffer();
      var workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);

      if (!workbook.worksheets || !workbook.worksheets.length) {
        throw new Error('NO_SHEETS');
      }

      state.file = file;
      state.workbook = workbook;
      state.worksheets = workbook.worksheets;
      state.sheet = null;
      state.sheetIndex = -1;
      state.records = [];
      state.stats = null;
      state.filter = 'all';
      state.mapping = { name: '', in: '', out: '' };

      el.fileName.textContent = name;
      el.fileName.title = name;
      el.fileCard.classList.remove('hidden');
      el.dropzone.classList.add('hidden');
      el.headerRowInput.value = '1';
      el.dataStartInput.value = '2';
      el.dataStartInput.min = '2';
      updateFileMeta();

      enableSettings(true);
      populateSheets();
      var preferred = pickDefaultSheet();
      el.sheetSelect.value = String(preferred);
      selectSheet(preferred, { autoMap: true });
      setFilter('all');

      showToast('已读取「' + name + '」，共 ' + workbook.worksheets.length + ' 个工作表');
    } catch (err) {
      console.error(err);
      showToast(formatLoadError(err), 'error');
    } finally {
      setLoading(false);
    }
  }

  function setCellStyle(cell, bad, zebra) {
    cell.font = {
      name: '等线',
      size: 11,
      bold: !!bad,
      color: { argb: bad ? 'FFC00000' : 'FF1F2937' }
    };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: false };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
    };
    if (bad) {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF1F0' } };
    } else if (zebra) {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFD' } };
    }
  }

  function setTimeCell(cell, parsed, bad, zebra) {
    var data = C.toExcelCell(parsed);
    if (data.value !== null && data.value !== undefined) cell.value = data.value;
    if (data.numFmt) cell.numFmt = data.numFmt;
    setCellStyle(cell, bad, zebra);
  }

  function downloadBlob(blob, fileName) {
    var url = URL.createObjectURL(blob);
    var link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(function () {
      URL.revokeObjectURL(url);
    }, 5000);
  }

  async function exportExcel() {
    flushRecalc();
    if (!state.stats || !state.stats.invalid) return;

    var invalidRecords = state.records.filter(function (record) { return record.invalid; });
    if (!invalidRecords.length) return;

    var today = C.todayInfo();
    var originalText = el.exportBtnText.textContent;
    el.exportBtn.disabled = true;
    el.exportBtnText.textContent = '正在生成…';

    try {
      var workbook = new ExcelJS.Workbook();
      workbook.creator = '打卡无效记录筛查';
      workbook.created = new Date();

      var worksheet = workbook.addWorksheet('打卡无效人员', {
        views: [{ state: 'frozen', ySplit: 1, showGridLines: false }],
        pageSetup: {
          paperSize: 9,
          orientation: 'landscape',
          fitToPage: true,
          fitToWidth: 1,
          fitToHeight: 0,
          margins: { left: 0.5, right: 0.5, top: 0.6, bottom: 0.6, header: 0.3, footer: 0.3 }
        }
      });
      worksheet.properties.defaultRowHeight = 20;
      worksheet.columns = [{ width: 22 }, { width: 24 }, { width: 24 }];

      var header = worksheet.getRow(1);
      header.height = 30;
      [today.header, '签到时间', '签退时间'].forEach(function (title, index) {
        var cell = header.getCell(index + 1);
        cell.value = title;
        cell.font = { name: '等线', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: false };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F3864' } };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FF1F3864' } },
          left: { style: 'thin', color: { argb: 'FF1F3864' } },
          bottom: { style: 'thin', color: { argb: 'FF1F3864' } },
          right: { style: 'thin', color: { argb: 'FF1F3864' } }
        };
      });

      invalidRecords.forEach(function (record, index) {
        var badIn = record.reasons.indexOf('late-in') !== -1 ||
          record.reasons.indexOf('missing-in') !== -1 ||
          record.reasons.indexOf('bad-in') !== -1;
        var badOut = record.reasons.indexOf('missing-out') !== -1 ||
          record.reasons.indexOf('bad-out') !== -1;
        var zebra = index % 2 === 1;

        var row = worksheet.addRow([record.name || '', null, null]);
        row.height = 22;
        setCellStyle(row.getCell(1), false, zebra);
        setTimeCell(row.getCell(2), record.signIn, badIn, zebra);
        setTimeCell(row.getCell(3), record.signOut, badOut, zebra);
      });

      worksheet.autoFilter = 'A1:C' + (invalidRecords.length + 1);
      worksheet.headerFooter.oddFooter = '&L' + today.full + '&C第 &P 页 / 共 &N 页&R打卡无效记录筛查';

      var buffer = await workbook.xlsx.writeBuffer();
      var blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      downloadBlob(blob, today.header + '.xlsx');
      showToast('已导出 ' + invalidRecords.length + ' 条无效记录');
    } catch (err) {
      console.error(err);
      showToast('导出失败：' + (err && err.message ? err.message : '未知错误'), 'error');
    } finally {
      el.exportBtnText.textContent = originalText;
      renderFooter();
    }
  }

  function fallbackCopy(text) {
    var textarea = document.createElement('textarea');
    textarea.value = text;
    textarea.setAttribute('readonly', 'readonly');
    textarea.style.position = 'fixed';
    textarea.style.left = '-9999px';
    textarea.style.top = '0';
    document.body.appendChild(textarea);
    textarea.select();
    var ok = false;
    try {
      ok = document.execCommand('copy');
    } catch (err) {
      ok = false;
    }
    textarea.remove();
    return ok;
  }

  async function copyNames() {
    flushRecalc();
    if (!state.stats || !state.stats.invalid) return;
    var allNames = state.records.filter(function (record) {
      return record.invalid;
    }).map(function (record) {
      return record.name;
    });
    var names = allNames.filter(function (name) { return !!name; });
    if (!names.length) {
      showToast('无效记录中没有可复制的姓名。', 'info');
      return;
    }

    var text = names.join('\n');
    var copied = false;
    if (navigator.clipboard && window.isSecureContext) {
      try {
        await navigator.clipboard.writeText(text);
        copied = true;
      } catch (err) {
        copied = false;
      }
    }
    if (!copied) copied = fallbackCopy(text);

    if (copied) {
      var skipped = allNames.length - names.length;
      showToast('已复制 ' + names.length + ' 个姓名' + (skipped ? '（跳过 ' + skipped + ' 条空姓名）' : ''));
    } else {
      showToast('复制失败，请手动选择表格内容后复制。', 'error');
    }
  }

  function bindEvents() {
    el.dropzone.addEventListener('click', function (event) {
      if (event.target === el.fileInput) return;
      el.fileInput.click();
    });
    el.dropzone.addEventListener('keydown', function (event) {
      if (event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        el.fileInput.click();
      }
    });

    var dragOn = function (event) {
      event.preventDefault();
      event.stopPropagation();
      el.dropzone.classList.add('is-drag');
    };
    var dragOff = function (event) {
      event.preventDefault();
      event.stopPropagation();
      el.dropzone.classList.remove('is-drag');
    };
    el.dropzone.addEventListener('dragenter', dragOn);
    el.dropzone.addEventListener('dragover', dragOn);
    el.dropzone.addEventListener('dragleave', function (event) {
      if (event.relatedTarget && el.dropzone.contains(event.relatedTarget)) return;
      el.dropzone.classList.remove('is-drag');
    });
    el.dropzone.addEventListener('drop', function (event) {
      dragOff(event);
      var file = event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files[0];
      if (file) loadFile(file);
    });

    window.addEventListener('dragover', function (event) { event.preventDefault(); });
    window.addEventListener('drop', function (event) {
      if (!el.dropzone.contains(event.target)) event.preventDefault();
    });

    el.fileInput.addEventListener('change', function () {
      var file = el.fileInput.files && el.fileInput.files[0];
      el.fileInput.value = '';
      if (file) loadFile(file);
    });

    el.changeFileBtn.addEventListener('click', function () {
      el.fileInput.click();
    });

    el.sheetSelect.addEventListener('change', function () {
      setFilter('all');
      selectSheet(Number(el.sheetSelect.value), { autoMap: true });
      showToast('已切换到工作表「' + state.sheet.name + '」', 'info');
    });

    el.headerRowInput.addEventListener('change', onHeaderRowChange);
    el.dataStartInput.addEventListener('change', onDataStartChange);
    el.nameCol.addEventListener('change', onMappingChange);
    el.inCol.addEventListener('change', onMappingChange);
    el.outCol.addEventListener('change', onMappingChange);

    el.remapBtn.addEventListener('click', function () {
      if (!state.workbook) return;
      var complete = applyAutoMapping();
      if (complete) showToast('已根据表头重新识别三列');
      else showToast('未能完全识别，请手动选择缺失的列。', 'info');
    });

    el.filterTabs.addEventListener('click', function (event) {
      var button = event.target.closest('.seg-btn');
      if (!button) return;
      var filter = button.getAttribute('data-filter');
      if (filter === state.filter) return;
      setFilter(filter);
      if (state.stats) renderResults();
      var scroller = el.tableShell.querySelector('.table-scroll');
      if (scroller) scroller.scrollTop = 0;
    });

    el.copyBtn.addEventListener('click', copyNames);
    el.exportBtn.addEventListener('click', exportExcel);
  }

  function init() {
    cacheDom();
    if (!C) {
      showToast('核心脚本未加载，请刷新页面重试。', 'error');
      return;
    }
    if (el.todayText) el.todayText.textContent = '今天 · ' + C.todayInfo().monthDay + ' ' + C.todayInfo().weekday;
    bindEvents();
    renderInitial();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
