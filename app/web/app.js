/* 座位轮转助手 · 主程序（原生 JS，零依赖）
 * 五步流程：导入 → 核对结构 → 轮转规则 → 预览 → 导出
 * 正式模式走 /api/*，?mock=1 时由 api.js 切换到内置样例与前端模拟。
 */
(function () {
  'use strict';

  var E = window.SeatEngine;
  var PARAMS = new URLSearchParams(window.location.search);
  // file:// 打开时 location.search 正常；某些旧浏览器兜底
  if (!PARAMS.has('mock') && /[?&]mock(=1)?(&|$)/.test(window.location.search)) {
    PARAMS = new URLSearchParams('mock=1');
  }

  var $ = function (sel, root) { return (root || document).querySelector(sel); };
  var $$ = function (sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); };

  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (ch) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch];
    });
  }
  function clone(o) { return o == null ? o : JSON.parse(JSON.stringify(o)); }
  function clamp(v, a, b) { return Math.max(a, Math.min(b, v)); }
  function debounce(fn, ms) {
    var t = null;
    return function () {
      var args = arguments, self = this;
      clearTimeout(t);
      t = setTimeout(function () { fn.apply(self, args); }, ms);
    };
  }
  function baseName(name) {
    return String(name || '座位表').replace(/\.xlsx$/i, '');
  }

  /* ================================================================ 常量 */
  var ROLE_META = {
    seat: { label: '座位', desc: '学生座位，参与轮转' },
    struct: { label: '结构', desc: '讲台、门窗、标题、排号、组标签等' },
    void: { label: '空位', desc: '写有「空」的占位格' },
    ignored: { label: '忽略', desc: '课代表区等非座位文字' }
  };
  var ROLE_ORDER = ['seat', 'struct', 'void', 'ignored'];
  var GROUP_COLORS = ['#2f6f5e', '#33547f', '#8a6a2f', '#6c4f74', '#3c6a7d', '#8a5a44', '#4f6b39', '#5b5a86', '#7b6240', '#41695a'];
  var FONT_MAP = {
    '宋体': '"Songti SC","SimSun",serif',
    '楷体': '"Kaiti SC","KaiTi","STKaiti",serif',
    'Times New Roman': '"Times New Roman","Songti SC",serif',
    'Arial': 'Arial,"PingFang SC",sans-serif',
    '微软雅黑': '"Microsoft YaHei","PingFang SC",sans-serif'
  };
  var FONT_DEFAULT = '"PingFang SC","Hiragino Sans GB","Microsoft YaHei",sans-serif';

  var PRESETS = [
    { id: 'mine', name: '我的常用', desc: '整组左移 1 组 + 组内后移 2 排', icon: '★',
      rules: { mode: 'combo', row_move: { steps: 2, toward: 'back' }, col_move: { steps: 0, toward: 'right' }, group_move: { steps: 1, toward: 'left' } } },
    { id: 'row', name: '前后轮转', desc: '每个大组内，全体前后移动', icon: '↕',
      rules: { mode: 'inner', row_move: { steps: 1, toward: 'back' }, col_move: { steps: 0, toward: 'right' }, group_move: { steps: 0, toward: 'right' } } },
    { id: 'col', name: '左右轮转', desc: '每个大组内，全体左右移动', icon: '↔',
      rules: { mode: 'inner', row_move: { steps: 0, toward: 'back' }, col_move: { steps: 1, toward: 'right' }, group_move: { steps: 0, toward: 'right' } } },
    { id: 'rowcol', name: '前后 + 左右', desc: '组内斜向轮转，换得更彻底', icon: '⤢',
      rules: { mode: 'inner', row_move: { steps: 1, toward: 'back' }, col_move: { steps: 1, toward: 'right' }, group_move: { steps: 0, toward: 'right' } } },
    { id: 'swap', name: '整组轮转', desc: '大组整体换位置，组内不动', icon: '⇄',
      rules: { mode: 'swap', row_move: { steps: 0, toward: 'back' }, col_move: { steps: 0, toward: 'right' }, group_move: { steps: 1, toward: 'right' } } },
    { id: 'combo', name: '整组 + 组内轮转', desc: '整组换位后，组内也转动', icon: '⟲',
      rules: { mode: 'combo', row_move: { steps: 1, toward: 'back' }, col_move: { steps: 1, toward: 'right' }, group_move: { steps: 1, toward: 'right' } } },
    { id: 'class', name: '全班整体轮转', desc: '忽略走道，全教室一起转', icon: '◎',
      rules: { mode: 'class', row_move: { steps: 1, toward: 'back' }, col_move: { steps: 1, toward: 'right' }, group_move: { steps: 0, toward: 'right' } } },
    { id: 'custom', name: '自定义', desc: '沿用当前参数，自由微调', icon: '✎', rules: null }
  ];

  /* ================================================================ 状态 */
  var CYCLE_KEY = 'seatRotation.cycleWeeks';
  function readCycleWeeks() {
    try {
      var v = parseInt(window.localStorage.getItem(CYCLE_KEY), 10);
      if (v >= 1 && v <= 4) return v;
    } catch (e) { /* file:// 下可能不可用 */ }
    return 2;
  }
  function writeCycleWeeks(n) {
    try { window.localStorage.setItem(CYCLE_KEY, String(n)); } catch (e) { /* ignore */ }
  }

  var state = {
    mock: false,
    step: 1,
    workbook: null,
    sheetIndex: 0,
    sheet: null,
    sheetMap: new Map(),
    sheetGeom: null,
    roles: {},
    seats: [],
    baseGroups: [],
    groups: [],
    groupMeta: { merges: [], names: {}, fixed: {} },
    groupLabels: [],
    podium: { r: null, c: null, relation: 'above' },
    fixedSeats: new Set(),
    detectWarnings: [],
    detectResult: null,
    rules: {
      mode: 'combo',
      row_move: { steps: 2, toward: 'back' },
      col_move: { steps: 0, toward: 'right' },
      group_move: { steps: 1, toward: 'left' },
      rounds: 1
    },
    rotate: null,
    rotateLoading: false,
    rotateError: null,
    rotateCheck: null,
    rotateToken: 0,
    previewRound: 1,
    compare: false,
    showRoles: true,
    zoom: { edit: 0, preview: 0 },
    scales: { edit: null, preview: null },
    history: [],
    groupNames: {},
    exportNameTouched: false,
    cycleWeeks: readCycleWeeks(),
    period: null,
    periodEnabled: true,
    periodStart: '',
    periodInfo: null,
    periodInfoKey: '',
    periodLoading: false,
    periodError: null,
    periodToken: 0,
    periodTimer: null,
    periodPromise: null,
    exportResult: null,
    exportError: null,
    exportBusy: false,
    importBusy: false,
    regroupTimer: null,
    rotateTimer: null,
    renameSnapshot: null
  };

  /* ================================================================ 启动 */
  function boot() {
    state.mock = PARAMS.has('mock');
    writeCycleWeeks(state.cycleWeeks); // 首次打开也把默认周期落盘，保证持久化
    window.SeatAPI.init({ mock: state.mock });
    bindEvents();
    renderModeChip();
    renderStepper();
    if (state.mock) {
      $('#import-demo').hidden = false;
      $('#import-error').hidden = true;
    }
    showStep(1);
    if (state.mock && PARAMS.get('demo') === '1') {
      setTimeout(autoDemo, 60);
    }
  }

  async function autoDemo() {
    try {
      await loadDemoWorkbook();
      var step = parseInt(PARAMS.get('step'), 10) || 2;
      showStep(clamp(step, 2, 5));
      if (step >= 4) {
        var round = parseInt(PARAMS.get('round'), 10);
        if (round > 0) {
          state.previewRound = round;
          renderStep4();
        }
      }
      if (PARAMS.get('exported') === '1') {
        showStep(5);
        await doExport();
      }
    } catch (e) {
      showImportError(e && e.message ? e.message : '演示数据加载失败');
    }
  }

  function renderModeChip() {
    var chip = $('#mode-chip'), label = $('#mode-label');
    chip.classList.toggle('mode-chip--mock', state.mock);
    label.textContent = state.mock ? '演示模式' : '本地服务';
    chip.title = state.mock
      ? '演示模式：使用内置真实样例，轮转与导出为前端模拟，不会生成真实文件'
      : '已连接本机服务：所有数据只在本机处理，不访问外网';
  }

  /* ================================================================ 步骤条 */
  var STEP_DEFS = [
    { n: 1, label: '导入' },
    { n: 2, label: '核对结构' },
    { n: 3, label: '轮转规则' },
    { n: 4, label: '预览' },
    { n: 5, label: '导出' }
  ];

  function canGoto(n) {
    if (n <= 1) return true;
    if (n === 2 || n === 3) return !!state.sheet;
    if (n === 4 || n === 5) return !!state.sheet && !!state.rotate;
  }

  function renderStepper() {
    var html = '';
    STEP_DEFS.forEach(function (s, i) {
      var active = state.step === s.n;
      var done = state.step > s.n && canGoto(s.n);
      var disabled = !canGoto(s.n) && !active;
      html += '<button class="step-item' + (active ? ' is-active' : '') + (done ? ' is-done' : '') +
        (disabled ? ' is-disabled' : '') + '" data-action="goto" data-step="' + s.n + '"' +
        (disabled ? ' disabled' : '') + ' aria-current="' + (active ? 'step' : 'false') + '">' +
        '<span class="step-num">' + (done ? '✓' : s.n) + '</span>' +
        '<span class="step-label">' + s.label + '</span></button>';
      if (i < STEP_DEFS.length - 1) html += '<span class="step-line' + (state.step > s.n ? ' is-done' : '') + '" aria-hidden="true"></span>';
    });
    $('#stepper').innerHTML = html;
  }

  function showStep(n) {
    n = clamp(n, 1, 5);
    if (!canGoto(n)) {
      toast(n === 4 || n === 5 ? '请先完成轮转规则，生成结果后再看预览' : '请先导入座位表', 'warn');
      return;
    }
    state.step = n;
    document.body.dataset.step = String(n);
    STEP_DEFS.forEach(function (s) {
      var el = $('#step-' + s.n);
      if (el) el.hidden = s.n !== n;
    });
    renderStepper();
    if (n === 1) renderStep1();
    if (n === 2) renderStep2();
    if (n === 3) renderStep3();
    if (n === 4) {
      if (state.previewRound <= 0 && state.rotate) state.previewRound = 1;
      renderStep4();
    }
    if (n === 5) renderStep5();
    closePopover();
    window.scrollTo(0, 0);
    if (n === 2 || n === 4) {
      // 布局稳定后再按可用空间算一次缩放
      setTimeout(renderCurrentGrid, 30);
    }
  }

  /* ================================================================ 导入 */
  function renderStep1() {
    $('#import-error').hidden = !$('#import-error-text').textContent;
    $('#import-demo').hidden = !state.mock;
  }

  function showImportBusy(on, title, sub) {
    state.importBusy = on;
    var box = $('#import-status');
    box.hidden = !on;
    if (on) {
      $('#import-status-title').textContent = title || '正在读取工作簿…';
      $('#import-status-sub').textContent = sub || '解析工作表与单元格格式';
      $('#import-error').hidden = true;
      $('#sheet-choose').hidden = true;
    }
    $('#dropzone').classList.toggle('is-busy', on);
  }

  function showImportError(msg) {
    state.importBusy = false;
    showImportBusy(false);
    $('#import-error').hidden = false;
    $('#import-error-text').textContent = msg || '未知错误';
    $('#import-demo').hidden = !state.mock;
  }

  async function handleFile(file) {
    if (!file) return;
    $('#import-error').hidden = true;
    $('#sheet-choose').hidden = true;
    if (!/\.xlsx$/i.test(file.name)) {
      showImportError('只支持 .xlsx 格式的 Excel 文件。如果原文件是 .xls，请先用 Excel 另存为 .xlsx。');
      return;
    }
    if (file.size > 20 * 1024 * 1024) {
      showImportError('文件过大（' + (file.size / 1024 / 1024).toFixed(1) + ' MB），最多支持 20 MB。请删除无关图片或另存一份精简版。');
      return;
    }
    if (file.size === 0) {
      showImportError('文件是空的，请检查后重新选择。');
      return;
    }
    if (state.mock) {
      toast('演示模式无法解析真实 Excel，已载入内置样例数据', 'info');
      await loadDemoWorkbook();
      return;
    }
    showImportBusy(true, '正在读取 ' + file.name, '上传到本机服务并解析工作表');
    try {
      var b64 = await window.SeatAPI.fileToBase64(file);
      var wb = await window.SeatAPI.importWorkbook({ filename: file.name, data: b64 });
      state.workbook = wb;
      renderWorkbookMeta();
      if (!wb.sheets || !wb.sheets.length) throw new Error('工作簿里没有可用的工作表');
      if (wb.sheets.length > 1) {
        showImportBusy(false);
        renderSheetChooser(wb);
      } else {
        await loadSheet(wb.sheets[0].index, { jump: true });
      }
    } catch (e) {
      showImportError(e && e.message ? e.message : '导入失败，请重试');
    }
  }

  function renderWorkbookMeta() {
    if (!state.workbook) return;
    $('#file-chip').textContent = state.workbook.filename;
    var wrap = $('#sheet-pick-wrap');
    var sel = $('#sheet-select');
    if (state.workbook.sheets && state.workbook.sheets.length > 1) {
      wrap.hidden = false;
      sel.innerHTML = state.workbook.sheets.map(function (s) {
        return '<option value="' + s.index + '"' + (s.index === state.sheetIndex ? ' selected' : '') + '>' +
          esc(s.name) + '（' + s.rows + '×' + s.cols + (s.hidden ? '，隐藏' : '') + '）</option>';
      }).join('');
    } else {
      wrap.hidden = true;
    }
  }

  function renderSheetChooser(wb) {
    $('#sheet-choose').hidden = false;
    $('#sheet-choose-file').textContent = wb.filename;
    $('#sheet-choose-list').innerHTML = wb.sheets.map(function (s) {
      return '<button class="sheet-item" data-action="pick-sheet" data-index="' + s.index + '">' +
        '<span class="sheet-name">' + esc(s.name) + '</span>' +
        '<span class="sheet-meta">' + s.rows + ' 行 × ' + s.cols + ' 列' + (s.hidden ? ' · 隐藏工作表' : '') + '</span>' +
        '<span class="sheet-arrow">选择 →</span></button>';
    }).join('');
  }

  async function loadSheet(index, opts) {
    opts = opts || {};
    if (!state.workbook) throw new Error('尚未导入工作簿');
    showImportBusy(true, '正在解析「' + sheetName(index) + '」', '读取单元格、合并区域、行高列宽与样式');
    try {
      var sheetRes = await window.SeatAPI.getSheet(state.workbook.workbook_id, index);
      var detectRes = await window.SeatAPI.detect(state.workbook.workbook_id, index);
      applySheet(sheetRes.sheet, detectRes);
      state.sheetIndex = index;
      renderWorkbookMeta();
      showImportBusy(false);
      $('#sheet-choose').hidden = true;
      $('#sheet-choose-list').innerHTML = '';
      showStep(2);
      scheduleRotate(0);
      if (opts.awaitRotate) await runRotate();
    } catch (e) {
      showImportBusy(false);
      showImportError(e && e.message ? e.message : '读取工作表失败');
      throw e;
    }
  }

  function sheetName(index) {
    if (!state.workbook || !state.workbook.sheets) return '工作表';
    var s = state.workbook.sheets.filter(function (x) { return x.index === index; })[0];
    return s ? s.name : '工作表';
  }

  async function loadDemoWorkbook() {
    var data = window.MOCK_DATA;
    if (!data) throw new Error('演示数据未加载（缺少 mock/mock.js）');
    showImportBusy(true, '正在载入示例座位表…', '使用内置的真实座位表数据');
    state.workbook = {
      workbook_id: 'mock_wb_001',
      filename: data.filename,
      sheets: [{ index: 0, name: data.sheet.name, hidden: false, rows: data.sheet.rows, cols: data.sheet.cols }],
      warnings: []
    };
    state.sheetIndex = 0;
    renderWorkbookMeta();
    try {
      var sheetRes = await window.SeatAPI.getSheet(state.workbook.workbook_id, 0);
      var detectRes = await window.SeatAPI.detect(state.workbook.workbook_id, 0);
      applySheet(sheetRes.sheet, detectRes);
      // 演示默认与老师常用一致：整组左移 1 组 + 组内后移 2 排，1 轮。
      setRules({
        mode: 'combo',
        row_move: { steps: 2, toward: 'back' },
        col_move: { steps: 0, toward: 'right' },
        group_move: { steps: 1, toward: 'left' },
        rounds: 1
      }, { silent: true });
      showImportBusy(false);
      showStep(2);
      await runRotate();
    } catch (e) {
      showImportBusy(false);
      showImportError(e && e.message ? e.message : '加载示例数据失败');
      throw e;
    }
  }

  function applySheet(sheet, detect) {
    state.sheet = sheet;
    state.sheetIndex = sheet.index;
    state.detectResult = detect;
    state.roles = Object.assign({}, (detect && detect.roles) || {});
    state.baseGroups = ((detect && detect.groups) || []).map(function (g) { return clone(g); });
    state.groupMeta = { merges: [], names: {}, fixed: {} };
    state.groupNames = {};
    state.fixedSeats = new Set();
    state.history = [];
    state.detectWarnings = (detect && detect.warnings) || [];
    state.period = (detect && detect.period) || null;
    state.periodEnabled = !!state.period;
    state.periodStart = '';
    state.periodInfo = null;
    state.periodInfoKey = '';
    state.periodLoading = false;
    state.periodError = null;
    state.periodToken++;
    state.podium = {
      r: detect && detect.podium ? detect.podium.r : null,
      c: detect && detect.podium ? detect.podium.c : null,
      relation: (detect && detect.podium && detect.podium.relation) || 'above'
    };
    prepareSheetGeometry();
    state.groupLabels = collectGroupLabels(detect);
    state.seats = buildSeatsFromRoles();
    state.rotate = null;
    state.rotateError = null;
    state.rotateCheck = null;
    state.previewRound = 1;
    state.exportResult = null;
    state.exportError = null;
    state.exportNameTouched = false;
    var nameInput = $('#export-name');
    if (nameInput) nameInput.value = '';
    recomputeGroups();
    renderAll();
  }

  function prepareSheetGeometry() {
    var sheet = state.sheet;
    state.sheetMap = new Map();
    (sheet.cells || []).forEach(function (c) {
      state.sheetMap.set(E.key(c.r, c.c), c);
    });
    var hiddenRows = new Set((sheet.hidden_rows || []).map(Number));
    var hiddenCols = new Set((sheet.hidden_cols || []).map(Number));
    var visibleRows = [], visibleCols = [];
    for (var r = 1; r <= sheet.rows; r++) if (!hiddenRows.has(r)) visibleRows.push(r);
    for (var c = 1; c <= sheet.cols; c++) if (!hiddenCols.has(c)) visibleCols.push(c);
    var colPx = {}, rowPx = {};
    var dcw = sheet.default_col_width != null ? sheet.default_col_width : 9;
    var drh = sheet.default_row_height != null ? sheet.default_row_height : 20;
    visibleCols.forEach(function (cc) {
      var w = sheet.col_widths && sheet.col_widths[cc] != null ? sheet.col_widths[cc] : dcw;
      colPx[cc] = Math.max(18, Math.round(w * 7 + 5));
    });
    visibleRows.forEach(function (rr) {
      var h = sheet.row_heights && sheet.row_heights[rr] != null ? sheet.row_heights[rr] : drh;
      rowPx[rr] = Math.max(18, Math.round(h * 96 / 72));
    });
    // 合并区域（映射到可见行列）
    var anchor = new Map(), covered = new Set();
    (sheet.merges || []).forEach(function (m) {
      var cs = visibleCols.filter(function (cc) { return cc >= m.c && cc < m.c + (m.colspan || 1); });
      var rs = visibleRows.filter(function (rr) { return rr >= m.r && rr < m.r + (m.rowspan || 1); });
      if (!cs.length || !rs.length) return;
      var info = { ref: m.ref, r: m.r, c: m.c, rowspan: rs.length, colspan: cs.length };
      anchor.set(E.key(m.r, m.c), info);
      rs.forEach(function (rr) {
        cs.forEach(function (cc) {
          if (rr === m.r && cc === m.c) return;
          covered.add(E.key(rr, cc));
        });
      });
    });
    state.sheetGeom = {
      visibleRows: visibleRows, visibleCols: visibleCols,
      rowPx: rowPx, colPx: colPx,
      anchor: anchor, covered: covered,
      hiddenRows: hiddenRows, hiddenCols: hiddenCols,
      natW: 0, natH: 0
    };
  }

  function collectGroupLabels(detect) {
    var labels = [], used = new Set();
    var cells = (state.sheet && state.sheet.cells) || [];
    ((detect && detect.groups) || []).forEach(function (g) {
      var found = null;
      for (var i = 0; i < cells.length; i++) {
        var cell = cells[i];
        if (String(cell.text == null ? '' : cell.text).trim() === g.name) { found = cell; break; }
      }
      if (!found) {
        // 兜底：组名形如「3组」，在组列范围内找同文本
        var re = /^(\d+|[一二三四五六七八九十]+)组$/;
        if (re.test(g.name)) {
          for (var j = 0; j < cells.length; j++) {
            var c2 = cells[j];
            if (String(c2.text || '').trim() === g.name && c2.c >= g.c1 && c2.c <= g.c2) { found = c2; break; }
          }
        }
      }
      if (!found) return;
      var k = E.key(found.r, found.c);
      if (used.has(k)) return;
      used.add(k);
      var merge = (state.sheet.merges || []).filter(function (m) { return m.r === found.r && m.c === found.c; })[0];
      labels.push({
        r: found.r, c: found.c, text: String(found.text).trim(),
        c_end: merge ? found.c + (merge.colspan || 1) - 1 : found.c
      });
    });
    return labels;
  }

  /* ================================================================ 座位 / 大组 */
  function buildSeatsFromRoles() {
    var seats = [];
    Object.keys(state.roles).forEach(function (k) {
      if (state.roles[k] !== 'seat') return;
      var p = E.parseKey(k);
      var cell = state.sheetMap.get(k);
      seats.push({
        r: p.r, c: p.c,
        text: cell && cell.text != null ? String(cell.text) : '',
        group_id: -1
      });
    });
    seats.sort(function (a, b) { return a.r - b.r || a.c - b.c; });
    return seats;
  }

  function recomputeGroups() {
    state.seats = buildSeatsFromRoles();
    var bases = (state.baseGroups || []).map(function (g) {
      return {
        name: g.name,
        r1: g.r1, r2: g.r2, c1: g.c1, c2: g.c2,
        rows: g.rows != null ? g.rows : g.r2 - g.r1 + 1,
        cols: g.cols != null ? g.cols : g.c2 - g.c1 + 1,
        baseKey: g.c1 + '-' + g.c2
      };
    });
    var parent = {};
    var exists = {};
    bases.forEach(function (b) { parent[b.baseKey] = b.baseKey; exists[b.baseKey] = true; });
    function find(k) { while (parent[k] !== k) { parent[k] = parent[parent[k]]; k = parent[k]; } return k; }
    function union(a, b) { a = find(a); b = find(b); if (a !== b) parent[b] = a; }
    (state.groupMeta.merges || []).forEach(function (rec) {
      var keys = rec.filter(function (k) { return exists[k]; });
      for (var i = 1; i < keys.length; i++) union(keys[0], keys[i]);
    });
    var buckets = {};
    bases.forEach(function (b) {
      var root = find(b.baseKey);
      if (!buckets[root]) buckets[root] = [];
      buckets[root].push(b);
    });
    var groups = Object.keys(buckets).map(function (root) {
      var bucket = buckets[root].slice().sort(function (a, b) { return a.c1 - b.c1; });
      var r1 = Infinity, r2 = -Infinity, c1 = Infinity, c2 = -Infinity;
      bucket.forEach(function (b) {
        r1 = Math.min(r1, b.r1); r2 = Math.max(r2, b.r2);
        c1 = Math.min(c1, b.c1); c2 = Math.max(c2, b.c2);
      });
      var memberKeys = bucket.map(function (b) { return b.baseKey; }).sort();
      var mergedKey = memberKeys.join('+');
      var name = state.groupMeta.names[mergedKey] || bucket.map(function (b) { return b.name; }).join('+');
      var fixed = memberKeys.length > 0 && memberKeys.every(function (k) { return !!state.groupMeta.fixed[k]; });
      return {
        memberKeys: memberKeys, key: mergedKey, name: name,
        r1: r1, r2: r2, c1: c1, c2: c2,
        rows: r2 - r1 + 1, cols: c2 - c1 + 1,
        fixed: fixed, seat_count: 0
      };
    });
    groups.sort(function (a, b) { return a.c1 - b.c1 || a.r1 - b.r1; });
    groups.forEach(function (g, i) { g.id = i; });
    state.groups = groups;
    assignSeatGroups();
  }

  function assignSeatGroups() {
    var byId = new Map();
    state.groups.forEach(function (g) { byId.set(g.id, 0); });
    state.seats.forEach(function (s) {
      var g = null;
      for (var i = 0; i < state.groups.length; i++) {
        var gg = state.groups[i];
        if (s.r >= gg.r1 && s.r <= gg.r2 && s.c >= gg.c1 && s.c <= gg.c2) { g = gg; break; }
      }
      if (!g && state.groups.length) {
        // 不落在任何组 bbox 内：按列就近归组，避免“无人管”的座位
        var best = null, bestD = Infinity;
        state.groups.forEach(function (gg) {
          var mid = (gg.c1 + gg.c2) / 2;
          var d = Math.abs(s.c - mid) + Math.abs(s.r - ((gg.r1 + gg.r2) / 2)) * 0.01;
          if (d < bestD) { bestD = d; best = gg; }
        });
        g = best;
      }
      s.group_id = g ? g.id : -1;
      if (g) byId.set(g.id, byId.get(g.id) + 1);
    });
    state.groups.forEach(function (g) {
      g.seat_count = byId.get(g.id) || 0;
      var list = state.seats.filter(function (s) { return s.group_id === g.id; });
      if (list.length) {
        var r1 = Infinity, r2 = -Infinity, c1 = Infinity, c2 = -Infinity;
        list.forEach(function (s) {
          r1 = Math.min(r1, s.r); r2 = Math.max(r2, s.r);
          c1 = Math.min(c1, s.c); c2 = Math.max(c2, s.c);
        });
        g.r1 = r1; g.r2 = r2; g.c1 = c1; g.c2 = c2;
        g.rows = r2 - r1 + 1; g.cols = c2 - c1 + 1;
      }
    });
  }

  function groupAt(r, c) {
    for (var i = 0; i < state.groups.length; i++) {
      var g = state.groups[i];
      if (r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2) return g;
    }
    return null;
  }

  function groupAtCol(c) {
    for (var i = 0; i < state.groups.length; i++) {
      var g = state.groups[i];
      if (c >= g.c1 && c <= g.c2) return g;
    }
    return null;
  }

  function groupColor(id) { return GROUP_COLORS[id % GROUP_COLORS.length]; }

  /* ================================================================ 结构变更 */
  function snapshot() {
    return {
      label: '结构调整',
      roles: clone(state.roles),
      baseGroups: clone(state.baseGroups),
      groupMeta: clone(state.groupMeta),
      groupNames: clone(state.groupNames),
      fixedSeats: Array.from(state.fixedSeats),
      podium: clone(state.podium)
    };
  }

  function pushHistory(label) {
    var snap = snapshot();
    snap.label = label || '结构调整';
    state.history.push(snap);
    if (state.history.length > 50) state.history.shift();
    updateUndoBtn();
  }

  function updateUndoBtn() {
    var btn = $('#undo-btn');
    if (!btn) return;
    btn.disabled = state.history.length === 0;
    btn.title = state.history.length
      ? '撤销「' + state.history[state.history.length - 1].label + '」（⌘Z），还剩 ' + state.history.length + ' 步'
      : '没有可撤销的操作';
  }

  function undo() {
    if (!state.history.length) return;
    var snap = state.history.pop();
    state.roles = snap.roles;
    state.baseGroups = snap.baseGroups;
    state.groupMeta = snap.groupMeta;
    state.groupNames = snap.groupNames || {};
    state.fixedSeats = new Set(snap.fixedSeats);
    state.podium = snap.podium;
    updateUndoBtn();
    recomputeGroups();
    renderAll();
    scheduleRegroupApi(0);
    scheduleRotate(120);
    toast('已撤销：' + snap.label, 'info');
  }

  function syncGroups(opts) {
    opts = opts || {};
    recomputeGroups();
    renderAll();
    if (!opts.skipServer) scheduleRegroupApi(150);
    scheduleRotate(200);
  }

  function scheduleRegroupApi(delay) {
    clearTimeout(state.regroupTimer);
    state.regroupTimer = setTimeout(callRegroupApi, delay == null ? 150 : delay);
  }

  async function callRegroupApi() {
    if (!state.sheet) return;
    var labels = state.groupLabels.filter(function (l) {
      return state.roles[E.key(l.r, l.c)] !== 'seat';
    });
    var seats = state.seats.map(function (s) { return { r: s.r, c: s.c, text: s.text }; });
    try {
      var res = await window.SeatAPI.regroup(seats, labels);
      if (res && res.groups) {
        state.baseGroups = res.groups.map(function (g) { return clone(g); });
        recomputeGroups();
        renderAll();
        scheduleRotate(120);
        if (res.warnings && res.warnings.length && state.step === 2) {
          // 仅提示一次，不打扰
        }
      }
    } catch (e) {
      // 服务端不可用时保留本地结果，不影响继续操作
      if (!state.mock && state.step === 2) toast('重算大组失败：' + e.message, 'error');
    }
  }

  /* ================================================================ 网格渲染 */
  function buildGridHtml(geom, opts) {
    var sheet = state.sheet;
    var rowHeadW = 38, colHeadH = 26, groupRowH = 34;
    var totalW = rowHeadW;
    geom.visibleCols.forEach(function (c) { totalW += geom.colPx[c]; });
    var totalH = colHeadH + groupRowH;
    geom.visibleRows.forEach(function (r) { totalH += geom.rowPx[r]; });

    var html = '<table class="xlsx" style="width:' + totalW + 'px">';
    html += '<colgroup><col style="width:' + rowHeadW + 'px">';
    geom.visibleCols.forEach(function (c) { html += '<col style="width:' + geom.colPx[c] + 'px">'; });
    html += '</colgroup><thead>';
    html += '<tr class="colhead-row"><th class="corner"></th>';
    geom.visibleCols.forEach(function (c) {
      html += '<th class="colhead">' + E.colLetter(c) + '</th>';
    });
    html += '</tr><tr class="grouprow-row"><th class="corner corner--group"></th>';
    html += groupStripHtml(geom);
    html += '</tr></thead><tbody>';

    geom.visibleRows.forEach(function (r) {
      html += '<tr style="height:' + geom.rowPx[r] + 'px">';
      html += '<th class="rowhead">' + r + '</th>';
      geom.visibleCols.forEach(function (c) {
        html += cellHtml(r, c, geom, opts);
      });
      html += '</tr>';
    });
    html += '</tbody></table>';
    return { html: html, natW: totalW, natH: totalH };
  }

  function groupStripHtml(geom) {
    var cols = geom.visibleCols;
    var out = '', i = 0;
    while (i < cols.length) {
      var c = cols[i];
      var g = groupAtCol(c);
      var span = 1;
      if (g) {
        while (i + span < cols.length && cols[i + span] === c + span && groupAtCol(cols[i + span]) === g) span++;
      }
      var first = true;
      if (g) {
        for (var j = 0; j < i; j++) {
          if (cols[j] >= g.c1 && cols[j] <= g.c2) { first = false; break; }
        }
      }
      if (!g) {
        out += '<th class="grp grp--none" colspan="' + span + '"></th>';
      } else {
        var btns = '';
        if (first) {
          if (g.memberKeys && g.memberKeys.length > 1) {
            btns = '<button class="grp-btn" data-action="split-group" data-gid="' + g.id + '" title="拆分「' + esc(g.name) + '」为原始大组">拆</button>';
          } else if (state.groups.length > 1) {
            btns = '<button class="grp-btn" data-action="merge-group" data-gid="' + g.id + '" title="与相邻大组合并">并</button>';
          }
        }
        out += '<th class="grp' + (first ? ' grp--first' : '') + (g.fixed ? ' grp--fixed' : '') +
          '" colspan="' + span + '" data-gid="' + g.id + '" style="--gc:' + groupColor(g.id) + '">' +
          (first ? '<span class="grp-name">' + esc(g.name) + (g.fixed ? ' · 固定' : '') + '</span>' + btns : '') +
          '</th>';
      }
      i += span;
    }
    return out;
  }

  function styleDecls(style) {
    var css = [];
    if (style.fill) css.push('background-color:#' + String(style.fill).replace(/^#/, ''));
    if (style.color) css.push('color:#' + String(style.color).replace(/^#/, ''));
    css.push('font-family:' + (FONT_MAP[style.font] || FONT_DEFAULT));
    if (style.size) css.push('font-size:' + (Math.round(style.size * 96 / 72 * 10) / 10) + 'px');
    if (style.bold) css.push('font-weight:700');
    css.push('text-align:' + ({ center: 'center', right: 'right', left: 'left' }[style.align] || 'center'));
    css.push('vertical-align:' + ({ center: 'middle', top: 'top', bottom: 'bottom' }[style.valign] || 'middle'));
    return css.join(';');
  }

  function cellHtml(r, c, geom, opts) {
    var k = E.key(r, c);
    if (geom.covered.has(k)) return '';
    var merge = geom.anchor.get(k);
    var cell = state.sheetMap.get(k);
    var style = (cell && cell.style) || {};
    var role = state.roles[k] || 'blank';
    // 预览时：组长标记（单元格底色）跟随学生移动到目标格
    if (opts.mode === 'preview' && opts.round > 0 && opts.marksFollow) {
      var markInfo = opts.assignByTo && opts.assignByTo.get(k);
      if (markInfo && markInfo.from) {
        var srcCell = state.sheetMap.get(E.key(markInfo.from[0], markInfo.from[1]));
        var srcFill = srcCell && srcCell.style ? srcCell.style.fill : null;
        // 注意：源座位没有底色时也要**清掉**目标位原有的底色（否则旧组长格会残留橙色）
        style = Object.assign({}, style, { fill: srcFill || null });
      }
    }
    var isSeat = role === 'seat';
    var fixed = state.fixedSeats.has(k);
    var cls = ['c'];
    if (style.border) cls.push('bd');
    if (role !== 'blank' && (state.showRoles || opts.mode === 'preview')) cls.push('role-' + role);
    if (style.wrap) cls.push('wrap');
    var text = cell && cell.text != null ? String(cell.text) : '';
    var content = esc(text);
    var badge = '';
    var origLine = '';

    if (opts.mode === 'preview' && isSeat) {
      var info = opts.round > 0 ? opts.assignByTo.get(k) : null;
      if (opts.round > 0) {
        if (info) {
          text = info.text != null ? String(info.text) : '';
          content = text ? esc(text) : '<span class="muted-text">（空）</span>';
          var moved = info.from[0] !== r || info.from[1] !== c;
          if (moved) {
            cls.push('pv-moved');
            badge = '<span class="move-badge" title="轮转前在 ' + E.a1(info.from[0], info.from[1]) + '">' + E.a1(info.from[0], info.from[1]) + '</span>';
          } else {
            cls.push('pv-stay');
          }
          if (opts.compare) {
            var origText = cell && cell.text != null ? String(cell.text) : '';
            if (moved) origLine = '<span class="orig-name">原 ' + esc(origText || '空') + '</span>';
          }
        } else {
          cls.push('pv-empty');
          content = '<span class="muted-text">空位</span>';
        }
      } else {
        cls.push('pv-seat');
        if (opts.compare && text) origLine = '';
      }
      if (fixed) {
        cls.push('pv-fixed');
        badge += '<span class="pin" title="固定座位：不参与轮转">固</span>';
      }
    } else if (isSeat) {
      if (fixed) {
        cls.push('is-fixed');
        badge = '<span class="pin" title="固定座位：不参与轮转">固</span>';
      }
    }

    var attrs = ' data-r="' + r + '" data-c="' + c + '" data-key="' + k + '" data-role="' + role + '"';
    if (fixed) attrs += ' data-fixed="1"';
    if (merge) {
      if (merge.rowspan > 1) attrs += ' rowspan="' + merge.rowspan + '"';
      if (merge.colspan > 1) attrs += ' colspan="' + merge.colspan + '"';
    }
    var title = E.a1(r, c) + (merge ? '（合并 ' + merge.ref + '）' : '') +
      (text ? '：' + text : '') + ' · ' + (ROLE_META[role] ? ROLE_META[role].label : '空白');
    var extra = badge + (origLine ? origLine : '');
    return '<td class="' + cls.join(' ') + '" style="' + styleDecls(style) + '"' + attrs +
      ' title="' + esc(title) + '">' + content + extra + '</td>';
  }

  function renderGrid(container, opts) {
    if (!state.sheet || !state.sheetGeom) { container.innerHTML = ''; return; }
    var geom = state.sheetGeom;
    var built = buildGridHtml(geom, opts);
    container.innerHTML =
      '<div class="grid-canvas" style="width:' + built.natW + 'px;height:' + built.natH + 'px">' +
      built.html.replace('<table class="xlsx"', '<table class="xlsx" style="width:' + built.natW + 'px;height:' + built.natH + 'px"') +
      '</div>';
    var canvas = container.firstElementChild;
    var table = canvas.querySelector('table');
    var availW = Math.max(80, container.clientWidth - 4);
    var availH = Math.max(80, container.clientHeight - 4);
    var fit = Math.min(availW / built.natW, availH / built.natH);
    if (!isFinite(fit) || fit <= 0) fit = 1;
    fit = clamp(fit, 0.22, 1.35);
    var zs = state.zoom[opts.key] || 0;
    var scale = zs === 0 ? fit : fit * Math.pow(1.14, zs);
    scale = clamp(scale, 0.18, 2.6);
    canvas.style.width = Math.round(built.natW * scale) + 'px';
    canvas.style.height = Math.round(built.natH * scale) + 'px';
    table.style.transform = 'translateZ(0) scale(' + scale + ')';
    state.scales[opts.key] = scale;
    updateZoomLabel(opts.key, scale, zs);
  }

  function updateZoomLabel(key, scale, zs) {
    var el = $(key === 'edit' ? '#zoom-value' : '#pzoom-value');
    if (!el) return;
    el.textContent = Math.round(scale * 100) + '%';
    el.title = zs === 0 ? '当前按窗口大小等比缩放，点击保持' : '点击恢复「适应窗口」';
    el.classList.toggle('is-fit', zs === 0);
  }

  function renderCurrentGrid() {
    if (state.step === 2) renderEditGrid();
    else if (state.step === 4) renderPreviewGrid();
  }

  function gridOptsFor(key, extra) {
    return Object.assign({ key: key, mode: key === 'edit' ? 'edit' : 'preview', round: 0, compare: false }, extra || {});
  }

  function renderEditGrid() {
    var container = $('#grid-edit');
    if (!container) return;
    renderGrid(container, gridOptsFor('edit'));
    var g = state.sheetGeom;
    if (g && (g.hiddenRows.size || g.hiddenCols.size)) {
      $('.grid-hint').textContent = '已折叠隐藏的 ' + g.hiddenRows.size + ' 行 / ' + g.hiddenCols.size +
        ' 列（与原表一致）。单击任意格子修改角色；标记为「座位」的格子可单独固定。';
    }
  }

  function renderPreviewGrid() {
    var container = $('#grid-preview');
    if (!container) return;
    var round = state.previewRound;
    var ctx = roundContext(round);
    var opts = gridOptsFor('preview', {
      round: round, compare: state.compare, assignByTo: ctx.assignByTo,
      marksFollow: moveMarksOn(),
    });
    renderGrid(container, opts);
    renderPreviewLegend(round);
  }

  function renderPreviewLegend(round) {
    var legend = $('#preview-legend');
    if (!legend) return;
    var on = moveMarksOn();
    var parts = ['<span class="lg lg--moved">移动过</span>',
      '<span class="lg lg--fixed">固定座位</span>'];
    parts.push(on
      ? '<span class="lg lg--marks">组长底色：<b>跟随学生</b></span>'
      : '<span class="lg lg--marks">组长底色：<b>留在原位</b></span>');
    legend.innerHTML = parts.join('');
  }

  function roundContext(round) {
    var ctx = { assignByTo: new Map(), empty: new Set(), unplaced: [], warnings: [], moved: 0, placed: 0 };
    if (round > 0 && state.rotate && state.rotate.rounds && state.rotate.rounds[round - 1]) {
      var rd = state.rotate.rounds[round - 1];
      (rd.assignments || []).forEach(function (a) {
        ctx.assignByTo.set(E.key(a.to[0], a.to[1]), a);
        if (a.from[0] !== a.to[0] || a.from[1] !== a.to[1]) ctx.moved++;
        ctx.placed++;
      });
      (rd.empty_seats || []).forEach(function (p) { ctx.empty.add(E.key(p[0], p[1])); });
      ctx.unplaced = rd.unplaced || [];
      ctx.warnings = rd.warnings || [];
    }
    return ctx;
  }

  /* ================================================================ 第 2 步 */
  function renderStep2() {
    renderWorkbookMeta();
    renderStats();
    renderGroupList();
    renderPodiumSeg();
    renderRoleLegend();
    renderEditGrid();
    updateUndoBtn();
  }

  function renderRoleLegend() {
    var el = $('#role-legend');
    if (!el) return;
    el.innerHTML = ROLE_ORDER.map(function (r) {
      return '<span class="lg lg--' + r + '">' + ROLE_META[r].label + '</span>';
    }).join('');
  }

  function renderPodiumSeg() {
    $$('#podium-seg button').forEach(function (btn) {
      var v = btn.dataset.podium;
      btn.classList.toggle('is-active', state.podium.relation === v);
    });
  }

  function renderStats() {
    var el = $('#statsbar');
    if (!el) return;
    var counts = state.groups.map(function (g) { return g.seat_count; }).filter(function (n) { return n > 0; });
    var min = counts.length ? Math.min.apply(null, counts) : 0;
    var max = counts.length ? Math.max.apply(null, counts) : 0;
    var per = counts.length ? (min === max ? min + ' 人' : min + '~' + max + ' 人') : '—';
    var warnings = collectWarnings();
    var fixedCount = state.fixedSeats.size;
    var html = '';
    html += statItem('座位', state.seats.length, '个');
    html += statItem('大组', state.groups.length, '组');
    html += statItem('每组人数', per, '');
    html += statItem('固定座位', fixedCount, '个');
    html += '<div class="stat stat--warnable">' +
      '<button class="stat-warn" data-action="toggle-warn-panel" aria-expanded="' + (state.warnPanelOpen ? 'true' : 'false') + '">' +
      '提示 <b>' + warnings.length + '</b> <span class="caret">' + (state.warnPanelOpen ? '▾' : '▸') + '</span></button></div>';
    html += '<div class="stats-spacer"></div>';
    if (fixedCount) html += '<button class="ghost-btn xs" data-action="clear-fixed">清除固定座位</button>';
    html += '<button class="ghost-btn xs" data-action="rerun-regroup" title="重新按组标签 / 连续座位列识别大组">重新识别大组</button>';
    if (state.warnPanelOpen && warnings.length) {
      html += '<div class="stats-warns">' + warnings.map(function (w) {
        var locate = firstRef(w.text);
        return '<button class="warn-item warn--' + (w.type || 'info') + '"' + (locate ? ' data-action="locate" data-ref="' + locate + '"' : ' data-action="noop"') + '>' +
          '<span class="warn-dot"></span><span class="warn-text">' + esc(w.text) + '</span>' +
          (locate ? '<span class="warn-locate">定位 ' + locate + '</span>' : '') + '</button>';
      }).join('') + '</div>';
    }
    el.innerHTML = html;
  }

  function statItem(label, value, unit) {
    return '<div class="stat"><span class="stat-k">' + label + '</span>' +
      '<span class="stat-v">' + esc(value) + (unit ? '<i>' + unit + '</i>' : '') + '</span></div>';
  }

  function collectWarnings() {
    var list = [];
    var seen = {};
    function add(type, text) {
      if (!text || seen[text]) return;
      seen[text] = true;
      list.push({ type: type, text: text });
    }
    state.detectWarnings.forEach(function (t) { add('info', t); });
    if (state.seats.length === 0) add('error', '当前没有任何「座位」格，请至少标记一个座位');
    var ungrouped = state.seats.filter(function (s) { return s.group_id < 0 || !state.groups.length; });
    if (ungrouped.length) {
      add('error', '有 ' + ungrouped.length + ' 个座位未归入大组（' +
        ungrouped.slice(0, 5).map(function (s) { return E.a1(s.r, s.c); }).join('、') + '），轮转时不会移动');
    }
    state.groups.forEach(function (g) {
      if (g.seat_count === 0) add('warn', '大组「' + g.name + '」里没有座位');
    });
    if (state.rotateError) add('error', '轮转计算失败：' + state.rotateError);
    if (state.rotate && state.rotate.warnings) {
      state.rotate.warnings.forEach(function (t) { add('warn', t); });
    }
    return list;
  }

  function firstRef(text) {
    var m = String(text || '').match(/\b([A-Z]{1,2}[1-9]\d?)\b/);
    return m ? m[1] : '';
  }

  function renderGroupList() {
    var el = $('#group-list');
    if (!el) return;
    $('#group-count-chip').textContent = state.groups.length + ' 组';
    if (!state.groups.length) {
      el.innerHTML = '<p class="empty-hint">还没有识别到大组。请先把座位标记好，再点「重新识别大组」。</p>';
      return;
    }
    el.innerHTML = state.groups.map(function (g) {
      return '<div class="group-item' + (g.fixed ? ' is-fixed' : '') + '" data-gid="' + g.id + '">' +
        '<span class="g-swatch" style="background:' + groupColor(g.id) + '"></span>' +
        '<div class="g-main">' +
        '<input class="g-name" value="' + esc(g.name) + '" data-action="rename-group" data-gid="' + g.id + '" aria-label="大组名称">' +
        '<div class="g-meta">' + g.seat_count + ' 座 · ' + g.cols + ' 列 × ' + g.rows + ' 行</div>' +
        '</div>' +
        '<div class="g-ops">' +
        '<label class="g-fixed" title="固定后组内和整组都不参与轮转"><input type="checkbox" data-action="fix-group" data-gid="' + g.id + '"' + (g.fixed ? ' checked' : '') + '><span>固定</span></label>' +
        '<div class="g-btns">' +
        (g.memberKeys && g.memberKeys.length > 1
          ? '<button class="mini-btn" data-action="split-group" data-gid="' + g.id + '">拆分</button>'
          : (state.groups.length > 1 ? '<button class="mini-btn" data-action="merge-group" data-gid="' + g.id + '">合并</button>' : '')) +
        '</div></div></div>';
    }).join('');
  }

  /* ================================================================ 角色浮层 */
  function openCellPopover(td) {
    var r = +td.dataset.r, c = +td.dataset.c;
    var k = E.key(r, c);
    var cell = state.sheetMap.get(k);
    var text = cell && cell.text != null ? String(cell.text) : '';
    var role = state.roles[k] || 'blank';
    var fixed = state.fixedSeats.has(k);
    var merge = state.sheetGeom.anchor.get(k);
    var g = groupAt(r, c);
    var el = $('#cell-popover');
    var html = '<div class="pop-head">' +
      '<span class="pop-ref">' + E.a1(r, c) + (merge ? ' · 合并 ' + merge.ref : '') + '</span>' +
      '<button class="icon-btn" data-action="close-popover" aria-label="关闭">×</button></div>' +
      '<div class="pop-text">' + (text ? esc(text) : '<span class="muted">（无文字）</span>') + '</div>' +
      '<div class="pop-sub">当前：' + (ROLE_META[role] ? ROLE_META[role].label : '空白') + (g ? ' · ' + esc(g.name) : '') + '</div>' +
      '<div class="role-opts">' + ROLE_ORDER.map(function (r2) {
        return '<button class="role-opt role-opt--' + r2 + (role === r2 ? ' is-active' : '') + '" data-action="set-role" data-role="' + r2 + '" data-key="' + k + '">' +
          '<span class="ro-dot"></span><span class="ro-label">' + ROLE_META[r2].label + '</span>' +
          '<span class="ro-desc">' + ROLE_META[r2].desc + '</span></button>';
      }).join('') + '</div>' +
      '<label class="pop-fixed' + (role === 'seat' ? '' : ' is-disabled') + '">' +
      '<input type="checkbox" data-action="toggle-fixed" data-key="' + k + '"' + (fixed ? ' checked' : '') + (role === 'seat' ? '' : ' disabled') + '>' +
      '<span>固定该座位（该学生不参与轮转）</span></label>' +
      '<p class="pop-hint">改错了可以按 ⌘Z 撤销。合并单元格整块修改。</p>';
    el.innerHTML = html;
    el.hidden = false;
    var rect = td.getBoundingClientRect();
    var pw = 300, ph = el.offsetHeight || 320;
    var left = clamp(rect.left + rect.width / 2 - pw / 2, 12, window.innerWidth - pw - 12);
    var top = rect.bottom + 8;
    if (top + ph > window.innerHeight - 12) top = Math.max(12, rect.top - ph - 8);
    el.style.left = Math.round(left) + 'px';
    el.style.top = Math.round(top) + 'px';
    state.popoverKey = k;
  }

  function refreshPopover() {
    if (!state.popoverKey) return;
    if ($('#cell-popover').hidden) { state.popoverKey = null; return; }
    var td = document.querySelector('#grid-edit td[data-key="' + state.popoverKey + '"]');
    if (td) openCellPopover(td);
    else closePopover();
  }

  function closePopover() {
    var el = $('#cell-popover');
    if (el) el.hidden = true;
    state.popoverKey = null;
  }

  function setRole(k, role) {
    if (!ROLE_META[role]) return;
    var old = state.roles[k] || 'blank';
    if (old === role) return;
    var pos = E.parseKey(k);
    pushHistory('把 ' + E.a1(pos.r, pos.c) + ' 设为「' + ROLE_META[role].label + '」');
    if (role === 'blank') delete state.roles[k];
    else state.roles[k] = role;
    if (role !== 'seat') state.fixedSeats.delete(k);
    syncGroups();
    refreshPopover();
  }

  function toggleFixedSeat(k) {
    pushHistory('切换固定座位 ' + E.a1(E.parseKey(k).r, E.parseKey(k).c));
    if (state.fixedSeats.has(k)) state.fixedSeats.delete(k);
    else state.fixedSeats.add(k);
    renderStats();
    renderEditGrid();
    updateUndoBtn();
    scheduleRotate(200);
    refreshPopover();
  }

  /* ================================================================ 大组操作 */
  function mergeAdjacentGroup(gid) {
    var idx = state.groups.findIndex(function (g) { return g.id === gid; });
    if (idx < 0 || state.groups.length < 2) return;
    var cur = state.groups[idx];
    var other = idx > 0 ? state.groups[idx - 1] : state.groups[idx + 1];
    if (!other) return;
    pushHistory('合并大组「' + cur.name + '」与「' + other.name + '」');
    var keys = cur.memberKeys.concat(other.memberKeys);
    state.groupMeta.merges.push(keys);
    var mergedKey = keys.slice().sort().join('+');
    state.groupMeta.names[mergedKey] = cur.name + '+' + other.name;
    var fixedKeys = cur.memberKeys.concat(other.memberKeys);
    var anyFixed = state.groups.filter(function (g) { return g.id === gid || g.id === other.id; })
      .some(function (g) { return g.fixed; });
    fixedKeys.forEach(function (k) { if (anyFixed) state.groupMeta.fixed[k] = true; });
    recomputeGroups();
    renderAll();
    scheduleRegroupApi(200);
    scheduleRotate(200);
  }

  function splitGroup(gid) {
    var g = state.groups.filter(function (x) { return x.id === gid; })[0];
    if (!g || !g.memberKeys || g.memberKeys.length <= 1) return;
    pushHistory('拆分大组「' + g.name + '」');
    state.groupMeta.merges = state.groupMeta.merges.filter(function (rec) {
      return !rec.some(function (k) { return g.memberKeys.indexOf(k) >= 0; });
    });
    delete state.groupMeta.names[g.key];
    recomputeGroups();
    renderAll();
    scheduleRegroupApi(200);
    scheduleRotate(200);
  }

  function renameGroup(gid, name) {
    var g = state.groups.filter(function (x) { return x.id === gid; })[0];
    if (!g) return;
    state.groupMeta.names[g.key] = name || g.name;
    if (state.renameSnapshot) {
      state.history.push(state.renameSnapshot);
      if (state.history.length > 50) state.history.shift();
      state.renameSnapshot = null;
      updateUndoBtn();
    }
    recomputeGroups();
    renderAll();
    scheduleRotate(150);
  }

  function toggleGroupFixed(gid) {
    var g = state.groups.filter(function (x) { return x.id === gid; })[0];
    if (!g) return;
    pushHistory((g.fixed ? '取消固定' : '固定') + '大组「' + g.name + '」');
    var next = !g.fixed;
    g.memberKeys.forEach(function (k) { state.groupMeta.fixed[k] = next; });
    recomputeGroups();
    renderAll();
    scheduleRotate(150);
  }

  function clearFixedSeats() {
    if (!state.fixedSeats.size) return;
    pushHistory('清除固定座位');
    state.fixedSeats.clear();
    renderAll();
    scheduleRotate(150);
  }

  /* ================================================================ 第 3 步：规则 */
  function normalizeRules(r) {
    var out = {
      mode: r.mode || 'inner',
      row_move: { steps: clamp(parseInt(r.row_move && r.row_move.steps, 10) || 0, 0, 3), toward: (r.row_move && r.row_move.toward) || 'back' },
      col_move: { steps: clamp(parseInt(r.col_move && r.col_move.steps, 10) || 0, 0, 3), toward: (r.col_move && r.col_move.toward) || 'right' },
      group_move: { steps: clamp(parseInt(r.group_move && r.group_move.steps, 10) || 0, 0, 3), toward: (r.group_move && r.group_move.toward) || 'right' },
      rounds: clamp(parseInt(r.rounds, 10) || 1, 1, 12)
    };
    // 组内 / 全班模式不使用整组移动；整组互换不使用前后 / 左右移动。
    // 这里直接清零，避免后端对「被忽略的参数」反复提示，摘要也更干净。
    if (out.mode === 'inner' || out.mode === 'class') out.group_move.steps = 0;
    if (out.mode === 'swap') { out.row_move.steps = 0; out.col_move.steps = 0; }
    return out;
  }

  function setRules(rules, opts) {
    opts = opts || {};
    state.rules = normalizeRules(rules);
    if (!opts.silent) {
      renderRulesUI();
      scheduleRotate(0);
    }
  }

  function presetMatches(p) {
    if (!p.rules) return false;
    var a = normalizeRules(p.rules), b = state.rules;
    return a.mode === b.mode &&
      a.row_move.steps === b.row_move.steps && a.row_move.toward === b.row_move.toward &&
      a.col_move.steps === b.col_move.steps && a.col_move.toward === b.col_move.toward &&
      a.group_move.steps === b.group_move.steps && a.group_move.toward === b.group_move.toward;
  }

  function activePresetId() {
    for (var i = 0; i < PRESETS.length; i++) {
      if (presetMatches(PRESETS[i])) return PRESETS[i].id;
    }
    return 'custom';
  }

  function applyPreset(id) {
    var p = PRESETS.filter(function (x) { return x.id === id; })[0];
    if (!p) return;
    if (!p.rules) { renderRulesUI(); return; }
    var next = clone(p.rules);
    next.rounds = state.rules.rounds;
    state.rules = normalizeRules(next);
    renderRulesUI();
    scheduleRotate(0);
  }

  function renderStep3() {
    renderWorkbookMeta();
    renderPresets();
    renderParams();
    renderFixedGroupChips();
    renderRuleSummary();
    renderRuleBadge();
  }

  function renderRulesUI() {
    renderPresets();
    renderParams();
    renderFixedGroupChips();
    renderRuleSummary();
  }

  function renderPresets() {
    var el = $('#preset-grid');
    if (!el) return;
    var active = activePresetId();
    el.innerHTML = PRESETS.map(function (p) {
      return '<button class="preset' + (p.id === 'mine' ? ' preset--mine' : '') + (active === p.id ? ' is-active' : '') + '" data-action="preset" data-preset="' + p.id + '">' +
        '<span class="preset-icon" aria-hidden="true">' + p.icon + '</span>' +
        '<span class="preset-name">' + p.name + '</span>' +
        '<span class="preset-desc">' + p.desc + '</span>' +
        (active === p.id ? '<span class="preset-check">✓</span>' : '') +
        '</button>';
    }).join('');
  }

  function renderParams() {
    var r = state.rules;
    $$('#seg-mode button').forEach(function (b) { b.classList.toggle('is-active', b.dataset.mode === r.mode); });
    setDirGroup('#row-dir', r.row_move);
    setDirGroup('#col-dir', r.col_move);
    setDirGroup('#group-dir', r.group_move);
    setStepper('#row-steps', r.row_move.steps);
    setStepper('#col-steps', r.col_move.steps);
    setStepper('#group-steps', r.group_move.steps);
    var range = $('#rounds-range'), out = $('#rounds-out');
    if (range) range.value = r.rounds;
    if (out) out.textContent = r.rounds + ' 轮';
    var rowColDim = (r.mode === 'swap');
    var groupDim = (r.mode === 'inner' || r.mode === 'class');
    dimRow($('#row-dir').closest('.param-row'), rowColDim);
    dimRow($('#col-dir').closest('.param-row'), rowColDim);
    dimRow($('#group-dir').closest('.param-row'), groupDim);
  }

  function dimRow(row, dim) {
    if (!row) return;
    row.classList.toggle('is-dim', !!dim);
    $$('button, input', row).forEach(function (b) {
      if (dim) b.disabled = true; else b.disabled = false;
    });
  }

  function setDirGroup(sel, move) {
    var active = move.steps === 0 ? 'none' : move.toward;
    $$(sel + ' button').forEach(function (b) { b.classList.toggle('is-active', b.dataset.dir === active); });
  }

  function setStepper(sel, steps) {
    $$(sel + ' button').forEach(function (b) { b.classList.toggle('is-active', +b.dataset.steps === steps); });
  }

  function renderFixedGroupChips() {
    var el = $('#fixed-groups');
    if (!el) return;
    if (!state.groups.length) { el.innerHTML = '<span class="muted">尚未识别大组</span>'; return; }
    el.innerHTML = state.groups.map(function (g) {
      return '<button class="chip-toggle' + (g.fixed ? ' is-active' : '') + '" data-action="toggle-group-fixed" data-gid="' + g.id + '">' +
        '<span class="chip-dot" style="background:' + groupColor(g.id) + '"></span>' + esc(g.name) + '</button>';
    }).join('');
  }

  function describeRules() {
    var r = state.rules;
    var parts = [];
    if (r.mode === 'inner') parts.push('每个大组内部各自轮转');
    else if (r.mode === 'swap') parts.push('各大组整体互换位置（组内不动）');
    else if (r.mode === 'class') parts.push('全班整体轮转（忽略大组边界，可跨走道）');
    else parts.push('先在各组内部轮转，再整组互换位置');

    var moves = [];
    if (r.row_move.steps) moves.push('向教室' + (r.row_move.toward === 'back' ? '后方' : '前方') + '移动 ' + r.row_move.steps + ' 排');
    if (r.col_move.steps) moves.push('向' + (r.col_move.toward === 'right' ? '右' : '左') + '移动 ' + r.col_move.steps + ' 列');
    if ((r.mode === 'swap' || r.mode === 'combo') && r.group_move.steps) {
      moves.push('整组向' + (r.group_move.toward === 'right' ? '右' : '左') + '平移 ' + r.group_move.steps + ' 组');
    }
    var fixed = state.groups.filter(function (g) { return g.fixed; }).map(function (g) { return '「' + g.name + '」'; });
    var text;
    if (!moves.length) {
      text = '当前没有设置任何移动，轮转结果和原表相同。请在下方选择前后 / 左右 / 整组移动中的至少一项。';
    } else {
      text = parts[0] + '：' + moves.join('、') + '；';
      if (fixed.length) text += fixed.join('') + '固定不动；';
      text += '共轮转 ' + r.rounds + ' 轮' + (r.rounds > 1 ? '（依次生成第 1～' + r.rounds + ' 轮）' : '') + '。';
    }
    return { text: text, moves: moves, fixed: fixed };
  }

  function renderRuleSummary() {
    var desc = describeRules();
    var p = $('#rule-desc');
    if (p) p.textContent = desc.text;
    var chips = $('#rule-chips');
    if (chips) {
      var items = [];
      if (state.podium.relation === 'below') items.push('<span class="rchip">讲台在下方</span>');
      else items.push('<span class="rchip">讲台在上方</span>');
      items.push('<span class="rchip">' + state.seats.length + ' 个座位</span>');
      items.push('<span class="rchip">' + state.groups.length + ' 个大组</span>');
      if (desc.fixed.length) items.push('<span class="rchip rchip--warn">固定 ' + desc.fixed.length + ' 组</span>');
      items.push('<span class="rchip">' + state.rules.rounds + ' 轮</span>');
      chips.innerHTML = items.join('');
    }
    var dh = $('#direction-hint');
    if (dh) {
      dh.textContent = state.podium.relation === 'below'
        ? '当前讲台在座位下方：向后方 = 远离讲台 = 表格行号减小（向上）。'
        : '当前讲台在座位上方：向后方 = 远离讲台 = 表格行号增大（向下）。';
    }
    renderRuleBadge();
    updateExportSummary();
  }

  function renderRuleBadge() {
    var badge = $('#rule-badge');
    if (!badge) return;
    badge.className = 'pill';
    if (state.rotateLoading) {
      badge.classList.add('busy');
      badge.innerHTML = '<span class="dot"></span>计算中…';
    } else if (state.rotateError) {
      badge.classList.add('bad');
      badge.innerHTML = '<span class="dot"></span>计算失败';
    } else if (state.rotateCheck) {
      badge.classList.add(state.rotateCheck.ok ? 'ok' : 'warn');
      badge.innerHTML = '<span class="dot"></span>' + (state.rotateCheck.ok ? '校验通过' : '有异常');
    } else {
      badge.classList.add('busy');
      badge.innerHTML = '<span class="dot"></span>等待计算';
    }
    renderValidation();
  }

  function renderValidation() {
    var el = $('#validation');
    if (!el) return;
    var html = '';
    if (state.rotateLoading) {
      html = '<div class="valid-line"><span class="spinner sm"></span> 正在计算第 1～' + state.rules.rounds + ' 轮…</div>';
    } else if (state.rotateError) {
      html = '<div class="valid-line valid-line--bad">⚠ ' + esc(state.rotateError) + '</div>';
    } else if (state.rotateCheck) {
      html = '<div class="valid-line ' + (state.rotateCheck.ok ? 'valid-line--ok' : 'valid-line--bad') + '">' +
        esc(state.rotateCheck.message) + '</div>';
      var warns = (state.rotate && state.rotate.warnings) || [];
      if (warns.length) {
        html += '<ul class="valid-warns">' + warns.slice(0, 4).map(function (t) {
          return '<li>' + esc(t) + '</li>';
        }).join('') + '</ul>';
      }
      if (state.rotate && state.rotate.rounds && state.rotate.rounds[0]) {
        var rd = state.rotate.rounds[0];
        var moved = (rd.assignments || []).filter(function (a) { return a.from[0] !== a.to[0] || a.from[1] !== a.to[1]; }).length;
        html += '<div class="valid-mini">第 1 轮：移动 ' + moved + ' 人 · 固定 ' + state.fixedSeats.size + ' 人' +
          (rd.empty_seats && rd.empty_seats.length ? ' · 空位 ' + rd.empty_seats.length : '') +
          (rd.unplaced && rd.unplaced.length ? ' · 未落座 ' + rd.unplaced.length : '') + '</div>';
      }
    }
    el.innerHTML = html;
  }

  /* ================================================================ 轮转调用 */
  function buildRotatePayload() {
    var fixedSeats = Array.from(state.fixedSeats).map(function (k) { return E.parseKey(k); }).map(function (p) { return [p.r, p.c]; });
    return {
      seats: state.seats.map(function (s) { return { r: s.r, c: s.c, text: s.text, group_id: s.group_id }; }),
      groups: state.groups.map(function (g) {
        return { id: g.id, name: g.name, r1: g.r1, r2: g.r2, c1: g.c1, c2: g.c2, rows: g.rows, cols: g.cols, seat_count: g.seat_count };
      }),
      podium: { r: state.podium.r, c: state.podium.c, relation: state.podium.relation || 'above' },
      rules: {
        mode: state.rules.mode,
        row_move: clone(state.rules.row_move),
        col_move: clone(state.rules.col_move),
        group_move: clone(state.rules.group_move),
        fixed_groups: state.groups.filter(function (g) { return g.fixed; }).map(function (g) { return g.id; }),
        fixed_seats: fixedSeats,
        rounds: state.rules.rounds
      }
    };
  }

  function scheduleRotate(delay) {
    clearTimeout(state.rotateTimer);
    state.rotateLoading = true;
    renderRuleBadge();
    state.rotateTimer = setTimeout(runRotate, delay == null ? 300 : delay);
  }

  async function runRotate() {
    if (!state.sheet) return null;
    if (!state.seats.length) {
      state.rotate = null;
      state.rotateLoading = false;
      state.rotateCheck = null;
      renderAll();
      return null;
    }
    var token = ++state.rotateToken;
    state.rotateLoading = true;
    state.rotateError = null;
    renderRuleBadge();
    try {
      var res = await window.SeatAPI.rotate(buildRotatePayload());
      if (token !== state.rotateToken) return null;
      state.rotate = res;
      state.rotateCheck = E.validate(res, state.seats, buildRotatePayload().rules.fixed_seats);
      state.rotateError = null;
      if (res && res.rounds) {
        state.previewRound = clamp(state.previewRound || 1, 0, res.rounds.length);
      }
      return res;
    } catch (e) {
      if (token !== state.rotateToken) return null;
      state.rotate = null;
      state.rotateCheck = null;
      state.rotateError = e && e.message ? e.message : '轮转计算失败';
      return null;
    } finally {
      if (token === state.rotateToken) {
        state.rotateLoading = false;
        renderAll();
      }
    }
  }

  /* ================================================================ 第 4 步：预览 */
  function renderStep4() {
    renderRoundTabs();
    renderPreviewGrid();
    renderPreviewStats();
    renderPreviewWarnings();
    renderPreviewSummary();
  }

  function renderRoundTabs() {
    var el = $('#round-tabs');
    if (!el) return;
    var rounds = (state.rotate && state.rotate.rounds) || [];
    var html = '<button class="tab' + (state.previewRound === 0 ? ' is-active' : '') + '" data-action="pick-round" data-round="0" role="tab">原始</button>';
    rounds.forEach(function (rd) {
      html += '<button class="tab' + (state.previewRound === rd.round ? ' is-active' : '') + '" data-action="pick-round" data-round="' + rd.round + '" role="tab">第 ' + rd.round + ' 轮</button>';
    });
    if (!rounds.length) html += '<span class="tabs-empty">还没有轮转结果</span>';
    el.innerHTML = html;
  }

  function renderPreviewStats() {
    var el = $('#preview-stats');
    if (!el) return;
    var round = state.previewRound;
    var badge = $('#preview-badge');
    if (!state.rotate || !state.rotate.rounds || !state.rotate.rounds.length) {
      el.innerHTML = '<p class="empty-hint">' + (state.rotateError ? esc(state.rotateError) : '正在等待轮转结果…') + '</p>';
      if (badge) { badge.className = 'pill busy'; badge.textContent = state.rotateError ? '计算失败' : '计算中…'; }
      return;
    }
    var ctx = roundContext(round);
    var fixedInRound = 0;
    if (round > 0) {
      (state.rotate.rounds[round - 1].assignments || []).forEach(function (a) {
        if (state.fixedSeats.has(E.key(a.to[0], a.to[1]))) fixedInRound++;
      });
    }
    var rd = round > 0 ? state.rotate.rounds[round - 1] : null;
    var items = [
      ['落座', round === 0 ? state.seats.length : ctx.placed],
      ['移动', round === 0 ? '—' : ctx.moved],
      ['固定', round === 0 ? state.fixedSeats.size : fixedInRound],
      ['空位', round === 0 ? '—' : (rd.empty_seats || []).length],
      ['未落座', round === 0 ? '—' : (rd.unplaced || []).length]
    ];
    el.innerHTML = items.map(function (it) {
      return '<div class="mini-stat"><span>' + it[0] + '</span><b>' + it[1] + '</b></div>';
    }).join('');
    if (badge) {
      var isOk = state.rotateCheck ? state.rotateCheck.ok : false;
      badge.className = 'pill ' + (isOk ? 'ok' : 'warn');
      badge.textContent = isOk ? '校验通过' : '有异常';
    }
  }

  function renderPreviewSummary() {
    var el = $('#preview-summary');
    if (!el) return;
    if (!state.rotate || !state.rotate.rounds || !state.rotate.rounds.length) {
      el.innerHTML = '<span class="sum-chip">预览不可用</span><span class="sum-text">请先在「轮转规则」里完成一次计算。</span>';
      return;
    }
    var ctx = roundContext(state.previewRound);
    var label = state.previewRound === 0 ? '原始座位' : '第 ' + state.previewRound + ' 轮';
    var seg = describeRules().text;
    el.innerHTML = '<span class="sum-chip">' + label + '</span>' +
      '<span class="sum-text">' + esc(seg) + '</span>' +
      '<span class="sum-stats">' + (state.previewRound === 0 ? ('共 ' + state.seats.length + ' 座') :
        ('移动 ' + ctx.moved + ' 人 · 固定 ' + state.fixedSeats.size + ' 人' +
          ((ctx.unplaced.length) ? ' · <b class="bad">未落座 ' + ctx.unplaced.length + ' 人</b>' : ''))) + '</span>' +
      '<button class="ghost-btn xs" data-action="goto" data-step="3">调整规则</button>';
  }

  function renderPreviewWarnings() {
    var list = $('#preview-warn-list');
    if (!list) return;
    var items = [];
    var seen = {};
    function add(type, text) {
      if (!text || seen[text]) return;
      seen[text] = true;
      items.push({ type: type, text: text });
    }
    if (state.rotateError) add('error', state.rotateError);
    if (state.rotate) (state.rotate.warnings || []).forEach(function (t) { add('warn', t); });
    if (state.previewRound > 0 && state.rotate && state.rotate.rounds[state.previewRound - 1]) {
      (state.rotate.rounds[state.previewRound - 1].warnings || []).forEach(function (t) { add('warn', t); });
      (state.rotate.rounds[state.previewRound - 1].unplaced || []).forEach(function (u) {
        add('error', u.text + ' 未落座（原 ' + E.a1(u.from[0], u.from[1]) + '）');
      });
    }
    $('#preview-warn-count').textContent = items.length;
    if (!items.length) {
      list.innerHTML = '<li class="warn-empty">没有警告，人数一一对应。</li>';
      return;
    }
    list.innerHTML = items.map(function (w) {
      var ref = firstRef(w.text);
      return '<li class="warn-item warn--' + w.type + '"' + (ref ? ' data-action="locate" data-ref="' + ref + '"' : '') + '>' +
        '<span class="warn-dot"></span><span class="warn-text">' + esc(w.text) + '</span>' +
        (ref ? '<span class="warn-locate">定位 ' + ref + '</span>' : '') + '</li>';
    }).join('');
  }

  function locateCell(ref) {
    var m = String(ref || '').match(/^([A-Z]{1,2})(\d{1,3})$/);
    if (!m) return;
    var c = 0;
    for (var i = 0; i < m[1].length; i++) c = c * 26 + (m[1].charCodeAt(i) - 64);
    var r = parseInt(m[2], 10);
    var container = state.step === 4 ? $('#grid-preview') : $('#grid-edit');
    if (!container) return;
    var td = container.querySelector('td[data-r="' + r + '"][data-c="' + c + '"]');
    if (!td) {
      var anchor = container.querySelector('td[data-r][data-c]');
      toast('这个座位在当前视图里不可见（可能被隐藏或已合并）', 'warn');
      return;
    }
    td.classList.add('is-flash');
    td.scrollIntoView({ block: 'center', inline: 'center', behavior: 'smooth' });
    setTimeout(function () { td.classList.remove('is-flash'); }, 1600);
  }

  function pickRound(n) {
    state.previewRound = n;
    renderStep4();
  }

  /* ================================================================ 第 5 步：导出 */
  function originalStem() {
    var name = state.workbook && state.workbook.filename ? state.workbook.filename : '座位表';
    return baseName(name);
  }

  function periodRoundInfo(roundNo) {
    var list = (state.periodInfo && state.periodInfo.rounds) || [];
    for (var i = 0; i < list.length; i++) {
      if (list[i].round === roundNo) return list[i];
    }
    return null;
  }

  /** 没有服务端 filename 时的本地兜底：原文件名 + 轮次/日期智能递进 */
  function fallbackNameForRound(roundNo) {
    var info = periodRoundInfo(roundNo);
    return window.SeatEngine.naming.suggestName(originalStem(), roundNo, info, state.rules.rounds);
  }

  /** 某一轮最终使用的文件名（不含 .xlsx） */
  function fileNameForRound(roundNo) {
    var info = periodRoundInfo(roundNo);
    var input = $('#export-name');
    if (!state.exportNameTouched) {
      if (info && info.filename) return String(info.filename).replace(/\.xlsx$/i, '');
      return fallbackNameForRound(roundNo);
    }
    var base = String((input && input.value) || '').trim().replace(/\.xlsx$/i, '') || fallbackNameForRound(roundNo);
    var cur = currentExportRound();
    var curRound = cur ? cur.round : 1;
    if (roundNo === curRound) return base;
    var naming = window.SeatEngine.naming;
    // 用户手改的名字里没有「第N轮 / 日期」时，多轮按任务约定兜底 _第N轮
    if (selectedScope() === 'all' && !naming.parseRound(base) && !naming.hasDate(base)) {
      return base + '_第' + roundNo + '轮';
    }
    var advanced = naming.advanceName(base, roundNo - curRound, info);
    if (!advanced || advanced === base) advanced = base + '_第' + roundNo + '轮';
    return advanced;
  }

  function syncFilenameInput() {
    if (state.exportNameTouched) return;
    var input = $('#export-name');
    if (!input) return;
    var cur = currentExportRound();
    var name = fileNameForRound(cur ? cur.round : 1);
    if (name) input.value = name;
  }

  function renderCycleUI() {
    $$('#cycle-seg button').forEach(function (b) {
      b.classList.toggle('is-active', parseInt(b.dataset.weeks, 10) === state.cycleWeeks);
    });
    var hint = $('#cycle-hint');
    if (hint) {
      hint.textContent = '每轮 ' + state.cycleWeeks + ' 周（' + (state.cycleWeeks * 7) + ' 天），影响表头与文件名日期';
    }
  }

  function setCycleWeeks(weeks) {
    var n = clamp(parseInt(weeks, 10) || 2, 1, 4);
    if (n === state.cycleWeeks) return;
    state.cycleWeeks = n;
    writeCycleWeeks(n);
    state.periodInfo = null;
    state.periodInfoKey = '';
    renderCycleUI();
    ensurePeriodPreview(true);
    renderFilePreview();
    renderExportSummary();
  }

  function renderStep5() {
    renderCycleUI();
    renderExportScopeLabels();
    renderPeriodUI();
    ensurePeriodPreview(false);
    renderMarksInfo();
    syncFilenameInput();
    renderFilePreview();
    renderExportSummary();
    renderExportResult();
  }

  /* ---------------- 表头「座位周期」 ---------------- */
  function periodDefaultStart() {
    if (!state.period || !state.period.end) return '';
    return window.SeatAPI.addDaysISO(state.period.end, 1) || '';
  }

  function renderPeriodUI() {
    var toggle = $('#export-period');
    var startInput = $('#period-start');
    var badge = $('#period-badge');
    var daysHint = $('#period-days-hint');
    var list = $('#period-rounds');
    var hint = $('#period-hint');
    if (!toggle) return;
    var has = !!state.period;
    var cycleDays = state.cycleWeeks * 7;
    toggle.disabled = !has;
    toggle.checked = has && !!state.periodEnabled;
    if (badge) {
      badge.hidden = !has;
      if (has) badge.textContent = state.period.ref + ' · ' + state.period.text + '（原周期 ' + (state.period.days || 15) + ' 天）';
    }
    if (!state.periodStart && has) state.periodStart = periodDefaultStart();
    if (startInput) {
      startInput.disabled = !has;
      startInput.value = state.periodStart || '';
    }
    if (daysHint) {
      daysHint.textContent = has
        ? '默认从原周期结束次日顺延'
        : '未检测到周期表头，仅文件名按轮次建议';
    }
    if (list) {
      var html;
      if (state.periodLoading) {
        html = '<li class="period-loading"><span class="spinner sm"></span> 正在计算每轮日期…</li>';
      } else if (state.periodError) {
        html = '<li class="period-error">' + esc(state.periodError) + '</li>';
      } else if (state.periodInfo && state.periodInfo.rounds && state.periodInfo.rounds.length) {
        var cur = currentExportRound();
        var curRound = cur ? cur.round : 1;
        html = state.periodInfo.rounds.map(function (r) {
          var isCur = selectedScope() === 'current' && r.round === curRound;
          var text;
          if (r.text) text = String(r.text).replace(/^座位周期：/, '');
          else if (r.start && r.end) text = String(r.start).replace(/-/g, '.') + '～' + String(r.end).replace(/-/g, '.');
          else text = '建议文件名：' + (r.filename || '—');
          return '<li class="' + (isCur ? 'is-current' : '') + '">' +
            '<span class="pr-round">第 ' + r.round + ' 轮</span>' +
            '<span class="pr-text">' + esc(text) + '</span>' +
            (isCur ? '<span class="pr-tag">当前导出</span>' : '') + '</li>';
        }).join('');
      } else if (!has) {
        html = '<li class="period-empty">未检测到表头「座位周期」，导出时不会改日期；文件名仍按轮次智能递进。</li>';
      } else {
        html = '<li class="period-empty">等待日期计算结果…</li>';
      }
      list.innerHTML = html;
    }
    if (hint) {
      hint.textContent = !has
        ? '表头日期不会变化；文件名按原文件名递进（第N轮→第N+1轮，带日期则替换日期）。'
        : (!state.periodEnabled
          ? '表头周期将保持原样，可能与合作老师看到的日期不一致；文件名仍按当前周期日期生成。'
          : '导出时把第 1 行表头替换为对应轮次的周期文字，合并单元格与样式保持不变。');
    }
  }

  function periodKey() {
    return state.rules.rounds + '|' + (state.periodStart || '') + '|' + state.cycleWeeks;
  }

  /** 拉取逐轮周期（含建议文件名）。即使没有「座位周期」表头也会调用，用于文件名建议。 */
  function ensurePeriodPreview(force) {
    if (!state.sheet) return Promise.resolve(null);
    if (!state.periodStart && state.period) state.periodStart = periodDefaultStart();
    var key = periodKey();
    if (!force && state.periodInfo && state.periodInfoKey === key) return Promise.resolve(state.periodInfo);
    if (state.periodPromise && !force) return state.periodPromise;
    var token = ++state.periodToken;
    state.periodLoading = true;
    state.periodError = null;
    renderPeriodUI();
    var promise = window.SeatAPI.period({
      workbook_id: state.workbook ? state.workbook.workbook_id : '',
      index: state.sheetIndex,
      rounds: state.rules.rounds,
      cycle_weeks: state.cycleWeeks,
      start_date: state.periodStart || undefined
    }).then(function (res) {
      if (token !== state.periodToken) return state.periodInfo;
      state.periodInfo = res;
      state.periodInfoKey = key;
      state.periodLoading = false;
      state.periodError = null;
      renderPeriodUI();
      syncFilenameInput();
      renderFilePreview();
      renderExportSummary();
      return res;
    }).catch(function (e) {
      if (token !== state.periodToken) return null;
      state.periodLoading = false;
      state.periodInfo = null;
      state.periodInfoKey = '';
      state.periodError = e && e.message ? e.message : '周期日期计算失败';
      renderPeriodUI();
      return null;
    }).then(function (res) {
      if (state.periodPromise === promise) state.periodPromise = null;
      return res;
    });
    state.periodPromise = promise;
    return promise;
  }

  function schedulePeriodPreview(force) {
    clearTimeout(state.periodTimer);
    state.periodTimer = setTimeout(function () { ensurePeriodPreview(force); }, 300);
  }

  function periodHeaderUpdates(round) {
    if (!state.periodEnabled || !state.period || !state.periodInfo) return [];
    var list = state.periodInfo.rounds || [];
    for (var i = 0; i < list.length; i++) {
      if (list[i].round === round) return [{ to: state.period.ref, text: list[i].text }];
    }
    return [];
  }

  /** 导出页开关：组长标记（底色）是否跟随学生移动，默认开启 */
  function moveMarksOn() {
    var el = $('#export-move-marks');
    return el ? !!el.checked : true;
  }

  /** 统计「有非默认底色」的座位数（真实样例如 8 个橙色组长格） */
  function countMarkedSeats() {
    var n = 0;
    state.seats.forEach(function (seat) {
      var cell = state.sheetMap.get(E.key(seat.r, seat.c));
      var fill = cell && cell.style && cell.style.fill;
      if (!fill) return;
      var f = String(fill).replace(/^#/, '').toUpperCase();
      if (f && f !== 'FFFFFF' && f !== 'FFFFFFFF' && f !== '00000000') n++;
    });
    return n;
  }

  function renderMarksInfo() {
    var badge = $('#export-marks-badge');
    var hint = $('#export-marks-hint');
    var n = countMarkedSeats();
    var on = moveMarksOn();
    if (badge) {
      badge.hidden = n === 0;
      badge.textContent = '检测到 ' + n + ' 个带底色的座位';
      badge.classList.toggle('is-off', !on);
    }
    if (hint) {
      hint.textContent = n === 0
        ? '未检测到带底色的座位，此开关不影响导出结果。'
        : (on
          ? '导出时这 ' + n + ' 个底色会跟着对应学生移动（预览也会跟随显示，可在预览页切换）。'
          : '已关闭：这 ' + n + ' 个底色会留在原坐标，不随学生移动。');
    }
  }

  function currentExportRound() {
    var rounds = (state.rotate && state.rotate.rounds) || [];
    if (!rounds.length) return null;
    var n = state.previewRound;
    if (n <= 0 || n > rounds.length) n = 1;
    return rounds[n - 1];
  }

  function selectedScope() {
    var el = document.querySelector('input[name="scope"]:checked');
    return el ? el.value : 'current';
  }

  function renderExportScopeLabels() {
    var rounds = (state.rotate && state.rotate.rounds) || [];
    var current = currentExportRound();
    var cur = $('#scope-current-label'), all = $('#scope-all-label');
    if (cur) cur.textContent = current ? ('导出第 ' + current.round + ' 轮') : '暂无可导出的轮次';
    if (all) all.textContent = rounds.length ? ('共 ' + rounds.length + ' 个文件（每轮一个）') : '暂无可导出的轮次';
  }

  function plannedFileNames() {
    var rounds = (state.rotate && state.rotate.rounds) || [];
    if (!rounds.length) return [];
    var list = selectedScope() === 'current' ? [currentExportRound()] : rounds;
    return list.filter(Boolean).map(function (rd) { return fileNameForRound(rd.round) + '.xlsx'; });
  }

  function renderFilePreview() {
    var el = $('#file-preview');
    if (!el) return;
    var names = plannedFileNames();
    if (!names.length) {
      el.innerHTML = '<li class="empty-hint">还没有轮转结果，无法导出。</li>';
      return;
    }
    el.innerHTML = names.map(function (n) {
      return '<li><span class="file-ico">XLSX</span><span class="file-name">' + esc(n) + '</span></li>';
    }).join('');
    var hint = $('#export-side-hint');
    if (hint) {
      hint.textContent = state.exportNameTouched
        ? '已按你输入的名字命名；多轮会在该名字上智能递进（无轮次/日期时加「_第N轮」）。'
        : (selectedScope() === 'all'
          ? '文件名按原文件名智能递进：第N轮→第N+1轮，带日期则替换为新轮日期。'
          : '文件名取自当前轮次的建议名，可在上方修改。');
    }
  }

  function renderExportSummary() {
    var el = $('#export-summary');
    if (!el) return;
    var rounds = (state.rotate && state.rotate.rounds) || [];
    var ok = state.rotateCheck && state.rotateCheck.ok;
    var cur = currentExportRound();
    var scopeText = selectedScope() === 'all'
      ? ('全部 ' + rounds.length + ' 轮')
      : (cur ? ('仅第 ' + cur.round + ' 轮') : '—');
    var items = [
      ['工作表', state.sheet ? state.sheet.name : '—'],
      ['导出范围', scopeText],
      ['可导出轮次', rounds.length ? rounds.length + ' 轮' : '—'],
      ['座位 / 大组', state.seats.length + ' / ' + state.groups.length],
      ['固定座位', state.fixedSeats.size],
      ['组长标记', moveMarksOn() ? '跟随学生' : '留在原位'],
      ['轮转周期', state.cycleWeeks + ' 周 / 轮'],
      ['表头周期', !state.period ? '未检测到'
        : (!state.periodEnabled ? '保持原样'
          : (state.periodLoading ? '计算中…' : (state.periodInfo ? '逐轮更新' : '待计算')))],
      ['人数校验', ok ? '通过' : '待确认']
    ];
    el.innerHTML = items.map(function (it) {
      return '<div class="sum-row"><span>' + it[0] + '</span><b>' + esc(it[1]) + '</b></div>';
    }).join('');
    var rs = $('#export-rule-summary');
    if (rs) rs.textContent = describeRules().text;
    renderExportScopeLabels();
  }

  function updateExportSummary() {
    if (state.step === 5) renderExportSummary();
  }

  async function doExport() {
    if (state.exportBusy) return;
    $('#export-error').hidden = true;
    var rounds = (state.rotate && state.rotate.rounds) || [];
    if (!rounds.length) {
      showExportError('还没有可导出的轮转结果，请先在「轮转规则」里完成计算。');
      return;
    }
    if (state.seats.length === 0) {
      showExportError('当前没有座位，无法导出。');
      return;
    }
    var input = $('#export-name');
    var touched = state.exportNameTouched;
    var base = touched ? String(input.value || '').trim().replace(/\.xlsx$/i, '') : '';
    if (touched) {
      base = base.replace(/[\\/:*?"<>|]/g, '').trim();
      if (!base) {
        showExportError('请填写文件名（不能包含 \\ / : * ? " < > | 等字符）。');
        return;
      }
      input.value = base;
    }
    var scope = selectedScope();
    var list = scope === 'current' ? [currentExportRound()] : rounds;
    list = list.filter(Boolean);
    if (!list.length) {
      showExportError('没有可导出的轮次。');
      return;
    }
    // 无论是否有表头周期，都先拿到 rounds[].filename 建议名
    var pinfo = await ensurePeriodPreview(false);
    if (!pinfo) {
      showExportError('文件名与轮次日期还没算好，请稍候重试。');
      return;
    }
    var fileNameList = list.map(function (rd) { return fileNameForRound(rd.round); });
    if (!touched) {
      base = fileNameList[0] || '';
      if (!base) {
        showExportError('无法生成建议文件名，请手动填写文件名后重试。');
        return;
      }
      var curForName = currentExportRound();
      if (input) input.value = fileNameForRound(curForName ? curForName.round : 1);
    }
    state.exportBusy = true;
    state.exportResult = null;
    state.lastExportMoveMarks = moveMarksOn();
    var btn = $('.export-actions .primary-btn');
    var oldText = btn ? btn.textContent : '';
    if (btn) { btn.disabled = true; btn.innerHTML = '<span class="spinner sm"></span> 正在导出…'; }
    renderExportResult();
    var payload = {
      workbook_id: state.workbook ? state.workbook.workbook_id : '',
      index: state.sheetIndex,
      filename_base: base,
      rounds: list.map(function (rd, idx) {
        var item = {
          round: rd.round,
          filename: fileNameList[idx] || base,
          moves: (rd.assignments || []).map(function (a) {
            return { to: [a.to[0], a.to[1]], from: [a.from[0], a.from[1]] };
          })
        };
        var updates = periodHeaderUpdates(rd.round);
        if (updates.length) item.header_updates = updates;
        return item;
      }),
      move_marks: moveMarksOn(),
      open_folder: !!(state.mock ? false : ($('#export-open-folder') && $('#export-open-folder').checked))
    };
    try {
      var res = await window.SeatAPI.exportWorkbook(payload);
      state.exportResult = res;
      state.exportError = null;
      toast('导出成功：' + ((res.files || []).length) + ' 个文件', 'ok');
    } catch (e) {
      state.exportError = e && e.message ? e.message : '导出失败，请重试';
      state.exportResult = null;
      showExportError(state.exportError);
    } finally {
      state.exportBusy = false;
      if (btn) { btn.disabled = false; btn.textContent = oldText || '导出 .xlsx'; }
      renderExportResult();
    }
  }

  function showExportError(msg) {
    var box = $('#export-error');
    if (!box) { toast(msg, 'error'); return; }
    box.hidden = false;
    $('#export-error-text').textContent = msg;
    state.exportError = msg;
  }

  function renderExportResult() {
    var box = $('#export-success');
    if (!box) return;
    if (state.exportBusy) {
      box.hidden = false;
      box.innerHTML = '<div class="success-head"><span class="spinner"></span><div><b>正在导出…</b><p>服务端会从原始工作簿逐轮生成文件</p></div></div>';
      return;
    }
    if (!state.exportResult) { box.hidden = true; box.innerHTML = ''; return; }
    var files = state.exportResult.files || [];
    var folder = state.exportResult.folder || '导出结果';
    box.hidden = false;
    var mockNote = (state.exportResult.mock || state.mock)
      ? '<p class="mock-note">演示模式：以上为模拟结果，不会真正生成文件。启动本地服务后导出会得到真实 .xlsx。</p>' : '';
    var marks = state.exportResult && typeof state.exportResult.move_marks === 'boolean'
      ? state.exportResult.move_marks : state.lastExportMoveMarks;
    var marksChip = marks == null ? '' :
      '<span class="success-marks">' + (marks ? '组长标记已跟随学生' : '组长标记留在原位') + '</span>';
    var applied = 0;
    files.forEach(function (f) { applied += parseInt(f.text_updates_applied, 10) || 0; });
    var headerChip = applied > 0
      ? '<span class="success-marks">表头周期已更新 ' + applied + ' 处</span>' : '';
    box.innerHTML = '<div class="success-head"><span class="success-check">✓</span>' +
      '<div><b>导出成功</b><p>共 ' + files.length + ' 个文件 · 保存到「' + esc(folder) + '」' + marksChip + headerChip + '</p></div></div>' +
      '<ul class="file-results">' + files.map(function (f) {
        var url = state.mock ? '#' : window.SeatAPI.downloadUrl(f.path);
        return '<li>' +
          '<span class="file-ico">XLSX</span>' +
          '<div class="fr-main"><b>' + esc(f.name) + '</b><span class="fr-path">' + esc(f.path) + '</span></div>' +
          '<div class="fr-actions">' +
          '<a class="btn btn--primary btn--sm" href="' + esc(url) + '" data-action="download" data-mock="' + (state.mock ? '1' : '0') + '" data-file="' + esc(f.name) + '" download>下载</a>' +
          '<button class="btn btn--ghost btn--sm" data-action="reveal" data-path="' + esc(f.path) + '">在文件夹中显示</button>' +
          '</div></li>';
      }).join('') + '</ul>' +
      (state.mock ? mockNote :
        '<div class="success-foot"><button class="btn btn--ghost btn--sm" data-action="reveal" data-path="' + esc(folder) + '">打开导出文件夹</button>' +
        '<span class="muted">找不到文件时，可点「在文件夹中显示」。</span></div>');
  }

  async function revealPath(path) {
    if (state.mock) { toast('演示模式不会打开访达；启动本地服务后可用。', 'info'); return; }
    try {
      await window.SeatAPI.reveal(path);
      toast('已在文件夹中显示', 'ok');
    } catch (e) {
      toast('无法在文件夹中显示：' + (e && e.message ? e.message : '未知原因'), 'error');
    }
  }

  /* ================================================================ 渲染总控 */
  function renderAll() {
    renderStepper();
    renderRuleSummary();
    if (state.step === 2) {
      renderWorkbookMeta();
      renderStats();
      renderGroupList();
      renderPodiumSeg();
      renderEditGrid();
      updateUndoBtn();
    } else if (state.step === 3) {
      renderParams();
      renderFixedGroupChips();
      renderPresets();
    } else if (state.step === 4) {
      renderRoundTabs();
      renderPreviewGrid();
      renderPreviewStats();
      renderPreviewWarnings();
      renderPreviewSummary();
    } else if (state.step === 5) {
      renderStep5();
    }
    renderRuleBadge();
  }

  /* ================================================================ 弹窗 / toast */
  function openModal(title, bodyHtml, footHtml) {
    $('#modal-title').textContent = title;
    $('#modal-body').innerHTML = bodyHtml;
    $('#modal-foot').innerHTML = footHtml || '<button class="btn btn--primary" data-action="close-modal">知道了</button>';
    $('#modal').hidden = false;
  }
  function closeModal() { $('#modal').hidden = true; }

  function toast(msg, type) {
    var box = $('#toasts');
    var el = document.createElement('div');
    el.className = 'toast toast--' + (type || 'info');
    el.innerHTML = '<span class="toast-dot"></span><span class="toast-text">' + esc(msg) + '</span>';
    box.appendChild(el);
    setTimeout(function () { el.classList.add('is-out'); }, 3000);
    setTimeout(function () { if (el.parentNode) el.parentNode.removeChild(el); }, 3500);
  }

  function showHelp() {
    openModal('使用说明', [
      '<div class="help-grid">',
      '<div class="help-item"><b>① 导入</b><p>把 .xlsx 座位表拖进来，或点击选择。文件只在本机处理，多工作表会先让你选择。</p></div>',
      '<div class="help-item"><b>② 核对结构</b><p>自动识别座位 / 结构 / 空位 / 忽略。单击格子可改角色；座位可单独固定；大组可改名、固定、相邻合并或拆分。支持 ⌘Z 撤销 50 步。</p></div>',
      '<div class="help-item"><b>③ 轮转规则</b><p>7 个预设 + 高级参数。讲台位置决定前后方向；规则改动后自动计算并校验人数。</p></div>',
      '<div class="help-item"><b>④ 预览</b><p>逐轮查看结果，角标是原座位号。固定座位、空位、未落座分别用不同颜色；警告可点击定位。</p></div>',
      '<div class="help-item"><b>⑤ 导出</b><p>文件名默认按原文件名智能递进（第N轮→第N+1轮；带日期则替换日期），不再加「_轮转后」。「轮转周期」可选每 1/2/3/4 周（默认 2 周，浏览器会记住），影响表头周期与文件名日期。「组长标记跟随学生」默认开启；「自动更新表头座位周期」也默认开启，跨年自动补年份。导出保留原表格式，成功后可直接下载或在文件夹中查看。</p></div>',
      '<div class="help-item"><b>完全离线</b><p>页面只访问本机 127.0.0.1，不加载任何外部资源。演示模式 <code>?mock=1</code> 使用内置真实样例，可在没有服务时体验完整流程。</p></div>',
      '</div>'
    ].join(''), '<button class="btn btn--primary" data-action="close-modal">知道了</button>');
  }

  /* ================================================================ 事件 */
  var resizeHandler = debounce(function () {
    if (state.step === 2 || state.step === 4) renderCurrentGrid();
    if (state.step === 2) renderStats();
  }, 160);

  function bindEvents() {
    document.addEventListener('click', onClick);
    document.addEventListener('change', onChange);
    document.addEventListener('input', onInput);
    document.addEventListener('keydown', onKeydown);
    window.addEventListener('resize', resizeHandler);
    document.addEventListener('dragover', function (e) { e.preventDefault(); });
    document.addEventListener('drop', function (e) { e.preventDefault(); });

    var dz = $('#dropzone');
    ['dragenter', 'dragover'].forEach(function (ev) {
      dz.addEventListener(ev, function (e) {
        e.preventDefault();
        dz.classList.add('is-over');
      });
    });
    ['dragleave', 'dragend'].forEach(function (ev) {
      dz.addEventListener(ev, function () { dz.classList.remove('is-over'); });
    });
    dz.addEventListener('drop', function (e) {
      e.preventDefault();
      dz.classList.remove('is-over');
      var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (f) handleFile(f);
    });
    dz.addEventListener('keydown', function (e) {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); e.stopPropagation(); $('#file-input').click(); }
    });
    $('#file-input').addEventListener('change', function (e) {
      var f = e.target.files && e.target.files[0];
      e.target.value = '';
      if (f) handleFile(f);
    });
    // 点击浮层外部关闭
    document.addEventListener('mousedown', function (e) {
      var pop = $('#cell-popover');
      if (pop.hidden) return;
      if (pop.contains(e.target) || (e.target.closest && e.target.closest('td.c'))) return;
      closePopover();
    });
    // 鼠标在网格内滚动时关闭浮层，避免错位
    document.addEventListener('wheel', function (e) {
      if (!$('#cell-popover').hidden && e.target.closest && e.target.closest('.grid-scroll')) closePopover();
    }, { passive: true });
  }

  function handleGridClick(t) {
    // 网格格子本身没有 data-action，单独处理：结构页打开角色浮层，预览页提示来源
    if (t.closest && t.closest('#cell-popover')) return;
    if (t.closest && t.closest('#grid-edit') && !t.closest('.grp-btn') && !t.closest('th')) {
      var td = t.closest('td.c');
      if (td) openCellPopover(td);
      return;
    }
    if (t.closest && t.closest('#grid-preview') && !t.closest('th')) {
      var ptd = t.closest('td.c');
      if (!ptd) return;
      var key = ptd.dataset.key;
      var info = state.previewRound > 0 ? roundContext(state.previewRound).assignByTo.get(key) : null;
      var text = (ptd.textContent || '').replace(/\s+/g, ' ').trim();
      var msg = E.a1(+ptd.dataset.r, +ptd.dataset.c) + '：' + (text || '（空）') +
        (info && (info.from[0] !== +ptd.dataset.r || info.from[1] !== +ptd.dataset.c)
          ? ('（轮转前在 ' + E.a1(info.from[0], info.from[1]) + '）') : '');
      toast(msg, 'info');
    }
  }

  function onClick(e) {
    var t = e.target;
    var actionEl = t.closest ? t.closest('[data-action]') : null;
    if (!actionEl) {
      handleGridClick(t);
      return;
    }
    var action = actionEl.dataset.action;
    if (action === 'noop') { e.preventDefault(); return; }
    switch (action) {
      case 'goto':
        e.preventDefault();
        showStep(parseInt(actionEl.dataset.step, 10));
        break;
      case 'pick-file':
        e.preventDefault();
        $('#file-input').click();
        break;
      case 'load-demo':
        e.preventDefault();
        loadDemoWorkbook();
        break;
      case 'pick-sheet':
        e.preventDefault();
        loadSheet(parseInt(actionEl.dataset.index, 10), { jump: true });
        break;
      case 'undo':
        e.preventDefault();
        undo();
        break;
      case 'open-help':
        e.preventDefault();
        showHelp();
        break;
      case 'close-modal':
        e.preventDefault();
        closeModal();
        break;
      case 'close-popover':
        e.preventDefault();
        closePopover();
        break;
      case 'set-role':
        e.preventDefault();
        setRole(actionEl.dataset.key, actionEl.dataset.role);
        break;
      case 'toggle-fixed':
        e.preventDefault();
        toggleFixedSeat(actionEl.dataset.key);
        break;
      case 'preset':
        e.preventDefault();
        applyPreset(actionEl.dataset.preset);
        break;
      case 'pick-round':
        e.preventDefault();
        pickRound(parseInt(actionEl.dataset.round, 10));
        break;
      case 'zoom-in': zoomBy('edit', 1); break;
      case 'zoom-out': zoomBy('edit', -1); break;
      case 'zoom-fit': zoomBy('edit', 0); break;
      case 'pzoom-in': zoomBy('preview', 1); break;
      case 'pzoom-out': zoomBy('preview', -1); break;
      case 'pzoom-fit': zoomBy('preview', 0); break;
      case 'toggle-warn-panel':
        state.warnPanelOpen = !state.warnPanelOpen;
        renderStats();
        break;
      case 'clear-fixed':
        clearFixedSeats();
        break;
      case 'rerun-regroup':
        scheduleRegroupApi(0);
        toast('正在按组标签 / 连续座位列重新识别大组', 'info');
        break;
      case 'merge-group':
        e.preventDefault();
        mergeAdjacentGroup(parseInt(actionEl.dataset.gid, 10));
        break;
      case 'split-group':
        e.preventDefault();
        splitGroup(parseInt(actionEl.dataset.gid, 10));
        break;
      case 'toggle-group-fixed':
        e.preventDefault();
        toggleGroupFixed(parseInt(actionEl.dataset.gid, 10));
        break;
      case 'locate':
        e.preventDefault();
        locateCell(actionEl.dataset.ref);
        break;
      case 'export':
        e.preventDefault();
        doExport();
        break;
      case 'reveal':
        e.preventDefault();
        revealPath(actionEl.dataset.path);
        break;
      case 'download':
        if (actionEl.dataset.mock === '1') {
          e.preventDefault();
          toast('演示模式不会产生真实文件；启动本地服务后即可下载。', 'info');
        }
        break;
      default:
        break;
    }
  }

  function zoomBy(key, dir) {
    if (dir === 0) state.zoom[key] = 0;
    else state.zoom[key] = clamp((state.zoom[key] || 0) + dir, -8, 10);
    if (key === 'edit') { if (state.step === 2) renderEditGrid(); }
    else { if (state.step === 4) renderPreviewGrid(); }
  }

  function onChange(e) {
    var el = e.target;
    var actionEl = el.closest ? el.closest('[data-action]') : null;
    var action = actionEl ? actionEl.dataset.action : '';
    if (action === 'fix-group') {
      toggleGroupFixed(parseInt(actionEl.dataset.gid, 10));
      return;
    }
    if (action === 'rename-group') {
      renameGroup(parseInt(actionEl.dataset.gid, 10), el.value.trim());
      return;
    }
    if (el.id === 'toggle-roles') {
      state.showRoles = el.checked;
      renderEditGrid();
      return;
    }
    if (el.id === 'toggle-compare') {
      state.compare = el.checked;
      renderPreviewGrid();
      return;
    }
    if (el.id === 'sheet-select') {
      loadSheet(parseInt(el.value, 10), { jump: false });
      return;
    }
    if (el.id === 'rounds-range') {
      state.rules.rounds = parseInt(el.value, 10) || 1;
      renderRulesUI();
      scheduleRotate(300);
      return;
    }
    if (el.name === 'scope') {
      syncFilenameInput();
      renderFilePreview();
      renderExportScopeLabels();
      renderPeriodUI();
      renderExportSummary();
      return;
    }
    if (el.id === 'export-name') {
      state.exportNameTouched = true;
      renderFilePreview();
      return;
    }
    if (el.id === 'export-move-marks') {
      var pv = $('#toggle-preview-marks');
      if (pv) pv.checked = el.checked;
      renderMarksInfo();
      renderExportSummary();
      if (state.step === 4) renderPreviewGrid();
      return;
    }
    if (el.id === 'toggle-preview-marks') {
      var ex = $('#export-move-marks');
      if (ex) ex.checked = el.checked;
      renderMarksInfo();
      renderExportSummary();
      if (state.step === 4) renderPreviewGrid();
      return;
    }
    if (el.id === 'export-period') {
      state.periodEnabled = el.checked;
      renderPeriodUI();
      renderExportSummary();
      if (el.checked) ensurePeriodPreview(true);
      return;
    }
    if (el.id === 'period-start') {
      state.periodStart = el.value;
      renderPeriodUI();
      schedulePeriodPreview(true);
      return;
    }
  }

  function onInput(e) {
    var el = e.target;
    if (el.id === 'export-name') {
      var cleaned = el.value.replace(/[\\/:*?"<>|]/g, '');
      if (cleaned !== el.value) el.value = cleaned;
      state.exportNameTouched = true;
      renderFilePreview();
      return;
    }
    if (el.dataset && el.dataset.action === 'rename-group') {
      return; // 改完（change）才提交，避免每次按键都重算
    }
  }

  function onKeydown(e) {
    if (e.key === 'Escape') {
      closePopover();
      if (!$('#modal').hidden) closeModal();
      return;
    }
    if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'z' && state.step === 2) {
      e.preventDefault();
      undo();
    }
  }

  // 记录大组改名前的快照，便于撤销
  document.addEventListener('focusin', function (e) {
    if (e.target && e.target.dataset && e.target.dataset.action === 'rename-group') {
      state.renameSnapshot = snapshot();
      state.renameSnapshot.label = '重命名大组';
    }
  });

  // 参数按钮（无 data-action，用 id 委派）
  document.addEventListener('click', function (e) {
    var t = e.target.closest ? e.target.closest('button') : null;
    if (!t) return;
    if (t.closest('#seg-mode') && t.dataset.mode) {
      state.rules.mode = t.dataset.mode;
      state.rules = normalizeRules(state.rules);
      if ((state.rules.mode === 'swap' || state.rules.mode === 'combo') && !state.rules.group_move.steps) {
        state.rules.group_move.steps = 1; // 切到整组模式时给一个有效默认，避免"没有移动"
      }
      renderRulesUI();
      scheduleRotate(300);
      return;
    }
    if (t.closest('#row-dir') && t.dataset.dir) {
      state.rules.row_move = { steps: t.dataset.dir === 'none' ? 0 : (state.rules.row_move.steps || 1), toward: t.dataset.dir === 'none' ? state.rules.row_move.toward : t.dataset.dir };
      renderRulesUI(); scheduleRotate(300);
      return;
    }
    if (t.closest('#col-dir') && t.dataset.dir) {
      state.rules.col_move = { steps: t.dataset.dir === 'none' ? 0 : (state.rules.col_move.steps || 1), toward: t.dataset.dir === 'none' ? state.rules.col_move.toward : t.dataset.dir };
      renderRulesUI(); scheduleRotate(300);
      return;
    }
    if (t.closest('#group-dir') && t.dataset.dir) {
      state.rules.group_move = { steps: t.dataset.dir === 'none' ? 0 : (state.rules.group_move.steps || 1), toward: t.dataset.dir === 'none' ? state.rules.group_move.toward : t.dataset.dir };
      renderRulesUI(); scheduleRotate(300);
      return;
    }
    if (t.closest('#row-steps') && t.dataset.steps) {
      state.rules.row_move.steps = parseInt(t.dataset.steps, 10);
      if (state.rules.row_move.toward !== 'front' && state.rules.row_move.toward !== 'back') state.rules.row_move.toward = 'back';
      renderRulesUI(); scheduleRotate(300);
      return;
    }
    if (t.closest('#col-steps') && t.dataset.steps) {
      state.rules.col_move.steps = parseInt(t.dataset.steps, 10);
      if (state.rules.col_move.toward !== 'left' && state.rules.col_move.toward !== 'right') state.rules.col_move.toward = 'right';
      renderRulesUI(); scheduleRotate(300);
      return;
    }
    if (t.closest('#group-steps') && t.dataset.steps) {
      state.rules.group_move.steps = parseInt(t.dataset.steps, 10);
      if (state.rules.group_move.toward !== 'left' && state.rules.group_move.toward !== 'right') state.rules.group_move.toward = 'right';
      renderRulesUI(); scheduleRotate(300);
      return;
    }
    if (t.closest('#cycle-seg') && t.dataset.weeks) {
      setCycleWeeks(parseInt(t.dataset.weeks, 10));
      return;
    }
    if (t.closest('#podium-seg') && t.dataset.podium) {
      state.podium.relation = t.dataset.podium;
      renderPodiumSeg();
      renderRuleSummary();
      scheduleRotate(200);
      return;
    }
  });

  window.addEventListener('DOMContentLoaded', boot);
  if (document.readyState !== 'loading') boot();
})();
