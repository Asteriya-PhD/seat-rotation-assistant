/* 座位轮转助手 · API 层（零依赖）
 * 契约见 ARCHITECTURE.md §3。真实模式一律相对路径请求本机服务；
 * ?mock=1 时切到 MockBackend，用内置真实样例（mock/mock.js）完整演示全流程。
 */
(function (global) {
  'use strict';

  function ApiError(message, detail) {
    this.name = 'ApiError';
    this.message = message || '请求失败';
    this.detail = detail || null;
  }
  ApiError.prototype = Object.create(Error.prototype);
  ApiError.prototype.constructor = ApiError;

  function sleep(ms) { return new Promise(function (res) { setTimeout(res, ms); }); }
  function clone(o) { return JSON.parse(JSON.stringify(o)); }

  /* ---------------- 日期工具（mock 周期计算；也供 app.js 复用） ---------------- */
  function pad2(n) { return n < 10 ? '0' + n : '' + n; }
  function parseISO(str) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(str || ''));
    if (!m) return null;
    var y = +m[1], mo = +m[2], d = +m[3];
    var dt = new Date(Date.UTC(y, mo - 1, d));
    if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== mo - 1 || dt.getUTCDate() !== d) return null;
    return dt;
  }
  function isoOf(dt) {
    return dt.getUTCFullYear() + '-' + pad2(dt.getUTCMonth() + 1) + '-' + pad2(dt.getUTCDate());
  }
  function addDaysISO(str, n) {
    var dt = parseISO(str);
    if (!dt) return null;
    dt.setUTCDate(dt.getUTCDate() + n);
    return isoOf(dt);
  }
  function fmtPeriodText(startISO, endISO) {
    var s = parseISO(startISO), e = parseISO(endISO);
    if (!s || !e) return '';
    var sy = s.getUTCFullYear(), sm = s.getUTCMonth() + 1, sd = s.getUTCDate();
    var ey = e.getUTCFullYear(), em = e.getUTCMonth() + 1, ed = e.getUTCDate();
    var startTxt = sy + '.' + sm + '.' + sd;
    var endTxt = (ey === sy ? '' : ey + '.') + em + '.' + ed;
    return '座位周期：' + startTxt + '～' + endTxt;
  }

  /* ============================================================ 真实后端 */
  var RealBackend = {
    label: '本地服务',
    async request(method, path, body) {
      var opts = { method: method, headers: {} };
      if (body !== undefined) {
        opts.headers['Content-Type'] = 'application/json; charset=utf-8';
        opts.body = JSON.stringify(body);
      }
      var res;
      try {
        res = await fetch(path, opts);
      } catch (e) {
        throw new ApiError('无法连接本地服务。请先双击「启动.command」启动服务，或打开 index.html?mock=1 使用演示模式。', e);
      }
      var text = '';
      try { text = await res.text(); } catch (e) { /* ignore */ }
      var data = null;
      if (text) { try { data = JSON.parse(text); } catch (e) { data = null; } }
      if (!res.ok) {
        var msg = (data && (data.error || data.message)) || ('服务返回错误（HTTP ' + res.status + '）');
        throw new ApiError(msg, data);
      }
      if (data === null) throw new ApiError('服务返回了无法解析的数据（不是 JSON）');
      return data;
    },
    async importWorkbook(payload) {
      return this.request('POST', '/api/import', { filename: payload.filename, data: payload.data });
    },
    async getSheet(workbookId, index) {
      return this.request('GET', '/api/sheet?workbook_id=' + encodeURIComponent(workbookId) + '&index=' + encodeURIComponent(index));
    },
    async detect(workbookId, index) {
      return this.request('POST', '/api/detect', { workbook_id: workbookId, index: index });
    },
    async period(payload) {
      return this.request('POST', '/api/period', payload);
    },
    async regroup(seats, groupLabels) {
      return this.request('POST', '/api/regroup', { seats: seats, group_labels: groupLabels });
    },
    async rotate(payload) {
      return this.request('POST', '/api/rotate', payload);
    },
    async exportWorkbook(payload) {
      return this.request('POST', '/api/export', payload);
    },
    async reveal(path) {
      return this.request('POST', '/api/reveal', { path: path });
    },
    downloadUrl(path) { return '/api/download?path=' + encodeURIComponent(path); }
  };

  /* ============================================================ mock 后端 */
  var MockBackend = {
    label: '演示模式',
    _data() {
      var d = global.MOCK_DATA;
      if (!d) throw new ApiError('演示数据未加载（缺少 mock/mock.js）');
      return d;
    },
    async importWorkbook() {
      await sleep(320);
      var d = this._data();
      return {
        workbook_id: 'mock_wb_001',
        filename: d.filename,
        sheets: [{ index: 0, name: d.sheet.name, hidden: false, rows: d.sheet.rows, cols: d.sheet.cols }],
        warnings: []
      };
    },
    async getSheet(workbookId, index) {
      await sleep(260);
      var d = this._data();
      return { workbook_id: workbookId || 'mock_wb_001', sheet: clone(d.sheet) };
    },
    async detect() {
      await sleep(420);
      return clone(this._data().detect);
    },
    async regroup(seats, groupLabels) {
      await sleep(160);
      return global.SeatEngine.regroup(seats, groupLabels);
    },
    async rotate(payload) {
      await sleep(280);
      var d = this._data();
      var podium = payload.podium || d.detect.podium;
      return global.SeatEngine.rotate(payload.seats, payload.groups, podium, payload.rules);
    },
    async period(payload) {
      await sleep(180);
      var d = this._data();
      var period = (d.detect && d.detect.period) || null;
      var rounds = Math.max(1, Math.min(12, parseInt(payload && payload.rounds, 10) || 1));
      var cycleWeeks = Math.max(1, Math.min(8, parseInt(payload && payload.cycle_weeks, 10) || 2));
      var duration = cycleWeeks * 7;
      var start = (payload && payload.start_date) || (period ? addDaysISO(period.end, 1) : null);
      if (start && !parseISO(start)) throw new ApiError('开始日期格式不正确，请使用 YYYY-MM-DD');
      var stem = String(d.filename || '座位表').replace(/\.xlsx$/i, '');
      var naming = global.SeatEngine.naming;
      var parsed = naming.parseRound(stem);
      var out = [], cur = start;
      for (var i = 1; i <= rounds; i++) {
        var end = cur ? addDaysISO(cur, duration - 1) : null;
        out.push({
          round: i,
          start: cur,
          end: end,
          text: (period && cur) ? fmtPeriodText(cur, end) : null,
          filename: naming.suggestName(stem, i, { start: cur, end: end }, rounds)
        });
        cur = end ? addDaysISO(end, 1) : null;
      }
      return {
        period: period,
        cycle_weeks: cycleWeeks,
        naming: { stem: stem, round: parsed ? parsed.num : null, date_mode: naming.hasDate(stem) },
        rounds: out
      };
    },
    async exportWorkbook(payload) {
      await sleep(900);
      var base = payload.filename_base || '座位表';
      var rounds = payload.rounds || [];
      var files = rounds.map(function (r) {
        var name = (r.filename ? String(r.filename).replace(/\.xlsx$/i, '') : base +
          (rounds.length > 1 ? '_第' + r.round + '轮' : '')) + '.xlsx';
        var path = '导出结果/' + name;
        return {
          name: name, path: path,
          download_url: '/api/download?path=' + encodeURIComponent(path),
          text_updates_applied: (r.header_updates || []).length
        };
      });
      return { files: files, folder: '导出结果', mock: true, move_marks: !!payload.move_marks };
    },
    async reveal() {
      await sleep(120);
      return { ok: true, mock: true };
    },
    downloadUrl(path) { return '#'; }
  };

  /* ============================================================ 门面 */
  var current = RealBackend;
  var isMock = false;

  var SeatAPI = {
    ApiError: ApiError,
    init(opts) {
      isMock = !!(opts && opts.mock);
      current = isMock ? MockBackend : RealBackend;
      return this;
    },
    get mock() { return isMock; },
    get label() { return current.label; },
    importWorkbook(payload) { return current.importWorkbook(payload); },
    getSheet(id, index) { return current.getSheet(id, index); },
    detect(id, index) { return current.detect(id, index); },
    period(payload) { return current.period(payload); },
    regroup(seats, labels) { return current.regroup(seats, labels); },
    rotate(payload) { return current.rotate(payload); },
    exportWorkbook(payload) { return current.exportWorkbook(payload); },
    reveal(path) { return current.reveal(path); },
    downloadUrl(path) { return current.downloadUrl(path); },
    addDaysISO: addDaysISO,
    fmtPeriodText: fmtPeriodText,

    /** File -> base64（不含 data: 前缀） */
    fileToBase64(file) {
      return new Promise(function (resolve, reject) {
        var reader = new FileReader();
        reader.onerror = function () { reject(new ApiError('读取文件失败，请重试或换一个文件')); };
        reader.onload = function () {
          var result = String(reader.result || '');
          var comma = result.indexOf(',');
          resolve(comma >= 0 ? result.slice(comma + 1) : result);
        };
        reader.readAsDataURL(file);
      });
    }
  };

  global.SeatAPI = SeatAPI;
})(window);
