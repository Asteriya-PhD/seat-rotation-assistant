/* 座位轮转助手 · 前端几何 / 轮转引擎（零依赖）
 * 用途：
 *   1) ?mock=1 演示模式下模拟 /api/rotate、/api/regroup（正式模式一律走服务端）；
 *   2) 结构核对时本地先算一版大组，让界面立即响应；
 *   3) 轮转结果校验（人数一一对应 / 警告），与服务端结论互相印证。
 * 语义严格对齐 ARCHITECTURE.md §3、§5。
 */
(function (global) {
  'use strict';

  var MOD = function (n, m) { return ((n % m) + m) % m; };

  function colLetter(c) {
    var s = '';
    while (c > 0) {
      var r = (c - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      c = Math.floor((c - 1) / 26);
    }
    return s;
  }
  function a1(r, c) { return colLetter(c) + r; }
  function key(r, c) { return r + ',' + c; }
  function parseKey(k) {
    var p = String(k).split(',');
    return { r: parseInt(p[0], 10), c: parseInt(p[1], 10) };
  }
  function clone(obj) { return JSON.parse(JSON.stringify(obj)); }
  function rowMajor(a, b) { return a.r - b.r || a.c - b.c; }
  function bbox(list) {
    var r1 = Infinity, r2 = -Infinity, c1 = Infinity, c2 = -Infinity;
    list.forEach(function (s) {
      r1 = Math.min(r1, s.r); r2 = Math.max(r2, s.r);
      c1 = Math.min(c1, s.c); c2 = Math.max(c2, s.c);
    });
    return { r1: r1, r2: r2, c1: c1, c2: c2, rows: r2 - r1 + 1, cols: c2 - c1 + 1 };
  }

  /* ---------------------------------------------------------------- 大组重算
   * labels: [{r, c, text, c_end?}]，c_end 为标签合并列的右边界（本地扩展，服务端忽略）。
   * 规则（§3.2）：有标签 → 标签列决定组；无标签 → 连续座位列成组，空列=走道。
   */
  function regroup(seats, labels) {
    var warnings = [];
    var labByCol = new Map();
    (labels || []).forEach(function (l) {
      if (!l || !l.text) return;
      var c0 = l.c, c1 = l.c_end || l.c;
      for (var c = c0; c <= c1; c++) if (!labByCol.has(c)) labByCol.set(c, l);
    });
    var seatCols = [];
    seats.forEach(function (s) { if (seatCols.indexOf(s.c) < 0) seatCols.push(s.c); });
    seatCols.sort(function (a, b) { return a - b; });

    var groups = [];
    var i = 0;
    while (i < seatCols.length) {
      var c0 = seatCols[i];
      var lab = labByCol.get(c0);
      var cols;
      if (lab) {
        var ce = lab.c_end || lab.c;
        cols = [];
        for (var cc = c0; cc <= ce; cc++) cols.push(cc);
        i += 1;
      } else {
        cols = [c0];
        var j = i + 1;
        while (j < seatCols.length && seatCols[j] === cols[cols.length - 1] + 1 && !labByCol.has(seatCols[j])) {
          cols.push(seatCols[j]); j += 1;
        }
        i = j;
      }
      var colSet = new Set(cols);
      var inCols = seats.filter(function (s) { return colSet.has(s.c); });
      if (!inCols.length) continue;
      var bb = bbox(inCols);
      groups.push({
        name: lab ? lab.text : '第' + (groups.length + 1) + '组',
        r1: bb.r1, r2: bb.r2, c1: bb.c1, c2: bb.c2,
        rows: bb.rows, cols: bb.cols, seat_count: inCols.length
      });
    }
    groups.forEach(function (g, idx) { g.id = idx; });

    if (!(labels || []).length) warnings.push('未找到组标签，已按「连续座位列」分组，可在右侧手动合并或拆分');
    var ungrouped = seats.filter(function (s) {
      return !groups.some(function (g) { return s.r >= g.r1 && s.r <= g.r2 && s.c >= g.c1 && s.c <= g.c2; });
    });
    if (ungrouped.length) {
      warnings.push('有 ' + ungrouped.length + ' 个座位未归入任何大组（' +
        ungrouped.slice(0, 6).map(function (s) { return a1(s.r, s.c); }).join('、') +
        (ungrouped.length > 6 ? ' 等' : '') + '），轮转时不会移动');
    }
    return { groups: groups, warnings: warnings };
  }

  /* ---------------------------------------------------------------- 轮转
   * rotate(seats, groups, podium, rules) -> {rounds:[{round,assignments,seats_after,unplaced,empty_seats,warnings}], warnings}
   */
  function rotate(seats, groups, podium, rules) {
    rules = rules || {};
    groups = groups || [];
    if (!seats || !seats.length) throw new Error('座位表为空，无法轮转');

    var relation = (podium && podium.relation) || 'above';
    var rm = rules.row_move || { steps: 0, toward: 'back' };
    var cm = rules.col_move || { steps: 0, toward: 'right' };
    var gm = rules.group_move || { steps: 0, toward: 'right' };
    var mode = rules.mode || 'inner';
    var roundCount = Math.max(1, Math.min(12, parseInt(rules.rounds, 10) || 1));
    var fixedSeatSet = new Set((rules.fixed_seats || []).map(function (p) {
      if (Array.isArray(p)) return key(p[0], p[1]);
      if (p && typeof p === 'object') return key(p.r, p.c);
      return normFixed(p);
    }));
    function normFixed(p) {
      var m = String(p).replace(/\s+/g, '').match(/^\$?([A-Z]+)\$?(\d+)$/i);
      if (!m) return String(p);
      var c = 0, letters = m[1].toUpperCase();
      for (var i = 0; i < letters.length; i++) c = c * 26 + (letters.charCodeAt(i) - 64);
      return key(parseInt(m[2], 10), c);
    }
    var fixedGroupSet = new Set((rules.fixed_groups || []).map(Number));

    var backSign = relation === 'below' ? -1 : 1;
    var dr = rm.steps ? parseInt(rm.steps, 10) * (rm.toward === 'back' ? backSign : -backSign) : 0;
    var dc = cm.steps ? parseInt(cm.steps, 10) * (cm.toward === 'right' ? 1 : -1) : 0;
    var gshift = gm.steps ? parseInt(gm.steps, 10) * (gm.toward === 'right' ? 1 : -1) : 0;

    var seatKeys = seats.map(function (s) { return key(s.r, s.c); });
    var seatByKey = new Map();
    seats.forEach(function (s) { seatByKey.set(key(s.r, s.c), s); });
    var warnings = [];

    /* 组内循环位移：固定座位视为空洞，跳过 */
    function cycleMap(list, bb, ddr, ddc) {
      var map = new Map();
      if (!ddr && !ddc) {
        list.forEach(function (s) { map.set(key(s.r, s.c), key(s.r, s.c)); });
        return map;
      }
      var pos = new Set(list.map(function (s) { return key(s.r, s.c); }));
      var H = bb.rows, W = bb.cols, maxIter = H * W + 2;
      list.forEach(function (t) {
        var r = t.r, c = t.c, found = null;
        for (var n = 0; n < maxIter; n++) {
          r = bb.r1 + MOD(r - bb.r1 - ddr, H);
          c = bb.c1 + MOD(c - bb.c1 - ddc, W);
          if (pos.has(key(r, c))) { found = { r: r, c: c }; break; }
        }
        map.set(key(t.r, t.c), found ? key(found.r, found.c) : key(t.r, t.c));
      });
      return map;
    }

    var perm = new Map(); // targetKey -> sourceKey（非固定座位）
    function groupSeats(g) {
      var list = seats.filter(function (s) { return s.group_id === g.id; });
      if (!list.length) {
        list = seats.filter(function (s) { return s.r >= g.r1 && s.r <= g.r2 && s.c >= g.c1 && s.c <= g.c2; });
      }
      return list;
    }
    function movable(list) { return list.filter(function (s) { return !fixedSeatSet.has(key(s.r, s.c)); }); }

    if (mode === 'class') {
      var all = movable(seats);
      perm = cycleMap(all, bbox(seats), dr, dc);
    } else if (mode === 'swap' || mode === 'combo') {
      var rot = groups.filter(function (g) { return !fixedGroupSet.has(g.id); })
        .slice().sort(function (a, b) { return a.c1 - b.c1 || a.r1 - b.r1; });
      if (!rot.length) {
        warnings.push('没有可轮转的大组（全部被固定），学生不会移动');
      }
      var byGroup = new Map();
      rot.forEach(function (g) { byGroup.set(g.id, movable(groupSeats(g))); });
      var n = rot.length;
      var innerMaps = new Map();
      if (mode === 'combo') {
        rot.forEach(function (g) {
          innerMaps.set(g.id, cycleMap(byGroup.get(g.id), bbox(groupSeats(g)), dr, dc));
        });
      }
      var shapeWarned = false;
      rot.forEach(function (gT, i) {
        if (!n) return;
        var gS = rot[MOD(i - gshift, n)];
        if (mode === 'combo' && gT.id === gS.id) {
          innerMaps.get(gT.id).forEach(function (src, tgt) { perm.set(tgt, src); });
          return;
        }
        var T = byGroup.get(gT.id), S = byGroup.get(gS.id);
        if (!T || !T.length) return;
        var srcInner = mode === 'combo' ? innerMaps.get(gS.id) : null;
        if (gT.rows === gS.rows && gT.cols === gS.cols) {
          T.forEach(function (t) {
            var p = key(gS.r1 + (t.r - gT.r1), gS.c1 + (t.c - gT.c1));
            if (!seatByKey.has(p)) return;
            if (fixedSeatSet.has(p)) return;
            if (mode === 'combo') {
              var s2 = srcInner.get(p);
              if (s2) perm.set(key(t.r, t.c), s2);
            } else {
              perm.set(key(t.r, t.c), p);
            }
          });
        } else {
          if (!shapeWarned) {
            warnings.push('部分大组形状不同，整组互换按「行优先序号」映射，可能出现空位或未落座');
            shapeWarned = true;
          }
          var Ts = T.slice().sort(rowMajor), Ss = S.slice().sort(rowMajor);
          Ts.forEach(function (t, idx) {
            if (idx >= Ss.length) return;
            var p = key(Ss[idx].r, Ss[idx].c);
            if (mode === 'combo') {
              var s3 = srcInner.get(p);
              if (s3) perm.set(key(t.r, t.c), s3);
            } else {
              perm.set(key(t.r, t.c), p);
            }
          });
        }
      });
    } else { // inner
      var innerGroups = groups.filter(function (g) { return !fixedGroupSet.has(g.id); });
      innerGroups.forEach(function (g) {
        var list = movable(groupSeats(g));
        if (!list.length) return;
        cycleMap(list, bbox(groupSeats(g)), dr, dc).forEach(function (src, tgt) { perm.set(tgt, src); });
      });
      var oneCol = groups.length > 0 && groups.every(function (g) { return g.cols === 1; });
      if (oneCol && dc !== 0) {
        warnings.push('每个大组只有 1 列，左右方向的组内轮转不会产生位移（可改用整组轮转）');
      }
      if (!groups.length) warnings.push('尚未识别到大组，无法做组内轮转');
    }

    if (!dr && !dc && !gshift) warnings.push('未设置任何移动：请至少设置前后 / 左右 / 整组移动中的一项');

    /* 固定座位：原地不动；没有映射的非固定座位视为空位 */
    seats.forEach(function (s) {
      var k = key(s.r, s.c);
      if (fixedSeatSet.has(k)) perm.set(k, k);
      else if (!perm.has(k)) perm.set(k, null);
    });

    /* 形差不一致时补位：把未能随组就位的学生按最近距离安排到空位。
     * 服务端对 combo 形状差异会这样兜底，mock 保持一致，避免演示时误报“丢人”。 */
    var placedSources = new Set();
    perm.forEach(function (src) { if (src) placedSources.add(src); });
    var emptyTargets = seats.filter(function (s) {
      var k = key(s.r, s.c);
      return !fixedSeatSet.has(k) && !perm.get(k);
    });
    var unplacedSources = seats.filter(function (s) {
      var k = key(s.r, s.c);
      return !fixedSeatSet.has(k) && !placedSources.has(k);
    });
    if (emptyTargets.length && unplacedSources.length) {
      var paired = 0;
      unplacedSources.slice().sort(rowMajor).forEach(function (src) {
        if (!emptyTargets.length) return;
        var bestIdx = -1, bestD = Infinity;
        emptyTargets.forEach(function (t, i) {
          var d = (t.r - src.r) * (t.r - src.r) + (t.c - src.c) * (t.c - src.c);
          if (d < bestD) { bestD = d; bestIdx = i; }
        });
        if (bestIdx >= 0) {
          var target = emptyTargets.splice(bestIdx, 1)[0];
          perm.set(key(target.r, target.c), key(src.r, src.c));
          paired++;
        }
      });
      if (paired) {
        warnings.push('大组形状不一致，' + paired + ' 名同学按最近距离自动补位；请在预览里核对这几个人的位置');
      }
    }

    /* 逐轮模拟，记录原始坐标 */
    var cur = new Map(); // seatKey -> {text, origin:{r,c}}
    seats.forEach(function (s) { cur.set(key(s.r, s.c), { text: s.text, origin: { r: s.r, c: s.c } }); });

    var rounds = [];
    var unplacedSeen = [];
    for (var round = 1; round <= roundCount; round++) {
      var next = new Map();
      var consumed = new Set();
      var assignments = [];
      var emptySeats = [];
      var roundWarnings = [];

      seats.slice().sort(rowMajor).forEach(function (s) {
        var t = key(s.r, s.c);
        var src = perm.get(t);
        if (src) {
          var v = cur.get(src);
          if (v) {
            next.set(t, v);
            consumed.add(src);
            assignments.push({ to: [s.r, s.c], from: [v.origin.r, v.origin.c], text: v.text });
          } else {
            emptySeats.push([s.r, s.c]);
          }
        } else {
          emptySeats.push([s.r, s.c]);
        }
      });

      var unplaced = [];
      cur.forEach(function (v, k) {
        if (!consumed.has(k) && seatByKey.has(k)) {
          unplaced.push({ text: v.text, from: [v.origin.r, v.origin.c] });
        }
      });
      if (unplaced.length && round === 1) {
        unplacedSeen = unplaced;
        roundWarnings.push(unplaced.length + ' 人无法落座（大组形状不同），已放入「未落座」列表');
      }
      if (emptySeats.length && round === 1) {
        roundWarnings.push('有 ' + emptySeats.length + ' 个座位空出');
      }

      var seatsAfter = [];
      seats.slice().sort(rowMajor).forEach(function (s) {
        var v = next.get(key(s.r, s.c));
        if (v) seatsAfter.push({ r: s.r, c: s.c, text: v.text });
      });

      rounds.push({
        round: round,
        assignments: assignments,
        seats_after: seatsAfter,
        unplaced: unplacedSeen.length ? clone(unplacedSeen) : [],
        empty_seats: emptySeats,
        warnings: roundWarnings
      });
      cur = next;
    }

    /* 一一对应自检 */
    var check = validate({ rounds: rounds }, seats, rules.fixed_seats);
    if (!check.ok) warnings.push(check.message);

    return { rounds: rounds, warnings: warnings };
  }

  /* ---------------------------------------------------------------- 校验 */
  function validate(result, seats, fixedSeats) {
    if (!result || !result.rounds || !result.rounds.length) {
      return { ok: false, message: '尚未计算轮转结果', warnings: [] };
    }
    var seatKeys = seats.map(function (s) { return key(s.r, s.c); });
    var seatSet = new Set(seatKeys);
    var fixedSet = new Set((fixedSeats || []).map(function (p) {
      return Array.isArray(p) ? key(p[0], p[1]) : (p && typeof p === 'object' ? key(p.r, p.c) : String(p));
    }));
    var issues = [];
    var totalUnplaced = 0, totalEmpty = 0;

    result.rounds.forEach(function (rd) {
      var all = rd.assignments || [];
      var seen = new Set();
      all.forEach(function (a) {
        var k = key(a.to[0], a.to[1]);
        if (seen.has(k)) issues.push('第 ' + rd.round + ' 轮：座位 ' + a1(a.to[0], a.to[1]) + ' 被重复占用');
        seen.add(k);
        if (!seatSet.has(k)) issues.push('第 ' + rd.round + ' 轮：目标 ' + a1(a.to[0], a.to[1]) + ' 不是已知座位');
      });
      seatKeys.forEach(function (k) {
        if (!seen.has(k)) {
          var p = parseKey(k);
          issues.push('第 ' + rd.round + ' 轮：座位 ' + a1(p.r, p.c) + ' 没有人落座');
        }
      });
      totalUnplaced += (rd.unplaced || []).length;
      totalEmpty += (rd.empty_seats || []).length;
    });

    if (totalUnplaced) issues.push('共有 ' + totalUnplaced + ' 人次无法落座');
    if (issues.length) {
      return {
        ok: false,
        issues: issues.slice(0, 12),
        message: '⚠ 有 ' + issues.length + ' 处异常：' + issues[0] + (issues.length > 1 ? '…' : ''),
        warnings: issues
      };
    }
    return {
      ok: true,
      issues: [],
      message: '✔ ' + seats.length + ' 人一一对应' + (fixedSet.size ? '（含 ' + fixedSet.size + ' 个固定座位）' : ''),
      warnings: []
    };
  }

  /* ---------------------------------------------------------------- 文件名智能递进
   * 与服务端 naming 规则对齐：
   *   - 原名带「第N轮」→ 替换为新的轮次（保留中文/阿拉伯数字写法）；
   *   - 原名带日期 → 替换成新轮次的开始/结束日期（保留分隔符与补零风格）；
   *   - 两者都有 → 都替换；都没有 → 单轮保持原名，多轮由调用方兜底 _第N轮。
   */
  var CN_DIGITS = { '零': 0, '一': 1, '二': 2, '两': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9 };
  var CN_ARR = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九'];
  var ROUND_RE = /(第\s*)(\d+|[零一二两三四五六七八九十]+)(\s*轮)/;
  var DATE_RE = /(\d{4})([.\-/年])(\d{1,2})([.\-/月])(\d{1,2})(日?)/g;

  function cnToInt(text) {
    if (/^\d+$/.test(text)) return parseInt(text, 10);
    if (text === '十') return 10;
    var m = /^([一二两三四五六七八九])?十([一二两三四五六七八九])?$/.exec(text);
    if (m) return (m[1] ? CN_DIGITS[m[1]] : 1) * 10 + (m[2] ? CN_DIGITS[m[2]] : 0);
    if (text.length === 1 && CN_DIGITS[text] != null) return CN_DIGITS[text];
    return NaN;
  }

  function intToCn(n) {
    if (n < 0) return String(n);
    if (n < 10) return CN_ARR[n];
    if (n === 10) return '十';
    if (n < 20) return '十' + CN_ARR[n % 10];
    var tens = Math.floor(n / 10), ones = n % 10;
    return CN_ARR[tens] + '十' + (ones ? CN_ARR[ones] : '');
  }

  function parseRound(name) {
    var m = ROUND_RE.exec(String(name || ''));
    if (!m) return null;
    var chinese = !/^\d+$/.test(m[2]);
    var num = chinese ? cnToInt(m[2]) : parseInt(m[2], 10);
    if (!isFinite(num)) return null;
    return { num: num, chinese: chinese, prefix: m[1], suffix: m[3] };
  }

  function withRound(name, num) {
    var r = parseRound(name);
    if (!r || num < 0) return name;
    var rep = r.chinese ? intToCn(num) : String(num);
    return String(name).replace(ROUND_RE, function () { return r.prefix + rep + r.suffix; });
  }

  function parseISODate(iso) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso || ''));
    if (!m) return null;
    return { y: +m[1], m: +m[2], d: +m[3] };
  }

  function pad2(n) { return n < 10 ? '0' + n : '' + n; }

  function formatDateLike(iso, style) {
    var d = parseISODate(iso);
    if (!d) return '';
    var month = style.padMonth ? pad2(d.m) : String(d.m);
    var day = style.padDay ? pad2(d.d) : String(d.d);
    return d.y + style.ysep + month + style.msep + day + style.daySuffix;
  }

  function hasDate(name) {
    DATE_RE.lastIndex = 0;
    return DATE_RE.test(String(name || ''));
  }

  function replaceDates(name, startISO, endISO) {
    var text = String(name || '');
    var matches = [];
    DATE_RE.lastIndex = 0;
    var m;
    while ((m = DATE_RE.exec(text)) !== null) {
      matches.push({
        raw: m[0], index: m.index, end: m.index + m[0].length,
        style: {
          ysep: m[2], msep: m[4], daySuffix: m[6] || '',
          padMonth: /^0\d$/.test(m[3]), padDay: /^0\d$/.test(m[5])
        }
      });
    }
    if (!matches.length || !startISO) return text;

    // 情况 1：两个完整日期（都带年份）
    if (matches.length >= 2) {
      var out = '', cursor = 0;
      matches.forEach(function (item, i) {
        var target = i === 0 ? startISO : (i === 1 && endISO ? endISO : null);
        out += text.slice(cursor, item.index);
        out += target ? formatDateLike(target, item.style) : item.raw;
        cursor = item.end;
      });
      return out + text.slice(cursor);
    }

    // 情况 2：只有开始日期带年份，结束日期写成「11.13」这种短格式
    var first = matches[0];
    var before = text.slice(0, first.index) + formatDateLike(startISO, first.style);
    var rest = text.slice(first.end);
    if (endISO) {
      var shortRe = /([～~—\-至])\s*(\d{1,2})([.\-/月])(\d{1,2})(日?)/;
      var sm = shortRe.exec(rest);
      if (sm) {
        var style = {
          ysep: first.style.ysep, msep: sm[3], daySuffix: sm[5] || '',
          padMonth: /^0\d$/.test(sm[2]), padDay: /^0\d$/.test(sm[4])
        };
        var endDate = parseISODate(endISO);
        var endText = endDate ? ((style.padMonth ? pad2(endDate.m) : String(endDate.m)) + style.msep +
          (style.padDay ? pad2(endDate.d) : String(endDate.d)) + style.daySuffix) : sm[4];
        rest = rest.slice(0, sm.index) + sm[1] + (sm[0].slice(sm[1].length).match(/^\s*/) || [''])[0] + endText +
          rest.slice(sm.index + sm[0].length);
      }
    }
    return before + rest;
  }

  /** 在已有名字上推进 delta 轮（用于用户手改过文件名后的多轮命名）。 */
  function advanceName(name, delta, roundInfo) {
    var out = String(name || '');
    var r = parseRound(out);
    if (r) out = withRound(out, r.num + delta);
    if (hasDate(out) && roundInfo && roundInfo.start) {
      out = replaceDates(out, roundInfo.start, roundInfo.end);
    }
    return out;
  }

  /** 根据原文件名 + 第几轮，生成建议文件名（不含 .xlsx）。 */
  function suggestName(originalStem, roundNo, roundInfo, totalRounds) {
    var stem = String(originalStem || '座位表');
    var out = stem;
    var r = parseRound(stem);
    if (r) out = withRound(out, r.num + roundNo);
    if (roundInfo && roundInfo.start) out = replaceDates(out, roundInfo.start, roundInfo.end);
    if (out === stem && (totalRounds || 1) > 1) out = stem + '_第' + roundNo + '轮';
    return out;
  }

  var naming = {
    parseRound: parseRound,
    withRound: withRound,
    hasDate: hasDate,
    replaceDates: replaceDates,
    advanceName: advanceName,
    suggestName: suggestName,
    cnToInt: cnToInt,
    intToCn: intToCn
  };

  var SeatEngine = {
    MOD: MOD, colLetter: colLetter, a1: a1, key: key, parseKey: parseKey, clone: clone,
    bbox: bbox, regroup: regroup, rotate: rotate, validate: validate, naming: naming
  };
  global.SeatEngine = SeatEngine;
})(window);
