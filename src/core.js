/* Lõi tính toán tiến độ CPM — không phụ thuộc giao diện */
const DAY = 86400000;
const pad = n => String(n).padStart(2, '0');
const mkDay = (y, m, d) => {
  const t = Date.UTC(y, m - 1, d); const dt = new Date(t);
  if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return null;
  return Math.round(t / DAY);
};
const toDay = iso => iso ? mkDay(+iso.slice(0, 4), +iso.slice(5, 7), +iso.slice(8, 10)) : null;
const isoOf = d => new Date(d * DAY).toISOString().slice(0, 10);
const dow = d => new Date(d * DAY).getUTCDay();
const fmtD = d => { const s = isoOf(d); return s.slice(8, 10) + '/' + s.slice(5, 7) + '/' + s.slice(0, 4); };
const fmtDM = d => { const s = isoOf(d); return s.slice(8, 10) + '/' + s.slice(5, 7); };
function parseDate(s, refYear) {
  s = String(s || '').trim(); if (!s) return null;
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/); if (m) return mkDay(+m[1], +m[2], +m[3]);
  m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})(?:[\/.\-](\d{2,4}))?$/);
  if (m) { let y = m[3] ? +m[3] : (refYear || new Date().getFullYear()); if (y < 100) y += 2000; return mkDay(y, +m[2], +m[1]); }
  return null;
}
const mondayOf = d => d - (dow(d) + 6) % 7;
function isoWeek(d) {
  const mon = mondayOf(d); const y = new Date((mon + 3) * DAY).getUTCFullYear();
  return 1 + Math.round((mon - mondayOf(mkDay(y, 1, 4))) / 7);
}

function expandHolidays(list) {
  const out = [];
  (list || []).forEach(h => { const a = toDay(h.s), b = toDay(h.f || h.s); if (a == null || b == null) return; for (let d = a; d <= Math.min(b, a + 60); d++) out.push(d); });
  return out;
}
function buildCal(P) {
  const hol = new Set(expandHolidays(P.holidays)); const ww = +P.workweek || 6;
  const isWork = d => { const w = dow(d); if (ww <= 6 && w === 0) return false; if (ww === 5 && w === 6) return false; return !hol.has(d); };
  const s0 = toDay(P.start); const days = [];
  for (let d = s0; days.length < 4000 && d - s0 < 9000; d++) if (isWork(d)) days.push(d);
  const idxOf = d => { // index of first workday >= d
    if (d <= days[0]) return 0; let lo = 0, hi = days.length;
    while (lo < hi) { const mid = (lo + hi) >> 1; if (days[mid] < d) lo = mid + 1; else hi = mid; }
    return lo;
  };
  const day = k => days[Math.max(0, Math.min(days.length - 1, k))];
  return { isWork, days, idxOf, day, hol };
}

function structure(T) {
  const n = T.length, parent = new Array(n).fill(-1), isSum = new Array(n).fill(false), end = new Array(n), wbs = new Array(n);
  const st = [], cnt = [];
  for (let i = 0; i < n; i++) {
    const lv = Math.max(0, Math.min(+T[i].level || 0, i ? T[i - 1].level + 1 : 0)); T[i].level = lv;
    while (st.length && T[st[st.length - 1]].level >= lv) st.pop();
    parent[i] = st.length ? st[st.length - 1] : -1; st.push(i);
    cnt.length = lv + 1; cnt[lv] = (cnt[lv] || 0) + 1; wbs[i] = cnt.slice(0, lv + 1).join('.');
  }
  for (let i = 0; i < n; i++) {
    isSum[i] = i + 1 < n && T[i + 1].level > T[i].level;
    let j = i; while (j + 1 < n && T[j + 1].level > T[i].level) j++; end[i] = j;
  }
  return { parent, isSum, end, wbs };
}

function parseRes(s) {
  const out = [];
  String(s || '').split(/[,;]+/).forEach(p => {
    p = p.trim(); if (!p) return;
    const m = p.match(/^([^:=\s]+)\s*[:=x×]?\s*(\d+(?:[.,]\d+)?)?$/i); if (!m) return;
    out.push({ code: m[1].toUpperCase(), qty: m[2] ? parseFloat(m[2].replace(',', '.')) : 1 });
  });
  return out;
}

function schedule(P, useDD) {
  const T = P.tasks, n = T.length, st = structure(T), cal = buildCal(P);
  const idIdx = new Map(T.map((t, i) => [t.id, i]));
  const leavesOf = i => { const r = []; for (let j = i; j <= st.end[i]; j++) if (!st.isSum[j]) r.push(j); return r; };
  const isAnc = (a, b) => { let p = st.parent[b]; while (p >= 0) { if (p === a) return true; p = st.parent[p]; } return false; };
  const badLink = new Array(n).fill(false), edges = [];
  T.forEach((t, q) => (t.preds || []).forEach(l => {
    const p = idIdx.get(l.id);
    if (p == null || p === q || isAnc(p, q) || isAnc(q, p)) { badLink[q] = true; return; }
    const F = st.isSum[p] ? leavesOf(p) : [p], G = st.isSum[q] ? leavesOf(q) : [q];
    F.forEach(f => G.forEach(g => { if (f !== g) edges.push({ f, g, type: l.type || 'FS', lag: +l.lag || 0 }); }));
  }));
  const inE = Array.from({ length: n }, () => []), outE = Array.from({ length: n }, () => []);
  edges.forEach(e => { inE[e.g].push(e); outE[e.f].push(e); });
  const leaves = []; for (let i = 0; i < n; i++) if (!st.isSum[i]) leaves.push(i);
  const indeg = new Array(n).fill(0); edges.forEach(e => indeg[e.g]++);
  const inQ = new Array(n).fill(false), cycle = new Array(n).fill(false), order = [];
  const queue = leaves.filter(i => indeg[i] === 0); queue.forEach(i => { inQ[i] = true; });
  while (order.length < leaves.length) {
    while (queue.length) { const i = queue.shift(); order.push(i); outE[i].forEach(e => { if (--indeg[e.g] === 0 && !inQ[e.g]) { inQ[e.g] = true; queue.push(e.g); } }); }
    if (order.length < leaves.length) { const r = leaves.find(i => !inQ[i]); cycle[r] = true; inQ[r] = true; queue.push(r); } // break a loop
  }

  const dataDay = toDay(P.dataDate) ?? toDay(P.start); const D = cal.idxOf(dataDay);
  const es = new Array(n).fill(0), ef = new Array(n).fill(0), du = new Array(n).fill(0), done = new Array(n).fill(false);
  const snetOf = i => { let v = null, k = i; while (k >= 0) { const s = toDay(T[k].snet); if (s != null) v = v == null ? s : Math.max(v, s); k = st.parent[k]; } return v; };
  for (const i of order) {
    const t = T[i], d = Math.max(0, Math.round(+t.dur || 0));
    let s = 0; const sn = snetOf(i); if (sn != null) s = Math.max(s, cal.idxOf(sn));
    let efMin = -Infinity;
    for (const e of inE[i]) {
      if (!done[e.f]) continue;
      if (e.type === 'FS') s = Math.max(s, ef[e.f] + e.lag);
      else if (e.type === 'SS') s = Math.max(s, es[e.f] + e.lag);
      else if (e.type === 'FF') efMin = Math.max(efMin, ef[e.f] + e.lag);
      else if (e.type === 'SF') efMin = Math.max(efMin, es[e.f] + e.lag);
    }
    if (efMin > -Infinity) s = Math.max(s, efMin - d);
    s = Math.max(0, s); let f = s + d;
    if (useDD) {
      const pc = +t.pct || 0;
      if (pc <= 0 && s < D) { s = D; f = s + d; }
      else if (pc > 0 && pc < 100) { const rem = Math.ceil(d * (1 - pc / 100)); if (f < D + rem) f = D + rem; }
    }
    es[i] = s; ef[i] = f; du[i] = f - s; done[i] = true;
  }
  let PE = 0; leaves.forEach(i => { PE = Math.max(PE, ef[i]); });
  const lf = new Array(n).fill(PE);
  for (let k = order.length - 1; k >= 0; k--) {
    const i = order[k]; let L = PE;
    for (const e of outE[i]) {
      const g = e.g; if (cycle[g] && cycle[i]) continue; const lsg = lf[g] - du[g]; let v;
      if (e.type === 'FS') v = lsg - e.lag; else if (e.type === 'SS') v = lsg - e.lag + du[i];
      else if (e.type === 'FF') v = lf[g] - e.lag; else v = lf[g] - e.lag + du[i];
      L = Math.min(L, v);
    }
    lf[i] = L;
  }
  const start = new Array(n), finish = new Array(n), tf = new Array(n).fill(null), crit = new Array(n).fill(false);
  const pct = new Array(n).fill(0), planned = new Array(n).fill(0), late = new Array(n).fill(false), vr = new Array(n).fill(null);
  const blS = new Array(n).fill(null), blF = new Array(n).fill(null);
  leaves.forEach(i => {
    start[i] = cal.day(es[i]); finish[i] = du[i] > 0 ? cal.day(ef[i] - 1) : cal.day(es[i]);
    tf[i] = lf[i] - du[i] - es[i]; pct[i] = Math.max(0, Math.min(100, +T[i].pct || 0));
    crit[i] = tf[i] <= 0 && pct[i] < 100;
    const bl = T[i].bl; let bES = es[i], bEF = ef[i];
    if (bl && toDay(bl.s) != null) {
      blS[i] = toDay(bl.s); blF[i] = toDay(bl.f);
      bES = cal.idxOf(blS[i]); const bd = Math.max(0, Math.round(+bl.d || 0));
      bEF = bd > 0 ? cal.idxOf(blF[i]) + 1 : bES;
      vr[i] = cal.idxOf(finish[i]) - cal.idxOf(blF[i]);
    }
    planned[i] = bEF > bES ? Math.max(0, Math.min(1, (D - bES) / (bEF - bES))) * 100 : (D > bES ? 100 : 0);
    late[i] = pct[i] < 100 && planned[i] - pct[i] >= 1;
  });
  for (let i = n - 1; i >= 0; i--) if (st.isSum[i]) {
    const L = leavesOf(i); let a = Infinity, b = -Infinity, w = 0, wp = 0, wpl = 0, mt = Infinity, ba = Infinity, bb = -Infinity, lt = false;
    L.forEach(j => { a = Math.min(a, start[j]); b = Math.max(b, finish[j]); const wt = Math.max(+T[j].dur || 0, 0); w += wt; wp += wt * pct[j]; wpl += wt * planned[j]; mt = Math.min(mt, tf[j]); if (blS[j] != null) { ba = Math.min(ba, blS[j]); bb = Math.max(bb, blF[j]); } lt = lt || late[j]; });
    start[i] = a; finish[i] = b; tf[i] = mt; late[i] = lt;
    pct[i] = w ? wp / w : (L.length ? L.reduce((s, j) => s + pct[j], 0) / L.length : 0);
    planned[i] = w ? wpl / w : 0;
    es[i] = Math.min(...L.map(j => es[j])); ef[i] = Math.max(...L.map(j => ef[j]));
    if (ba < Infinity) { blS[i] = ba; blF[i] = bb; vr[i] = cal.idxOf(b) - cal.idxOf(bb); }
  }
  const projStart = leaves.length ? Math.min(...leaves.map(i => start[i])) : cal.day(0);
  const projFinish = leaves.length ? Math.max(...leaves.map(i => finish[i])) : cal.day(0);
  let W = 0, WP = 0, WPL = 0, bMax = null;
  leaves.forEach(i => { const wt = Math.max(+T[i].dur || 0, 0); W += wt; WP += wt * pct[i]; WPL += wt * planned[i]; if (blF[i] != null) bMax = bMax == null ? blF[i] : Math.max(bMax, blF[i]); });
  return { st, cal, idIdx, leaves, es, ef, du, start, finish, tf, crit, pct, planned, late, vr, blS, blF, cycle, badLink,
    projStart, projFinish, blFinish: bMax, dataDay, D, actualPct: W ? WP / W : 0, plannedPct: W ? WPL / W : 0, workDays: PE };
}

function fmtPreds(t, idIdx) {
  return (t.preds || []).map(l => {
    const k = idIdx.get(l.id); if (k == null) return null;
    let s = String(k + 1); const ty = l.type || 'FS', lag = +l.lag || 0;
    if (ty !== 'FS' || lag) s += ty; if (lag) s += (lag > 0 ? '+' : '') + lag; return s;
  }).filter(Boolean).join(', ');
}
function parsePreds(str, T, selfIdx) {
  const links = [], errs = [];
  String(str || '').split(/[,;]+/).forEach(tok => {
    const raw = tok.trim(); if (!raw) return;
    const m = raw.replace(/\s+/g, '').toUpperCase().match(/^(\d+)(FS|SS|FF|SF)?([+-]\d+)?(D|NGÀY|N)?$/);
    if (!m) { errs.push(raw); return; }
    const k = +m[1] - 1; if (k < 0 || k >= T.length || k === selfIdx) { errs.push(raw); return; }
    links.push({ id: T[k].id, type: m[2] || 'FS', lag: m[3] ? parseInt(m[3], 10) : 0 });
  });
  return { links, errs };
}
let _idn = 0;
const newId = () => 't' + Date.now().toString(36).slice(-4) + (++_idn).toString(36) + Math.random().toString(36).slice(2, 5);

function buildTasks(rows) { // rows: [level,name,dur,preds,res,pct]
  const T = rows.map(r => ({ id: newId(), level: r[0], name: r[1], dur: r[2], preds: [], res: r[4] || '', pct: r[5] || 0, snet: null, bl: null }));
  rows.forEach((r, i) => { if (r[3]) T[i].preds = parsePreds(r[3], T, i).links; });
  return T;
}
const VN_HOLIDAYS = [
  { s: '2026-01-01', f: '2026-01-01', note: 'Tết Dương lịch' },
  { s: '2026-02-14', f: '2026-02-22', note: 'Tết Nguyên đán Bính Ngọ' },
  { s: '2026-04-27', f: '2026-04-27', note: 'Giỗ Tổ Hùng Vương (nghỉ bù)' },
  { s: '2026-04-30', f: '2026-05-01', note: '30/4 – 1/5' },
  { s: '2026-09-01', f: '2026-09-02', note: 'Quốc khánh' },
  { s: '2027-01-01', f: '2027-01-01', note: 'Tết Dương lịch' },
  { s: '2027-02-04', f: '2027-02-12', note: 'Tết Nguyên đán Đinh Mùi (nghỉ công trường)' },
  { s: '2027-04-16', f: '2027-04-16', note: 'Giỗ Tổ Hùng Vương' },
  { s: '2027-04-30', f: '2027-05-01', note: '30/4 – 1/5' },
  { s: '2027-09-02', f: '2027-09-03', note: 'Quốc khánh' }
];
const DEFAULT_RES = { NC: { name: 'Nhân công xây dựng', unit: 'người', limit: 24 }, MAY: { name: 'Máy thi công', unit: 'máy', limit: 2 } };

function sampleProject() {
  const rows = [
    [0, 'NHÀ PHỐ 1 TRỆT 3 LẦU – 5 × 20 m', 0, '', '', 0],
    [1, 'CHUẨN BỊ MẶT BẰNG', 0, '', '', 0],
    [2, 'Nhận mặt bằng, định vị tim cốt', 2, '', 'NC:2', 100],
    [2, 'Lán trại, rào chắn, điện nước thi công', 4, '3', 'NC:4', 100],
    [1, 'PHẦN NGẦM (MÓNG)', 0, '', '', 0],
    [2, 'Ép cọc BTCT 250×250', 6, '4', 'NC:4, MAY:1', 100],
    [2, 'Đào đất hố móng', 4, '6', 'NC:3, MAY:1', 100],
    [2, 'Đập đầu cọc, bê tông lót', 2, '7', 'NC:5', 100],
    [2, 'Cốt thép, cốp pha móng – đà kiềng', 6, '8', 'NC:8', 100],
    [2, 'Đổ bê tông móng, đà kiềng', 1, '9', 'NC:10, MAY:1', 100],
    [2, 'Xây hầm phân, bể nước ngầm', 5, '10', 'NC:4', 100],
    [2, 'Lấp đất, bê tông nền trệt', 3, '11', 'NC:6', 100],
    [1, 'KẾT CẤU PHẦN THÂN', 0, '', '', 0],
    [2, 'Cột, vách tầng trệt', 4, '10FS+2', 'NC:8', 100],
    [2, 'Dầm sàn lầu 1', 7, '14', 'NC:10', 100],
    [2, 'Cột, vách lầu 1', 4, '15FS+2', 'NC:8', 100],
    [2, 'Dầm sàn lầu 2', 7, '16', 'NC:10', 100],
    [2, 'Cột, vách lầu 2', 4, '17FS+2', 'NC:8', 100],
    [2, 'Dầm sàn lầu 3', 7, '18', 'NC:10', 40],
    [2, 'Cột, vách lầu 3', 4, '19FS+2', 'NC:8', 0],
    [2, 'Dầm sàn mái, sân thượng', 7, '20', 'NC:10', 0],
    [2, 'Cất nóc', 0, '21', '', 0],
    [1, 'XÂY TÔ & BAO CHE', 0, '', '', 0],
    [2, 'Xây tường trệt – lầu 1', 8, '17', 'NC:6', 70],
    [2, 'Xây tường lầu 2 – lầu 3', 8, '24, 22', 'NC:6', 0],
    [2, 'Tô trát trong nhà', 15, '24SS+5', 'NC:8', 10],
    [2, 'Tô trát ngoài, chống thấm', 12, '25', 'NC:6, MAY:1', 0],
    [1, 'HỆ THỐNG CƠ ĐIỆN (MEP)', 0, '', '', 0],
    [2, 'Đi ống âm sàn, âm tường', 10, '24SS+2', 'NC:3', 40],
    [2, 'Kéo dây, lắp tủ điện', 8, '29, 26SS+5', 'NC:3', 0],
    [2, 'Cấp thoát nước, thiết bị vệ sinh', 6, '30SS+2', 'NC:3', 0],
    [1, 'HOÀN THIỆN & BÀN GIAO', 0, '', '', 0],
    [2, 'Trần thạch cao', 8, '30SS+3, 26', 'NC:4', 0],
    [2, 'Ốp lát gạch nền, vệ sinh', 14, '26', 'NC:6', 0],
    [2, 'Sơn nước trong – ngoài', 12, '33, 27', 'NC:6', 0],
    [2, 'Lắp cửa, lan can, cầu thang', 8, '34SS+6', 'NC:4', 0],
    [2, 'Lắp thiết bị điện, đèn, vệ sinh', 5, '35SS+4, 31', 'NC:3', 0],
    [2, 'Vệ sinh công nghiệp', 3, '36, 37', 'NC:6', 0],
    [2, 'Nghiệm thu, bàn giao', 0, '38', '', 0]
  ];
  const P = { v: 1, name: 'Nhà phố 1 trệt 3 lầu – mẫu ví dụ', contractor: 'Gói thầu xây dựng phần thô & hoàn thiện', start: '2026-09-07', workweek: 6,
    dataDate: '2026-11-21', useDataDate: true, holidays: VN_HOLIDAYS.map(h => ({ ...h })), resources: JSON.parse(JSON.stringify(DEFAULT_RES)), tasks: buildTasks(rows), sample: true };
  setBaseline(P, schedule(P, false)); P.baselineAt = '2026-09-04';
  return P;
}
function setBaseline(P, S) {
  P.tasks.forEach((t, i) => { if (!S.st.isSum[i]) t.bl = { s: isoOf(S.start[i]), f: isoOf(S.finish[i]), d: Math.max(0, Math.round(+t.dur || 0)) }; else t.bl = null; });
  P.baselineAt = new Date().toISOString().slice(0, 10);
}

export { DAY, DEFAULT_RES, VN_HOLIDAYS, _idn, buildCal, buildTasks, dow, expandHolidays, fmtD, fmtDM, fmtPreds, isoOf, isoWeek, mkDay, mondayOf, newId, pad, parseDate, parsePreds, parseRes, sampleProject, schedule, setBaseline, structure, toDay };
