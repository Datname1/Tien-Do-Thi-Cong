import './styles.css';
import * as XLSX from 'xlsx';
import { DAY, DEFAULT_RES, VN_HOLIDAYS, _idn, buildCal, buildTasks, dow, expandHolidays, fmtD, fmtDM, fmtPreds, isoOf, isoWeek, mkDay, mondayOf, newId, pad, parseDate, parsePreds, parseRes, sampleProject, schedule, setBaseline, structure, toDay } from './core.js';
/* ================= UI ================= */
const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const ICONS = {
  plus: '<path d="M12 5v14M5 12h14"/>', child: '<path d="M6 4v9a3 3 0 0 0 3 3h9"/><path d="m15 13 3 3-3 3"/>',
  diamond: '<path d="M12 3 21 12 12 21 3 12Z"/>', indent: '<path d="M3 5h18M11 12h10M11 19h10"/><path d="m3 9 4 3-4 3"/>',
  outdent: '<path d="M3 5h18M11 12h10M11 19h10"/><path d="m7 9-4 3 4 3"/>', up: '<path d="M12 19V5M6 11l6-6 6 6"/>',
  down: '<path d="M12 5v14M6 13l6 6 6-6"/>', copy: '<rect x="8" y="8" width="12" height="12" rx="2"/><path d="M16 8V6a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2v8a2 2 0 0 0 2 2h2"/>',
  trash: '<path d="M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3"/>', undo: '<path d="M9 14 4 9l5-5"/><path d="M4 9h11a5 5 0 0 1 0 10h-3"/>',
  redo: '<path d="m15 14 5-5-5-5"/><path d="M20 9H9a5 5 0 0 0 0 10h3"/>', folder: '<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2Z"/>',
  save: '<path d="M5 3h11l5 5v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2Z"/><path d="M7 3v5h8V3M7 21v-7h10v7"/>',
  sheet: '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M3 15h18M9 3v18"/>',
  gear: '<circle cx="12" cy="12" r="3"/><path d="M12 2v3M12 19v3M4.9 4.9l2.1 2.1M17 17l2.1 2.1M2 12h3M19 12h3M4.9 19.1 7 17M17 7l2.1-2.1"/>',
  flag: '<path d="M5 21V4h12l-2.5 4L17 12H5"/>', paste: '<rect x="6" y="4" width="12" height="17" rx="2"/><path d="M9 4V3h6v1M9 11h6M9 15h4"/>',
  file: '<path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9Z"/><path d="M14 3v6h6M12 12v6M9 15h6"/>',
  target: '<circle cx="12" cy="12" r="8"/><circle cx="12" cy="12" r="3"/>', help: '<circle cx="12" cy="12" r="9"/><path d="M9.5 9a2.5 2.5 0 1 1 3.5 2.3c-.6.3-1 .9-1 1.6V14M12 17.5v.01"/>',
  chev: '<path d="m9 6 6 6-6 6"/>', pin: '<path d="M12 17v5M8 3h8l-1 6 3 3H6l3-3Z"/>', x: '<path d="M6 6l12 12M18 6 6 18"/>',
  zoom: '<path d="M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5"/>'
};
const icon = (n, cls = '') => `<svg class="ic ${cls}" viewBox="0 0 24 24" aria-hidden="true">${ICONS[n]}</svg>`;

const LS_KEY = 'sotiendo.project.v1', LS_UI = 'sotiendo.ui.v1';
const lsGet = k => { try { return localStorage.getItem(k); } catch { return null; } };
const lsSet = (k, v) => { try { localStorage.setItem(k, v); return true; } catch { return false; } };

let P, S, VIS = [], G = null;
const UI = { view: 'plan', scale: 'week', filter: 'all', sel: null, collapsed: [], lw: 560, resOpen: true, res: null };
try { Object.assign(UI, JSON.parse(lsGet(LS_UI) || '{}')); } catch {}
UI.view = 'plan';
const RH = 30;
const SCALES = { day: { px: 30, label: 'Ngày' }, week: { px: 13, label: 'Tuần' }, month: { px: 4, label: 'Tháng' } };
const undoStack = [], redoStack = [];
let pendingFocus = null, saveTimer = null;

function loadInitial() {
  try { const raw = lsGet(LS_KEY); if (raw) { const p = JSON.parse(raw); if (p && Array.isArray(p.tasks)) return normalize(p); } } catch {}
  return sampleProject();
}
function normalize(p) {
  const q = Object.assign({ v: 1, name: 'Dự án mới', contractor: '', start: isoOf(todayDay()), workweek: 6, dataDate: isoOf(todayDay()), useDataDate: true, holidays: [], resources: {}, tasks: [] }, p);
  q.tasks = q.tasks.map(t => Object.assign({ id: newId(), level: 0, name: '', dur: 1, preds: [], res: '', pct: 0, snet: null, bl: null }, t));
  return q;
}
const todayDay = () => { const n = new Date(); return mkDay(n.getFullYear(), n.getMonth() + 1, n.getDate()); };

function pushUndo() { undoStack.push(JSON.stringify(P)); if (undoStack.length > 100) undoStack.shift(); redoStack.length = 0; }
function undo() { if (!undoStack.length) return; redoStack.push(JSON.stringify(P)); P = JSON.parse(undoStack.pop()); renderAll(); toast('Đã hoàn tác'); }
function redo() { if (!redoStack.length) return; undoStack.push(JSON.stringify(P)); P = JSON.parse(redoStack.pop()); renderAll(); toast('Đã làm lại'); }

function persist() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => {
    const ok = lsSet(LS_KEY, JSON.stringify(P)); lsSet(LS_UI, JSON.stringify({ scale: UI.scale, filter: UI.filter, collapsed: UI.collapsed, lw: UI.lw, resOpen: UI.resOpen, res: UI.res }));
    $('#saved').textContent = ok ? 'Đã lưu trên trình duyệt này' : 'Chưa lưu được trên trình duyệt (hãy dùng Lưu .json)';
  }, 300);
}

function toast(msg, err) {
  const el = document.createElement('div'); el.className = 'toast' + (err ? ' err' : ''); el.textContent = msg;
  $('#toasts').appendChild(el); setTimeout(() => el.remove(), err ? 4200 : 2600);
}

/* ---------- visible rows ---------- */
function computeVisible() {
  const T = P.tasks, st = S.st, n = T.length, col = new Set(UI.collapsed);
  let keep = null;
  if (UI.filter !== 'all') {
    keep = new Set(); const a = S.dataDay, b = S.dataDay + 13;
    S.leaves.forEach(i => {
      let ok = false;
      if (UI.filter === 'crit') ok = S.crit[i];
      else if (UI.filter === 'late') ok = S.late[i];
      else if (UI.filter === 'week') ok = S.pct[i] < 100 && S.start[i] <= b && S.finish[i] >= a;
      if (ok) { keep.add(i); let p = st.parent[i]; while (p >= 0) { keep.add(p); p = st.parent[p]; } }
    });
  }
  const out = [];
  for (let i = 0; i < n; i++) {
    if (keep && !keep.has(i)) continue;
    let p = st.parent[i], hide = false; while (p >= 0) { if (col.has(T[p].id)) { hide = true; break; } p = st.parent[p]; }
    if (!hide) out.push(i);
  }
  VIS = out;
}

/* ---------- ribbon ---------- */
function renderRibbon() {
  const nCrit = S.leaves.filter(i => S.crit[i]).length, nLate = S.leaves.filter(i => S.late[i]).length;
  const f = (k, l, c) => `<button class="rb ${UI.filter === k ? 'on' : ''}" data-filter="${k}">${l}${c != null ? ` <span class="cnt">${c}</span>` : ''}</button>`;
  const sc = k => `<button class="rb ${UI.scale === k ? 'on' : ''}" data-scale="${k}">${SCALES[k].label}</button>`;
  $('#ribbon').innerHTML = `
    <div class="grp"><span class="gl">Thêm</span>
      <button class="rb" data-act="addBelow" title="Thêm công việc bên dưới, tự nối FS với việc đang chọn (Insert)">${icon('plus')}Việc dưới</button>
      <button class="rb" data-act="addChild" title="Thêm việc con của dòng đang chọn">${icon('child')}Việc con</button>
      <button class="rb" data-act="addMs" title="Thêm mốc (0 ngày)">${icon('diamond')}Mốc</button>
      <button class="rb" data-act="paste" title="Dán bảng copy từ Excel">${icon('paste')}Dán Excel</button>
    </div>
    <div class="grp"><span class="gl">Cấu trúc</span>
      <button class="rb" data-act="indent" title="Thụt lề (Alt+Shift+→)">${icon('indent')}</button>
      <button class="rb" data-act="outdent" title="Lùi lề (Alt+Shift+←)">${icon('outdent')}</button>
      <button class="rb" data-act="up" title="Lên trên (Alt+Shift+↑)">${icon('up')}</button>
      <button class="rb" data-act="down" title="Xuống dưới (Alt+Shift+↓)">${icon('down')}</button>
      <button class="rb" data-act="dup" title="Nhân đôi dòng (kèm việc con)">${icon('copy')}</button>
      <button class="rb danger" data-act="del" title="Xóa dòng đang chọn (Delete)">${icon('trash')}</button>
    </div>
    <div class="grp"><span class="gl">Lọc</span>${f('all', 'Tất cả')}${f('crit', 'Đường găng', nCrit)}${f('late', 'Chậm', nLate)}${f('week', '2 tuần tới')}</div>
    <div class="grp"><span class="gl">Thang đo</span>${sc('day')}${sc('week')}${sc('month')}</div>
    <div class="grp"><span class="gl">Ngày BC</span>
      <input type="date" id="dataDate" value="${esc(P.dataDate)}" title="Ngày báo cáo (data date): mốc để tính kế hoạch và dời việc chưa làm">
      <button class="rb" data-act="goData" title="Cuộn Gantt tới ngày báo cáo">${icon('target')}</button>
      <button class="rb" data-act="fit" title="Vừa màn hình">${icon('zoom')}</button>
      <button class="rb" data-act="baseline" title="Chốt kế hoạch gốc để so sánh">${icon('flag')}Lưu KH gốc</button>
    </div>`;
}

/* ---------- table ---------- */
const COLS = [['STT', 'c'], ['WBS', ''], ['Tên công việc', ''], ['Ngày', 'c'], ['Bắt đầu', ''], ['Kết thúc', ''], ['LK trước', ''], ['Dự trữ', 'c'], ['Nguồn lực', ''], ['%', 'c'], ['Lệch KHG', 'c']];
const COL_TIPS = ['Số thứ tự – dùng để khai báo liên kết', 'Mã WBS tự đánh theo cấp', '', 'Thời gian (ngày làm việc)', 'Gõ ngày để ghim “bắt đầu không sớm hơn”', 'Tính tự động', 'Ví dụ: 3, 5SS+2, 7FF-1', 'Dự trữ toàn phần (ngày làm việc)', 'Ví dụ: NC:8, MAY:1', '% hoàn thành thực tế', 'Chênh lệch ngày kết thúc so với kế hoạch gốc'];
function renderHead() {
  $('#lHead').innerHTML = `<div class="grid">${COLS.map((c, k) => `<div class="th ${c[1]}" title="${esc(COL_TIPS[k])}">${c[0]}</div>`).join('')}</div>`;
}
function renderTable() {
  const T = P.tasks, st = S.st, col = new Set(UI.collapsed);
  if (!T.length) {
    $('#lBody').innerHTML = `<div class="empty"><b>Chưa có công việc nào.</b><br>Bấm <b>Việc dưới</b> để thêm từng dòng, <b>Dán Excel</b> để dán bảng có sẵn, hoặc <b>Dự án mới</b> để tạo khung WBS 6 phần.</div>`;
    return;
  }
  let h = '';
  for (const i of VIS) {
    const t = T[i], sum = st.isSum[i], ms = !sum && (+t.dur || 0) === 0;
    const crit = S.crit[i] && !sum, sel = t.id === UI.sel;
    const tf = S.tf[i];
    const tfTxt = sum ? '<span class="muted">—</span>' : (tf <= 0 ? `<span class="tf-crit">0d</span>` : `<span class="ro">+${tf}d</span>`);
    const v = S.vr[i];
    const vTxt = v == null ? '<span class="muted">—</span>' : v > 0 ? `<span class="var-bad">+${v}d</span>` : v < 0 ? `<span class="var-good">−${-v}d</span>` : '<span class="muted">0d</span>';
    const pv = Math.round(S.pct[i]);
    h += `<div class="tr grid ${sum ? 'sum' : ''} ${sel ? 'sel' : ''}" data-id="${t.id}" data-i="${i}">
      <div class="td c c-stt" title="Chọn dòng">${i + 1}</div>
      <div class="td c-wbs">${st.wbs[i]}</div>
      <div class="td c-name" style="--lvl:${t.level}">${sum ? `<button class="chev ${col.has(t.id) ? '' : 'open'}" data-tog="${t.id}" title="Thu gọn / mở rộng">${icon('chev')}</button>` : '<span class="chev-sp"></span>'}${ms ? '<span class="dia"></span>' : ''}${crit ? '<span class="badge crit">GĂNG</span>' : (!sum && S.late[i] ? '<span class="badge late">CHẬM</span>' : '')}<input data-col="name" value="${esc(t.name)}" aria-label="Tên công việc"></div>
      <div class="td c num">${sum ? `<span class="ro">${S.ef[i] - S.es[i]}</span>` : `<input data-col="dur" value="${+t.dur || 0}" inputmode="numeric" aria-label="Số ngày">`}</div>
      <div class="td num">${sum ? `<span class="ro">${fmtD(S.start[i])}</span>` : `<input data-col="start" value="${fmtD(S.start[i])}" aria-label="Bắt đầu">${t.snet ? `<span title="Ghim bắt đầu không sớm hơn ${fmtD(toDay(t.snet))} – xóa ô để bỏ ghim">${icon('pin', 'pin')}</span>` : ''}`}</div>
      <div class="td"><span class="ro">${fmtD(S.finish[i])}</span></div>
      <div class="td num" ${S.badLink[i] ? 'title="Có liên kết không hợp lệ (tự liên kết hoặc nối với việc cha/con)"' : ''}><input data-col="preds" value="${esc(fmtPreds(t, S.idIdx))}" aria-label="Liên kết trước" ${S.badLink[i] || S.cycle[i] ? 'style="color:var(--crit)"' : ''}></div>
      <div class="td c num">${tfTxt}</div>
      <div class="td num">${sum ? '' : `<input data-col="res" value="${esc(t.res || '')}" aria-label="Nguồn lực">`}</div>
      <div class="td c num">${sum ? `<span class="ro">${pv}%</span>` : `<input data-col="pct" value="${pv}" inputmode="numeric" aria-label="Phần trăm hoàn thành">`}<span class="pbar"><i style="width:${pv}%"></i></span></div>
      <div class="td c num">${vTxt}</div>
    </div>`;
  }
  $('#lBody').innerHTML = h;
}

/* ---------- gantt ---------- */
function ganttRange() {
  let a = S.projStart, b = S.projFinish;
  if (S.blFinish) b = Math.max(b, S.blFinish);
  if (Math.abs(S.dataDay - a) < 900) { a = Math.min(a, S.dataDay); b = Math.max(b, S.dataDay); }
  a = mondayOf(a - 7); b = b + 28;
  return { rs: a, re: b };
}
function renderGantt() {
  const T = P.tasks, st = S.st, px = SCALES[UI.scale].px, { rs, re } = ganttRange();
  const x = d => (d - rs) * px, W = (re - rs + 1) * px;
  G = { rs, re, px, x, W };
  const body = $('#gBody'), H = Math.max(VIS.length * RH, body.clientHeight - 2);
  let s = `<svg class="gsvg" width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg"><defs>
    <marker id="ah" viewBox="0 0 8 8" refX="7" refY="4" markerWidth="7" markerHeight="7" orient="auto"><path d="M0 0 8 4 0 8Z" class="g-ah"/></marker>
    <marker id="ahc" viewBox="0 0 8 8" refX="7" refY="4" markerWidth="7" markerHeight="7" orient="auto"><path d="M0 0 8 4 0 8Z" class="g-ah crit"/></marker></defs>`;
  if (UI.scale !== 'month') for (let d = rs; d <= re; d++) if (!S.cal.isWork(d)) s += `<rect class="g-off" x="${x(d)}" y="0" width="${px}" height="${H}"/>`;
  const selR = VIS.findIndex(i => T[i].id === UI.sel);
  s += `<rect id="gSel" class="g-selrow" x="0" y="${selR * RH}" width="${W}" height="${RH}" ${selR < 0 ? 'visibility="hidden"' : ''}/>`;
  for (let d = rs; d <= re; d++) {
    const dm = +isoOf(d).slice(8, 10);
    if (UI.scale === 'day') s += `<line class="g-grid ${dow(d) === 1 ? 'strong' : ''}" x1="${x(d) + .5}" x2="${x(d) + .5}" y1="0" y2="${H}"/>`;
    else if (UI.scale === 'week' && dow(d) === 1) s += `<line class="g-grid" x1="${x(d) + .5}" x2="${x(d) + .5}" y1="0" y2="${H}"/>`;
    else if (UI.scale === 'month' && dm === 1) s += `<line class="g-grid strong" x1="${x(d) + .5}" x2="${x(d) + .5}" y1="0" y2="${H}"/>`;
  }
  VIS.forEach((i, r) => { s += `<line class="g-row" x1="0" x2="${W}" y1="${(r + 1) * RH - .5}" y2="${(r + 1) * RH - .5}"/>`; });
  const rowOf = new Map(VIS.map((i, r) => [i, r]));
  const ends = i => { const ms = !st.isSum[i] && (+T[i].dur || 0) === 0; return ms ? { a: x(S.start[i]), b: x(S.start[i]) } : { a: x(S.start[i]), b: x(S.finish[i] + 1) }; };
  // links
  VIS.forEach(q => (T[q].preds || []).forEach(l => {
    const p = S.idIdx.get(l.id); if (p == null || !rowOf.has(p)) return;
    const ty = l.type || 'FS', ep = ends(p), eq = ends(q);
    const y1 = rowOf.get(p) * RH + RH / 2, y2 = rowOf.get(q) * RH + RH / 2;
    const fromEnd = ty[0] === 'F', toStart = ty[1] === 'S';
    const x1 = fromEnd ? ep.b : ep.a, x2 = toStart ? eq.a - 1 : eq.b + 1;
    const dOut = fromEnd ? 1 : -1, dIn = toStart ? 1 : -1;
    const p1x = x1 + dOut * 6, qx = x2 - dIn * 8;
    let path;
    if ((dIn === 1 && qx >= p1x) || (dIn === -1 && qx <= p1x)) path = `M${x1} ${y1}H${p1x}V${y2}H${x2}`;
    else { const ym = y2 + (y2 > y1 ? -RH / 2 : RH / 2); path = `M${x1} ${y1}H${p1x}V${ym}H${qx}V${y2}H${x2}`; }
    const c = S.crit[p] && S.crit[q];
    s += `<path class="g-link ${c ? 'crit' : ''}" d="${path}" marker-end="url(#${c ? 'ahc' : 'ah'})"/>`;
  }));
  // bars
  VIS.forEach((i, r) => {
    const t = T[i], y = r * RH, sum = st.isSum[i], ms = !sum && (+t.dur || 0) === 0;
    if (S.blS[i] != null && !sum) {
      const bms = (t.bl && +t.bl.d === 0);
      if (bms) s += `<rect class="g-bl" x="${x(S.blS[i]) - 3}" y="${y + RH - 8}" width="6" height="4" rx="1"/>`;
      else s += `<rect class="g-bl" x="${x(S.blS[i])}" y="${y + RH - 8}" width="${Math.max(2, x(S.blF[i] + 1) - x(S.blS[i]))}" height="4" rx="1"/>`;
    }
    const xa = x(S.start[i]), xb = x(S.finish[i] + 1);
    if (sum) {
      s += `<rect class="g-sum" x="${xa}" y="${y + 10}" width="${Math.max(2, xb - xa)}" height="6"/><path class="g-sum" d="M${xa} ${y + 10}h6l-6 8Z M${xb} ${y + 10}h-6l6 8Z"/>`;
      s += `<text class="g-lbl" x="${xb + 8}" y="${y + 18}">${fmtDM(S.start[i])} → ${fmtDM(S.finish[i])}</text>`;
    } else if (ms) {
      const cx = xa, cy = y + RH / 2 - 1;
      s += `<path class="g-ms ${S.crit[i] ? 'crit' : ''}" d="M${cx} ${cy - 7}L${cx + 7} ${cy}L${cx} ${cy + 7}L${cx - 7} ${cy}Z"/>`;
      s += `<text class="g-lbl" x="${cx + 11}" y="${cy + 4}">${fmtDM(S.start[i])}${S.pct[i] >= 100 ? ' ✓' : ''}</text>`;
    } else {
      const w = Math.max(3, xb - xa), cls = S.pct[i] >= 100 ? 'g-done' : S.crit[i] ? 'g-crit' : 'g-norm';
      s += `<rect class="${cls}" x="${xa}" y="${y + 7}" width="${w}" height="13" rx="3" fill-opacity="${S.pct[i] >= 100 ? 0.85 : 0.38}"/>`;
      if (S.pct[i] > 0 && S.pct[i] < 100) s += `<rect class="${cls}" x="${xa}" y="${y + 7}" width="${w * S.pct[i] / 100}" height="13" rx="3"/>`;
      if (S.cycle[i]) s += `<rect class="g-bad" x="${xa - 2}" y="${y + 5}" width="${w + 4}" height="17" rx="4"/>`;
      s += `<text class="g-lbl" x="${xb + 6}" y="${y + 18}">${fmtDM(S.finish[i])}${S.pct[i] > 0 && S.pct[i] < 100 ? ' · ' + Math.round(S.pct[i]) + '%' : ''}</text>`;
    }
  });
  const tday = todayDay();
  if (tday >= rs && tday <= re) s += `<line class="g-today" x1="${x(tday) + .5}" x2="${x(tday) + .5}" y1="0" y2="${H}"/>`;
  if (S.dataDay >= rs && S.dataDay <= re) s += `<line class="g-dd" x1="${x(S.dataDay) + .5}" x2="${x(S.dataDay) + .5}" y1="0" y2="${H}"/>`;
  s += '</svg>';
  body.innerHTML = s;
  renderGanttHead();
}
function renderGanttHead() {
  const { rs, re, px, x, W } = G; let s = `<svg class="gsvg" width="${W}" height="44" xmlns="http://www.w3.org/2000/svg">`;
  // months
  let d = rs;
  while (d <= re) {
    const iso = isoOf(d), y = +iso.slice(0, 4), m = +iso.slice(5, 7);
    const next = mkDay(m === 12 ? y + 1 : y, m === 12 ? 1 : m + 1, 1), e = Math.min(next - 1, re);
    const w = x(e + 1) - x(d);
    s += `<rect class="h-cell" x="${x(d) + .5}" y=".5" width="${w}" height="21"/>`;
    const lbl = w > 90 ? `Tháng ${pad(m)}/${y}` : w > 40 ? `${pad(m)}/${String(y).slice(2)}` : '';
    if (lbl) s += `<text class="h-txt" x="${x(d) + 6}" y="15">${lbl}</text>`;
    d = next;
  }
  const WD = ['CN', 'T2', 'T3', 'T4', 'T5', 'T6', 'T7'];
  if (UI.scale === 'day') {
    for (let k = rs; k <= re; k++) {
      const off = !S.cal.isWork(k);
      s += `<rect class="h-cell" x="${x(k) + .5}" y="22.5" width="${px}" height="21"/><text class="h-sub ${off ? 'off' : ''}" x="${x(k) + px / 2}" y="31" text-anchor="middle">${isoOf(k).slice(8, 10)}</text><text class="h-sub ${off ? 'off' : ''}" x="${x(k) + px / 2}" y="41" text-anchor="middle" style="font-size:8.5px">${WD[dow(k)]}</text>`;
    }
  } else {
    for (let k = mondayOf(rs); k <= re; k += 7) {
      const xa = Math.max(0, x(k));
      s += `<rect class="h-cell" x="${xa + .5}" y="22.5" width="${7 * px}" height="21"/>`;
      const lbl = UI.scale === 'week' ? `T${isoWeek(k)} · ${fmtDM(k)}` : isoOf(k).slice(8, 10);
      s += `<text class="h-sub" x="${xa + (UI.scale === 'week' ? 5 : 3)}" y="37">${lbl}</text>`;
    }
  }
  if (S.dataDay >= rs && S.dataDay <= re) { const xd = x(S.dataDay); s += `<rect class="h-pill" x="${xd - 31}" y="3" width="62" height="16" rx="8"/><text class="h-pill-t" x="${xd}" y="14.5" text-anchor="middle">BC ${fmtDM(S.dataDay)}</text>`; }
  s += '</svg>';
  $('#gHead').innerHTML = s;
  $('#gHead').scrollLeft = $('#gBody').scrollLeft;
}

/* ---------- resources ---------- */
function resCodes() {
  const set = new Set(Object.keys(P.resources || {}));
  P.tasks.forEach((t, i) => { if (!S.st.isSum[i]) parseRes(t.res).forEach(r => set.add(r.code)); });
  return [...set];
}
function renderRes() {
  const box = $('#res'); box.classList.toggle('closed', !UI.resOpen);
  $('#resToggle .chev').innerHTML = icon('chev');
  if (!UI.resOpen) return;
  const codes = resCodes();
  if (!codes.length) { $('#resLeft').innerHTML = '<div class="none">Chưa có nguồn lực. Ghi vào cột Nguồn lực, ví dụ <span class="code">NC:8</span>.</div>'; $('#resChart').innerHTML = ''; return; }
  if (!UI.res || !codes.includes(UI.res)) UI.res = codes[0];
  const info = (P.resources || {})[UI.res] || { name: UI.res, unit: '', limit: 0 };
  const load = new Map();
  S.leaves.forEach(i => parseRes(P.tasks[i].res).forEach(r => {
    if (r.code !== UI.res) return;
    for (let d = S.start[i]; d <= S.finish[i]; d++) if (S.cal.isWork(d)) load.set(d, (load.get(d) || 0) + r.qty);
  }));
  const vals = [...load.values()], peak = vals.length ? Math.max(...vals) : 0, lim = +info.limit || 0;
  const over = lim ? vals.filter(v => v > lim).length : 0, avg = vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : 0;
  const u = info.unit ? ' ' + esc(info.unit) : '';
  $('#resLeft').innerHTML = `<select id="resSel" aria-label="Chọn nguồn lực">${codes.map(c => `<option value="${esc(c)}" ${c === UI.res ? 'selected' : ''}>${esc(c)} — ${esc((P.resources[c] || {}).name || 'Chưa đặt tên')}</option>`).join('')}</select>
    <div class="stats">
      <div class="stat"><div class="k">Giới hạn/ngày</div><input id="resLim" value="${lim || ''}" placeholder="—" inputmode="numeric" aria-label="Giới hạn mỗi ngày"></div>
      <div class="stat ${lim && peak > lim ? 'bad' : ''}"><div class="k">Nhu cầu đỉnh</div><div class="v">${+peak.toFixed(1)}${u}</div></div>
      <div class="stat ${over ? 'bad' : 'good'}"><div class="k">Ngày vượt</div><div class="v">${lim ? (over ? over + ' ngày' : '✓ 0') : '—'}</div></div>
      <div class="stat"><div class="k">TB/ngày</div><div class="v">${+avg.toFixed(1)}${u}</div></div>
    </div>`;
  const { rs, re, px, x, W } = G, Hc = Math.max(90, $('#resChart').clientHeight || 128), top = 14, bot = 14, ph = Hc - top - bot;
  const scaleMax = Math.max(peak, lim, 1) * 1.1;
  let s = `<svg class="gsvg" width="${W}" height="${Hc}" xmlns="http://www.w3.org/2000/svg">`;
  for (const [d, v] of load) {
    if (d < rs || d > re) continue; const h = v / scaleMax * ph;
    s += `<rect class="r-bar ${lim && v > lim ? 'over' : ''}" x="${x(d) + (px > 6 ? 1 : 0)}" y="${top + ph - h}" width="${Math.max(1, px - (px > 6 ? 2 : 0))}" height="${h}"/>`;
    if (UI.scale === 'day') s += `<text class="r-lbl" x="${x(d) + px / 2}" y="${top + ph - h - 3}" text-anchor="middle">${+v.toFixed(1)}</text>`;
  }
  if (lim) { const yl = top + ph - lim / scaleMax * ph; s += `<line class="r-lim" x1="0" x2="${W}" y1="${yl}" y2="${yl}"/><text class="r-lbl" x="${($('#gBody').scrollLeft || 0) + 6}" y="${yl - 4}">Giới hạn ${lim}${u}</text>`; }
  s += '</svg>';
  $('#resChart').innerHTML = s; $('#resChart').scrollLeft = $('#gBody').scrollLeft;
}

/* ---------- status ---------- */
function renderStatus() {
  const n = S.leaves.length, c = S.leaves.filter(i => S.crit[i]).length;
  const cal = S.projFinish - S.projStart + 1;
  $('#status').innerHTML = `<span><b>${n}</b> việc</span><span class="crit"><b class="crit">${c}</b> việc găng</span>
    <span>Khởi công <b>${fmtD(S.projStart)}</b> → Hoàn thành <b>${fmtD(S.projFinish)}</b></span>
    <span><b>${S.workDays}</b> ngày làm việc · ${cal} ngày lịch</span>
    <span>Thực tế <b>${S.actualPct.toFixed(1)}%</b> / KH <b>${S.plannedPct.toFixed(1)}%</b></span>
    <span class="legend">
      <span class="lg"><i class="sw" style="background:var(--crit)"></i>Găng</span>
      <span class="lg"><i class="sw" style="background:var(--norm)"></i>Có dự trữ</span>
      <span class="lg"><i class="sw" style="background:var(--ok)"></i>Xong</span>
      <span class="lg"><i class="sw" style="background:var(--bl);height:4px"></i>KH gốc</span>
      <span class="lg"><i class="sw" style="background:var(--sum);height:5px"></i>Hạng mục</span>
      <span class="lg"><i class="sw" style="border-top:2px dashed var(--data);height:0;border-radius:0"></i>Ngày BC</span>
      <span class="lg"><i class="sw" style="background:var(--today);width:2px"></i>Hôm nay</span>
    </span>`;
}

/* ---------- report ---------- */
function sCurve() {
  const T = P.tasks, a = S.projStart, b = Math.max(S.projFinish, S.blFinish || 0);
  const pts = []; const W = S.leaves.reduce((s, i) => s + Math.max(+T[i].dur || 0, 0), 0) || 1;
  const fracPlan = (i, d) => { const s0 = S.blS[i] ?? S.start[i], f0 = S.blF[i] ?? S.finish[i]; if (d < s0) return 0; if (d > f0) return 1; return (d - s0 + 1) / (f0 - s0 + 1); };
  const fracFc = (i, d) => { const p = S.pct[i] / 100; const s0 = Math.max(S.start[i], S.dataDay), f0 = S.finish[i]; if (p >= 1) return 1; if (d < s0) return p; if (d > f0) return 1; return p + (1 - p) * (d - s0 + 1) / (f0 - s0 + 1); };
  for (let d = a; d <= b + 7; d += 7) {
    const dd = Math.min(d, b);
    let pl = 0, fc = 0; S.leaves.forEach(i => { const w = Math.max(+T[i].dur || 0, 0); pl += w * fracPlan(i, dd); fc += w * fracFc(i, dd); });
    pts.push({ d: dd, pl: pl / W * 100, fc: fc / W * 100 });
    if (dd === b) break;
  }
  return { a, b, pts };
}
function renderReport() {
  const T = P.tasks, D = S.dataDay;
  const gap = S.actualPct - S.plannedPct;
  const pfVar = S.blFinish != null ? S.cal.idxOf(S.projFinish) - S.cal.idxOf(S.blFinish) : null;
  const lateL = S.leaves.filter(i => S.late[i]).sort((x, y) => (S.planned[y] - S.pct[y]) - (S.planned[x] - S.pct[x]));
  const doing = S.leaves.filter(i => S.pct[i] > 0 && S.pct[i] < 100 || (S.pct[i] < 100 && S.start[i] <= D && S.finish[i] >= D && (+T[i].dur || 0) > 0));
  const ahead = S.leaves.filter(i => S.pct[i] < 100 && S.start[i] > D && S.start[i] <= D + 14).sort((x, y) => S.start[x] - S.start[y]);
  const critLeft = S.leaves.filter(i => S.crit[i]).length;
  const pill = (v, unit = '%', inv) => { const bad = inv ? v > 0 : v < 0; const t = (v > 0 ? '+' : v < 0 ? '−' : '') + Math.abs(v).toFixed(unit === '%' ? 1 : 0) + unit; return `<span class="pill ${Math.abs(v) < 0.05 ? 'neu' : bad ? 'bad' : 'good'}">${t}</span>`; };
  // S curve svg
  const sc = sCurve(); const Wd = 720, Hd = 240, L = 40, R = 16, Tp = 12, B = 28;
  const sx = d => L + (sc.b === sc.a ? 0 : (d - sc.a) / (sc.b - sc.a)) * (Wd - L - R), sy = v => Tp + (1 - v / 100) * (Hd - Tp - B);
  let svg = `<svg class="s-svg" viewBox="0 0 ${Wd} ${Hd}" role="img" aria-label="Đường cong S tiến độ">`;
  [0, 25, 50, 75, 100].forEach(v => { svg += `<line class="s-grid" x1="${L}" x2="${Wd - R}" y1="${sy(v)}" y2="${sy(v)}"/><text class="s-lbl" x="${L - 6}" y="${sy(v) + 4}" text-anchor="end">${v}%</text>`; });
  const step = Math.max(1, Math.ceil(sc.pts.length / 8));
  sc.pts.forEach((p, k) => { if (k % step === 0) svg += `<text class="s-lbl" x="${sx(p.d)}" y="${Hd - 8}" text-anchor="middle">${fmtDM(p.d)}</text>`; });
  svg += `<path class="s-plan" d="${sc.pts.map((p, k) => (k ? 'L' : 'M') + sx(p.d).toFixed(1) + ' ' + sy(p.pl).toFixed(1)).join('')}"/>`;
  const fc = sc.pts.filter(p => p.d >= D);
  if (D >= sc.a && D <= sc.b) {
    svg += `<path class="s-fc" d="M${sx(D).toFixed(1)} ${sy(S.actualPct).toFixed(1)}${fc.map(p => 'L' + sx(p.d).toFixed(1) + ' ' + sy(p.fc).toFixed(1)).join('')}"/>`;
    svg += `<line class="g-dd" x1="${sx(D)}" x2="${sx(D)}" y1="${Tp}" y2="${Hd - B}"/><circle class="s-act" cx="${sx(D)}" cy="${sy(S.actualPct)}" r="5"/>`;
    svg += `<text class="s-lbl" x="${sx(D) + 8}" y="${sy(S.actualPct) + 14}">Thực tế ${S.actualPct.toFixed(1)}%</text>`;
  }
  svg += '</svg>';
  const row = i => `<tr><td>${esc(T[i].name)}${S.crit[i] ? ' <span class="badge crit">GĂNG</span>' : ''}</td><td class="n">${fmtDM(S.start[i])} → ${fmtDM(S.finish[i])}</td><td class="n">${Math.round(S.pct[i])}% / ${Math.round(S.planned[i])}%</td></tr>`;
  const tbl = (list, empty) => list.length ? `<div class="tbl-wrap"><table class="rt"><thead><tr><th>Công việc</th><th>Thời gian</th><th>TT / KH</th></tr></thead><tbody>${list.map(row).join('')}</tbody></table></div>` : `<div class="none">${empty}</div>`;
  $('#report').innerHTML = `<div class="rep-in">
    <div class="rep-head"><div><h2>${esc(P.name)}</h2><p>Báo cáo tại ngày <b>${fmtD(D)}</b>${P.baselineAt ? ` · so với kế hoạch gốc chốt ngày ${fmtD(toDay(P.baselineAt))}` : ' · chưa chốt kế hoạch gốc'}</p></div></div>
    <div class="kpis">
      <div class="kpi"><div class="k">Hoàn thành thực tế</div><div class="v">${S.actualPct.toFixed(1)}%</div><div class="s">Kế hoạch ${S.plannedPct.toFixed(1)}% ${pill(gap)}</div></div>
      <div class="kpi"><div class="k">Dự kiến hoàn thành</div><div class="v">${fmtD(S.projFinish)}</div><div class="s">${pfVar == null ? 'Chưa có kế hoạch gốc' : `So với KH gốc ${pill(pfVar, ' ngày', true)}`}</div></div>
      <div class="kpi"><div class="k">Việc đang chậm</div><div class="v" style="color:${lateL.length ? 'var(--crit)' : 'var(--ok)'}">${lateL.length}</div><div class="s">trên ${S.leaves.length} công việc</div></div>
      <div class="kpi"><div class="k">Việc găng còn lại</div><div class="v">${critLeft}</div><div class="s">${S.workDays} ngày làm việc · tuần ${P.workweek} ngày</div></div>
    </div>
    <div class="two">
      <div class="card"><h3>Đường cong S <span class="pill neu">trọng số theo thời lượng</span></h3>${svg}
        <div class="s-key"><span class="lg"><i class="sw" style="background:var(--faint);height:3px"></i>Kế hoạch ${P.baselineAt ? 'gốc' : 'hiện tại'}</span><span class="lg"><i class="sw" style="border-top:2px dashed var(--accent-strong);height:0;border-radius:0"></i>Dự báo từ ngày BC</span><span class="lg"><i class="sw" style="background:var(--ok);width:8px;height:8px;border-radius:50%"></i>Thực tế</span></div></div>
      <div class="card"><h3>Tin nhắn báo cáo nhanh</h3><textarea class="zalo" id="zaloTxt" aria-label="Nội dung báo cáo">${esc(zaloText(lateL, doing, ahead, gap, pfVar))}</textarea>
        <div style="display:flex;gap:8px;margin-top:8px;flex-wrap:wrap"><button class="tb primary" id="copyZalo">${icon('copy')}Sao chép</button></div></div>
    </div>
    <div class="two">
      <div class="card"><h3>Việc đang chậm <span class="pill ${lateL.length ? 'bad' : 'good'}">${lateL.length}</span></h3>${tbl(lateL, 'Không có việc chậm so với kế hoạch tại ngày báo cáo.')}</div>
      <div class="card"><h3>Đang thi công</h3>${tbl(doing, 'Không có việc nào đang thi công.')}</div>
    </div>
    <div class="card"><h3>Kế hoạch 2 tuần tới (${fmtDM(D)} – ${fmtDM(D + 14)})</h3>${tbl(ahead, 'Không có việc mới bắt đầu trong 2 tuần tới.')}</div>
  </div>`;
}
function zaloText(lateL, doing, ahead, gap, pfVar) {
  const T = P.tasks; const L = [];
  L.push(`BÁO CÁO TIẾN ĐỘ – ${P.name.toUpperCase()}`);
  L.push(`Ngày báo cáo: ${fmtD(S.dataDay)}`);
  L.push(`• Hoàn thành: ${S.actualPct.toFixed(1)}% (kế hoạch ${S.plannedPct.toFixed(1)}%, ${gap >= 0 ? 'vượt' : 'chậm'} ${Math.abs(gap).toFixed(1)}%)`);
  L.push(`• Dự kiến hoàn thành: ${fmtD(S.projFinish)}${pfVar == null ? '' : pfVar > 0 ? ` (trễ ${pfVar} ngày so với KH gốc)` : pfVar < 0 ? ` (sớm ${-pfVar} ngày so với KH gốc)` : ' (đúng KH gốc)'}`);
  if (doing.length) { L.push('', 'ĐANG THI CÔNG:'); doing.slice(0, 8).forEach(i => L.push(`- ${T[i].name}: ${Math.round(S.pct[i])}%`)); }
  if (lateL.length) { L.push('', 'VIỆC CHẬM CẦN XỬ LÝ:'); lateL.slice(0, 8).forEach(i => L.push(`- ${T[i].name}: TT ${Math.round(S.pct[i])}% / KH ${Math.round(S.planned[i])}%`)); }
  if (ahead.length) { L.push('', '2 TUẦN TỚI:'); ahead.slice(0, 8).forEach(i => L.push(`- ${T[i].name} (${fmtDM(S.start[i])} – ${fmtDM(S.finish[i])})`)); }
  return L.join('\n');
}

/* ---------- render all ---------- */
function renderAll(opts = {}) {
  S = schedule(P, P.useDataDate !== false);
  computeVisible();
  $('#projName').textContent = P.name || 'Dự án chưa đặt tên';
  $('#projSub').textContent = `${S.leaves.length} công việc · ${P.contractor || 'Tuần ' + P.workweek + ' ngày'}`;
  $('#bUndo').disabled = !undoStack.length; $('#bRedo').disabled = !redoStack.length;
  $('#board').style.setProperty('--lw', UI.lw + 'px');
  renderRibbon();
  if (UI.view === 'plan') { renderTable(); renderGantt(); renderRes(); renderStatus(); }
  else renderReport();
  persist();
}

/* ---------- mutations ---------- */
const idxById = id => P.tasks.findIndex(t => t.id === id);
function selIdx() { const i = UI.sel ? idxById(UI.sel) : -1; return i; }
function select(id) {
  UI.sel = id;
  document.querySelectorAll('#lBody .tr.sel').forEach(r => r.classList.remove('sel'));
  const row = document.querySelector(`#lBody .tr[data-id="${id}"]`); if (row) row.classList.add('sel');
  const r = VIS.findIndex(i => P.tasks[i].id === id), g = document.getElementById('gSel');
  if (g) { if (r < 0) g.setAttribute('visibility', 'hidden'); else { g.removeAttribute('visibility'); g.setAttribute('y', r * RH); } }
}
function insertAt(pos, task) { P.tasks.splice(pos, 0, task); }
function newTask(o) { return Object.assign({ id: newId(), level: 1, name: 'Công việc mới', dur: 3, preds: [], res: '', pct: 0, snet: null, bl: null }, o); }
function act(a) {
  const T = P.tasks, i = selIdx(), st = S.st;
  const needSel = () => { if (i < 0) { toast('Hãy chọn một dòng trước (bấm vào số STT).', true); return false; } return true; };
  if (a === 'addBelow' || a === 'addMs') {
    pushUndo(); const ms = a === 'addMs';
    let pos, lvl, preds = [];
    if (i < 0) { pos = T.length; lvl = T.length ? Math.max(1, T[T.length - 1].level) : 0; if (T.length) { const last = T[T.length - 1]; if (!st.isSum[T.length - 1]) preds = [{ id: last.id, type: 'FS', lag: 0 }]; } }
    else if (st.isSum[i] && !UI.collapsed.includes(T[i].id)) { pos = i + 1; lvl = T[i].level + 1; }
    else { pos = st.end[i] + 1; lvl = T[i].level; if (!st.isSum[i]) preds = [{ id: T[i].id, type: 'FS', lag: 0 }]; }
    const t = newTask({ level: lvl, preds, name: ms ? 'Mốc hoàn thành' : 'Công việc mới', dur: ms ? 0 : 3 });
    insertAt(pos, t); UI.sel = t.id; if (UI.filter !== 'all') UI.filter = 'all';
    pendingFocus = { id: t.id, col: 'name', selectAll: true }; renderAll(); flushFocus(); return;
  }
  if (a === 'addChild') {
    if (!needSel()) return; pushUndo();
    if (UI.collapsed.includes(T[i].id)) UI.collapsed = UI.collapsed.filter(x => x !== T[i].id);
    const t = newTask({ level: T[i].level + 1 }); insertAt(st.end[i] + 1, t); UI.sel = t.id;
    pendingFocus = { id: t.id, col: 'name', selectAll: true }; renderAll(); flushFocus(); return;
  }
  if (a === 'indent') {
    if (!needSel()) return; if (i === 0 || T[i - 1].level < T[i].level) { toast('Không thụt lề được: dòng phía trên phải cùng cấp hoặc sâu hơn.', true); return; }
    pushUndo(); for (let j = i; j <= st.end[i]; j++) T[j].level++; renderAll(); return;
  }
  if (a === 'outdent') {
    if (!needSel()) return; if (T[i].level === 0) return; pushUndo(); for (let j = i; j <= st.end[i]; j++) T[j].level--; renderAll(); return;
  }
  if (a === 'up' || a === 'down') {
    if (!needSel()) return; const lv = T[i].level, e = st.end[i];
    if (a === 'up') {
      let k = i - 1; while (k >= 0 && T[k].level > lv) k--; if (k < 0 || T[k].level !== lv) return;
      pushUndo(); const blk = T.splice(i, e - i + 1); T.splice(k, 0, ...blk);
    } else {
      const k = e + 1; if (k >= T.length || T[k].level !== lv) return; const e2 = st.end[k];
      pushUndo(); const blk = T.splice(i, e - i + 1); T.splice(e2 - (e - i + 1) + 1, 0, ...blk);
    }
    renderAll(); return;
  }
  if (a === 'dup') {
    if (!needSel()) return; pushUndo(); const e = st.end[i], map = new Map();
    const copies = T.slice(i, e + 1).map(t => { const c = JSON.parse(JSON.stringify(t)); c.id = newId(); map.set(t.id, c.id); return c; });
    copies.forEach(c => { c.preds = c.preds.map(l => map.has(l.id) ? { ...l, id: map.get(l.id) } : l); });
    copies[0].name += ' (bản sao)'; T.splice(e + 1, 0, ...copies); UI.sel = copies[0].id; renderAll(); toast('Đã nhân đôi ' + copies.length + ' dòng'); return;
  }
  if (a === 'del') {
    if (!needSel()) return; pushUndo(); const e = st.end[i], gone = new Set(T.slice(i, e + 1).map(t => t.id));
    T.splice(i, e - i + 1); T.forEach(t => { t.preds = (t.preds || []).filter(l => !gone.has(l.id)); });
    UI.sel = T[Math.min(i, T.length - 1)]?.id || null; renderAll(); toast(`Đã xóa ${gone.size} dòng · Ctrl+Z để hoàn tác`); return;
  }
  if (a === 'paste') return openPaste();
  if (a === 'goData') return scrollToDay(S.dataDay);
  if (a === 'fit') return fitGantt();
  if (a === 'baseline') {
    const has = P.tasks.some(t => t.bl);
    const doIt = () => { pushUndo(); setBaseline(P, S); renderAll(); toast('Đã chốt kế hoạch gốc'); };
    if (has) confirmDlg('Chốt lại kế hoạch gốc?', 'Kế hoạch gốc hiện tại sẽ bị thay bằng ngày bắt đầu/kết thúc đang hiển thị. Cột “Lệch KHG” và đường cong S sẽ so với mốc mới.', 'Chốt lại', doIt);
    else doIt();
  }
}
function scrollToDay(d) { if (!G) return; $('#gBody').scrollLeft = Math.max(0, G.x(d) - 160); }
function fitGantt() {
  const span = S.projFinish - S.projStart + 50, w = $('#gBody').clientWidth || 600;
  const need = w / span; UI.scale = need >= 22 ? 'day' : need >= 8 ? 'week' : 'month'; renderAll(); scrollToDay(S.projStart - 5);
}

/* ---------- cell editing ---------- */
function commitCell(inp) {
  const row = inp.closest('.tr'); if (!row) return; const id = row.dataset.id, i = idxById(id); if (i < 0) return;
  const t = P.tasks[i], col = inp.dataset.col, v = inp.value.trim();
  let change = null;
  if (col === 'name') { if (v !== t.name) change = () => { t.name = v; }; }
  else if (col === 'dur') {
    const n = Math.round(parseFloat(v.replace(',', '.')));
    if (!Number.isFinite(n) || n < 0 || n > 3000) { toast('Số ngày phải là số nguyên từ 0 trở lên (0 = mốc).', true); inp.value = inp.defaultValue; return; }
    if (n !== +t.dur) change = () => { t.dur = n; };
  } else if (col === 'start') {
    if (!v) { if (t.snet) change = () => { t.snet = null; }; }
    else {
      const d = parseDate(v, +P.start.slice(0, 4)); if (d == null) { toast('Ngày không hợp lệ. Gõ theo dạng dd/mm/yyyy.', true); inp.value = inp.defaultValue; return; }
      if (v !== inp.defaultValue) change = () => { t.snet = isoOf(d); };
    }
  } else if (col === 'preds') {
    const r = parsePreds(v, P.tasks, i);
    if (r.errs.length) { toast(`Không hiểu liên kết “${r.errs.join(', ')}”. Ví dụ đúng: 3, 5SS+2, 7FF-1`, true); inp.value = inp.defaultValue; return; }
    if (v !== inp.defaultValue) change = () => { t.preds = r.links; };
  } else if (col === 'res') { if (v !== (t.res || '')) change = () => { t.res = parseRes(v).map(x => x.code + ':' + x.qty).join(', '); }; }
  else if (col === 'pct') {
    let n = parseFloat(v.replace('%', '').replace(',', '.')); if (!Number.isFinite(n)) { inp.value = inp.defaultValue; return; }
    n = Math.max(0, Math.min(100, Math.round(n))); if (n !== +t.pct) change = () => { t.pct = n; };
  }
  if (!change) return;
  pushUndo(); change();
  setTimeout(() => {
    let f = pendingFocus; pendingFocus = null;
    const a = document.activeElement;
    if (!f && a && a.dataset && a.dataset.col && $('#lBody').contains(a)) f = { id: a.closest('.tr').dataset.id, col: a.dataset.col, range: [a.selectionStart, a.selectionEnd] };
    renderAll(); if (f) focusCell(f);
  }, 0);
}
function focusCell(f) {
  const row = document.querySelector(`#lBody .tr[data-id="${f.id}"]`); if (!row) return;
  const el = row.querySelector(`input[data-col="${f.col}"]`) || row.querySelector('input[data-col="name"]'); if (!el) return;
  el.focus({ preventScroll: false });
  try { if (f.range && f.range[0] != null) el.setSelectionRange(f.range[0], f.range[1]); else el.select(); } catch {}
}
function flushFocus() { if (pendingFocus) { const f = pendingFocus; pendingFocus = null; focusCell(f); } }

/* ---------- dialogs ---------- */
const dlg = $('#dlg');
function openDlg(title, body, foot, onMount) {
  dlg.innerHTML = `<div class="dlg-h"><h3>${title}</h3><button class="ib" data-close title="Đóng">${icon('x')}</button></div><div class="dlg-b">${body}</div>${foot ? `<div class="dlg-f">${foot}</div>` : ''}`;
  dlg.querySelectorAll('[data-close]').forEach(b => b.onclick = () => dlg.close());
  if (!dlg.open) dlg.showModal(); if (onMount) onMount(dlg);
}
function confirmDlg(title, msg, ok, fn) {
  openDlg(esc(title), `<p style="margin:0;line-height:1.6">${esc(msg)}</p>`, `<button class="tb" data-close>Hủy</button><button class="tb primary" id="cOk">${esc(ok)}</button>`,
    d => { d.querySelector('#cOk').onclick = () => { dlg.close(); fn(); }; });
}
function wwSeg(id, val) { return `<div class="seg" id="${id}">${[6, 7, 5].map(w => `<button type="button" data-ww="${w}" class="${+val === w ? 'on' : ''}">${w} ngày</button>`).join('')}</div>`; }
function bindSeg(d, id) { const seg = d.querySelector('#' + id); seg.onclick = e => { const b = e.target.closest('button'); if (!b) return; seg.querySelectorAll('button').forEach(x => x.classList.toggle('on', x === b)); }; return () => +seg.querySelector('.on').dataset.ww; }

function holToText(list) { return (list || []).map(h => { const a = fmtD(toDay(h.s)), b = h.f && h.f !== h.s ? ' - ' + fmtD(toDay(h.f)) : ''; return a + b + (h.note ? '  ' + h.note : ''); }).join('\n'); }
function textToHol(txt) {
  const out = [], bad = [];
  txt.split('\n').forEach(line => {
    const l = line.trim(); if (!l) return; const ds = l.match(/\d{4}-\d{2}-\d{2}|\d{1,2}\/\d{1,2}\/\d{4}/g);
    if (!ds) { bad.push(l); return; }
    const a = parseDate(ds[0]), b = ds[1] ? parseDate(ds[1]) : a; if (a == null || b == null) { bad.push(l); return; }
    const note = l.replace(/\d{4}-\d{2}-\d{2}|\d{1,2}\/\d{1,2}\/\d{4}/g, '').replace(/^[\s\-–.]+/, '').trim();
    out.push({ s: isoOf(Math.min(a, b)), f: isoOf(Math.max(a, b)), note });
  });
  return { out, bad };
}

function openSettings() {
  const R = P.resources || {};
  const resRows = Object.keys(R).map(k => `<tr><td><input data-r="code" value="${esc(k)}"></td><td><input data-r="name" value="${esc(R[k].name || '')}"></td><td><input data-r="unit" value="${esc(R[k].unit || '')}"></td><td><input data-r="limit" value="${R[k].limit || ''}" inputmode="numeric"></td></tr>`).join('');
  openDlg('Cấu hình dự án', `
    <div class="fld"><label for="sName">Tên công trình</label><input type="text" id="sName" value="${esc(P.name)}"></div>
    <div class="fld"><label for="sCon">Gói thầu / ghi chú</label><input type="text" id="sCon" value="${esc(P.contractor || '')}"></div>
    <div class="row2">
      <div class="fld"><label for="sStart">Ngày khởi công</label><input type="date" id="sStart" value="${esc(P.start)}"></div>
      <div class="fld"><span class="lb">Tuần làm việc</span>${wwSeg('sWW', P.workweek)}</div>
    </div>
    <label class="chk"><input type="checkbox" id="sDD" ${P.useDataDate !== false ? 'checked' : ''}> Dời phần việc chưa làm ra sau ngày báo cáo (cập nhật tiến độ theo thực tế)</label>
    <div class="fld"><label for="sHol">Ngày nghỉ lễ</label><textarea id="sHol" spellcheck="false">${esc(holToText(P.holidays))}</textarea>
      <div class="hint">Mỗi dòng một ngày hoặc một khoảng: <span class="code">04/02/2027 - 12/02/2027  Tết</span>. Lịch Tết có sẵn chỉ là gợi ý, hãy chỉnh theo lịch nghỉ của công trường.</div></div>
    <div class="fld"><span class="lb">Nguồn lực (mã dùng trong cột Nguồn lực)</span>
      <table class="rtab"><thead><tr><td class="lb">Mã</td><td class="lb">Tên</td><td class="lb">Đơn vị</td><td class="lb">Giới hạn/ngày</td></tr></thead><tbody id="sRes">${resRows}</tbody></table>
      <div><button class="tb" id="sAddRes" type="button">${icon('plus')}Thêm nguồn lực</button></div></div>
    <div class="fld"><span class="lb">Kế hoạch gốc</span><div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap"><span class="hint" style="font-size:12.5px">${P.baselineAt ? 'Đã chốt ngày ' + fmtD(toDay(P.baselineAt)) : 'Chưa chốt'}</span>${P.baselineAt ? `<button class="tb" id="sClrBl" type="button"><span class="danger-txt">Xóa kế hoạch gốc</span></button>` : ''}</div></div>`,
    `<button class="tb" data-close>Hủy</button><button class="tb primary" id="sOk">Lưu cấu hình</button>`,
    d => {
      const getWW = bindSeg(d, 'sWW');
      d.querySelector('#sAddRes').onclick = () => d.querySelector('#sRes').insertAdjacentHTML('beforeend', `<tr><td><input data-r="code" placeholder="VD: THEP"></td><td><input data-r="name"></td><td><input data-r="unit"></td><td><input data-r="limit" inputmode="numeric"></td></tr>`);
      const clr = d.querySelector('#sClrBl'); if (clr) clr.onclick = () => { pushUndo(); P.tasks.forEach(t => t.bl = null); P.baselineAt = null; dlg.close(); renderAll(); toast('Đã xóa kế hoạch gốc'); };
      d.querySelector('#sOk').onclick = () => {
        const st = d.querySelector('#sStart').value; if (!toDay(st)) { toast('Chọn ngày khởi công hợp lệ.', true); return; }
        const h = textToHol(d.querySelector('#sHol').value); if (h.bad.length) { toast('Dòng ngày lễ chưa đúng: ' + h.bad[0], true); return; }
        pushUndo();
        P.name = d.querySelector('#sName').value.trim() || 'Dự án mới'; P.contractor = d.querySelector('#sCon').value.trim();
        P.start = st; P.workweek = getWW(); P.useDataDate = d.querySelector('#sDD').checked; P.holidays = h.out;
        const res = {}; d.querySelectorAll('#sRes tr').forEach(tr => { const g = k => tr.querySelector(`[data-r="${k}"]`).value.trim(); const c = g('code').toUpperCase(); if (c) res[c] = { name: g('name'), unit: g('unit'), limit: +g('limit') || 0 }; });
        P.resources = res; dlg.close(); renderAll(); toast('Đã lưu cấu hình');
      };
    });
}

const WBS6 = ['Chuẩn bị mặt bằng', 'Phần ngầm (móng & hầm)', 'Kết cấu phần thân', 'Xây tô & bao che', 'Hệ thống cơ điện (MEP)', 'Hoàn thiện & bàn giao'];
function openNew() {
  openDlg('Khởi tạo dự án mới', `
    <div class="fld"><label for="nName">Tên công trình</label><input type="text" id="nName" value="Dự án mới" placeholder="VD: Nhà phố anh Minh – Q.7"></div>
    <div class="row2">
      <div class="fld"><label for="nStart">Ngày khởi công</label><input type="date" id="nStart" value="${isoOf(todayDay())}"></div>
      <div class="fld"><span class="lb">Tuần làm việc</span>${wwSeg('nWW', 6)}</div>
    </div>
    <div class="fld"><span class="lb">Bắt đầu từ</span>
      <label class="opt"><input type="radio" name="nTpl" value="wbs6" checked><div><b>Khung WBS 6 phần</b><span>${WBS6.join(' · ')}</span></div></label>
      <label class="opt"><input type="radio" name="nTpl" value="sample"><div><b>Mẫu nhà phố 1 trệt 3 lầu</b><span>39 dòng có liên kết, nguồn lực và % hoàn thành để tham khảo</span></div></label>
      <label class="opt"><input type="radio" name="nTpl" value="blank"><div><b>Bảng trống</b><span>Tự nhập, hoặc dùng Dán Excel / Mở tệp Excel sau đó</span></div></label>
    </div>
    <div class="hint" style="font-size:12px;color:var(--faint)">Dự án đang mở sẽ được thay thế. Bấm Hoàn tác (Ctrl+Z) để quay lại, hoặc Lưu .json trước khi tạo mới.</div>`,
    `<button class="tb" data-close>Hủy</button><button class="tb primary" id="nOk">Tạo dự án</button>`,
    d => {
      const getWW = bindSeg(d, 'nWW');
      d.querySelector('#nOk').onclick = () => {
        const tpl = d.querySelector('input[name=nTpl]:checked').value, st = d.querySelector('#nStart').value || isoOf(todayDay());
        const name = d.querySelector('#nName').value.trim() || 'Dự án mới';
        pushUndo();
        if (tpl === 'sample') { P = sampleProject(); P.name = name === 'Dự án mới' ? P.name : name; }
        else {
          let tasks = [];
          if (tpl === 'wbs6') {
            const rows = [[0, name.toUpperCase(), 0, '', '', 0]]; let last = null;
            WBS6.forEach((w, k) => { rows.push([1, w.toUpperCase(), 0, '', '', 0]); rows.push([2, 'Công việc ' + (k + 1) + '.1', 5, last ? String(last) : '', '', 0]); last = rows.length; });
            tasks = buildTasks(rows);
          }
          P = { v: 1, name, contractor: '', start: st, workweek: getWW(), dataDate: st, useDataDate: true, holidays: VN_HOLIDAYS.map(h => ({ ...h })), resources: JSON.parse(JSON.stringify(DEFAULT_RES)), tasks };
        }
        if (tpl !== 'sample') P.workweek = getWW();
        UI.collapsed = []; UI.sel = null; UI.filter = 'all'; dlg.close(); renderAll(); fitGantt(); toast('Đã tạo dự án mới');
      };
    });
}

function openPaste() {
  openDlg('Dán bảng từ Excel', `
    <div class="fld"><label for="pTxt">Copy vùng bảng trong Excel rồi dán vào đây (Ctrl+V)</label>
      <textarea id="pTxt" spellcheck="false" style="min-height:200px" placeholder="Tên công việc&#9;Ngày&#9;LK trước&#9;Nguồn lực"></textarea>
      <div class="hint">Nhận cột theo tiêu đề: <b>Tên công việc</b>, <b>Ngày</b> (thời gian), <b>LK trước</b> (theo STT trong bảng dán), <b>WBS</b> hoặc <b>Cấp</b> (để dựng cây), <b>Nguồn lực</b>, <b>%</b>. Không có tiêu đề thì hiểu là: Tên · Ngày · LK trước.</div></div>
    <label class="chk"><input type="radio" name="pMode" value="append" checked> Thêm vào cuối bảng</label>
    <label class="chk"><input type="radio" name="pMode" value="replace"> Thay toàn bộ bảng hiện tại</label>`,
    `<button class="tb" data-close>Hủy</button><button class="tb primary" id="pOk">Nhập</button>`,
    d => {
      setTimeout(() => d.querySelector('#pTxt').focus(), 30);
      d.querySelector('#pOk').onclick = () => {
        const txt = d.querySelector('#pTxt').value; const rows = txt.split(/\r?\n/).filter(l => l.trim()).map(l => l.split('\t'));
        const r = importRows(rows, d.querySelector('input[name=pMode]:checked').value === 'replace'); if (r) dlg.close();
      };
    });
}
function importRows(rows, replace) {
  if (!rows.length) { toast('Không có dữ liệu để nhập.', true); return false; }
  const norm = s => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').trim();
  let hi = -1, map = null;
  for (let r = 0; r < Math.min(rows.length, 8); r++) {
    const h = rows[r].map(norm); const find = (...keys) => h.findIndex(c => keys.some(k => c === k || c.startsWith(k)));
    const name = find('ten cong viec', 'ten', 'cong viec', 'noi dung', 'hang muc', 'task name', 'name');
    if (name >= 0) { hi = r; map = { name, dur: find('ngay', 'so ngay', 'thoi gian', 'tg', 'duration'), pred: find('lk', 'lien ket', 'tien nhiem', 'predecessor', 'cv truoc'), wbs: find('wbs', 'ma'), lvl: find('cap', 'level', 'outline'), res: find('nguon luc', 'resource'), pct: find('%', 'phan tram', 'hoan thanh', 'percent'), stt: find('stt', 'id', 'tt') }; break; }
  }
  if (!map) map = { name: 0, dur: 1, pred: 2, wbs: -1, lvl: -1, res: 3, pct: -1, stt: -1 };
  const body = rows.slice(hi + 1).filter(r => String(r[map.name] ?? '').trim());
  if (!body.length) { toast('Không tìm thấy dòng công việc nào.', true); return false; }
  const get = (r, k) => k >= 0 ? String(r[k] ?? '').trim() : '';
  const base = replace ? 0 : P.tasks.length;
  const sttMap = new Map(); body.forEach((r, k) => { const s = get(r, map.stt); if (s) sttMap.set(s, k); });
  const T = body.map(r => {
    let lvl = 1; const w = get(r, map.wbs), lv = get(r, map.lvl);
    if (lv && /^\d+$/.test(lv)) lvl = +lv; else if (w && /^\d+(\.\d+)*$/.test(w)) lvl = w.split('.').length - 1;
    else { const lead = String(r[map.name]).match(/^\s*/)[0].length; if (lead) lvl = 1 + Math.floor(lead / 2); }
    const durRaw = get(r, map.dur).replace(/[^\d.,]/g, '').replace(',', '.');
    return { id: newId(), level: lvl, name: get(r, map.name), dur: durRaw ? Math.round(parseFloat(durRaw)) : 1, preds: [], res: get(r, map.res), pct: Math.max(0, Math.min(100, Math.round(parseFloat(get(r, map.pct).replace('%', '').replace(',', '.')) || 0))), snet: null, bl: null };
  });
  let bad = 0;
  body.forEach((r, k) => {
    const s = get(r, map.pred); if (!s) return;
    s.split(/[,;]+/).forEach(tok => {
      const m = tok.replace(/\s+/g, '').toUpperCase().match(/^(\d+)(FS|SS|FF|SF)?([+-]\d+)?/); if (!m) { bad++; return; }
      const j = sttMap.size ? sttMap.get(m[1]) : (+m[1] - 1); if (j == null || j < 0 || j >= T.length || j === k) { bad++; return; }
      T[k].preds.push({ id: T[j].id, type: m[2] || 'FS', lag: m[3] ? parseInt(m[3], 10) : 0 });
    });
  });
  const minL = Math.min(...T.map(t => t.level)); T.forEach(t => t.level -= minL - (replace ? 0 : Math.min(1, minL)));
  pushUndo();
  if (replace) P.tasks = T; else P.tasks.push(...T);
  UI.filter = 'all'; renderAll(); toast(`Đã nhập ${T.length} dòng${bad ? ` · ${bad} liên kết không hiểu được đã bỏ qua` : ''}`);
  return true;
}

function openHelp() {
  openDlg('Phím tắt & cách nhập', `
    <div class="help-list">
      <span class="kbd">Enter</span><span>Lưu ô và xuống dòng dưới (Shift+Enter: lên)</span>
      <span class="kbd">↑ ↓</span><span>Di chuyển giữa các dòng khi đang sửa</span>
      <span class="kbd">Esc</span><span>Bỏ thay đổi trong ô</span>
      <span class="kbd">Insert</span><span>Thêm việc dưới dòng đang chọn (tự nối FS)</span>
      <span class="kbd">Delete</span><span>Xóa dòng đang chọn (khi không sửa ô)</span>
      <span class="kbd">Alt+Shift+→ / ←</span><span>Thụt lề / lùi lề (tạo hạng mục cha – con)</span>
      <span class="kbd">Alt+Shift+↑ / ↓</span><span>Đưa dòng lên / xuống</span>
      <span class="kbd">Ctrl+Z / Ctrl+Y</span><span>Hoàn tác / làm lại</span>
      <span class="kbd">Ctrl+S / Ctrl+O</span><span>Lưu tệp .json / mở tệp</span>
    </div>
    <div class="fld"><span class="lb">Cột LK trước (liên kết)</span><div class="hint" style="font-size:12.5px;color:var(--ink);line-height:1.7">
      Ghi theo số STT, cách nhau bằng dấu phẩy. <span class="code">3</span> = bắt đầu sau khi việc 3 xong (FS).
      <span class="code">5SS+2</span> = bắt đầu sau khi việc 5 bắt đầu 2 ngày. <span class="code">7FF-1</span> = kết thúc trước việc 7 một ngày. <span class="code">4FS+3</span> = chờ 3 ngày sau khi việc 4 xong (ví dụ chờ bê tông đủ cường độ).
      Số ngày tính theo ngày làm việc của tuần đã chọn, trừ ngày lễ.</div></div>
    <div class="fld"><span class="lb">Cột khác</span><div class="hint" style="font-size:12.5px;color:var(--ink);line-height:1.7">
      <b>Ngày = 0</b> là mốc. Gõ vào <b>Bắt đầu</b> để ghim ngày không sớm hơn, xóa trống ô để bỏ ghim.
      <b>Nguồn lực</b> dạng <span class="code">NC:8, MAY:1</span> (số lượng mỗi ngày). <b>Dự trữ 0d</b> là việc găng: trễ việc này là trễ cả công trình.
      <b>Lệch KHG</b> so ngày kết thúc với kế hoạch gốc đã chốt.</div></div>`, `<button class="tb primary" data-close>Đã hiểu</button>`);
}

/* ---------- files ---------- */
const safeName = s => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[đĐ]/g, 'd').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 60) || 'tien-do';
function offerFile(filename, data) {
  const blob = data instanceof Blob ? data : new Blob([data], { type: 'application/json' });
  const url = URL.createObjectURL(blob); const a = document.createElement('a');
  a.href = url; a.download = filename; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000); toast('Đã tải ' + filename);
}
function saveJson() { offerFile(safeName(P.name) + '.tiendo.json', JSON.stringify(P, null, 1)); }
function exportXlsx() {
  const T = P.tasks, aoa = [[P.name], [`Khởi công ${fmtD(S.projStart)} · Hoàn thành ${fmtD(S.projFinish)} · Ngày BC ${fmtD(S.dataDay)}`], [],
    ['STT', 'WBS', 'Tên công việc', 'Cấp', 'Ngày', 'Bắt đầu', 'Kết thúc', 'LK trước', 'Dự trữ', 'Găng', 'Nguồn lực', '% HT', 'KH gốc BĐ', 'KH gốc KT', 'Lệch KHG']];
  T.forEach((t, i) => {
    const sum = S.st.isSum[i];
    aoa.push([i + 1, S.st.wbs[i], '  '.repeat(t.level) + t.name, t.level, sum ? S.ef[i] - S.es[i] : +t.dur || 0, fmtD(S.start[i]), fmtD(S.finish[i]), fmtPreds(t, S.idIdx), sum ? '' : S.tf[i], !sum && S.crit[i] ? 'Có' : '', t.res || '', Math.round(S.pct[i]),
      S.blS[i] != null ? fmtD(S.blS[i]) : '', S.blF[i] != null ? fmtD(S.blF[i]) : '', S.vr[i] == null ? '' : S.vr[i]]);
  });
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [5, 8, 44, 5, 6, 11, 11, 14, 7, 6, 16, 6, 11, 11, 9].map(w => ({ wch: w }));
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Tien do');
  const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  offerFile(safeName(P.name) + '.xlsx', new Blob([buf]));
}
function copyText(ta) {
  const v = ta.value;
  if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(v).then(() => toast('Đã sao chép'), () => { ta.focus(); ta.select(); toast('Đã chọn sẵn nội dung, bấm Ctrl+C để sao chép'); });
  else { ta.focus(); ta.select(); toast('Đã chọn sẵn nội dung, bấm Ctrl+C để sao chép'); }
}
$('#fileIn').addEventListener('change', async e => {
  const f = e.target.files[0]; e.target.value = ''; if (!f) return;
  try {
    if (/\.json$/i.test(f.name)) {
      const p = JSON.parse(await f.text()); if (!p || !Array.isArray(p.tasks)) throw new Error('fmt');
      pushUndo(); P = normalize(p); UI.collapsed = []; UI.sel = null; renderAll(); fitGantt(); toast('Đã mở ' + f.name);
    } else {
          const wb = XLSX.read(await f.arrayBuffer(), { type: 'array' }); const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
      openDlg('Nhập từ Excel', `<p style="margin:0;line-height:1.6">Đọc được <b>${rows.length}</b> dòng từ trang “${esc(wb.SheetNames[0])}”. Thêm vào bảng hiện tại hay thay thế toàn bộ?</p>`,
        `<button class="tb" data-close>Hủy</button><button class="tb" id="xAppend">Thêm vào cuối</button><button class="tb primary" id="xReplace">Thay thế</button>`,
        d => { d.querySelector('#xAppend').onclick = () => { if (importRows(rows, false)) dlg.close(); }; d.querySelector('#xReplace').onclick = () => { if (importRows(rows, true)) { dlg.close(); fitGantt(); } }; });
    }
  } catch { toast('Không đọc được tệp này. Chọn tệp .tiendo.json hoặc .xlsx.', true); }
});

/* ---------- events ---------- */
$('#bUndo').innerHTML = icon('undo'); $('#bRedo').innerHTML = icon('redo');
$('#bNew').innerHTML = icon('file') + '<span class="lbl">Dự án mới</span>';
$('#bOpen').innerHTML = icon('folder') + '<span class="lbl">Mở tệp</span>';
$('#bSave').innerHTML = icon('save') + '<span class="lbl">Lưu .json</span>';
$('#bXlsx').innerHTML = icon('sheet') + '<span class="lbl">Xuất Excel</span>';
$('#bSettings').innerHTML = icon('gear'); $('#bHelp').innerHTML = icon('help');
$('#bUndo').onclick = undo; $('#bRedo').onclick = redo; $('#bNew').onclick = openNew; $('#bOpen').onclick = () => $('#fileIn').click();
$('#bSave').onclick = saveJson; $('#bXlsx').onclick = exportXlsx; $('#bSettings').onclick = openSettings; $('#projBtn').onclick = openSettings; $('#bHelp').onclick = openHelp;

$('#tabs').addEventListener('click', e => {
  const b = e.target.closest('button[data-view]'); if (!b) return; UI.view = b.dataset.view;
  document.querySelectorAll('#tabs button').forEach(x => x.classList.toggle('on', x === b));
  $('#vPlan').hidden = UI.view !== 'plan'; $('#vReport').hidden = UI.view !== 'report'; renderAll();
});
$('#ribbon').addEventListener('click', e => {
  const b = e.target.closest('button'); if (!b) return;
  if (b.dataset.act) act(b.dataset.act);
  else if (b.dataset.filter) { UI.filter = b.dataset.filter; renderAll(); }
  else if (b.dataset.scale) { const mid = $('#gBody').scrollLeft + $('#gBody').clientWidth / 2, d = G ? G.rs + mid / G.px : S.dataDay; UI.scale = b.dataset.scale; renderAll(); $('#gBody').scrollLeft = Math.max(0, G.x(Math.round(d)) - $('#gBody').clientWidth / 2); }
});
$('#ribbon').addEventListener('change', e => {
  if (e.target.id === 'dataDate') { const v = e.target.value; if (!toDay(v)) return; pushUndo(); P.dataDate = v; renderAll(); }
});
const lBody = $('#lBody'), gBody = $('#gBody');
lBody.addEventListener('change', e => { const inp = e.target.closest('input[data-col]'); if (inp) commitCell(inp); });
lBody.addEventListener('focusin', e => { const row = e.target.closest('.tr'); if (row && row.dataset.id !== UI.sel) select(row.dataset.id); });
lBody.addEventListener('click', e => {
  const tog = e.target.closest('[data-tog]');
  if (tog) { const id = tog.dataset.tog; UI.collapsed = UI.collapsed.includes(id) ? UI.collapsed.filter(x => x !== id) : [...UI.collapsed, id]; renderAll(); return; }
  const row = e.target.closest('.tr'); if (row) select(row.dataset.id);
});
lBody.addEventListener('keydown', e => {
  const inp = e.target.closest('input[data-col]'); if (!inp) return;
  if (e.key === 'Escape') { inp.value = inp.defaultValue; inp.blur(); e.stopPropagation(); return; }
  if (e.key === 'Enter' || ((e.key === 'ArrowDown' || e.key === 'ArrowUp') && !e.altKey)) {
    e.preventDefault(); const dir = (e.key === 'ArrowUp' || (e.key === 'Enter' && e.shiftKey)) ? -1 : 1;
    const rows = [...lBody.querySelectorAll('.tr')], k = rows.indexOf(inp.closest('.tr')) + dir;
    pendingFocus = rows[k] ? { id: rows[k].dataset.id, col: inp.dataset.col } : null;
    inp.blur(); setTimeout(flushFocus, 0);
  }
});
lBody.addEventListener('scroll', () => { gBody.scrollTop = lBody.scrollTop; $('#lHead').scrollLeft = lBody.scrollLeft; });
gBody.addEventListener('scroll', () => { lBody.scrollTop = gBody.scrollTop; $('#gHead').scrollLeft = gBody.scrollLeft; $('#resChart').scrollLeft = gBody.scrollLeft; });
gBody.addEventListener('click', e => { const r = Math.floor((e.clientY - gBody.getBoundingClientRect().top + gBody.scrollTop) / RH); if (VIS[r] != null) select(P.tasks[VIS[r]].id); });
$('#resToggle').onclick = () => { UI.resOpen = !UI.resOpen; renderRes(); persist(); };
$('#res').addEventListener('change', e => {
  if (e.target.id === 'resSel') { UI.res = e.target.value; renderRes(); persist(); }
  if (e.target.id === 'resLim') { pushUndo(); P.resources = P.resources || {}; const r = P.resources[UI.res] || (P.resources[UI.res] = { name: '', unit: '', limit: 0 }); r.limit = +e.target.value || 0; renderAll(); }
});
document.addEventListener('click', e => { if (e.target.id === 'copyZalo') copyText($('#zaloTxt')); });

// splitter
(() => {
  const sp = $('#splitter'); let x0 = 0, w0 = 0;
  const mv = e => { UI.lw = Math.max(240, Math.min(w0 + e.clientX - x0, $('#board').clientWidth - 200)); $('#board').style.setProperty('--lw', UI.lw + 'px'); };
  const up = () => { sp.classList.remove('drag'); document.removeEventListener('pointermove', mv); document.removeEventListener('pointerup', up); renderGantt(); renderRes(); persist(); };
  sp.addEventListener('pointerdown', e => { x0 = e.clientX; w0 = UI.lw; sp.classList.add('drag'); document.addEventListener('pointermove', mv); document.addEventListener('pointerup', up); e.preventDefault(); });
})();

document.addEventListener('keydown', e => {
  const k = e.key, ctrl = e.ctrlKey || e.metaKey, inField = e.target.matches('input,textarea,select');
  if (dlg.open) return;
  if (ctrl && !e.shiftKey && k.toLowerCase() === 'z') { if (inField && e.target.value !== e.target.defaultValue) return; e.preventDefault(); undo(); return; }
  if (ctrl && (k.toLowerCase() === 'y' || (e.shiftKey && k.toLowerCase() === 'z'))) { e.preventDefault(); redo(); return; }
  if (ctrl && k.toLowerCase() === 's') { e.preventDefault(); saveJson(); return; }
  if (ctrl && k.toLowerCase() === 'o') { e.preventDefault(); $('#fileIn').click(); return; }
  if (k === 'F1') { e.preventDefault(); openHelp(); return; }
  if (UI.view !== 'plan') return;
  if (k === 'Insert') { e.preventDefault(); if (inField) e.target.blur(); setTimeout(() => act('addBelow'), 0); return; }
  if (e.altKey && e.shiftKey && ['ArrowRight', 'ArrowLeft', 'ArrowUp', 'ArrowDown'].includes(k)) {
    e.preventDefault(); if (inField) e.target.blur();
    setTimeout(() => act({ ArrowRight: 'indent', ArrowLeft: 'outdent', ArrowUp: 'up', ArrowDown: 'down' }[k]), 0); return;
  }
  if (k === 'Delete' && !inField && UI.sel) { e.preventDefault(); act('del'); }
});
let rzT; window.addEventListener('resize', () => { clearTimeout(rzT); rzT = setTimeout(() => { if (UI.view === 'plan') { renderGantt(); renderRes(); } }, 150); });

/* ---------- boot ---------- */
P = loadInitial();
renderHead();
renderAll();
requestAnimationFrame(() => scrollToDay(S.projStart - 3));
if (P.sample) setTimeout(() => toast('Đang mở dự án mẫu. Bấm “Dự án mới” để bắt đầu công trình của bạn.'), 400);
