import './styles.css';
import * as XLSX from 'xlsx';
import { DAY, DEFAULT_RES, VN_HOLIDAYS, _idn, buildCal, buildTasks, dow, expandHolidays, fmtD, fmtDM, fmtPreds, isoOf, isoWeek, mkDay, mondayOf, newId, pad, parseDate, parsePreds, parseRes, sampleProject, schedule, setBaseline, structure, toDay } from './core.js';
import { cashFlow, evm, defaultPayments, lookAhead, snapshot } from './core.js';
import { normalizeBuilding, buildElements, linksFor, progressAt } from './model3d.js';
import { parseQuote, quoteSections, scheduleRowsFromQuote, matchSections, estimateDays, niceName, num as vnNum, fmtM } from './boq.js';
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
  zoom: '<path d="M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5"/>',
  dup: '<rect x="4" y="4" width="11" height="11" rx="2"/><rect x="9" y="9" width="11" height="11" rx="2"/>',
  cut: '<circle cx="6" cy="6" r="3"/><circle cx="6" cy="18" r="3"/><path d="M20 4 8.1 15.9M14.5 14.5 20 20M8.1 8.1 12 12"/>',
  above: '<path d="M12 19V9M7 14l5-5 5 5M5 4h14"/>', group: '<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2Z"/><path d="M12 10v6M9 13h6"/>',
  fold: '<path d="m7 15 5-5 5 5M7 20l5-5 5 5M5 4h14"/>', unfold: '<path d="m7 9 5 5 5-5M7 4l5 5 5-5M5 20h14"/>',
  coins: '<ellipse cx="12" cy="6" rx="7" ry="3"/><path d="M5 6v6c0 1.7 3.1 3 7 3s7-1.3 7-3V6M5 12v6c0 1.7 3.1 3 7 3s7-1.3 7-3v-6"/>',
  check: '<path d="M9 11l3 3 8-8"/><path d="M20 12v6a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h9"/>'
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
    <div class="grp"><span class="gl">Bảng tạm</span>
      <button class="rb" data-act="cut" title="Cắt dòng đã chọn (Ctrl+X)" aria-label="Cắt">${icon('cut')}</button>
      <button class="rb" data-act="copy" title="Chép dòng đã chọn, dán được sang Excel (Ctrl+C)" aria-label="Chép">${icon('copy')}</button>
      <button class="rb" data-act="pasteRows" title="Dán dòng đã chép vào dưới dòng đang chọn (Ctrl+V). Ctrl+V cũng dán được bảng copy từ Excel" aria-label="Dán">${icon('paste')}</button>
    </div>
    <div class="grp"><span class="gl">Thêm</span>
      <button class="rb" data-act="addBelow" title="Thêm công việc bên dưới, tự nối FS với việc đang chọn (Insert)">${icon('plus')}Việc dưới</button>
      <button class="rb" data-act="addAbove" title="Thêm công việc phía trên dòng đang chọn">${icon('above')}Việc trên</button>
      <button class="rb" data-act="addChild" title="Thêm việc con của dòng đang chọn">${icon('child')}Việc con</button>
      <button class="rb" data-act="addGroup" title="Thêm hạng mục tổng hợp (kèm một việc con)">${icon('group')}Hạng mục</button>
      <button class="rb" data-act="addMs" title="Thêm mốc (0 ngày)">${icon('diamond')}Mốc</button>
      <button class="rb" data-act="paste" title="Dán bảng copy từ Excel">${icon('sheet')}Dán Excel</button>
    </div>
    <div class="grp"><span class="gl">Cấu trúc</span>
      <button class="rb" data-act="indent" title="Thụt lề (Alt+Shift+→)" aria-label="Thụt lề">${icon('indent')}</button>
      <button class="rb" data-act="outdent" title="Lùi lề (Alt+Shift+←)" aria-label="Lùi lề">${icon('outdent')}</button>
      <button class="rb" data-act="up" title="Lên trên (Alt+↑)" aria-label="Lên trên">${icon('up')}</button>
      <button class="rb" data-act="down" title="Xuống dưới (Alt+↓)" aria-label="Xuống dưới">${icon('down')}</button>
      <button class="rb" data-act="dup" title="Nhân đôi dòng đã chọn, kèm việc con (Ctrl+D)" aria-label="Nhân đôi">${icon('dup')}</button>
      <button class="rb danger" data-act="del" title="Xóa dòng đã chọn (Delete)" aria-label="Xóa">${icon('trash')}</button>
    </div>
    <div class="grp"><span class="gl">Xem</span>
      <button class="rb" data-act="collapseAll" title="Thu gọn hết các hạng mục (Alt+-)" aria-label="Thu gọn hết">${icon('fold')}</button>
      <button class="rb" data-act="expandAll" title="Mở rộng hết (Alt+=)" aria-label="Mở rộng hết">${icon('unfold')}</button>
      <button class="rb ${showMoney() ? 'on' : ''}" data-act="money" title="Hiện / ẩn cột Giá trị và Chi phí">${icon('coins')}Tiền</button>
      <button class="rb" data-act="check" title="Kiểm tra dữ liệu: liên kết hở, vòng lặp, % sai trình tự">${icon('check')}Kiểm tra</button>
    </div>
    <div class="grp"><span class="gl">Lọc</span>${f('all', 'Tất cả')}${f('crit', 'Đường găng', nCrit)}${f('late', 'Chậm', nLate)}${f('week', '2 tuần tới')}</div>
    <div class="grp"><span class="gl">Thang đo</span>${sc('day')}${sc('week')}${sc('month')}</div>
    <div class="grp"><span class="gl">Ngày BC</span>
      <input type="date" id="dataDate" value="${esc(P.dataDate)}" title="Ngày báo cáo (data date): mốc để tính kế hoạch và dời việc chưa làm">
      <button class="rb" data-act="goData" title="Cuộn Gantt tới ngày báo cáo">${icon('target')}</button>
      <button class="rb" data-act="fit" title="Vừa màn hình">${icon('zoom')}</button>
      <button class="rb" data-act="baseline" title="Chốt kế hoạch gốc để so sánh">${icon('flag')}Lưu KH gốc</button>
      <button class="rb" data-act="weekly" title="Cập nhật % hoàn thành, ghi chú và chuyển ngày báo cáo">${icon('check')}Cập nhật tuần</button>
      <button class="rb" data-act="meeting" title="Trình chiếu họp giao ban toàn màn hình">${icon('target')}Họp giao ban</button>
    </div>`;
}

/* ---------- table ---------- */
const COLS = [['STT', 'c'], ['WBS', ''], ['Tên công việc', ''], ['Ngày', 'c'], ['Bắt đầu', ''], ['Kết thúc', ''], ['LK trước', ''], ['Dự trữ', 'c'], ['Nguồn lực', ''], ['%', 'c'], ['Lệch KHG', 'c'], ['Giá trị (tr)', 'r'], ['Chi phí (tr)', 'r']];
const showMoney = () => UI.money != null ? UI.money : !!(S && (S.totalVal || S.totalCost));
const fmtTr = v => v ? (v / 1e6).toLocaleString('vi-VN', { maximumFractionDigits: 1 }) : '';
const COL_TIPS = ['Số thứ tự – dùng để khai báo liên kết', 'Mã WBS tự đánh theo cấp', '', 'Thời gian (ngày làm việc)', 'Gõ ngày để ghim “bắt đầu không sớm hơn”', 'Tính tự động', 'Ví dụ: 3, 5SS+2, 7FF-1', 'Dự trữ toàn phần (ngày làm việc)', 'Ví dụ: NC:8, MAY:1', '% hoàn thành thực tế', 'Chênh lệch ngày kết thúc so với kế hoạch gốc', 'Giá bán / giá trị hợp đồng của việc, đơn vị triệu đồng', 'Giá vốn / chi phí dự kiến, đơn vị triệu đồng'];
function renderHead() {
  const m = showMoney(), cols = m ? COLS : COLS.slice(0, 11);
  $('#lHead').innerHTML = `<div class="grid ${m ? 'money' : ''}">${cols.map((c, k) => `<div class="th ${c[1]}" title="${esc(COL_TIPS[k])}">${c[0]}</div>`).join('')}</div>`;
}
function renderTable() {
  const T = P.tasks, st = S.st, col = new Set(UI.collapsed), money = showMoney();
  if (!T.length) {
    $('#lBody').innerHTML = `<div class="empty"><b>Chưa có công việc nào.</b><br>Bấm <b>Việc dưới</b> để thêm từng dòng, <b>Dán Excel</b> để dán bảng có sẵn, hoặc <b>Dự án mới</b> để tạo khung WBS 6 phần.</div>`;
    return;
  }
  let h = '';
  for (const i of VIS) {
    const t = T[i], sum = st.isSum[i], ms = !sum && (+t.dur || 0) === 0;
    const crit = S.crit[i] && !sum, sel = SELS.has(t.id), cur = sel && SELS.size > 1 && t.id === UI.sel;
    const tf = S.tf[i];
    const tfTxt = sum ? '<span class="muted">—</span>' : (tf <= 0 ? `<span class="tf-crit">0d</span>` : `<span class="ro">+${tf}d</span>`);
    const v = S.vr[i];
    const vTxt = v == null ? '<span class="muted">—</span>' : v > 0 ? `<span class="var-bad">+${v}d</span>` : v < 0 ? `<span class="var-good">−${-v}d</span>` : '<span class="muted">0d</span>';
    const pv = Math.round(S.pct[i]);
    h += `<div class="tr grid ${money ? 'money' : ''} ${sum ? 'sum' : ''} ${sel ? 'sel' : ''} ${cur ? 'cur' : ''}" data-id="${t.id}" data-i="${i}">
      <div class="td c c-stt" title="Chọn dòng">${i + 1}</div>
      <div class="td c-wbs">${st.wbs[i]}</div>
      <div class="td c-name" style="--lvl:${t.level}">${sum ? `<button class="chev ${col.has(t.id) ? '' : 'open'}" data-tog="${t.id}" title="Thu gọn / mở rộng">${icon('chev')}</button>` : '<span class="chev-sp"></span>'}${ms ? '<span class="dia"></span>' : ''}${crit ? '<span class="badge crit">GĂNG</span>' : (!sum && S.late[i] ? '<span class="badge late">CHẬM</span>' : '')}<input data-col="name" value="${esc(t.name)}" aria-label="Tên công việc">${t.note ? `<span class="note-dot" title="${esc(t.note)}">!</span>` : ''}</div>
      <div class="td c num">${sum ? `<span class="ro">${S.ef[i] - S.es[i]}</span>` : `<input data-col="dur" value="${+t.dur || 0}" inputmode="numeric" aria-label="Số ngày">`}</div>
      <div class="td num">${sum ? `<span class="ro">${fmtD(S.start[i])}</span>` : `<input data-col="start" value="${fmtD(S.start[i])}" aria-label="Bắt đầu">${t.snet ? `<span title="Ghim bắt đầu không sớm hơn ${fmtD(toDay(t.snet))} – xóa ô để bỏ ghim">${icon('pin', 'pin')}</span>` : ''}`}</div>
      <div class="td"><span class="ro">${fmtD(S.finish[i])}</span></div>
      <div class="td num" ${S.badLink[i] ? 'title="Có liên kết không hợp lệ (tự liên kết hoặc nối với việc cha/con)"' : ''}><input data-col="preds" value="${esc(fmtPreds(t, S.idIdx))}" aria-label="Liên kết trước" ${S.badLink[i] || S.cycle[i] ? 'style="color:var(--crit)"' : ''}></div>
      <div class="td c num">${tfTxt}</div>
      <div class="td num">${sum ? '' : `<input data-col="res" value="${esc(t.res || '')}" aria-label="Nguồn lực">`}</div>
      <div class="td c num">${sum ? `<span class="ro">${pv}%</span>` : `<input data-col="pct" value="${pv}" inputmode="numeric" aria-label="Phần trăm hoàn thành">`}<span class="pbar"><i style="width:${pv}%"></i></span></div>
      <div class="td c num">${vTxt}</div>${money ? (sum ? `<div class="td r num"><span class="ro">${fmtTr(S.val[i])}</span></div><div class="td r num"><span class="ro">${fmtTr(S.cost[i])}</span></div>`
        : `<div class="td r num"><input data-col="val" value="${fmtTr(t.val)}" inputmode="decimal" aria-label="Giá trị (triệu)"></div><div class="td r num"><input data-col="cost" value="${fmtTr(t.cost)}" inputmode="decimal" aria-label="Chi phí (triệu)"></div>`) : ''}
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
  $('#status').innerHTML = `<span id="selCount" class="selc">${SELS.size > 1 ? `Đang chọn ${SELS.size} dòng` : ''}</span><span><b>${n}</b> việc</span><span class="crit"><b class="crit">${c}</b> việc găng</span>
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
  const pts = []; const W = S.leaves.reduce((s, i) => s + (S.hasMoney ? S.val[i] : Math.max(+T[i].dur || 0, 0)), 0) || 1;
  const fracPlan = (i, d) => { const s0 = S.blS[i] ?? S.start[i], f0 = S.blF[i] ?? S.finish[i]; if (d < s0) return 0; if (d > f0) return 1; return (d - s0 + 1) / (f0 - s0 + 1); };
  const fracFc = (i, d) => { const p = S.pct[i] / 100; const s0 = Math.max(S.start[i], S.dataDay), f0 = S.finish[i]; if (p >= 1) return 1; if (d < s0) return p; if (d > f0) return 1; return p + (1 - p) * (d - s0 + 1) / (f0 - s0 + 1); };
  for (let d = a; d <= b + 7; d += 7) {
    const dd = Math.min(d, b);
    let pl = 0, fc = 0; S.leaves.forEach(i => { const w = (S.hasMoney ? S.val[i] : Math.max(+T[i].dur || 0, 0)); pl += w * fracPlan(i, dd); fc += w * fracFc(i, dd); });
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
  const hist = (P.history || []).map(h => ({ d: toDay(h.date), v: h.actualPct })).filter(h => h.d != null && h.d >= sc.a && h.d <= sc.b && h.d <= D);
  if (hist.length) svg += `<path class="s-hist" d="M${sx(sc.a).toFixed(1)} ${sy(0)}${hist.map(h => 'L' + sx(h.d).toFixed(1) + ' ' + sy(h.v).toFixed(1)).join('')}"/>` + hist.map(h => `<circle class="s-hpt" cx="${sx(h.d)}" cy="${sy(h.v)}" r="3"><title>${fmtD(h.d)}: ${h.v.toFixed(1)}%</title></circle>`).join('');
  svg += '</svg>';
  const row = i => `<tr><td>${esc(T[i].name)}${S.crit[i] ? ' <span class="badge crit">GĂNG</span>' : ''}</td><td class="n">${fmtDM(S.start[i])} → ${fmtDM(S.finish[i])}</td><td class="n">${Math.round(S.pct[i])}% / ${Math.round(S.planned[i])}%</td></tr>`;
  const tbl = (list, empty) => list.length ? `<div class="tbl-wrap"><table class="rt"><thead><tr><th>Công việc</th><th>Thời gian</th><th>TT / KH</th></tr></thead><tbody>${list.map(row).join('')}</tbody></table></div>` : `<div class="none">${empty}</div>`;
  $('#report').innerHTML = `<div class="rep-in">
    <div class="rep-head"><div><h2>${esc(P.name)}</h2><p>Báo cáo tại ngày <b>${fmtD(D)}</b>${P.baselineAt ? ` · so với kế hoạch gốc chốt ngày ${fmtD(toDay(P.baselineAt))}` : ' · chưa chốt kế hoạch gốc'}</p></div>
      <div class="rep-act"><button class="tb" data-top="weekly">${icon('check')}Cập nhật tuần</button><button class="tb primary" data-top="meeting">${icon('target')}Họp giao ban</button></div></div>
    <div class="kpis">
      <div class="kpi"><div class="k">Hoàn thành thực tế</div><div class="v">${S.actualPct.toFixed(1)}%</div><div class="s">Kế hoạch ${S.plannedPct.toFixed(1)}% ${pill(gap)}</div></div>
      <div class="kpi"><div class="k">Dự kiến hoàn thành</div><div class="v">${fmtD(S.projFinish)}</div><div class="s">${pfVar == null ? 'Chưa có kế hoạch gốc' : `So với KH gốc ${pill(pfVar, ' ngày', true)}`}</div></div>
      <div class="kpi"><div class="k">Việc đang chậm</div><div class="v" style="color:${lateL.length ? 'var(--crit)' : 'var(--ok)'}">${lateL.length}</div><div class="s">trên ${S.leaves.length} công việc</div></div>
      <div class="kpi"><div class="k">Việc găng còn lại</div><div class="v">${critLeft}</div><div class="s">${S.workDays} ngày làm việc · tuần ${P.workweek} ngày</div></div>
    </div>
    ${moneySection()}
    <div class="two">
      <div class="card"><h3>Đường cong S <span class="pill neu">trọng số theo ${S.hasMoney ? 'giá trị' : 'thời lượng'}</span></h3>${svg}
        <div class="s-key"><span class="lg"><i class="sw" style="background:var(--faint);height:3px"></i>Kế hoạch ${P.baselineAt ? 'gốc' : 'hiện tại'}</span><span class="lg"><i class="sw" style="border-top:2px dashed var(--accent-strong);height:0;border-radius:0"></i>Dự báo từ ngày BC</span><span class="lg"><i class="sw" style="background:var(--ok);width:8px;height:8px;border-radius:50%"></i>Thực tế</span></div></div>
      <div class="card"><h3>Tin nhắn báo cáo nhanh</h3><textarea class="zalo" id="zaloTxt" aria-label="Nội dung báo cáo">${esc(zaloText(lateL, doing, ahead, gap, pfVar))}</textarea>
        <div style="display:flex;gap:8px;margin-top:8px;flex-wrap:wrap"><button class="tb primary" id="copyZalo">${icon('copy')}Sao chép</button></div></div>
    </div>
    <div class="two">
      <div class="card"><h3>Việc đang chậm <span class="pill ${lateL.length ? 'bad' : 'good'}">${lateL.length}</span></h3>${tbl(lateL, 'Không có việc chậm so với kế hoạch tại ngày báo cáo.')}</div>
      <div class="card"><h3>Đang thi công</h3>${tbl(doing, 'Không có việc nào đang thi công.')}</div>
    </div>
    ${lookAheadCard()}
  </div>`;
}
function zaloText(lateL, doing, ahead, gap, pfVar) {
  const T = P.tasks; const L = [];
  L.push(`BÁO CÁO TIẾN ĐỘ – ${P.name.toUpperCase()}`);
  L.push(`Ngày báo cáo: ${fmtD(S.dataDay)}`);
  L.push(`• Hoàn thành: ${S.actualPct.toFixed(1)}% (kế hoạch ${S.plannedPct.toFixed(1)}%, ${gap >= 0 ? 'vượt' : 'chậm'} ${Math.abs(gap).toFixed(1)}%)`);
  if (S.hasMoney) { const e = evm(P, S); L.push(`• Giá trị đã thực hiện: ${fmtTy(S.leaves.reduce((a, i) => a + S.val[i] * S.pct[i] / 100, 0))} / ${fmtTy(S.totalVal)}${e.SPI != null ? ` (SPI ${e.SPI.toFixed(2)})` : ''}`); }
  L.push(`• Dự kiến hoàn thành: ${fmtD(S.projFinish)}${pfVar == null ? '' : pfVar > 0 ? ` (trễ ${pfVar} ngày so với KH gốc)` : pfVar < 0 ? ` (sớm ${-pfVar} ngày so với KH gốc)` : ' (đúng KH gốc)'}`);
  if (doing.length) { L.push('', 'ĐANG THI CÔNG:'); doing.slice(0, 8).forEach(i => L.push(`- ${T[i].name}: ${Math.round(S.pct[i])}%`)); }
  if (lateL.length) { L.push('', 'VIỆC CHẬM CẦN XỬ LÝ:'); lateL.slice(0, 8).forEach(i => L.push(`- ${T[i].name}: TT ${Math.round(S.pct[i])}% / KH ${Math.round(S.planned[i])}%${T[i].note ? ` (${T[i].note})` : ''}`)); }
  if (ahead.length) { L.push('', '2 TUẦN TỚI:'); ahead.slice(0, 8).forEach(i => L.push(`- ${T[i].name} (${fmtDM(S.start[i])} – ${fmtDM(S.finish[i])})`)); }
  return L.join('\n');
}

/* ---------- render all ---------- */
function renderAll(opts = {}) {
  S = schedule(P, P.useDataDate !== false);
  computeVisible(); syncSel();
  $('#projName').textContent = P.name || 'Dự án chưa đặt tên';
  $('#projSub').textContent = `${S.leaves.length} công việc · ${P.contractor || 'Tuần ' + P.workweek + ' ngày'}`;
  $('#bUndo').disabled = !undoStack.length; $('#bRedo').disabled = !redoStack.length;
  $('#board').style.setProperty('--lw', UI.lw + 'px');
  renderRibbon();
  if (UI.view === 'plan') { renderHead(); renderTable(); renderGantt(); renderRes(); renderStatus(); }
  else if (UI.view === '3d') render3d();
  else renderReport();
  persist();
}

/* ---------- selection ---------- */
const idxById = id => P.tasks.findIndex(t => t.id === id);
function selIdx() { const i = UI.sel ? idxById(UI.sel) : -1; return i; }
let SELS = new Set(), anchorId = null;
function syncSel() {
  const ids = new Set(P.tasks.map(t => t.id));
  SELS = new Set([...SELS].filter(id => ids.has(id)));
  if (UI.sel && !ids.has(UI.sel)) UI.sel = null;
  if (UI.sel && !SELS.has(UI.sel)) SELS = new Set([UI.sel]);
  if (!UI.sel && SELS.size) UI.sel = [...SELS][0];
}
function select(id, mode = 'single') {
  if (mode === 'toggle') { if (SELS.has(id) && SELS.size > 1) { SELS.delete(id); if (UI.sel === id) UI.sel = [...SELS][0]; } else { SELS.add(id); UI.sel = id; } anchorId = id; }
  else if (mode === 'range' && anchorId) {
    const a = VIS.findIndex(i => P.tasks[i].id === anchorId), b = VIS.findIndex(i => P.tasks[i].id === id);
    if (a >= 0 && b >= 0) SELS = new Set(VIS.slice(Math.min(a, b), Math.max(a, b) + 1).map(i => P.tasks[i].id)); else SELS = new Set([id]);
    UI.sel = id;
  } else { SELS = new Set([id]); anchorId = id; UI.sel = id; }
  paintSel();
}
function paintSel() {
  document.querySelectorAll('#lBody .tr').forEach(r => { r.classList.toggle('sel', SELS.has(r.dataset.id)); r.classList.toggle('cur', r.dataset.id === UI.sel && SELS.size > 1); });
  const r = VIS.findIndex(i => P.tasks[i].id === UI.sel), g = document.getElementById('gSel');
  if (g) { if (r < 0) g.setAttribute('visibility', 'hidden'); else { g.removeAttribute('visibility'); g.setAttribute('y', r * RH); } }
  const n = $('#selCount'); if (n) n.textContent = SELS.size > 1 ? `Đang chọn ${SELS.size} dòng` : '';
}
// Top-level blocks [start, end] covering the selected rows (a summary carries its children)
function selectedBlocks() {
  const st = S.st, idx = [...SELS].map(idxById).filter(i => i >= 0).sort((a, b) => a - b), out = []; let last = -1;
  for (const i of idx) { if (i <= last) continue; out.push([i, st.end[i]]); last = st.end[i]; }
  return out;
}
// Where a new row goes: below the selected row (as first child when it is an open summary)
function insertionPoint() {
  const T = P.tasks, st = S.st, i = selIdx();
  if (i < 0) return { pos: T.length, lvl: T.length ? Math.max(1, T[T.length - 1].level) : 0, after: T.length ? T.length - 1 : -1 };
  if (st.isSum[i] && !UI.collapsed.includes(T[i].id)) return { pos: i + 1, lvl: T[i].level + 1, after: -1 };
  return { pos: st.end[i] + 1, lvl: T[i].level, after: i };
}
function insertAt(pos, task) { P.tasks.splice(pos, 0, task); }
function newTask(o) { return Object.assign({ id: newId(), level: 1, name: 'Công việc mới', dur: 3, preds: [], res: '', pct: 0, snet: null, bl: null }, o); }
function focusNew(id) { UI.sel = id; SELS = new Set([id]); anchorId = id; if (UI.filter !== 'all') UI.filter = 'all'; pendingFocus = { id, col: 'name' }; renderAll(); flushFocus(); }

/* ---------- clipboard (rows) ---------- */
let CLIP = null, CLIP_TSV = '', CLIP_CUT = false;
function blockRows(B) {
  const rows = [];
  B.forEach(([a, b]) => { const base = P.tasks[a].level; for (let j = a; j <= b; j++) { const t = JSON.parse(JSON.stringify(P.tasks[j])); t.level -= base; rows.push(t); } });
  return rows;
}
function rowsToTsv(rows) {
  const pos = new Map(rows.map((t, k) => [t.id, k + 1]));
  const out = ['STT\tCấp\tTên công việc\tNgày\tLK trước\tNguồn lực\t%'];
  rows.forEach((t, k) => {
    const lk = (t.preds || []).filter(l => pos.has(l.id)).map(l => { let s = String(pos.get(l.id)); if (l.type !== 'FS' || l.lag) s += l.type; if (l.lag) s += (l.lag > 0 ? '+' : '') + l.lag; return s; }).join(', ');
    out.push([k + 1, t.level, t.name, +t.dur || 0, lk, t.res || '', +t.pct || 0].join('\t'));
  });
  return out.join('\n');
}
function copySel(cut) {
  const B = selectedBlocks(); if (!B.length) { toast('Hãy chọn dòng trước (bấm vào số STT, giữ Shift hoặc Ctrl để chọn nhiều).', true); return; }
  CLIP = blockRows(B); CLIP_TSV = rowsToTsv(CLIP); CLIP_CUT = !!cut;
  try { navigator.clipboard && navigator.clipboard.writeText(CLIP_TSV).catch(() => {}); } catch {}
  if (cut) { deleteBlocks(B, true); toast(`Đã cắt ${CLIP.length} dòng · chọn vị trí rồi bấm Dán (Ctrl+V)`); }
  else toast(`Đã chép ${CLIP.length} dòng · Ctrl+V để dán, hoặc dán sang Excel`);
}
function pasteRows(rows, label, keepBl) {
  const T = P.tasks, { pos, lvl } = insertionPoint(), map = new Map(), exists = new Set(T.map(t => t.id));
  const copies = rows.map(t => { const c = JSON.parse(JSON.stringify(t)); c.id = newId(); map.set(t.id, c.id); c.level = (t.level || 0) + lvl; if (!keepBl) c.bl = null; return c; });
  copies.forEach(c => { c.preds = (c.preds || []).map(l => map.has(l.id) ? { ...l, id: map.get(l.id) } : (exists.has(l.id) ? l : null)).filter(Boolean); });
  pushUndo(); T.splice(pos, 0, ...copies);
  SELS = new Set(copies.filter((c, k) => (rows[k].level || 0) === 0).map(c => c.id)); UI.sel = copies[0].id; anchorId = UI.sel;
  if (UI.filter !== 'all') UI.filter = 'all';
  renderAll(); toast(`Đã dán ${copies.length} dòng${label || ''}`);
}
function pasteClip() { if (!CLIP) { openPaste(); return; } pasteRows(CLIP, '', CLIP_CUT); }
function deleteBlocks(B, silent) {
  const T = P.tasks, gone = new Set(); B.forEach(([a, b]) => { for (let j = a; j <= b; j++) gone.add(T[j].id); });
  const first = B[0][0]; pushUndo();
  P.tasks = T.filter(t => !gone.has(t.id)); P.tasks.forEach(t => { t.preds = (t.preds || []).filter(l => !gone.has(l.id)); });
  const nx = P.tasks[Math.min(first, P.tasks.length - 1)]; UI.sel = nx ? nx.id : null; SELS = new Set(UI.sel ? [UI.sel] : []); anchorId = UI.sel;
  renderAll(); if (!silent) toast(`Đã xóa ${gone.size} dòng · Ctrl+Z để hoàn tác`);
}

/* ---------- row actions ---------- */
function act(a) {
  const T = P.tasks, i = selIdx(), st = S.st;
  const needSel = () => { if (i < 0) { toast('Hãy chọn một dòng trước (bấm vào số STT).', true); return false; } return true; };
  if (a === 'addBelow' || a === 'addMs') {
    const ms = a === 'addMs', { pos, lvl, after } = insertionPoint();
    const preds = after >= 0 && !st.isSum[after] ? [{ id: T[after].id, type: 'FS', lag: 0 }] : [];
    pushUndo(); const t = newTask({ level: lvl, preds, name: ms ? 'Mốc hoàn thành' : 'Công việc mới', dur: ms ? 0 : 3 });
    insertAt(pos, t); focusNew(t.id); return;
  }
  if (a === 'addAbove') {
    if (!needSel()) return; pushUndo(); const t = newTask({ level: T[i].level }); insertAt(i, t); focusNew(t.id); return;
  }
  if (a === 'addChild') {
    if (!needSel()) return; pushUndo();
    if (UI.collapsed.includes(T[i].id)) UI.collapsed = UI.collapsed.filter(x => x !== T[i].id);
    const t = newTask({ level: T[i].level + 1 }); insertAt(st.end[i] + 1, t); focusNew(t.id); return;
  }
  if (a === 'addGroup') {
    const { pos, lvl } = insertionPoint(); pushUndo();
    const g = newTask({ level: lvl, name: 'HẠNG MỤC MỚI', dur: 0 }), c = newTask({ level: lvl + 1 });
    T.splice(pos, 0, g, c); focusNew(g.id); return;
  }
  if (a === 'cut' || a === 'copy') return copySel(a === 'cut');
  if (a === 'pasteRows') return pasteClip();
  if (a === 'indent' || a === 'outdent') {
    const B = selectedBlocks(); if (!B.length) { needSel(); return; }
    const ok = B.filter(([s]) => a === 'indent' ? s > 0 && T[s - 1].level >= T[s].level : T[s].level > 0);
    if (!ok.length) { if (a === 'indent') toast('Không thụt lề được: dòng phía trên phải cùng cấp hoặc sâu hơn.', true); return; }
    pushUndo(); ok.forEach(([s, e]) => { for (let j = s; j <= e; j++) T[j].level += a === 'indent' ? 1 : -1; }); renderAll(); return;
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
    const B = selectedBlocks(); if (!B.length) { needSel(); return; }
    const rows = blockRows(B), map = new Map(), lastEnd = B[B.length - 1][1];
    const copies = rows.map(t => { const c = JSON.parse(JSON.stringify(t)); c.id = newId(); map.set(t.id, c.id); c.bl = null; return c; });
    // keep each block's own base level
    let k = 0; B.forEach(([s, e2]) => { const base = T[s].level; for (let j = s; j <= e2; j++, k++) copies[k].level = rows[k].level + base; });
    copies.forEach(c => { c.preds = c.preds.map(l => map.has(l.id) ? { ...l, id: map.get(l.id) } : l); });
    if (B.length === 1) copies[0].name += ' (bản sao)';
    pushUndo(); T.splice(lastEnd + 1, 0, ...copies); UI.sel = copies[0].id; SELS = new Set([UI.sel]); renderAll(); toast('Đã nhân đôi ' + copies.length + ' dòng'); return;
  }
  if (a === 'del') { const B = selectedBlocks(); if (!B.length) { needSel(); return; } deleteBlocks(B); return; }
  if (a === 'collapseAll') { UI.collapsed = P.tasks.filter((t, j) => st.isSum[j] && t.level > 0).map(t => t.id); renderAll(); return; }
  if (a === 'expandAll') { UI.collapsed = []; renderAll(); return; }
  if (a === 'check') return openCheck();
  if (a === 'money') { UI.money = !showMoney(); renderAll(); return; }
  if (a === 'weekly') return openWeekly();
  if (a === 'meeting') return openMeeting();
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

/* ---------- data check ---------- */
function dataIssues() {
  const T = P.tasks, st = S.st, out = [], add = (i, sev, msg) => out.push({ i, sev, msg });
  const succ = new Set();
  T.forEach(t => (t.preds || []).forEach(l => { const p = S.idIdx.get(l.id); if (p == null) return; for (let j = p; j <= st.end[p]; j++) succ.add(j); }));
  const hasPred = i => { for (let k = i; k >= 0; k = st.parent[k]) if ((T[k].preds || []).length) return true; return false; };
  const endEF = S.leaves.length ? Math.max(...S.leaves.map(i => S.ef[i])) : 0;
  const firstLeaf = S.leaves[0];
  const known = new Set(Object.keys(P.resources || {}));
  S.leaves.forEach(i => {
    const t = T[i];
    if (!String(t.name || '').trim()) add(i, 'err', 'Chưa có tên công việc');
    if (S.cycle[i]) add(i, 'err', 'Liên kết bị vòng lặp: một liên kết đã được tạm bỏ qua để tính tiến độ');
    if (S.badLink[i]) add(i, 'err', 'Liên kết không hợp lệ: tự nối với chính nó hoặc nối với hạng mục cha/con');
    if (i !== firstLeaf && !hasPred(i)) add(i, 'warn', 'Đầu hở: chưa có việc trước, đang bắt đầu ngay từ ngày khởi công');
    if (!succ.has(i) && S.ef[i] < endEF) add(i, 'warn', 'Đuôi hở: chưa có việc sau, trễ việc này không làm trễ việc nào khác');
    if ((+t.dur || 0) > 60) add(i, 'warn', `Thời gian ${t.dur} ngày khá dài, nên chia nhỏ để dễ theo dõi`);
    if (S.pct[i] > 0) (t.preds || []).forEach(l => { const p = S.idIdx.get(l.id); if (p != null && (l.type || 'FS') === 'FS' && S.pct[p] < 100) add(i, 'warn', `Đã có ${Math.round(S.pct[i])}% trong khi việc trước (STT ${p + 1}) mới ${Math.round(S.pct[p])}%`); });
    if (S.pct[i] > 0 && S.start[i] > S.dataDay) add(i, 'warn', 'Đã có % hoàn thành nhưng ngày bắt đầu sau ngày báo cáo');
    if (t.snet) add(i, 'info', `Đang ghim bắt đầu không sớm hơn ${fmtD(toDay(t.snet))}`);
    parseRes(t.res).forEach(r => { if (!known.has(r.code)) add(i, 'info', `Mã nguồn lực “${r.code}” chưa khai báo trong Cấu hình (chưa có giới hạn/ngày)`); });
  });
  const rank = { err: 0, warn: 1, info: 2 }; out.sort((a, b) => rank[a.sev] - rank[b.sev] || a.i - b.i);
  return out;
}
function openCheck() {
  const list = dataIssues(), T = P.tasks;
  const n = s => list.filter(x => x.sev === s).length;
  const lbl = { err: 'Lỗi', warn: 'Cảnh báo', info: 'Lưu ý' };
  const body = list.length ? `
    <div style="display:flex;gap:8px;flex-wrap:wrap"><span class="pill bad">${n('err')} lỗi</span><span class="pill neu" style="color:var(--accent);background:var(--accent-bg)">${n('warn')} cảnh báo</span><span class="pill neu">${n('info')} lưu ý</span></div>
    <div class="tbl-wrap"><table class="rt"><thead><tr><th>STT</th><th>Công việc</th><th>Vấn đề</th><th></th></tr></thead><tbody>
    ${list.map(x => `<tr><td class="n">${x.i + 1}</td><td>${esc(T[x.i].name || '(chưa đặt tên)')}</td><td><span class="sev ${x.sev}">${lbl[x.sev]}</span> ${esc(x.msg)}</td><td><button class="tb" data-goto="${T[x.i].id}" style="height:26px">Tới dòng</button></td></tr>`).join('')}
    </tbody></table></div>` : `<p style="margin:0;line-height:1.6">Không thấy vấn đề nào: mọi công việc đều có liên kết trước/sau hợp lý, không có vòng lặp hay liên kết sai.</p>`;
  openDlg('Kiểm tra dữ liệu tiến độ', body, `<button class="tb primary" data-close>Đóng</button>`, d => {
    d.querySelectorAll('[data-goto]').forEach(b => b.onclick = () => { dlg.close(); gotoRow(b.dataset.goto); });
  });
}
function gotoRow(id) {
  const i = idxById(id); if (i < 0) return; UI.filter = 'all';
  for (let p = S.st.parent[i]; p >= 0; p = S.st.parent[p]) UI.collapsed = UI.collapsed.filter(x => x !== P.tasks[p].id);
  UI.sel = id; SELS = new Set([id]); anchorId = id; renderAll();
  const r = VIS.indexOf(i); if (r >= 0) { lBody.scrollTop = Math.max(0, r * RH - lBody.clientHeight / 3); if (S.start[i] != null) scrollToDay(S.start[i]); }
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
  } else if (col === 'val' || col === 'cost') {
    const n = v ? Math.round(vnNum(v.replace(/tr(iệu)?/i, '')) * 1e6) : 0;
    if (v && !n && !/^0([.,]0*)?$/.test(v)) { toast('Nhập số tiền theo triệu đồng, ví dụ 88,2 cho 88.200.000 đ.', true); inp.value = inp.defaultValue; return; }
    if (n !== (+t[col] || 0)) change = () => { t[col] = n; };
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
    <label class="chk"><input type="radio" name="pMode" value="insert" checked> Chèn dưới dòng đang chọn</label>
    <label class="chk"><input type="radio" name="pMode" value="append"> Thêm vào cuối bảng</label>
    <label class="chk"><input type="radio" name="pMode" value="replace"> Thay toàn bộ bảng hiện tại</label>`,
    `<button class="tb" data-close>Hủy</button><button class="tb primary" id="pOk">Nhập</button>`,
    d => {
      setTimeout(() => d.querySelector('#pTxt').focus(), 30);
      d.querySelector('#pOk').onclick = () => {
        const txt = d.querySelector('#pTxt').value; const rows = txt.split(/\r?\n/).filter(l => l.trim()).map(l => l.split('\t'));
        const r = importRows(rows, d.querySelector('input[name=pMode]:checked').value); if (r) dlg.close();
      };
    });
}
function importRows(rows, mode) {
  const replace = mode === true || mode === 'replace';
  if (!rows.length) { toast('Không có dữ liệu để nhập.', true); return false; }
  const norm = s => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').trim();
  let hi = -1, map = null;
  for (let r = 0; r < Math.min(rows.length, 8); r++) {
    const h = rows[r].map(norm); const find = (...keys) => h.findIndex(c => keys.some(k => c === k || c.startsWith(k)));
    const name = find('ten cong viec', 'ten', 'cong viec', 'noi dung', 'hang muc', 'task name', 'name');
    if (name >= 0) { hi = r; map = { name, dur: find('ngay', 'so ngay', 'thoi gian', 'tg', 'duration'), pred: find('lk', 'lien ket', 'tien nhiem', 'predecessor', 'cv truoc'), wbs: find('wbs', 'ma'), lvl: find('cap', 'level', 'outline'), res: find('nguon luc', 'resource'), pct: find('%', 'phan tram', 'hoan thanh', 'percent'), stt: find('stt', 'id', 'tt'), val: find('gia tri', 'thanh tien', 'gia ban'), cost: find('chi phi', 'gia von') }; break; }
  }
  if (!map) map = { name: 0, dur: 1, pred: 2, wbs: -1, lvl: -1, res: 3, pct: -1, stt: -1, val: -1, cost: -1 };
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
    return { id: newId(), level: lvl, name: get(r, map.name), dur: durRaw ? Math.round(parseFloat(durRaw)) : 1, preds: [], res: get(r, map.res), pct: Math.max(0, Math.min(100, Math.round(parseFloat(get(r, map.pct).replace('%', '').replace(',', '.')) || 0))), snet: null, bl: null, val: Math.round(vnNum(map.val >= 0 ? r[map.val] : 0)), cost: Math.round(vnNum(map.cost >= 0 ? r[map.cost] : 0)) };
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
  const minL = Math.min(...T.map(t => t.level));
  if (mode === 'insert') { T.forEach(t => { t.level -= minL; }); pasteRows(T, ' từ Excel' + (bad ? ` · ${bad} liên kết không hiểu được đã bỏ qua` : '')); return true; }
  T.forEach(t => t.level -= minL - (replace ? 0 : Math.min(1, minL)));
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
      <span class="kbd">Ctrl+Z / Ctrl+Y</span><span>Hoàn tác / làm lại</span>
      <span class="kbd">Ctrl+S / Ctrl+O</span><span>Lưu tệp .json / mở tệp</span>
      <span class="kbd">Bấm STT + Shift / Ctrl</span><span>Chọn nhiều dòng liền nhau / chọn thêm từng dòng (Ctrl+A: chọn hết)</span>
      <span class="kbd">Ctrl+C / Ctrl+X / Ctrl+V</span><span>Chép / cắt / dán dòng (khi không sửa ô). Dán được cả bảng copy từ Excel; dòng đã chép dán sang Excel được</span>
      <span class="kbd">Ctrl+D</span><span>Nhân đôi dòng đã chọn</span>
      <span class="kbd">Alt+↑ / Alt+↓</span><span>Đưa dòng lên / xuống</span>
      <span class="kbd">Alt+- / Alt+=</span><span>Thu gọn / mở rộng tất cả hạng mục</span>
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

/* ---------- tiền: báo giá, đợt thanh toán, dòng tiền ---------- */
const fmtTy = v => { const a = Math.abs(v), sg = v < 0 ? '−' : ''; return sg + (a >= 1e9 ? (a / 1e9).toLocaleString('vi-VN', { maximumFractionDigits: 2 }) + ' tỷ' : (a / 1e6).toLocaleString('vi-VN', { maximumFractionDigits: 1 }) + ' tr'); };
const monthLabel = k => `T${+k.slice(5, 7)}/${k.slice(2, 4)}`;
function ensurePayments() { if (P.payments == null && (S.totalVal || S.totalCost)) P.payments = defaultPayments(P, S); }

function moneySection() {
  if (!(S.totalVal || S.totalCost)) return `<div class="card"><h3>Dòng tiền</h3><p style="margin:0 0 10px;color:var(--muted);line-height:1.6;max-width:640px">Chưa có giá trị cho các công việc. Nhập file báo giá (BOQ) để có dòng tiền thu – chi theo tháng, chỉ số SPI/CPI và kế hoạch thanh toán theo đợt. Cũng có thể gõ trực tiếp vào cột Giá trị và Chi phí trong bảng.</p><button class="tb primary" data-top="quote">${icon('coins')}Nhập báo giá</button></div>`;
  ensurePayments();
  const cf = cashFlow(P, S), e = evm(P, S), contract = +P.contractValue || S.totalVal;
  const margin = S.totalVal ? (S.totalVal - S.totalCost) / S.totalVal * 100 : 0;
  const minCum = Math.min(0, ...cf.months.map(m => m.fcCum)), minM = cf.months.find(m => m.fcCum === minCum);
  const spiTxt = e.SPI == null ? '—' : e.SPI.toFixed(2), cpiTxt = e.CPI == null ? '—' : e.CPI.toFixed(2);
  const spiCls = e.SPI == null ? '' : e.SPI >= 0.98 ? 'var(--ok)' : 'var(--crit)';
  // biểu đồ cột thu/chi theo tháng + đường lũy kế
  const M = cf.months, Wd = Math.max(1080, M.length * 64 + 90), Hd = 300, L = 70, R = 14, Tp = 14, B = 30;
  const vals = M.flatMap(m => [m.fcIn, m.fcOut, m.fcCum, m.planCum]); const rawHi = Math.max(1, ...vals), rawLo = Math.min(0, ...vals);
  const step = (() => { const r = (rawHi - rawLo) / 4, p = Math.pow(10, Math.floor(Math.log10(r))), f = r / p; return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10) * p; })();
  const hi = Math.ceil(rawHi / step) * step, lo = Math.floor(rawLo / step) * step;
  const sy = v => Tp + (hi - v) / (hi - lo) * (Hd - Tp - B), bw = (Wd - L - R) / Math.max(1, M.length);
  const cx = k => L + bw * k + bw / 2;
  let svg = `<svg class="s-svg" viewBox="0 0 ${Wd} ${Hd}" style="min-width:${Math.min(Wd, 760)}px" role="img" aria-label="Dòng tiền theo tháng">`;
  for (let v = lo; v <= hi + step / 2; v += step) { svg += `<line class="s-grid" x1="${L}" x2="${Wd - R}" y1="${sy(v)}" y2="${sy(v)}"/><text class="s-lbl" x="${L - 6}" y="${sy(v) + 4}" text-anchor="end">${fmtTy(v)}</text>`; }
  svg += `<line class="s-axis" x1="${L}" x2="${Wd - R}" y1="${sy(0)}" y2="${sy(0)}"/>`;
  M.forEach((m, k) => {
    const w = Math.max(3, Math.min(22, bw / 2 - 6));
    svg += `<rect class="cf-in" x="${cx(k) - w - 1}" y="${sy(m.fcIn)}" width="${w}" height="${Math.max(0, sy(0) - sy(m.fcIn))}"><title>${monthLabel(m.key)} thu ${fmtTy(m.fcIn)}</title></rect>`;
    svg += `<rect class="cf-out" x="${cx(k) + 1}" y="${sy(m.fcOut)}" width="${w}" height="${Math.max(0, sy(0) - sy(m.fcOut))}"><title>${monthLabel(m.key)} chi ${fmtTy(m.fcOut)}</title></rect>`;
    svg += `<text class="s-lbl" x="${cx(k)}" y="${Hd - 10}" text-anchor="middle">${monthLabel(m.key)}</text>`;
  });
  const line = f => M.map((m, k) => (k ? 'L' : 'M') + cx(k).toFixed(1) + ' ' + sy(m[f]).toFixed(1)).join('');
  svg += `<path class="cf-plan" d="${line('planCum')}"/><path class="cf-cum" d="${line('fcCum')}"/>`;
  M.forEach((m, k) => { svg += `<circle class="${m.fcCum < 0 ? 'cf-neg' : 'cf-dot'}" cx="${cx(k)}" cy="${sy(m.fcCum)}" r="3.5"><title>Lũy kế ${fmtTy(m.fcCum)}</title></circle>`; });
  svg += '</svg>';
  const rows = M.map(m => `<tr><td class="n">${monthLabel(m.key)}</td><td class="n r">${fmtTy(m.planIn)}</td><td class="n r">${fmtTy(m.planOut)}</td><td class="n r">${fmtTy(m.fcIn)}</td><td class="n r">${fmtTy(m.fcOut)}</td><td class="n r ${m.fcNet < 0 ? 'neg' : ''}">${fmtTy(m.fcNet)}</td><td class="n r ${m.fcCum < 0 ? 'neg' : ''}">${fmtTy(m.fcCum)}</td></tr>`).join('');
  const stages = cf.byStages ? cf.stages.map(x => `<tr><td>${esc(x.name)}</td><td class="n r">${x.pct}%</td><td class="n r">${fmtTy(x.amount)}</td><td class="n">${fmtD(x.day)}</td></tr>`).join('') : '';
  const payPct = (P.payments || []).reduce((a, p) => a + (+p.pct || 0), 0);
  return `
    <div class="kpis">
      <div class="kpi"><div class="k">Giá trị hợp đồng</div><div class="v">${fmtTy(contract)}</div><div class="s">Giá vốn ${fmtTy(S.totalCost)} · lãi gộp ${margin.toLocaleString('vi-VN', { maximumFractionDigits: 1 })}%</div></div>
      <div class="kpi"><div class="k">Chỉ số tiến độ SPI</div><div class="v" style="color:${spiCls}">${spiTxt}</div><div class="s">Đã làm ${fmtTy(e.EV)} / kế hoạch ${fmtTy(e.PV)}${e.SPI != null ? (e.SPI >= 0.98 ? ' · đạt kế hoạch' : ' · chậm hơn kế hoạch') : ''}</div></div>
      <div class="kpi"><div class="k">Chỉ số chi phí CPI</div><div class="v" style="color:${e.CPI == null ? '' : e.CPI >= 1 ? 'var(--ok)' : 'var(--crit)'}">${cpiTxt}</div><div class="s"><label for="acIn">Chi thực tế đến ngày BC (tr)</label> <input id="acIn" class="ac-in" inputmode="decimal" value="${e.AC != null ? fmtTr(e.AC) : ''}" placeholder="nhập để tính"></div></div>
      <div class="kpi"><div class="k">Vốn cần ứng tối đa</div><div class="v" style="color:${minCum < 0 ? 'var(--crit)' : 'var(--ok)'}">${minCum < 0 ? fmtTy(-minCum) : '0'}</div><div class="s">${minCum < 0 ? `Lũy kế thấp nhất vào ${monthLabel(minM.key)}` : 'Tiền thu luôn đi trước tiền chi'}</div></div>
    </div>
    <div class="card"><h3>Dòng tiền theo tháng <span class="pill neu">${cf.byStages ? 'thu theo đợt thanh toán' : 'thu theo giá trị làm được'}</span></h3>
      <div class="tbl-wrap">${svg}</div>
      <div class="s-key"><span class="lg"><i class="sw" style="background:var(--ok)"></i>Thu (dự báo)</span><span class="lg"><i class="sw" style="background:var(--crit)"></i>Chi (dự báo)</span><span class="lg"><i class="sw" style="background:var(--accent-strong);height:3px"></i>Lũy kế thu − chi</span><span class="lg"><i class="sw" style="border-top:2px dashed var(--faint);height:0;border-radius:0"></i>Lũy kế theo kế hoạch ${P.baselineAt ? 'gốc' : ''}</span></div>
    </div>
    <div class="two">
      <div class="card"><h3>Bảng dòng tiền</h3><div class="tbl-wrap"><table class="rt"><thead><tr><th>Tháng</th><th class="r">Thu KH</th><th class="r">Chi KH</th><th class="r">Thu dự báo</th><th class="r">Chi dự báo</th><th class="r">Chênh lệch</th><th class="r">Lũy kế</th></tr></thead><tbody>${rows}</tbody></table></div></div>
      <div class="card"><h3>Đợt thanh toán <span class="pill ${Math.abs(payPct - 100) < 0.01 || !cf.byStages ? 'neu' : 'bad'}">${cf.byStages ? payPct + '%' : 'chưa khai báo'}</span></h3>
        ${cf.byStages ? `<div class="tbl-wrap"><table class="rt"><thead><tr><th>Đợt</th><th class="r">%</th><th class="r">Số tiền</th><th>Ngày thu dự kiến</th></tr></thead><tbody>${stages}</tbody></table></div>` : '<div class="none">Đang tính thu theo giá trị làm được hằng tháng.</div>'}
        <div style="margin-top:10px"><button class="tb" data-top="payments">${icon('gear')}Sửa đợt thanh toán</button></div></div>
    </div>`;
}

function openPayments() {
  ensurePayments();
  const T = P.tasks, opts = sel => T.map((t, i) => `<option value="${t.id}" ${t.id === sel ? 'selected' : ''}>${i + 1} · ${'  '.repeat(t.level)}${esc(t.name).slice(0, 60)}</option>`).join('');
  const rowHtml = p => `<tr>
    <td><input data-p="name" value="${esc(p.name || '')}" placeholder="Tên đợt"></td>
    <td style="width:70px"><input data-p="pct" value="${p.pct ?? ''}" inputmode="decimal"></td>
    <td><select data-p="kind"><option value="start" ${p.kind === 'start' ? 'selected' : ''}>Khi khởi công</option><option value="task" ${p.kind === 'task' ? 'selected' : ''}>Khi xong việc</option><option value="date" ${p.kind === 'date' ? 'selected' : ''}>Ngày cố định</option></select></td>
    <td><select data-p="taskId" ${p.kind === 'task' ? '' : 'hidden'}>${opts(p.taskId)}</select><input type="date" data-p="date" value="${esc(p.date || '')}" ${p.kind === 'date' ? '' : 'hidden'}></td>
    <td style="width:64px"><input data-p="lag" value="${p.lag ?? 0}" inputmode="numeric" title="Số ngày sau mốc thì nhận được tiền"></td>
    <td style="width:36px"><button class="ib" data-del title="Xóa đợt">${icon('trash')}</button></td></tr>`;
  openDlg('Đợt thanh toán', `
    <div class="row2"><div class="fld"><label for="cvIn">Giá trị hợp đồng (triệu đồng)</label><input type="text" id="cvIn" inputmode="decimal" value="${fmtTr(+P.contractValue || S.totalVal)}"><div class="hint">Tổng giá trị các công việc: ${fmtTy(S.totalVal)}. Ghi giá trị đã gồm VAT nếu chủ đầu tư trả cả VAT.</div></div>
    <div class="fld"><span class="lb">Tổng các đợt</span><div id="payTotal" class="pay-total"></div></div></div>
    <div class="tbl-wrap"><table class="rtab pay"><thead><tr><td class="lb">Tên đợt</td><td class="lb">%</td><td class="lb">Thu khi</td><td class="lb">Mốc</td><td class="lb">+ ngày</td><td></td></tr></thead><tbody id="payRows">${(P.payments || []).map(rowHtml).join('')}</tbody></table></div>
    <div style="display:flex;gap:8px;flex-wrap:wrap"><button class="tb" id="payAdd" type="button">${icon('plus')}Thêm đợt</button><button class="tb" id="payTpl" type="button">Dùng mẫu 6 đợt</button><button class="tb" id="payNone" type="button">Thu theo giá trị làm được</button></div>`,
    `<button class="tb" data-close>Hủy</button><button class="tb primary" id="payOk">Lưu</button>`,
    d => {
      const body = d.querySelector('#payRows');
      const total = () => { const t = [...body.querySelectorAll('[data-p="pct"]')].reduce((a, x) => a + (vnNum(x.value) || 0), 0); const el = d.querySelector('#payTotal'); el.textContent = t.toLocaleString('vi-VN') + '%'; el.classList.toggle('bad', Math.abs(t - 100) > 0.01); };
      body.addEventListener('input', total); total();
      body.addEventListener('change', e => { if (e.target.dataset.p === 'kind') { const tr = e.target.closest('tr'); tr.querySelector('[data-p="taskId"]').hidden = e.target.value !== 'task'; tr.querySelector('[data-p="date"]').hidden = e.target.value !== 'date'; } });
      body.addEventListener('click', e => { const b = e.target.closest('[data-del]'); if (b) { b.closest('tr').remove(); total(); } });
      d.querySelector('#payAdd').onclick = () => { body.insertAdjacentHTML('beforeend', rowHtml({ name: '', pct: 0, kind: 'task', taskId: T[T.length - 1]?.id, lag: 7 })); total(); };
      d.querySelector('#payTpl').onclick = () => { body.innerHTML = defaultPayments(P, S).map(rowHtml).join(''); total(); };
      d.querySelector('#payNone').onclick = () => { body.innerHTML = ''; total(); };
      d.querySelector('#payOk').onclick = () => {
        const list = [...body.querySelectorAll('tr')].map(tr => { const g = k => tr.querySelector(`[data-p="${k}"]`).value; return { name: g('name').trim() || 'Đợt thanh toán', pct: vnNum(g('pct')), kind: g('kind'), taskId: g('taskId'), date: g('date'), lag: parseInt(g('lag'), 10) || 0 }; });
        if (list.some(p => p.kind === 'date' && !toDay(p.date))) { toast('Chọn ngày cho các đợt “Ngày cố định”.', true); return; }
        pushUndo(); P.payments = list; const cv = Math.round(vnNum(d.querySelector('#cvIn').value) * 1e6); P.contractValue = cv && Math.abs(cv - S.totalVal) > 1e5 ? cv : null;
        dlg.close(); renderAll(); toast(list.length ? `Đã lưu ${list.length} đợt thanh toán` : 'Đã chuyển sang thu theo giá trị làm được');
      };
    });
}

/* ---------- nhập báo giá ---------- */
let QUOTE = null;
async function readQuoteFile(f) {
  try {
    const wb = XLSX.read(await f.arrayBuffer(), { type: 'array' });
    const sheets = wb.SheetNames.map(n => ({ name: n, rows: XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null }) }));
    const q = parseQuote(sheets);
    if (!q.sheets.length) { toast('Không tìm thấy bảng báo giá trong file: cần cột “Nội dung/Tên công việc” và “Thành tiền”.', true); return; }
    QUOTE = { q, file: f.name }; openQuote();
  } catch { toast('Không đọc được file này. Hãy chọn file Excel (.xlsx) của báo giá.', true); }
}
function openQuote() {
  const { q, file } = QUOTE, secs = quoteSections(q);
  const sell = secs.reduce((a, s) => a + s.sell, 0), cost = secs.reduce((a, s) => a + s.cost, 0);
  const warnN = secs.reduce((a, s) => a + s.warnings.length, 0);
  const diff = q.grandTotal ? q.grandTotal - sell : 0;
  const hasTasks = P.tasks.some((t, i) => !S.st.isSum[i]) && !P.sample;
  openDlg('Nhập báo giá', `
    <div class="qsum"><div><span class="k">File</span><b>${esc(file)}</b></div><div><span class="k">Giá bán</span><b>${fmtTy(sell)}</b></div><div><span class="k">Giá vốn</span><b>${fmtTy(cost)}</b></div><div><span class="k">Lãi gộp</span><b>${sell ? ((sell - cost) / sell * 100).toLocaleString('vi-VN', { maximumFractionDigits: 1 }) : 0}%</b></div></div>
    ${q.grandTotal ? `<div class="hint" style="font-size:12.5px">Bảng tổng hợp ghi tổng <b>${fmtTy(q.grandTotal)}</b>${Math.abs(diff) > 1e6 ? `, chênh ${fmtTy(diff)} so với tổng các hạng mục (thường là VAT hoặc khoản ngoài BOQ). Giá trị hợp đồng dùng cho dòng thu mặc định lấy theo bảng tổng hợp.` : ', khớp với tổng các hạng mục.'}</div>` : ''}
    <div class="fld"><span class="lb">Sheet đưa vào</span><div class="chips">${q.sheets.map((s, k) => `<label class="chk"><input type="checkbox" data-sheet="${k}" ${s.use ? 'checked' : ''}> ${esc(s.name)} <span class="muted">(${s.sections.length} hạng mục)</span></label>`).join('')}</div></div>
    <div class="tbl-wrap qtbl"><table class="rt"><thead><tr><th>Hạng mục</th><th class="r">Giá bán</th><th class="r">Giá vốn</th><th class="r">Đầu việc</th><th>Kiểm tra</th></tr></thead><tbody>
      ${secs.map(s => `<tr><td>${esc(niceName(s.name))}<div class="muted" style="font-size:11px">${esc(s.sheet)} · ${s.code}</div></td><td class="n r">${fmtTy(s.sell)}</td><td class="n r">${fmtTy(s.cost)}</td><td class="n r">${s.items.length}</td><td>${s.warnings.length ? s.warnings.map(w => `<div class="qwarn">${esc(w)}</div>`).join('') : '<span class="muted">✓</span>'}</td></tr>`).join('')}
    </tbody></table></div>
    ${warnN ? `<div class="hint" style="font-size:12.5px">${warnN} chỗ cần xem lại trong file. Ứng dụng dùng tổng của các dòng con, không dùng dòng tổng.</div>` : ''}
    <div class="fld"><span class="lb">Cách dùng báo giá</span>
      <label class="opt"><input type="radio" name="qMode" value="create" ${hasTasks ? '' : 'checked'}><div><b>Tạo tiến độ mới từ báo giá</b><span>Mỗi hạng mục thành một công việc có sẵn giá trị, chi phí, thời gian gợi ý (theo chi phí nhân công) và liên kết theo trình tự thi công. Dự án đang mở sẽ được thay (Ctrl+Z để quay lại).</span></div></label>
      <label class="opt"><input type="radio" name="qMode" value="attach" ${hasTasks ? 'checked' : ''}><div><b>Gắn giá trị vào tiến độ đang mở</b><span>Tự ghép hạng mục với công việc theo tên, bạn chỉnh lại ở bước sau.</span></div></label>
    </div>
    <div class="row3" id="qCreate">
      <div class="fld"><label for="qName">Tên công trình</label><input type="text" id="qName" value="${esc(P.sample ? 'Công trình mới' : P.name)}"></div>
      <div class="fld"><label for="qStart">Ngày khởi công</label><input type="date" id="qStart" value="${isoOf(todayDay() + 7)}"></div>
      <div class="fld"><span class="lb">Mức chi tiết</span><label class="chk"><input type="checkbox" id="qDetail"> Tách từng đầu việc</label></div>
    </div>`,
    `<button class="tb" data-close>Hủy</button><button class="tb primary" id="qOk">Tiếp tục</button>`,
    d => {
      d.querySelectorAll('[data-sheet]').forEach(c => c.onchange = () => { q.sheets[+c.dataset.sheet].use = c.checked; openQuote(); });
      const sync = () => { d.querySelector('#qCreate').hidden = d.querySelector('input[name=qMode]:checked').value !== 'create'; };
      d.querySelectorAll('input[name=qMode]').forEach(r => r.onchange = sync); sync();
      d.querySelector('#qOk').onclick = () => {
        if (!secs.length) { toast('Chưa chọn sheet nào có hạng mục.', true); return; }
        if (d.querySelector('input[name=qMode]:checked').value === 'create') createFromQuote(secs, d.querySelector('#qName').value.trim() || 'Công trình mới', d.querySelector('#qStart').value || isoOf(todayDay()), d.querySelector('#qDetail').checked);
        else openAttach(secs);
      };
    });
}
function createFromQuote(secs, title, start, detail) {
  const { q } = QUOTE;
  pushUndo();
  P = { v: 1, name: title, contractor: '', start, workweek: P.workweek || 6, dataDate: start, useDataDate: true, holidays: P.holidays && P.holidays.length ? P.holidays : VN_HOLIDAYS.map(h => ({ ...h })),
    resources: P.resources || JSON.parse(JSON.stringify(DEFAULT_RES)), tasks: buildTasks(scheduleRowsFromQuote(secs, { title, detail })), contractValue: q.grandTotal || null, building: { floors: q.floors } };
  P.payments = defaultPayments(P, schedule(P, true));
  UI.collapsed = []; UI.sel = null; UI.filter = 'all'; UI.money = null; dlg.close(); renderAll(); fitGantt();
  toast(`Đã tạo ${P.tasks.length} dòng từ báo giá. Thời gian là gợi ý, hãy chỉnh lại cho sát thực tế rồi bấm Lưu KH gốc.`);
}
function openAttach(secs) {
  const T = P.tasks, m = matchSections(secs, T.map(t => t.name));
  const opts = k => `<option value="">— Bỏ qua —</option><option value="__new">+ Thêm thành việc mới</option>` + T.map((t, i) => `<option value="${t.id}" ${m[k].best === i ? 'selected' : ''}>${i + 1} · ${'  '.repeat(t.level)}${esc(t.name).slice(0, 60)}</option>`).join('');
  openDlg('Gắn báo giá vào công việc', `
    <div class="hint" style="font-size:12.5px">Chọn công việc nhận giá trị của từng hạng mục. Gắn vào hạng mục tổng hợp thì tiền được chia cho các việc con theo số ngày. Những dòng tự ghép đã được chọn sẵn.</div>
    <div class="tbl-wrap qtbl"><table class="rt"><thead><tr><th>Hạng mục báo giá</th><th class="r">Giá bán</th><th>Gắn vào công việc</th></tr></thead><tbody>
      ${secs.map((s, k) => `<tr><td>${esc(niceName(s.name))}</td><td class="n r">${fmtTy(s.sell)}</td><td><select data-k="${k}" class="qsel">${opts(k)}</select></td></tr>`).join('')}
    </tbody></table></div>
    <label class="chk"><input type="checkbox" id="aClear" checked> Xóa giá trị, chi phí cũ của các công việc trước khi gắn</label>`,
    `<button class="tb" data-close>Hủy</button><button class="tb primary" id="aOk">Gắn giá trị</button>`,
    d => {
      d.querySelector('#aOk').onclick = () => {
        const pick = [...d.querySelectorAll('select[data-k]')].map(x => x.value), st = S.st;
        pushUndo(); if (d.querySelector('#aClear').checked) P.tasks.forEach(t => { t.val = 0; t.cost = 0; });
        let n = 0, added = 0;
        secs.forEach((s, k) => {
          const v = pick[k]; if (!v) return; n++;
          if (v === '__new') { P.tasks.push(newTask({ level: 1, name: niceName(s.name), dur: estimateDays(s.labor, s.sell), val: Math.round(s.sell), cost: Math.round(s.cost) })); added++; return; }
          const i = idxById(v); if (i < 0) return;
          const L = st.isSum[i] ? (() => { const r = []; for (let j = i; j <= st.end[i]; j++) if (!st.isSum[j]) r.push(j); return r; })() : [i];
          const W = L.reduce((a, j) => a + Math.max(1, +P.tasks[j].dur || 0), 0);
          L.forEach(j => { const w = Math.max(1, +P.tasks[j].dur || 0) / W; P.tasks[j].val = Math.round((+P.tasks[j].val || 0) + s.sell * w); P.tasks[j].cost = Math.round((+P.tasks[j].cost || 0) + s.cost * w); });
        });
        if (QUOTE.q.grandTotal) P.contractValue = QUOTE.q.grandTotal;
        if (QUOTE.q.floors.length) P.building = { floors: QUOTE.q.floors };
        UI.money = null; dlg.close(); renderAll();
        toast(`Đã gắn ${n} hạng mục${added ? `, thêm ${added} việc mới ở cuối bảng (chưa có liên kết)` : ''}.`);
      };
    });
}

/* ---------- cập nhật tuần ---------- */
function openWeekly() {
  const T = P.tasks, today = todayDay(), old = S.dataDay;
  let newD = today > old ? today : old + 7, showAll = false;
  const entered = new Map(); // id -> { pct, note }
  const render = d => {
    const Pn = Object.assign({}, P, { dataDate: isoOf(newD) }), Sn = schedule(Pn, P.useDataDate !== false);
    const list = S.leaves.filter(i => showAll || (S.pct[i] < 100 && Sn.start[i] <= newD + 7) || entered.has(T[i].id)).sort((a, b) => S.start[a] - S.start[b] || a - b);
    d.querySelector('#wRows').innerHTML = list.length ? list.map(i => {
      const t = T[i], e = entered.get(t.id) || {}, pv = e.pct ?? Math.round(S.pct[i]), plan = Math.round(Sn.planned[i]);
      const behind = pv < plan - 0.5;
      return `<tr data-id="${t.id}">
        <td class="n">${i + 1}</td>
        <td>${esc(t.name)}${S.crit[i] ? ' <span class="badge crit">GĂNG</span>' : ''}<div class="muted" style="font-size:11px">${fmtDM(S.start[i])} → ${fmtDM(S.finish[i])}</div></td>
        <td class="n r">${Math.round(S.pct[i])}%</td>
        <td class="n r">${plan}%</td>
        <td><div class="pctin ${behind ? 'behind' : ''}"><input data-w="pct" value="${pv}" inputmode="numeric" aria-label="% mới"><span>%</span><button type="button" class="tb" data-q="25">+25</button><button type="button" class="tb" data-q="100">Xong</button></div></td>
        <td><input data-w="note" value="${esc(e.note ?? t.note ?? '')}" placeholder="Vướng mắc, ghi chú" class="wnote"></td></tr>`;
    }).join('') : '<tr><td colspan="6" class="none">Không có công việc nào cần cập nhật trong tuần này.</td></tr>';
  };
  const stash = d => d.querySelectorAll('#wRows tr[data-id]').forEach(tr => {
    const id = tr.dataset.id, t = P.tasks.find(x => x.id === id); if (!t) return;
    const pct = Math.max(0, Math.min(100, Math.round(vnNum(tr.querySelector('[data-w="pct"]').value)))), note = tr.querySelector('[data-w="note"]').value.trim();
    if (pct !== Math.round(+t.pct || 0) || note !== (t.note || '')) entered.set(id, { pct, note }); else entered.delete(id);
  });
  openDlg('Cập nhật tiến độ tuần', `
    <div class="row3">
      <div class="fld"><label for="wDate">Ngày báo cáo mới</label><input type="date" id="wDate" value="${isoOf(newD)}"><div class="hint">Lần trước: ${fmtD(old)}</div></div>
      <div class="fld"><label for="wAc">Chi phí thực tế lũy kế (triệu)</label><input type="text" id="wAc" inputmode="decimal" value="${P.actualCost != null ? fmtTr(P.actualCost) : ''}" placeholder="Không bắt buộc"><div class="hint">Dùng để tính CPI</div></div>
      <div class="fld"><span class="lb">Danh sách</span><label class="chk"><input type="checkbox" id="wAll"> Hiện tất cả công việc</label></div>
    </div>
    <div class="tbl-wrap qtbl" style="max-height:52vh"><table class="rt wtbl"><thead><tr><th>STT</th><th>Công việc</th><th class="r">Trước</th><th class="r">KH tới ngày</th><th>% thực tế</th><th>Ghi chú</th></tr></thead><tbody id="wRows"></tbody></table></div>
    <div class="hint" style="font-size:12px">Ô % tô đỏ khi thấp hơn kế hoạch tới ngày báo cáo mới. Ghi chú sẽ hiện trong Họp giao ban và tin nhắn báo cáo.</div>`,
    `<button class="tb" data-close>Hủy</button><button class="tb primary" id="wOk">Lưu cập nhật</button>`,
    d => {
      render(d);
      d.querySelector('#wDate').onchange = e => { const v = toDay(e.target.value); if (v != null) { stash(d); newD = v; render(d); } };
      d.querySelector('#wAll').onchange = e => { stash(d); showAll = e.target.checked; render(d); };
      d.querySelector('#wRows').addEventListener('click', e => {
        const b = e.target.closest('[data-q]'); if (!b) return; const inp = b.closest('tr').querySelector('[data-w="pct"]');
        inp.value = b.dataset.q === '100' ? 100 : Math.min(100, Math.round(vnNum(inp.value)) + 25);
      });
      d.querySelector('#wOk').onclick = () => {
        stash(d); pushUndo();
        entered.forEach((e, id) => { const t = P.tasks.find(x => x.id === id); if (t) { t.pct = e.pct; t.note = e.note; } });
        P.dataDate = isoOf(newD);
        const ac = d.querySelector('#wAc').value.trim(); P.actualCost = ac ? Math.round(vnNum(ac) * 1e6) : (P.actualCost ?? null);
        S = schedule(P, P.useDataDate !== false);
        const snap = snapshot(P, S); P.history = (P.history || []).filter(h => h.date !== snap.date); P.history.push(snap); P.history.sort((a, b) => a.date < b.date ? -1 : 1);
        dlg.close(); renderAll(); toast(`Đã cập nhật ${entered.size} công việc tới ngày ${fmtD(newD)}`);
      };
    });
}

/* ---------- look-ahead ---------- */
function lookAheadCard() {
  const W = UI.laWeeks || 3, la = lookAhead(P, S, W), T = P.tasks;
  const seg = [3, 4, 6].map(w => `<button type="button" data-la="${w}" class="${w === W ? 'on' : ''}">${w} tuần</button>`).join('');
  let body = '', lastG = null;
  la.rows.forEach(r => {
    if (r.group !== lastG) { lastG = r.group; body += `<tr class="la-g"><td colspan="${3 + W}">${r.group >= 0 ? esc(T[r.group].name) : 'Khác'}</td></tr>`; }
    const i = r.i;
    body += `<tr><td>${esc(T[i].name)}${S.crit[i] ? ' <span class="badge crit">GĂNG</span>' : ''}${S.late[i] ? ' <span class="badge late">CHẬM</span>' : ''}${T[i].note ? `<div class="la-note">${esc(T[i].note)}</div>` : ''}</td><td class="n">${fmtDM(S.start[i])} → ${fmtDM(S.finish[i])}</td><td class="n r">${Math.round(S.pct[i])}%</td>${r.cells.map(c => `<td class="la-c">${c ? `<i class="${S.crit[i] ? 'crit' : ''}"></i>` : ''}</td>`).join('')}</tr>`;
  });
  return `<div class="card" id="laCard"><h3>Look-ahead ${W} tuần <span class="pill neu">${la.rows.length} việc</span><span style="margin-left:auto;display:flex;gap:8px;align-items:center"><span class="seg">${seg}</span><button class="tb" data-top="laXlsx">${icon('sheet')}Xuất Excel</button></span></h3>
    ${la.rows.length ? `<div class="tbl-wrap"><table class="rt la"><thead><tr><th>Công việc</th><th>Thời gian</th><th class="r">%</th>${la.cols.map(c => `<th class="la-h">T${c.week}<div>${fmtDM(c.s)}</div></th>`).join('')}</tr></thead><tbody>${body}</tbody></table></div>` : '<div class="none">Không có việc nào trong khoảng này.</div>'}</div>`;
}
function exportLookAhead() {
  const W = UI.laWeeks || 3, la = lookAhead(P, S, W), T = P.tasks;
  const aoa = [[`LOOK-AHEAD ${W} TUẦN – ${P.name}`], [`Ngày báo cáo ${fmtD(S.dataDay)}`], [], ['STT', 'Hạng mục', 'Công việc', 'Bắt đầu', 'Kết thúc', '% HT', 'Găng', 'Nguồn lực', 'Ghi chú', ...la.cols.map(c => `Tuần ${c.week} (${fmtDM(c.s)})`)]];
  la.rows.forEach(r => aoa.push([r.i + 1, r.group >= 0 ? T[r.group].name : '', T[r.i].name, fmtD(S.start[r.i]), fmtD(S.finish[r.i]), Math.round(S.pct[r.i]), S.crit[r.i] ? 'Có' : '', T[r.i].res || '', T[r.i].note || '', ...r.cells.map(c => c ? '■' : '')]));
  const ws = XLSX.utils.aoa_to_sheet(aoa); ws['!cols'] = [5, 22, 40, 11, 11, 6, 6, 14, 30, ...la.cols.map(() => ({ wch: 13 }))].map(w => typeof w === 'number' ? { wch: w } : w);
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Look-ahead');
  offerFile(safeName(P.name) + '-look-ahead.xlsx', new Blob([XLSX.write(wb, { bookType: 'xlsx', type: 'array' })]));
}

/* ---------- họp giao ban ---------- */
let MEET = null;
function meetingSlides() {
  const T = P.tasks, D = S.dataDay, e = evm(P, S);
  const gap = S.actualPct - S.plannedPct, pfVar = S.blFinish != null ? S.cal.idxOf(S.projFinish) - S.cal.idxOf(S.blFinish) : null;
  const prev = (P.history || []).filter(h => h.date < P.dataDate).slice(-1)[0];
  const moved = S.leaves.map(i => ({ i, from: prev ? (prev.pcts[T[i].id] || 0) : null, to: Math.round(S.pct[i]) }))
    .filter(x => prev ? x.to > x.from : (x.to > 0 && x.to < 100));
  const late = S.leaves.filter(i => S.late[i]).sort((a, b) => (S.planned[b] - S.pct[b]) - (S.planned[a] - S.pct[a]));
  const next = S.leaves.filter(i => S.pct[i] < 100 && S.start[i] <= D + 13 && S.finish[i] >= D).sort((a, b) => S.start[a] - S.start[b]);
  const crit = S.leaves.filter(i => S.crit[i]).sort((a, b) => S.start[a] - S.start[b]);
  const li = (i, extra = '') => `<li><span class="m-n">${esc(T[i].name)}</span>${extra}</li>`;
  const more = (arr, n) => arr.length > n ? `<li class="m-more">… và ${arr.length - n} việc khác</li>` : '';
  const slides = [
    `<div class="m-kicker">Họp giao ban · ${fmtD(D)}</div><h1>${esc(P.name)}</h1>
     <div class="m-kpis">
       <div><span>Thực tế / Kế hoạch</span><b>${S.actualPct.toFixed(1)}% <small>/ ${S.plannedPct.toFixed(1)}%</small></b><em class="${gap < -0.05 ? 'bad' : 'good'}">${gap >= 0 ? 'Vượt' : 'Chậm'} ${Math.abs(gap).toFixed(1)}%</em></div>
       <div><span>Dự kiến hoàn thành</span><b>${fmtD(S.projFinish)}</b><em class="${pfVar > 0 ? 'bad' : 'good'}">${pfVar == null ? 'Chưa chốt KH gốc' : pfVar > 0 ? `Trễ ${pfVar} ngày` : pfVar < 0 ? `Sớm ${-pfVar} ngày` : 'Đúng KH gốc'}</em></div>
       <div><span>Việc chậm</span><b>${late.length}</b><em>${S.leaves.filter(i => S.crit[i]).length} việc găng còn lại</em></div>
       ${S.hasMoney ? `<div><span>SPI / CPI</span><b class="nowrap">${e.SPI == null ? '—' : e.SPI.toFixed(2)} <small>/ ${e.CPI == null ? '—' : e.CPI.toFixed(2)}</small></b><em>Đã làm ${fmtTy(e.EV)}</em></div>` : ''}
     </div>`,
    `<h2>${prev ? `Tuần qua (từ ${fmtD(toDay(prev.date))})` : 'Đang thi công'}</h2>${moved.length ? `<ul class="m-list">${moved.slice(0, 10).map(x => li(x.i, `<span class="m-p">${x.from != null ? x.from + '% → ' : ''}${x.to}%</span>`)).join('')}${more(moved, 10)}</ul>` : '<p class="m-empty">Chưa có công việc nào thay đổi.</p>'}`,
    `<h2>Việc chậm &amp; vướng mắc</h2>${late.length ? `<ul class="m-list">${late.slice(0, 9).map(i => li(i, `<span class="m-p bad">TT ${Math.round(S.pct[i])}% / KH ${Math.round(S.planned[i])}%</span>${T[i].note ? `<div class="m-note">${esc(T[i].note)}</div>` : ''}`)).join('')}${more(late, 9)}</ul>` : '<p class="m-empty">Không có việc chậm so với kế hoạch.</p>'}`,
    `<h2>Kế hoạch 2 tuần tới (${fmtDM(D)} – ${fmtDM(D + 13)})</h2>${next.length ? `<ul class="m-list">${next.slice(0, 10).map(i => li(i, `<span class="m-p">${fmtDM(S.start[i])} → ${fmtDM(S.finish[i])}${S.crit[i] ? ' · găng' : ''}</span>`)).join('')}${more(next, 10)}</ul>` : '<p class="m-empty">Không có việc trong 2 tuần tới.</p>'}`
  ];
  if (S.hasMoney) {
    const cf = cashFlow(P, S), now = isoOf(D).slice(0, 7), M = cf.months.filter(m => m.key >= now).slice(0, 3);
    const up = cf.stages.filter(x => x.day >= D).slice(0, 3);
    slides.push(`<h2>Dòng tiền sắp tới</h2><div class="m-cols"><div><h3>Theo tháng</h3><ul class="m-list">${M.map(m => `<li><span class="m-n">${monthLabel(m.key)}</span><span class="m-p">Thu ${fmtTy(m.fcIn)} · Chi ${fmtTy(m.fcOut)}</span><div class="m-note ${m.fcCum < 0 ? 'bad' : ''}">Lũy kế ${fmtTy(m.fcCum)}</div></li>`).join('')}</ul></div>
      <div><h3>Đợt thanh toán kế tiếp</h3>${up.length ? `<ul class="m-list">${up.map(x => `<li><span class="m-n">${esc(x.name)}</span><span class="m-p">${fmtTy(x.amount)} · ${fmtD(x.day)}</span></li>`).join('')}</ul>` : '<p class="m-empty">Không còn đợt nào.</p>'}</div></div>`);
  }
  slides.push(`<h2>Đường găng còn lại</h2>${crit.length ? `<ul class="m-list">${crit.slice(0, 10).map(i => li(i, `<span class="m-p">${fmtDM(S.start[i])} → ${fmtDM(S.finish[i])}</span>`)).join('')}${more(crit, 10)}</ul><p class="m-foot">Trễ bất kỳ việc nào ở đây là trễ ngày bàn giao.</p>` : '<p class="m-empty">Không còn việc găng.</p>'}`);
  return slides;
}
function openMeeting() {
  const slides = meetingSlides(); MEET = { k: 0, slides };
  let el = document.getElementById('meet');
  if (!el) { el = document.createElement('div'); el.id = 'meet'; el.setAttribute('role', 'dialog'); el.setAttribute('aria-label', 'Họp giao ban'); document.body.appendChild(el); }
  const draw = () => {
    el.innerHTML = `<div class="m-bar"><span>${esc(P.name)}</span><span class="m-dots">${slides.map((_, k) => `<button type="button" data-go="${k}" class="${k === MEET.k ? 'on' : ''}" aria-label="Trang ${k + 1}"></button>`).join('')}</span>
      <span class="m-act"><button type="button" class="tb" data-m="full">Toàn màn hình</button><button type="button" class="tb" data-m="close">Đóng (Esc)</button></span></div>
      <div class="m-slide">${slides[MEET.k]}</div>
      <div class="m-nav"><button type="button" class="tb" data-m="prev" ${MEET.k ? '' : 'disabled'}>← Trước</button><span class="num">${MEET.k + 1} / ${slides.length}</span><button type="button" class="tb primary" data-m="next" ${MEET.k < slides.length - 1 ? '' : 'disabled'}>Tiếp →</button></div>`;
  };
  el.onclick = e => {
    const b = e.target.closest('[data-m],[data-go]'); if (!b) return;
    if (b.dataset.go) MEET.k = +b.dataset.go;
    else if (b.dataset.m === 'next') MEET.k = Math.min(slides.length - 1, MEET.k + 1);
    else if (b.dataset.m === 'prev') MEET.k = Math.max(0, MEET.k - 1);
    else if (b.dataset.m === 'full') { const r = document.documentElement.requestFullscreen; if (r) document.documentElement.requestFullscreen().catch(() => {}); }
    else if (b.dataset.m === 'close') return closeMeeting();
    draw();
  };
  el.hidden = false; draw();
  MEET.key = e => {
    if (e.key === 'Escape') { closeMeeting(); return; }
    if (['ArrowRight', 'PageDown', ' '].includes(e.key)) { e.preventDefault(); MEET.k = Math.min(slides.length - 1, MEET.k + 1); draw(); }
    if (['ArrowLeft', 'PageUp'].includes(e.key)) { e.preventDefault(); MEET.k = Math.max(0, MEET.k - 1); draw(); }
  };
  document.addEventListener('keydown', MEET.key, true);
}
function closeMeeting() {
  const el = document.getElementById('meet'); if (el) el.hidden = true;
  if (MEET && MEET.key) document.removeEventListener('keydown', MEET.key, true); MEET = null;
  if (document.fullscreenElement && document.exitFullscreen) document.exitFullscreen().catch(() => {});
}

/* ---------- 3D ---------- */
let V3 = null, V3sig = '', V3play = 0, V3pick = null, V3ctx = null;
function d3Context() {
  const B = normalizeBuilding(P), { els, height } = buildElements(B), L = linksFor(B, P, S);
  const starts = S.leaves.map(i => Math.min(S.blS[i] ?? S.start[i], S.start[i])), fins = S.leaves.map(i => Math.max(S.blF[i] ?? S.finish[i], S.finish[i]));
  const a = (starts.length ? Math.min(...starts) : S.projStart) - 3, b = (fins.length ? Math.max(...fins) : S.projFinish) + 3;
  return { B, els, height, L, a, b };
}
function d3Progress(ctx) {
  const mode = UI.d3mode || 'plan', day = mode === 'plan' ? UI.d3day : S.dataDay, prog = {};
  ctx.els.forEach(e => { prog[e.key] = progressAt(ctx.L[e.key], P, S, day, mode); });
  // tường nhiều tầng cùng gắn một công việc: xây lần lượt từ dưới lên
  const groups = new Map();
  ctx.els.filter(e => e.kind === 'wall' && prog[e.key]).forEach(e => { const sig = [...ctx.L[e.key]].sort().join(','); if (!groups.has(sig)) groups.set(sig, []); groups.get(sig).push(e); });
  groups.forEach(list => { if (list.length < 2) return; const n = list.length; list.forEach((e, k) => { const p = prog[e.key]; prog[e.key] = { ...p, p: Math.max(0, Math.min(1, p.p * n - k)) }; }); });
  return prog;
}
function render3d() {
  const ctx = V3ctx = d3Context(), mode = UI.d3mode || 'plan';
  if (UI.d3day == null || UI.d3day < ctx.a || UI.d3day > ctx.b) UI.d3day = Math.min(Math.max(S.dataDay, ctx.a), ctx.b);
  $('#d3bar').innerHTML = `
    <span class="seg"><button type="button" data-d3mode="plan" class="${mode === 'plan' ? 'on' : ''}">Kế hoạch theo ngày</button><button type="button" data-d3mode="actual" class="${mode === 'actual' ? 'on' : ''}">Thực tế tại ngày BC</button></span>
    <button class="tb" data-d3="play" ${mode === 'plan' ? '' : 'disabled'}>${V3play ? '❚❚ Dừng' : '▶ Phát'}</button>
    <input type="range" id="d3range" min="${ctx.a}" max="${ctx.b}" step="1" value="${mode === 'plan' ? UI.d3day : S.dataDay}" ${mode === 'plan' ? '' : 'disabled'} aria-label="Ngày xem mô hình">
    <b class="num d3date" id="d3date">${fmtD(mode === 'plan' ? UI.d3day : S.dataDay)}</b>
    <span class="d3sp"></span>
    <button class="tb" data-d3="config">${icon('gear')}Cấu hình 3D</button>
    <button class="tb" data-d3="links">${icon('child')}Gán công việc</button>
    <button class="tb" data-d3="shot">${icon('save')}Chụp ảnh</button>
    <button class="ib" data-d3="reset" title="Về góc nhìn ban đầu">${icon('zoom')}</button>`;
  $('#d3legend').innerHTML = `<b>${esc(P.name)}</b><span class="muted">${ctx.B.floors.filter(f => f.kind === 'floor').length} tầng · mặt tiền ${ctx.B.width} m · cao ${ctx.height.toFixed(1)} m</span>
    <span class="lg"><i class="sw lg-todo"></i>Chưa làm</span><span class="lg"><i class="sw" style="background:var(--accent-strong)"></i>Đang làm</span><span class="lg"><i class="sw" style="background:#b9bcc2"></i>Xong (khung)</span><span class="lg"><i class="sw" style="background:#c8764f"></i>Tường xây</span><span class="lg"><i class="sw" style="background:#efe7d8;border:1px solid var(--line)"></i>Đã hoàn thiện</span>${mode === 'actual' ? '<span class="lg"><i class="sw" style="background:var(--crit)"></i>Đang chậm</span>' : ''}
    <span class="muted d3hint">Kéo để xoay · cuộn để phóng to · bấm vào khối để xem công việc</span>`;
  if (!V3) {
    $('#stage').innerHTML = '<div class="stage-load">Đang tải mô hình 3D…</div>';
    import('./view3d.js').then(m => {
      if (UI.view !== '3d') return;
      $('#stage').innerHTML = '';
      try { V3 = m.mount($('#stage'), { onPick: k => { V3pick = k; d3Info(); } }); }
      catch { $('#stage').innerHTML = '<div class="stage-load">Trình duyệt này không hiển thị được 3D (WebGL bị tắt). Hãy thử Chrome hoặc Edge bản mới.</div>'; return; }
      V3sig = ''; render3d();
    }).catch(() => { $('#stage').innerHTML = '<div class="stage-load">Không tải được phần 3D. Kiểm tra kết nối mạng rồi mở lại tab.</div>'; });
    return;
  }
  const sig = JSON.stringify([ctx.B.width, ctx.B.floors.map(f => [f.name, f.area, f.h, f.kind])]);
  if (sig !== V3sig) { V3.build(ctx.B, ctx.els); V3sig = sig; }
  V3.apply(d3Progress(ctx), { mode }); d3Info();
}
function d3Apply() { if (!V3 || !V3ctx) return; V3.apply(d3Progress(V3ctx), { mode: UI.d3mode || 'plan' }); $('#d3date').textContent = fmtD(UI.d3mode === 'actual' ? S.dataDay : UI.d3day); d3Info(); }
function d3Info() {
  const box = $('#d3info'); if (!box) return;
  if (!V3pick || !V3ctx) { box.hidden = true; return; }
  const e = V3ctx.els.find(x => x.key === V3pick); if (!e) { box.hidden = true; return; }
  const ids = V3ctx.L[e.key] || [], T = P.tasks;
  const prog = d3Progress(V3ctx)[e.key];
  box.hidden = false;
  box.innerHTML = `<div class="d3i-h"><b>${esc(e.label)}</b><button class="ib" data-d3="closeinfo" title="Đóng">${icon('x')}</button></div>
    ${prog ? `<div class="d3i-p"><span style="width:${Math.round(prog.p * 100)}%"></span></div><div class="muted">${UI.d3mode === 'actual' ? 'Thực tế' : 'Kế hoạch tới ' + fmtD(UI.d3day)}: <b>${Math.round(prog.p * 100)}%</b> · ${fmtDM(prog.start)} → ${fmtDM(prog.finish)}</div>` : '<div class="muted">Chưa gắn công việc nào.</div>'}
    <ul>${ids.map(id => { const i = S.idIdx.get(id); return i == null ? '' : `<li><span>${i + 1} · ${esc(T[i].name)}</span><span class="num">${Math.round(S.pct[i])}%</span></li>`; }).join('')}</ul>
    <button class="tb" data-d3="links">Đổi công việc gắn</button>`;
}
function d3Play() {
  if (V3play) { clearInterval(V3play); V3play = 0; render3d(); return; }
  const ctx = V3ctx; if (!ctx) return; if (UI.d3day >= ctx.b) UI.d3day = ctx.a;
  const step = Math.max(1, Math.round((ctx.b - ctx.a) / 160));
  V3play = setInterval(() => {
    UI.d3day = Math.min(ctx.b, UI.d3day + step); const r = $('#d3range'); if (r) r.value = UI.d3day; d3Apply();
    if (UI.d3day >= ctx.b || UI.view !== '3d') { clearInterval(V3play); V3play = 0; if (UI.view === '3d') render3d(); }
  }, 70);
  render3d();
}
function open3dConfig() {
  const B = normalizeBuilding(P);
  const kinds = [['found', 'Móng / hầm'], ['floor', 'Tầng'], ['roof', 'Mái / sân thượng'], ['roofbox', 'Tum / mái thang']];
  const row = f => `<tr><td><input data-f="name" value="${esc(f.name)}"></td><td><select data-f="kind">${kinds.map(([k, l]) => `<option value="${k}" ${f.kind === k ? 'selected' : ''}>${l}</option>`).join('')}</select></td><td style="width:90px"><input data-f="area" value="${f.area ? String(+f.area.toFixed(2)).replace('.', ',') : ''}" inputmode="decimal" placeholder="m²"></td><td style="width:80px"><input data-f="h" value="${String(f.h).replace('.', ',')}" inputmode="decimal"></td><td style="width:36px"><button class="ib" data-del type="button" title="Xóa">${icon('trash')}</button></td></tr>`;
  openDlg('Cấu hình mô hình 3D', `
    <div class="row2"><div class="fld"><label for="bW">Chiều rộng mặt tiền (m)</label><input type="text" id="bW" inputmode="decimal" value="${String(B.width).replace('.', ',')}"><div class="hint">Chiều sâu từng tầng = diện tích ÷ mặt tiền. Để trống diện tích thì lấy chiều sâu lớn nhất.</div></div>
    <div class="fld"><span class="lb">Nguồn số liệu</span><div class="hint" style="font-size:12.5px">${(P.building && P.building.floors && P.building.floors.length) ? 'Đang dùng danh sách tầng đã lưu (từ báo giá hoặc tự nhập).' : 'Đang đoán số tầng từ tên công việc. Nhập báo giá có bảng diện tích để lấy diện tích từng tầng.'}</div></div></div>
    <div class="tbl-wrap"><table class="rtab pay"><thead><tr><td class="lb">Tên tầng</td><td class="lb">Loại</td><td class="lb">Diện tích</td><td class="lb">Cao (m)</td><td></td></tr></thead><tbody id="bRows">${B.floors.map(row).join('')}</tbody></table></div>
    <div style="display:flex;gap:8px;flex-wrap:wrap"><button class="tb" id="bAdd" type="button">${icon('plus')}Thêm tầng</button><button class="tb" id="bAuto" type="button">Đoán lại từ tên công việc</button></div>`,
    `<button class="tb" data-close>Hủy</button><button class="tb primary" id="bOk">Lưu</button>`,
    d => {
      const body = d.querySelector('#bRows');
      body.addEventListener('click', e => { const b = e.target.closest('[data-del]'); if (b) b.closest('tr').remove(); });
      d.querySelector('#bAdd').onclick = () => { const trs = body.querySelectorAll('tr'); const ref = trs[trs.length - 1]; const html = row({ name: 'Lầu mới', kind: 'floor', area: 0, h: 3.4 }); const roofRow = [...trs].find(tr => ['roof', 'roofbox'].includes(tr.querySelector('[data-f="kind"]').value)); if (roofRow) roofRow.insertAdjacentHTML('beforebegin', html); else if (ref) ref.insertAdjacentHTML('afterend', html); else body.innerHTML = html; };
      d.querySelector('#bAuto').onclick = () => { pushUndo(); P.building = { ...(P.building || {}), floors: [] }; dlg.close(); renderAll(); toast('Đã đoán lại số tầng từ tên công việc'); };
      d.querySelector('#bOk').onclick = () => {
        const floors = [...body.querySelectorAll('tr')].map(tr => { const g = k => tr.querySelector(`[data-f="${k}"]`).value; return { name: g('name').trim(), kind: g('kind'), area: vnNum(g('area')), h: vnNum(g('h')) || 0 }; }).filter(f => f.name);
        if (!floors.some(f => f.kind === 'floor')) { toast('Cần ít nhất một tầng loại “Tầng”.', true); return; }
        pushUndo(); P.building = { ...(P.building || {}), width: Math.max(2, vnNum(d.querySelector('#bW').value) || 5), floors };
        dlg.close(); renderAll(); toast('Đã cập nhật mô hình 3D');
      };
    });
}
function open3dLinks() {
  const ctx = V3ctx || d3Context(), T = P.tasks;
  const opts = sel => T.map((t, i) => `<option value="${t.id}" ${sel.includes(t.id) ? 'selected' : ''}>${i + 1} · ${'  '.repeat(t.level)}${esc(t.name).slice(0, 70)}</option>`).join('');
  const els = ctx.els;
  openDlg('Gán công việc cho mô hình 3D', `
    <div class="hint" style="font-size:12.5px">Mỗi khối lấy mức hoàn thành từ các công việc được chọn (giữ Ctrl để chọn nhiều). Chọn một hạng mục tổng hợp là lấy toàn bộ việc con. Ứng dụng đã tự gắn theo tên tầng; chỉ cần sửa chỗ chưa đúng.</div>
    <div class="links3d">${els.map(e => `<div class="fld"><label>${esc(e.label)} <span class="muted">(${(ctx.L[e.key] || []).length} việc)</span></label><select multiple size="4" data-key="${e.key}">${opts(ctx.L[e.key] || [])}</select></div>`).join('')}</div>`,
    `<button class="tb" id="lAuto">Tự gắn lại theo tên</button><button class="tb" data-close>Hủy</button><button class="tb primary" id="lOk">Lưu</button>`,
    d => {
      d.querySelector('#lAuto').onclick = () => { pushUndo(); P.building = { ...(P.building || {}), links: null }; dlg.close(); renderAll(); toast('Đã gắn lại công việc theo tên'); };
      d.querySelector('#lOk').onclick = () => {
        const links = {}; d.querySelectorAll('select[data-key]').forEach(s => { links[s.dataset.key] = [...s.selectedOptions].map(o => o.value); });
        pushUndo(); P.building = { ...(P.building || {}), links }; dlg.close(); renderAll(); toast('Đã lưu cách gắn công việc');
      };
    });
}
function d3Shot() {
  if (!V3) return; const url = V3.snapshot();
  fetch(url).then(r => r.blob()).then(b => offerFile(safeName(P.name) + '-3d-' + isoOf(UI.d3mode === 'actual' ? S.dataDay : UI.d3day) + '.png', b));
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
    ['STT', 'WBS', 'Tên công việc', 'Cấp', 'Ngày', 'Bắt đầu', 'Kết thúc', 'LK trước', 'Dự trữ', 'Găng', 'Nguồn lực', '% HT', 'KH gốc BĐ', 'KH gốc KT', 'Lệch KHG', 'Giá trị (đ)', 'Chi phí (đ)']];
  T.forEach((t, i) => {
    const sum = S.st.isSum[i];
    aoa.push([i + 1, S.st.wbs[i], '  '.repeat(t.level) + t.name, t.level, sum ? S.ef[i] - S.es[i] : +t.dur || 0, fmtD(S.start[i]), fmtD(S.finish[i]), fmtPreds(t, S.idIdx), sum ? '' : S.tf[i], !sum && S.crit[i] ? 'Có' : '', t.res || '', Math.round(S.pct[i]),
      S.blS[i] != null ? fmtD(S.blS[i]) : '', S.blF[i] != null ? fmtD(S.blF[i]) : '', S.vr[i] == null ? '' : S.vr[i], S.val[i] || '', S.cost[i] || '']);
  });
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [5, 8, 44, 5, 6, 11, 11, 14, 7, 6, 16, 6, 11, 11, 9, 15, 15].map(w => ({ wch: w }));
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
$('#bQuote').innerHTML = icon('coins') + '<span class="lbl">Báo giá</span>';
$('#bQuote').onclick = () => $('#quoteIn').click();
$('#quoteIn').addEventListener('change', e => { const f = e.target.files[0]; e.target.value = ''; if (f) readQuoteFile(f); });
$('#report').addEventListener('click', e => { const b = e.target.closest('[data-top]'); if (!b) return; if (b.dataset.top === 'quote') $('#quoteIn').click(); if (b.dataset.top === 'payments') openPayments(); if (b.dataset.top === 'weekly') openWeekly(); if (b.dataset.top === 'meeting') openMeeting(); if (b.dataset.top === 'laXlsx') exportLookAhead(); });
$('#report').addEventListener('click', e => { const b = e.target.closest('[data-la]'); if (b) { UI.laWeeks = +b.dataset.la; renderAll(); } });
$('#report').addEventListener('change', e => {
  if (e.target.id !== 'acIn') return; const v = e.target.value.trim();
  pushUndo(); P.actualCost = v ? Math.round(vnNum(v) * 1e6) : null; renderAll();
});
$('#bSettings').innerHTML = icon('gear'); $('#bHelp').innerHTML = icon('help');
$('#bUndo').onclick = undo; $('#bRedo').onclick = redo; $('#bNew').onclick = openNew; $('#bOpen').onclick = () => $('#fileIn').click();
$('#bSave').onclick = saveJson; $('#bXlsx').onclick = exportXlsx; $('#bSettings').onclick = openSettings; $('#projBtn').onclick = openSettings; $('#bHelp').onclick = openHelp;

$('#tabs').addEventListener('click', e => {
  const b = e.target.closest('button[data-view]'); if (!b) return; UI.view = b.dataset.view;
  document.querySelectorAll('#tabs button').forEach(x => x.classList.toggle('on', x === b));
  $('#vPlan').hidden = UI.view !== 'plan'; $('#vReport').hidden = UI.view !== 'report'; $('#v3d').hidden = UI.view !== '3d'; renderAll();
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
const selMode = e => e.shiftKey ? 'range' : (e.ctrlKey || e.metaKey) ? 'toggle' : 'single';
lBody.addEventListener('focusin', e => { const row = e.target.closest('.tr'); if (row && !SELS.has(row.dataset.id)) select(row.dataset.id); });
lBody.addEventListener('click', e => {
  const tog = e.target.closest('[data-tog]');
  if (tog) { const id = tog.dataset.tog; UI.collapsed = UI.collapsed.includes(id) ? UI.collapsed.filter(x => x !== id) : [...UI.collapsed, id]; renderAll(); return; }
  const row = e.target.closest('.tr'); if (!row) return;
  const mode = selMode(e); if (mode !== 'single' || e.target.closest('.c-stt') || !e.target.matches('input')) select(row.dataset.id, mode);
  else if (SELS.size > 1 || UI.sel !== row.dataset.id) select(row.dataset.id);
});
lBody.addEventListener('mousedown', e => { if (e.shiftKey && e.target.closest('.c-stt')) e.preventDefault(); });
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
gBody.addEventListener('click', e => { const r = Math.floor((e.clientY - gBody.getBoundingClientRect().top + gBody.scrollTop) / RH); if (VIS[r] != null) select(P.tasks[VIS[r]].id, selMode(e)); });
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
  if (dlg.open || MEET) return;
  if (ctrl && !e.shiftKey && k.toLowerCase() === 'z') { if (inField && e.target.value !== e.target.defaultValue) return; e.preventDefault(); undo(); return; }
  if (ctrl && (k.toLowerCase() === 'y' || (e.shiftKey && k.toLowerCase() === 'z'))) { e.preventDefault(); redo(); return; }
  if (ctrl && k.toLowerCase() === 's') { e.preventDefault(); saveJson(); return; }
  if (ctrl && k.toLowerCase() === 'o') { e.preventDefault(); $('#fileIn').click(); return; }
  if (k === 'F1') { e.preventDefault(); openHelp(); return; }
  if (UI.view !== 'plan') return;
  if (k === 'Insert') { e.preventDefault(); if (inField) e.target.blur(); setTimeout(() => act('addBelow'), 0); return; }
  if (e.altKey && (e.shiftKey || k === 'ArrowUp' || k === 'ArrowDown') && ['ArrowRight', 'ArrowLeft', 'ArrowUp', 'ArrowDown'].includes(k)) {
    e.preventDefault(); if (inField) e.target.blur();
    setTimeout(() => act({ ArrowRight: 'indent', ArrowLeft: 'outdent', ArrowUp: 'up', ArrowDown: 'down' }[k]), 0); return;
  }
  if (e.altKey && (k === '-' || k === '=' || k === '+')) { e.preventDefault(); act(k === '-' ? 'collapseAll' : 'expandAll'); return; }
  if (inField) return;
  const lk = k.toLowerCase();
  if (ctrl && lk === 'c') { e.preventDefault(); act('copy'); return; }
  if (ctrl && lk === 'x') { e.preventDefault(); act('cut'); return; }
  if (ctrl && lk === 'd') { e.preventDefault(); act('dup'); return; }
  if (ctrl && lk === 'a') { e.preventDefault(); SELS = new Set(VIS.map(i => P.tasks[i].id)); if (!UI.sel && VIS.length) UI.sel = P.tasks[VIS[0]].id; paintSel(); return; }
  if (k === 'Escape' && SELS.size > 1) { select(UI.sel); return; }
  if (k === 'Delete' && UI.sel) { e.preventDefault(); act('del'); }
});
// Ctrl+V outside a cell: rows copied in this app, or a table copied from Excel
document.addEventListener('paste', e => {
  if (dlg.open || UI.view !== 'plan' || e.target.matches('input,textarea,select')) return;
  const txt = (e.clipboardData && e.clipboardData.getData('text/plain')) || '';
  e.preventDefault();
  if (CLIP && (!txt || txt.replace(/\r/g, '') === CLIP_TSV)) { pasteClip(); return; }
  if (/\t|\n/.test(txt.trim())) { const rows = txt.split(/\r?\n/).filter(l => l.trim()).map(l => l.split('\t')); importRows(rows, 'insert'); return; }
  if (CLIP) pasteClip(); else toast('Chưa có dòng nào được chép. Chọn dòng rồi bấm Ctrl+C, hoặc copy một bảng trong Excel.', true);
});
$('#v3d').addEventListener('click', e => {
  const m = e.target.closest('[data-d3mode]'); if (m) { UI.d3mode = m.dataset.d3mode; if (V3play) { clearInterval(V3play); V3play = 0; } render3d(); return; }
  const b = e.target.closest('[data-d3]'); if (!b) return;
  const a = b.dataset.d3;
  if (a === 'play') d3Play(); else if (a === 'config') open3dConfig(); else if (a === 'links') open3dLinks(); else if (a === 'shot') d3Shot();
  else if (a === 'reset') { V3 && V3.resetView(); V3sig = ''; render3d(); } else if (a === 'closeinfo') { V3pick = null; d3Info(); }
});
$('#v3d').addEventListener('input', e => { if (e.target.id === 'd3range') { UI.d3day = +e.target.value; d3Apply(); } });
let rzT; window.addEventListener('resize', () => { clearTimeout(rzT); rzT = setTimeout(() => { if (UI.view === 'plan') { renderGantt(); renderRes(); } }, 150); });

/* ---------- boot ---------- */
P = loadInitial();
renderHead();
renderAll();
requestAnimationFrame(() => scrollToDay(S.projStart - 3));
if (P.sample) setTimeout(() => toast('Đang mở dự án mẫu. Bấm “Dự án mới” để bắt đầu công trình của bạn.'), 400);
