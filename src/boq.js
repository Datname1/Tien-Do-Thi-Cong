/* Đọc bảng báo giá / BOQ (dạng mảng các dòng của từng sheet Excel) và dựng tiến độ từ đó.
   Không phụ thuộc giao diện: nhận dữ liệu thô, trả về cấu trúc thuần để kiểm thử được. */

const ROMAN = /^(?=[IVXLC]+$)M{0,3}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$/;
const INT = /^\d{1,3}(\.\d{1,2})?$/;

export const fold = s => String(s ?? '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/\s+/g, ' ').trim();

// Số kiểu Việt Nam: "1.234.567", "1.234,5", "12,5", số thật của Excel
export function num(v) {
  if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
  let s = String(v ?? '').trim().replace(/\s|đ|vnđ|vnd/gi, '');
  if (!s || !/^-?[\d.,]+$/.test(s)) return 0;
  if (s.includes('.') && s.includes(',')) s = s.lastIndexOf(',') > s.lastIndexOf('.') ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
  else if (s.includes(',')) s = /^-?\d{1,3}(,\d{3})+$/.test(s) ? s.replace(/,/g, '') : s.replace(',', '.');
  else if (/^-?\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, '');
  const n = parseFloat(s); return Number.isFinite(n) ? n : 0;
}

const cellText = v => (v == null ? '' : String(v)).replace(/\s+/g, ' ').trim();

// Tìm dòng tiêu đề và vị trí các cột trong một sheet báo giá
function findColumns(rows) {
  for (let r = 0; r < Math.min(rows.length, 45); r++) {
    const row = rows[r] || [];
    const nameCol = row.findIndex(c => /^(noi dung|ten cong viec|noi dung cong viec|hang muc|dien giai)/.test(fold(c)));
    if (nameCol < 0) continue;
    const band = [r, r + 1, r + 2].map(k => rows[k] || []);
    const cols = { header: r, name: nameCol, no: -1, unit: -1, qty: -1, sell: -1, cost: -1, labor: -1 };
    const width = Math.max(...band.map(b => b.length));
    const total = [];
    for (let c = 0; c < width; c++) {
      const t = band.map(b => fold(b[c])).join(' | ');
      if (cols.no < 0 && /(^|\| )(no\.?|stt|tt)( |$|\|)/.test(t) && c < nameCol) cols.no = c;
      if (cols.unit < 0 && /don vi/.test(t) && !/don gia/.test(t)) cols.unit = c;
      if (cols.qty < 0 && /khoi luong/.test(t) && !/tinh toan/.test(t)) cols.qty = c;
      if (/thanh tien/.test(t)) total.push(c);
      if (cols.labor < 0 && /nhan cong/.test(t)) cols.labor = c;
    }
    if (!total.length) continue;
    if (cols.no < 0) cols.no = Math.max(0, nameCol - 1);
    cols.sell = total[0]; cols.cost = total.length > 1 ? total[1] : -1;
    if (cols.labor > cols.sell) cols.labor = -1; // chỉ lấy đơn giá nhân công của khối giá bán
    return cols;
  }
  return null;
}

const ACRO = new Set(['btct', 'bt', 'mep', 'pccc', 'hd', 'ups', 'led', 'ppr', 'upvc', 'pvc', 'inox', 'wc', 'tk', 'gcld']);
export function niceName(s) {
  const t = cellText(s); if (!t) return '';
  if (t !== t.toUpperCase()) return t;
  const lower = t.toLowerCase().split(' ').map(w => ACRO.has(fold(w)) ? w.toUpperCase() : w).join(' ');
  return lower.charAt(0).toUpperCase() + lower.slice(1);
}

/* Đọc một sheet BOQ: hạng mục (số La Mã) và đầu việc (số thứ tự) */
export function parseBoqSheet(name, rows) {
  const cols = findColumns(rows); if (!cols) return null;
  const sections = [];
  let cur = null;
  for (let r = cols.header + 1; r < rows.length; r++) {
    const row = rows[r] || [], no = cellText(row[cols.no]), title = cellText(row[cols.name]);
    if (!title) continue;
    if (ROMAN.test(no.toUpperCase()) && /[a-zA-ZÀ-ỹ]/.test(title)) {
      cur = { sheet: name, row: r + 1, code: no.toUpperCase(), name: title, headSell: num(row[cols.sell]), headCost: cols.cost >= 0 ? num(row[cols.cost]) : 0, items: [] };
      sections.push(cur); continue;
    }
    if (cur && INT.test(no)) {
      const qty = cols.qty >= 0 ? num(row[cols.qty]) : 0;
      cur.items.push({
        row: r + 1, no, name: title, unit: cols.unit >= 0 ? cellText(row[cols.unit]) : '', qty,
        sell: num(row[cols.sell]), cost: cols.cost >= 0 ? num(row[cols.cost]) : 0,
        labor: cols.labor >= 0 ? qty * num(row[cols.labor]) : 0
      });
    }
  }
  sections.forEach(s => {
    const sum = k => s.items.reduce((a, it) => a + it[k], 0);
    s.itemSell = sum('sell'); s.itemCost = sum('cost'); s.labor = sum('labor');
    s.sell = s.items.length ? s.itemSell : s.headSell;
    s.cost = s.items.length ? s.itemCost : s.headCost;
    s.warnings = [];
    const off = (a, b) => Math.abs(a - b) > Math.max(1000, Math.abs(b) * 0.001);
    if (s.items.length && s.headSell && off(s.headSell, s.itemSell)) s.warnings.push(`Giá bán ở dòng tổng (${fmtM(s.headSell)}) khác tổng các dòng con (${fmtM(s.itemSell)})`);
    if (cols.cost >= 0 && s.items.length) {
      if (!s.headCost && s.itemCost) s.warnings.push(`Dòng tổng giá vốn để trống, tổng các dòng con là ${fmtM(s.itemCost)}`);
      else if (s.headCost && off(s.headCost, s.itemCost)) s.warnings.push(`Giá vốn ở dòng tổng (${fmtM(s.headCost)}) khác tổng các dòng con (${fmtM(s.itemCost)})`);
    }
    if (cols.cost >= 0 && !s.cost && s.sell) s.warnings.push('Chưa có giá vốn');
    if (s.cost > s.sell && s.sell > 0) s.warnings.push('Giá vốn lớn hơn giá bán');
  });
  return { name, cols, sections: sections.filter(s => s.sell || s.cost || s.items.length), hasCost: cols.cost >= 0 };
}

export const fmtM = v => (v / 1e6).toLocaleString('vi-VN', { maximumFractionDigits: 1 }) + ' tr';

// Tổng giá trị dự kiến ghi trong bảng tổng hợp (nếu có) để đối chiếu
function findGrandTotal(rows) {
  let best = 0;
  rows.forEach(row => (row || []).forEach((c, k) => {
    const t = fold(c);
    if (/^tong (chi phi|gia tri) du kien/.test(t) || /^tong cong (chi phi|gia tri)/.test(t)) {
      for (let j = k + 1; j < row.length; j++) { const v = num(row[j]); if (v > best) { best = v; break; } }
    }
  }));
  return best;
}

// Diện tích từng tầng từ "bảng tính diện tích" (dùng cho mô hình khối 3D sau này)
function findFloors(rows) {
  const out = []; let areaCol = -1, nameCol = -1, noCol = 0, start = -1;
  for (let r = 0; r < Math.min(rows.length, 40) && start < 0; r++) {
    const row = rows[r] || [];
    const a = row.findIndex(c => fold(c) === 'dien tich'); const n = row.findIndex(c => /^ten cong viec|^noi dung|^hang muc/.test(fold(c)));
    if (a >= 0 && n >= 0) { areaCol = a; nameCol = n; noCol = Math.max(0, n - 1); start = r + 1; }
  }
  if (start < 0) return out;
  for (let r = start; r < rows.length; r++) {
    const row = rows[r] || [], no = cellText(row[noCol]), name = cellText(row[nameCol]);
    if (!name) { if (out.length) break; continue; }
    if (/^\d+$/.test(no)) out.push({ name, area: num(row[areaCol]) });
    else if (out.length && /^[IVX]+\.?/.test(name)) break;
  }
  return out;
}

const SKIP_SHEET = /t\.?\s*ph[aẩ]m|copy|\bold\b|\bcu\b|nhap|draft/i;

/* sheets: [{ name, rows }] → kết quả đọc cả file */
export function parseQuote(sheets) {
  const out = { sheets: [], grandTotal: 0, floors: [] };
  sheets.forEach(({ name, rows }) => {
    const f = fold(name);
    if (/bang th|tong hop/.test(f)) out.grandTotal = Math.max(out.grandTotal, findGrandTotal(rows));
    if (/bang dt|dien tich/.test(f) && !out.floors.length) out.floors = findFloors(rows);
    const p = parseBoqSheet(name, rows);
    if (p && p.sections.length) { p.use = !SKIP_SHEET.test(name); out.sheets.push(p); }
  });
  return out;
}
export const quoteSections = q => q.sheets.filter(s => s.use).flatMap(s => s.sections);

/* ---------- ước lượng thời gian & dựng tiến độ ---------- */
export const DAY_RATE = 580000; // đơn giá nhân công/ngày công tham khảo
export function estimateDays(labor, sell, crew = 6) {
  if (labor > 0) return Math.max(2, Math.min(45, Math.round(labor / DAY_RATE / crew)));
  return Math.max(2, Math.min(10, Math.round(sell / 60e6) + 2)); // cung cấp/lắp đặt: chủ yếu là vật tư
}
export function classify(name) {
  const f = fold(name);
  if (/ve sinh cong nghiep/.test(f)) return 'clean';
  if (/chuan bi|dao dat|mat bang/.test(f)) return 'prep';
  if (/\bcoc\b/.test(f)) return 'pile';
  if (/mong|da kieng|tang ham|\bham\b/.test(f)) return 'found';
  if (/khung|btct|dam san|cot vach/.test(f)) return 'frame';
  if (/xay|to trat|hoan thien|phan tho/.test(f)) return 'mason';
  if (/dien nuoc cong trinh|ong ngam|dien nuoc/.test(f)) return 'rough';
  if (/thang may/.test(f)) return 'lift';
  return 'finish';
}

/* Trả về các dòng [cấp, tên, ngày, LK trước, nguồn lực, %, giá trị, chi phí] cho buildTasks */
export function scheduleRowsFromQuote(sections, { title = 'Công trình', detail = false } = {}) {
  const ORDER = ['prep', 'pile', 'found', 'frame', 'mason', 'rough', 'finish', 'lift', 'clean'];
  sections = sections.map((s, k) => ({ s, k, o: ORDER.indexOf(classify(s.name)) })).sort((x, y) => x.o - y.o || x.k - y.k).map(x => x.s);
  const rows = [[0, title.toUpperCase(), 0, '', '', 0, 0, 0]];
  const pos = []; // số STT (1-based) của từng hạng mục
  const info = [];
  sections.forEach(s => {
    const kind = classify(s.name), lvl = 1;
    const dur = estimateDays(s.labor, s.sell);
    if (detail && s.items.length > 1) {
      rows.push([lvl, niceName(s.name), 0, '', '', 0, 0, 0]); const at = rows.length; pos.push(at);
      // các đầu việc trong một hạng mục chạy song song, không dài hơn cả hạng mục
      s.items.forEach(it => {
        rows.push([lvl + 1, niceName(it.name), Math.min(dur, estimateDays(it.labor, it.sell, 4)), '', '', 0, Math.round(it.sell), Math.round(it.cost)]);
      });
      info.push({ kind, at, dur });
    } else {
      rows.push([lvl, niceName(s.name), dur, '', '', 0, Math.round(s.sell), Math.round(s.cost)]);
      pos.push(rows.length); info.push({ kind, at: rows.length, dur });
    }
  });
  const of = k => info.filter(x => x.kind === k);
  const structure = info.filter(x => ['prep', 'pile', 'found', 'frame'].includes(x.kind));
  const link = (x, s) => { const r = rows[x.at - 1]; r[3] = r[3] ? r[3] + ', ' + s : s; };
  structure.forEach((x, k) => { if (k) link(x, String(structure[k - 1].at)); });
  const frames = of('frame'), lastStruct = structure[structure.length - 1];
  const masonry = of('mason')[0];
  if (masonry) link(masonry, String((frames[Math.min(2, frames.length - 1)] || lastStruct || {}).at || ''));
  of('mason').slice(1).forEach(x => link(x, `${masonry.at}SS+${Math.max(2, Math.round(masonry.dur / 3))}`));
  of('rough').forEach(x => link(x, masonry ? `${masonry.at}SS+3` : String((lastStruct || {}).at || '')));
  const lateStart = masonry ? `${masonry.at}SS+${Math.max(3, Math.round(masonry.dur / 2))}` : '';
  of('finish').forEach(x => { if (lateStart) link(x, lateStart); if (lastStruct) link(x, String(lastStruct.at)); });
  of('lift').forEach(x => { if (lastStruct) link(x, String(lastStruct.at)); });
  const tail = info.filter(x => ['finish', 'lift', 'rough', 'mason'].includes(x.kind));
  of('clean').forEach(x => link(x, (tail.length ? tail : structure.slice(-1)).map(t => t.at).join(', ')));
  // bàn giao sau mọi hạng mục chưa có việc nối tiếp
  const used = new Set(rows.flatMap(r => String(r[3] || '').split(', ').map(v => parseInt(v, 10)).filter(Boolean)));
  const ends = info.filter(x => !used.has(x.at));
  rows.push([1, 'Nghiệm thu, bàn giao', 0, ends.map(t => t.at).join(', '), '', 0, 0, 0]);
  rows.forEach(r => { r[3] = String(r[3] || '').split(', ').filter(Boolean).filter((v, k, a) => a.indexOf(v) === k).join(', '); });
  return rows;
}

/* Ghép hạng mục báo giá với công việc có sẵn theo tên */
const STOP = new Set(['thi', 'cong', 'cung', 'cap', 'va', 'lap', 'dat', 'he', 'thong', 'vat', 'tu', 'cac', 'phan', 'cho', 'tang', 'hang', 'muc', 'cong', 'tac', 'nha', 'tri']);
const tokens = s => new Set(fold(s).replace(/[^a-z0-9 ]/g, ' ').split(' ').filter(w => (w.length > 1 || /\d/.test(w)) && !STOP.has(w)));
export function similarity(a, b) {
  const A = tokens(a), B = tokens(b); if (!A.size || !B.size) return 0;
  let inter = 0; A.forEach(w => { if (B.has(w)) inter++; });
  return inter / (A.size + B.size - inter);
}
export function matchSections(sections, taskNames) {
  return sections.map(s => {
    let best = -1, score = 0;
    taskNames.forEach((n, i) => { const v = similarity(s.name, n); if (v > score) { score = v; best = i; } });
    return { best: score >= 0.2 ? best : -1, score };
  });
}
