import test from 'node:test';
import assert from 'node:assert/strict';
import { parseQuote, quoteSections, scheduleRowsFromQuote, num, niceName, matchSections, classify } from '../src/boq.js';
import { buildTasks, schedule, cashFlow, evm, defaultPayments, setBaseline } from '../src/core.js';

// Sheet BOQ tối giản theo mẫu: No | Nội dung | Đơn vị | ... | Khối lượng (J) | Vật tư | Nhân công | ... | Thành tiền (P) | ... | Thành tiền giá vốn
const head = () => {
  const r1 = []; r1[0] = 'No.'; r1[1] = 'Nội dung'; r1[2] = 'Đơn vị'; r1[9] = 'Khối lượng'; r1[10] = 'Đơn giá (VNĐ)'; r1[15] = 'Thành tiền\n(VNĐ)'; r1[24] = 'Thành tiền\n(VNĐ)';
  const r2 = []; r2[10] = 'Vật tư'; r2[11] = 'Nhân công'; r2[18] = 'Mã công tác'; r2[19] = 'Vật tư'; r2[20] = 'Nhân công';
  return [r1, r2];
};
const line = (no, name, qty, labor, sell, cost) => { const r = []; r[0] = no; r[1] = name; r[2] = 'm3'; r[9] = qty; r[11] = labor; r[15] = sell; r[24] = cost; return r; };
const sheet = [
  ['TIÊN LƯỢNG'], ...head(), ['-1', '-2'],
  ['I', 'CÔNG TÁC CHUẨN BỊ - ĐÀO ĐẤT', null, null, null, null, null, null, null, null, null, null, null, null, null, 30e6, null, null, null, null, null, null, null, null, 20e6],
  line('1', 'Đào đất', 100, 50000, '30.000.000', 20e6),
  ['Chi tiết bỏ qua'],
  ['II', 'THI CÔNG MÓNG - ĐÀ KIỀNG', null, null, null, null, null, null, null, null, null, null, null, null, null, 90e6, null, null, null, null, null, null, null, null, 999e6],
  line('1', 'Bê tông móng', 20, 400000, 50e6, 40e6),
  line('2', 'Cốt thép móng', 2, 1000000, 40e6, 30e6),
  ['III', 'VỆ SINH CÔNG NGHIỆP', null, null, null, null, null, null, null, null, null, null, null, null, null, 5e6],
  line('1', 'Vệ sinh', 1, 0, 5e6, 4e6)
];

test('số kiểu Việt Nam', () => {
  assert.equal(num('1.234.567'), 1234567); assert.equal(num('1.234,5'), 1234.5); assert.equal(num('12,5'), 12.5);
  assert.equal(num(42), 42); assert.equal(num('abc'), 0); assert.equal(num('12400000'), 12400000);
});
test('đọc hạng mục, cộng từ dòng con và báo lệch dòng tổng', () => {
  const q = parseQuote([{ name: 'BoQ_BTCT', rows: sheet }, { name: 'BoQ_BTCT (T.PHẨM)', rows: sheet }]);
  assert.equal(q.sheets.length, 2); assert.equal(q.sheets[1].use, false);
  const s = quoteSections(q);
  assert.equal(s.length, 3);
  assert.equal(s[0].sell, 30e6); assert.equal(s[0].cost, 20e6); assert.equal(s[0].labor, 5e6);
  assert.equal(s[1].cost, 70e6);
  assert.ok(s[1].warnings.some(w => /Giá vốn ở dòng tổng/.test(w)));
  assert.ok(s[2].warnings.some(w => /để trống/.test(w)));
});
test('tên đẹp và phân loại', () => {
  assert.equal(niceName('THI CÔNG KHUNG BTCT LẦU 1'), 'Thi công khung BTCT lầu 1');
  assert.equal(classify('THI CÔNG MÓNG - ĐÀ KIỀNG'), 'found');
  assert.equal(classify('VỆ SINH CÔNG NGHIỆP'), 'clean');
});
test('dựng tiến độ từ báo giá: tuần tự hợp lý, không vòng lặp, giữ tiền', () => {
  const q = parseQuote([{ name: 'BoQ_BTCT', rows: sheet }]);
  const T = buildTasks(scheduleRowsFromQuote(quoteSections(q), { title: 'Nhà' }));
  const P = { start: '2026-12-07', workweek: 6, dataDate: '2026-12-07', holidays: [], tasks: T };
  const S = schedule(P, false);
  assert.equal(S.cycle.filter(Boolean).length, 0);
  assert.equal(S.totalVal, 125e6); assert.equal(S.totalCost, 94e6);
  const iPrep = T.findIndex(t => /đào đất/i.test(t.name)), iFound = T.findIndex(t => /móng/i.test(t.name)), iEnd = T.findIndex(t => /bàn giao/i.test(t.name));
  assert.ok(S.start[iFound] > S.finish[iPrep]); assert.equal(S.finish[iEnd], S.projFinish);
});
test('ghép theo tên', () => {
  const m = matchSections([{ name: 'THI CÔNG MÓNG - ĐÀ KIỀNG' }, { name: 'THANG MÁY' }], ['Ép cọc', 'Đổ bê tông móng, đà kiềng', 'Sơn nước']);
  assert.equal(m[0].best, 1); assert.equal(m[1].best, -1);
});
test('dòng tiền: chi theo tiến độ, thu theo đợt, EVM', () => {
  const T = buildTasks([[0, 'A', 10, '', '', 100, 100e6, 80e6], [0, 'B', 10, '1', '', 0, 100e6, 80e6]]);
  const P = { start: '2026-01-05', workweek: 7, dataDate: '2026-01-15', holidays: [], tasks: T, useDataDate: true };
  let S = schedule(P, true); setBaseline(P, S); S = schedule(P, true);
  const cf0 = cashFlow(P, S);
  const out = cf0.months.reduce((a, m) => a + m.planOut, 0), inn = cf0.months.reduce((a, m) => a + m.planIn, 0);
  assert.ok(Math.abs(out - 160e6) < 1); assert.ok(Math.abs(inn - 200e6) < 1); // chưa có đợt: thu theo giá trị làm được
  P.payments = [{ name: 'Tạm ứng', pct: 30, kind: 'start', lag: 0 }, { name: 'Xong', pct: 70, kind: 'task', taskId: T[1].id, lag: 0 }];
  const cf = cashFlow(P, S);
  assert.equal(cf.byStages, true);
  assert.ok(Math.abs(cf.months.reduce((a, m) => a + m.fcIn, 0) - 200e6) < 1);
  assert.equal(cf.months[0].key, '2026-01');
  const e = evm(P, S);
  assert.equal(e.use, 'cost'); assert.equal(e.BAC, 160e6); assert.equal(e.EV, 80e6); assert.ok(e.SPI > 0.99);
  assert.ok(defaultPayments(P, S).reduce((a, p) => a + p.pct, 0) <= 100);
});
