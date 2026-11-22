import test from 'node:test';
import assert from 'node:assert/strict';
import { schedule, buildTasks, toDay, fmtD, parsePreds, sampleProject } from '../src/core.js';

const base = rows => ({ name: 't', start: '2026-10-05', workweek: 6, dataDate: '2026-10-05', holidays: [], tasks: buildTasks(rows) });

test('FS nối tiếp, bỏ qua Chủ nhật', () => {
  const P = base([[0, 'A', 6, '', '', 0], [0, 'B', 2, '1', '', 0]]);
  const S = schedule(P, false);
  assert.equal(fmtD(S.finish[0]), '10/10/2026'); // T2..T7
  assert.equal(fmtD(S.start[1]), '12/10/2026');  // bỏ CN 11/10
  assert.ok(S.crit[0] && S.crit[1]);
});
test('SS+lag và dự trữ', () => {
  const P = base([[0, 'A', 10, '', '', 0], [0, 'B', 3, '1SS+2', '', 0]]);
  const S = schedule(P, false);
  assert.equal(S.es[1], 2); assert.equal(S.tf[1], 5); assert.equal(S.crit[1], false);
});
test('FF và SF', () => {
  const P = base([[0, 'A', 5, '', '', 0], [0, 'B', 2, '1FF+1', '', 0], [0, 'C', 3, '1SF', '', 0]]);
  const S = schedule(P, false);
  assert.equal(S.ef[1], 6); assert.equal(S.ef[2], 3);
});
test('ngày lễ bị bỏ qua', () => {
  const P = base([[0, 'A', 3, '', '', 0]]); P.holidays = [{ s: '2026-10-06', f: '2026-10-07' }];
  const S = schedule(P, false);
  assert.equal(fmtD(S.finish[0]), '09/10/2026');
});
test('dời việc chưa làm ra sau ngày báo cáo', () => {
  const P = base([[0, 'A', 4, '', '', 0]]); P.dataDate = '2026-10-12';
  const S = schedule(P, true);
  assert.equal(fmtD(S.start[0]), '12/10/2026');
});
test('vòng lặp không treo, chỉ đánh dấu điểm cắt', () => {
  const P = base([[0, 'A', 2, '2', '', 0], [0, 'B', 2, '1', '', 0], [0, 'C', 2, '2', '', 0]]);
  const S = schedule(P, false);
  assert.equal(S.cycle.filter(Boolean).length, 1);
});
test('cú pháp liên kết', () => {
  const T = buildTasks([[0, 'a', 1], [0, 'b', 1], [0, 'c', 1]]);
  const r = parsePreds('1, 2SS+3, 1ff-1, x', T, 2);
  assert.deepEqual(r.links.map(l => l.type + l.lag), ['FS0', 'SS3', 'FF-1']); assert.deepEqual(r.errs, ['x']);
});
test('dự án mẫu tính được', () => {
  const P = sampleProject(); const S = schedule(P, true);
  assert.equal(S.leaves.length, 32); assert.ok(S.projFinish > toDay('2027-01-01'));
});
