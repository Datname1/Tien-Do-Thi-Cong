import test from 'node:test';
import assert from 'node:assert/strict';
import { normalizeBuilding, buildElements, autoLinks, progressAt, floorToken, floorKind } from '../src/model3d.js';
import { sampleProject, schedule, toDay } from '../src/core.js';

test('mã tầng và loại tầng', () => {
  assert.equal(floorToken('Dầm sàn lầu 2'), 'lau 2'); assert.equal(floorToken('Cột, vách tầng trệt'), 'tret');
  assert.equal(floorKind('Móng'), 'found'); assert.equal(floorKind('Tầng Mái BTCT : Chống thấm, lát gạch'), 'roof');
  assert.equal(floorKind('Mái thang máy'), 'roofbox'); assert.equal(floorKind('Lầu 1'), 'floor');
});
test('đoán tầng từ tên công việc và gắn cấu kiện', () => {
  const P = sampleProject(), S = schedule(P, true), B = normalizeBuilding(P);
  assert.deepEqual(B.floors.filter(f => f.kind === 'floor').map(f => f.name), ['Trệt', 'Lầu 1', 'Lầu 2', 'Lầu 3']);
  const { els, height } = buildElements(B); assert.ok(height > 13 && height < 15);
  assert.ok(els.find(e => e.key === 'found') && els.find(e => e.key === 'roof'));
  const L = autoLinks(B, P, S), name = id => P.tasks.find(t => t.id === id).name;
  assert.deepEqual(L['s:tret'].map(name).sort(), ['Cột, vách tầng trệt', 'Dầm sàn lầu 1'].sort());
  assert.deepEqual(L['s:lau-3'].map(name).sort(), ['Cột, vách lầu 3', 'Dầm sàn mái, sân thượng'].sort());
  assert.ok(L['w:tret'].map(name).includes('Xây tường trệt – lầu 1'));
  assert.ok(L['w:lau-2'].map(name).includes('Xây tường lầu 2 – lầu 3'));
  assert.ok(L.found.length >= 4);
});
test('tầng từ báo giá: chiều sâu theo diện tích, có tum', () => {
  const P = { tasks: [], building: { width: 5, floors: [{ name: 'Móng', area: 150 }, { name: 'Trệt', area: 100 }, { name: 'Lầu 1', area: 80 }, { name: 'Tầng Mái BTCT', area: 60 }, { name: 'Mái thang máy', area: 4 }] } };
  const B = normalizeBuilding(P);
  assert.equal(B.maxDepth, 20); assert.equal(B.floors.find(f => f.name === 'Lầu 1').depth, 16);
  const { els } = buildElements(B); assert.ok(els.find(e => e.key === 'roofbox')); assert.equal(els.find(e => e.key === 'roof').depth, 12);
});
test('mức hoàn thành theo kế hoạch và thực tế', () => {
  const P = sampleProject(), S = schedule(P, true), B = normalizeBuilding(P), L = autoLinks(B, P, S);
  const before = progressAt(L['s:tret'], P, S, toDay('2026-09-01'), 'plan'), after = progressAt(L['s:tret'], P, S, toDay('2027-03-01'), 'plan');
  assert.equal(before.p, 0); assert.equal(after.p, 1);
  const act = progressAt(L['s:lau-3'], P, S, 0, 'actual'); assert.ok(act.p < 0.5);
  assert.equal(progressAt([], P, S, 0, 'plan'), null);
});
