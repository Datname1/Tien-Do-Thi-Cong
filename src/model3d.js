/* Mô hình khối 3D: tầng, cấu kiện, tự gắn công việc và tính mức hoàn thành theo ngày.
   Không phụ thuộc three.js để kiểm thử được. Đơn vị: mét. */
import { fold } from './boq.js';

export function floorKind(name) {
  const f = fold(name);
  if (/mong|tang ham|\bham\b|coc/.test(f)) return 'found';
  if (/mai thang|thang may|\btum\b|mai che cau thang/.test(f)) return 'roofbox';
  if (/\bmai\b|san thuong|tang thuong/.test(f) && !/lau|tret|lung/.test(f)) return 'roof';
  return 'floor';
}
// Mã tầng để so tên công việc: "tret", "lung", "lau 2", "mai"
export function floorToken(name) {
  const f = fold(name), m = f.match(/lau\s*(\d+)/);
  if (m) return 'lau ' + m[1];
  if (/tret/.test(f)) return 'tret';
  if (/lung/.test(f)) return 'lung';
  if (/\bmai\b|san thuong|cat noc/.test(f)) return 'mai';
  return null;
}
const hasToken = (f, tok) => tok && new RegExp(`(^|[^a-z0-9])${tok.replace(' ', '\\s*')}([^0-9]|$)`).test(f);

/* Tầng của công trình: lấy từ báo giá (bảng diện tích), nếu không có thì đoán từ tên công việc */
export function normalizeBuilding(P) {
  const B = P.building || {}, width = +B.width || 5;
  let floors = (B.floors || []).filter(x => x && String(x.name || '').trim()).map(x => ({ name: String(x.name).trim(), area: +x.area || 0, h: +x.h || 0, kind: x.kind || floorKind(x.name) }));
  if (!floors.some(x => x.kind === 'floor')) {
    const names = P.tasks.map(t => fold(t.name)), set = new Set();
    names.forEach(n => { const m = n.match(/lau\s*(\d+)/g); if (m) m.forEach(x => set.add(+x.replace(/\D/g, ''))); });
    const maxLau = set.size ? Math.max(...set) : 2;
    floors = [{ name: 'Móng', area: 0, kind: 'found' }, { name: 'Trệt', area: 0, kind: 'floor' }];
    if (names.some(n => /\blung\b/.test(n))) floors.push({ name: 'Lửng', area: 0, kind: 'floor' });
    for (let k = 1; k <= maxLau; k++) floors.push({ name: 'Lầu ' + k, area: 0, kind: 'floor' });
    floors.push({ name: 'Sân thượng, mái', area: 0, kind: 'roof' });
  }
  if (!floors.some(x => x.kind === 'found')) floors.unshift({ name: 'Móng', area: 0, kind: 'found' });
  let firstFloor = true;
  floors.forEach(x => {
    if (!x.h) x.h = x.kind === 'found' ? 1.2 : x.kind === 'roof' ? 1.1 : x.kind === 'roofbox' ? 2.8 : (firstFloor ? 3.8 : /lung/.test(fold(x.name)) ? 3.0 : 3.4);
    if (x.kind === 'floor') firstFloor = false;
    x.key = fold(x.name).replace(/[^a-z0-9]+/g, '-');
  });
  const depths = floors.filter(x => x.kind === 'floor' && x.area > 0).map(x => x.area / width);
  const maxDepth = Math.max(8, Math.min(40, depths.length ? Math.max(...depths) : 18));
  floors.forEach(x => { x.depth = x.area > 0 ? Math.max(3, Math.min(maxDepth, x.area / width)) : maxDepth; });
  return { width, maxDepth, floors, links: B.links || null };
}

/* Cấu kiện theo tầng: móng, khung (cột + sàn phía trên), tường, mái, tum; "finish" là lớp hoàn thiện mặt ngoài */
export function buildElements(B) {
  const els = [], ups = B.floors.filter(x => x.kind === 'floor');
  let z = 0;
  const found = B.floors.find(x => x.kind === 'found');
  els.push({ key: 'found', kind: 'found', label: found ? found.name : 'Móng', z: -(found ? found.h : 1.2), h: found ? found.h : 1.2, depth: B.maxDepth });
  ups.forEach((x, k) => {
    els.push({ key: 's:' + x.key, kind: 'struct', label: 'Khung ' + x.name, floor: x.name, z, h: x.h, depth: x.depth, k });
    els.push({ key: 'w:' + x.key, kind: 'wall', label: 'Tường ' + x.name, floor: x.name, z, h: x.h, depth: x.depth, k });
    z += x.h;
  });
  const roof = B.floors.find(x => x.kind === 'roof');
  els.push({ key: 'roof', kind: 'roof', label: roof ? roof.name : 'Mái', z, h: roof ? roof.h : 1.1, depth: roof && roof.area ? roof.depth : (ups.length ? ups[ups.length - 1].depth : B.maxDepth) });
  const box = B.floors.find(x => x.kind === 'roofbox');
  if (box) els.push({ key: 'roofbox', kind: 'roofbox', label: box.name, z, h: box.h, depth: Math.min(4, B.maxDepth / 3) });
  els.push({ key: 'finish', kind: 'finish', label: 'Hoàn thiện mặt ngoài', z: 0, h: z, depth: B.maxDepth });
  return { els, height: z };
}

const STRUCT = /khung|cot|vach|dam|san|be tong|btct|cop pha|cot thep/;
/* Tự gắn công việc cho từng cấu kiện theo tên */
export function autoLinks(B, P, S) {
  const T = P.tasks, names = T.map(t => fold(t.name)), leaf = i => !S.st.isSum[i];
  const ups = B.floors.filter(x => x.kind === 'floor'), links = {};
  const pick = test => T.map((t, i) => i).filter(i => leaf(i) && test(names[i])).map(i => T[i].id);
  links.found = pick(n => /mong|da kieng|coc|dao dat|tang ham|be ngam|ham phan/.test(n) && !/mai/.test(n));
  ups.forEach((x, k) => {
    const tok = floorToken(x.name), next = k + 1 < ups.length ? floorToken(ups[k + 1].name) : 'mai';
    links['s:' + x.key] = pick(n => (STRUCT.test(n) && /khung|cot|vach/.test(n) && hasToken(n, tok)) || (/dam|san/.test(n) && !/khung/.test(n) && hasToken(n, next) && STRUCT.test(n)));
    links['w:' + x.key] = pick(n => /xay|tuong/.test(n) && hasToken(n, tok));
  });
  const genericWall = pick(n => /xay|to trat trong|phan tho|hoan thien/.test(n) && !floorToken(n) && !/ban giao|son/.test(n));
  ups.forEach(x => { if (!links['w:' + x.key].length) links['w:' + x.key] = genericWall.slice(0, 1); });
  links.roof = pick(n => (/mai|san thuong|cat noc/.test(n) && STRUCT.test(n)) || /chong tham mai|cat noc/.test(n));
  links.roofbox = pick(n => /tum|thang may/.test(n) && !/lap dat thiet bi/.test(n));
  links.finish = pick(n => /son|to trat ngoai|op lat|cua|hoan thien|nhom kinh/.test(n) && !/ban giao/.test(n));
  return links;
}
export const linksFor = (B, P, S) => Object.assign(autoLinks(B, P, S), B.links || {});

/* Mức hoàn thành 0..1 của một cấu kiện: theo kế hoạch tại một ngày, hoặc thực tế tại ngày báo cáo */
export function progressAt(ids, P, S, day, mode) {
  const st = S.st, idx = new Set();
  (ids || []).forEach(id => { const i = S.idIdx.get(id); if (i == null) return; if (st.isSum[i]) { for (let j = i; j <= st.end[i]; j++) if (!st.isSum[j]) idx.add(j); } else idx.add(i); });
  if (!idx.size) return null;
  let W = 0, done = 0, late = false, a = Infinity, b = -Infinity;
  idx.forEach(i => {
    const w = Math.max(1, +P.tasks[i].dur || 0); W += w;
    const s = mode === 'plan' ? (S.blS[i] ?? S.start[i]) : S.start[i], f = mode === 'plan' ? (S.blF[i] ?? S.finish[i]) : S.finish[i];
    a = Math.min(a, s); b = Math.max(b, f);
    if (mode === 'plan') done += w * Math.max(0, Math.min(1, (day - s + 1) / (f - s + 1)));
    else { done += w * S.pct[i] / 100; late = late || S.late[i]; }
  });
  return { p: done / W, late, start: a, finish: b, n: idx.size };
}
