/* Hiển thị mô hình khối 3D bằng three.js. Được tải riêng khi mở tab 3D. */
import * as THREE from 'three';
import { OrbitControls } from 'three/examples/jsm/controls/OrbitControls.js';

const css = n => getComputedStyle(document.documentElement).getPropertyValue(n).trim() || '#888';
const col = v => new THREE.Color(v);

export function mount(container, { onPick } = {}) {
  const renderer = new THREE.WebGLRenderer({ antialias: true, preserveDrawingBuffer: true });
  renderer.setPixelRatio(Math.min(window.devicePixelRatio || 1, 2));
  renderer.shadowMap.enabled = true; renderer.shadowMap.type = THREE.PCFSoftShadowMap;
  container.appendChild(renderer.domElement);
  const scene = new THREE.Scene();
  const camera = new THREE.PerspectiveCamera(38, 1, 0.1, 500);
  const controls = new OrbitControls(camera, renderer.domElement);
  controls.enableDamping = true; controls.maxPolarAngle = Math.PI * 0.49; controls.minDistance = 6; controls.maxDistance = 160;
  scene.add(new THREE.HemisphereLight(0xffffff, 0x8a8f99, 1.1));
  const sun = new THREE.DirectionalLight(0xffffff, 1.6); sun.position.set(-16, 34, -26); sun.castShadow = true;
  sun.shadow.mapSize.set(2048, 2048); Object.assign(sun.shadow.camera, { left: -40, right: 40, top: 40, bottom: -40, near: 1, far: 120 });
  scene.add(sun);
  const root = new THREE.Group(); scene.add(root);
  let ground = null, grid = null, model = null, raf = 0, dirty = true;

  const mats = {};
  const mat = (key, color, opts = {}) => {
    if (!mats[key]) mats[key] = new THREE.MeshStandardMaterial({ color, roughness: 0.85, metalness: 0.02, ...opts });
    return mats[key];
  };
  const theme = () => {
    const dark = matchMedia('(prefers-color-scheme: dark)').matches && document.documentElement.dataset.theme !== 'light' || document.documentElement.dataset.theme === 'dark';
    scene.background = col(dark ? '#121821' : '#eef2f6');
    return { dark, crit: css('--crit'), accent: css('--accent-strong'), ok: css('--ok') };
  };

  function box(w, h, d, material, x, y, z) {
    const m = new THREE.Mesh(new THREE.BoxGeometry(w, h, d), material);
    m.position.set(x + w / 2, y + h / 2, z + d / 2); m.castShadow = true; m.receiveShadow = true; return m;
  }
  function ghost(w, h, d, x, y, z, color) {
    const g = new THREE.LineSegments(new THREE.EdgesGeometry(new THREE.BoxGeometry(w, h, d)), new THREE.LineBasicMaterial({ color, transparent: true, opacity: 0.35 }));
    g.position.set(x + w / 2, y + h / 2, z + d / 2); return g;
  }

  /* Dựng lại toàn bộ khối khi cấu hình tầng thay đổi */
  function build(B, els) {
    root.clear(); model = { parts: {}, B, els };
    const W = B.width, t = theme();
    const lineCol = t.dark ? '#7d8aa0' : '#8794a8';
    if (ground) scene.remove(ground); if (grid) scene.remove(grid);
    ground = new THREE.Mesh(new THREE.PlaneGeometry(120, 120), new THREE.MeshStandardMaterial({ color: t.dark ? '#1b232e' : '#dfe5ec', transparent: true, opacity: 0.72, roughness: 1 }));
    ground.rotation.x = -Math.PI / 2; ground.receiveShadow = true; scene.add(ground);
    grid = new THREE.GridHelper(120, 60, t.dark ? 0x2a3442 : 0xc9d1db, t.dark ? 0x222b37 : 0xd5dce4); grid.position.y = 0.005; scene.add(grid);
    els.forEach(e => {
      const g = new THREE.Group(); g.userData.key = e.key; const parts = { group: g, e, solids: [] };
      const D = e.depth;
      if (e.kind === 'found') {
        const s = box(W + 0.8, e.h, D + 0.8, mat('found', '#8d8a83'), -0.4, e.z, -0.4); parts.solids.push(s);
        g.add(ghost(W + 0.8, e.h, D + 0.8, -0.4, e.z, -0.4, lineCol));
      } else if (e.kind === 'struct') {
        const xs = W > 7 ? [0, W / 2 - 0.12, W - 0.25] : [0, W - 0.25], nz = Math.max(2, Math.round(D / 4) + 1);
        const cols = [];
        xs.forEach(x => { for (let k = 0; k < nz; k++) { const z = k * (D - 0.25) / (nz - 1); const c = box(0.25, e.h - 0.2, 0.25, mat('struct', '#b9bcc2'), x, e.z, z); cols.push(c); } });
        const slab = box(W, 0.2, D, mat('struct', '#b9bcc2'), 0, e.z + e.h - 0.2, 0);
        parts.cols = cols; parts.slab = slab; parts.solids.push(...cols, slab);
        g.add(ghost(W, e.h, D, 0, e.z, 0, lineCol));
      } else if (e.kind === 'wall') {
        const th = 0.12, h = e.h - 0.2, wm = mat('wall:' + e.key, '#c8764f');
        const walls = [box(th, h, D, wm, 0.02, e.z, 0), box(th, h, D, wm, W - th - 0.02, e.z, 0), box(W, h, th, wm, 0, e.z, D - th - 0.02), box(W, 1.0, th, wm, 0, e.z, 0.02)];
        const glass = box(W - 0.6, h - 1.2, 0.05, mat('glass', '#8fb8d8', { transparent: true, opacity: 0.55, roughness: 0.1, metalness: 0.3 }), 0.3, e.z + 1.0, 0.06);
        parts.walls = walls; parts.glass = glass; parts.solids.push(...walls, glass); parts.wm = wm;
      } else if (e.kind === 'roof') {
        const rm = mat('roof', '#b9bcc2');
        const pieces = [box(W, 0.12, D, rm, 0, e.z, 0), box(0.15, e.h, D, rm, 0, e.z, 0), box(0.15, e.h, D, rm, W - 0.15, e.z, 0), box(W, e.h, 0.15, rm, 0, e.z, 0), box(W, e.h, 0.15, rm, 0, e.z, D - 0.15)];
        parts.solids.push(...pieces); g.add(ghost(W, e.h, D, 0, e.z, 0, lineCol));
      } else if (e.kind === 'roofbox') {
        const bw = Math.min(3, W * 0.6), s = box(bw, e.h, e.depth, mat('roofbox', '#b9bcc2'), (W - bw) / 2, e.z, Math.max(0, (model.B.maxDepth - e.depth) * 0.6));
        parts.solids.push(s); g.add(ghost(bw, e.h, e.depth, (W - bw) / 2, e.z, Math.max(0, (model.B.maxDepth - e.depth) * 0.6), lineCol));
      }
      parts.solids.forEach(s => { s.userData.key = e.key; g.add(s); });
      root.add(g); model.parts[e.key] = parts;
    });
    const H = Math.max(6, els.reduce((a, e) => Math.max(a, e.z + e.h), 0));
    root.position.set(-W / 2, 0, -B.maxDepth / 2);
    // canh camera cho vừa khối nhà: khoảng cách theo bán kính bao quanh
    const r = 0.5 * Math.hypot(W, B.maxDepth, H), dist = r / Math.sin(THREE.MathUtils.degToRad(camera.fov / 2)) * 0.95;
    const dir = new THREE.Vector3(0.8, 0.62, -1.15).normalize(); // nhìn từ mặt tiền (z = 0)
    controls.target.set(0, H * 0.4, 0); camera.position.copy(controls.target).addScaledVector(dir, dist); controls.update();
    dirty = true;
  }

  /* Tô màu và dựng dần theo mức hoàn thành */
  function apply(prog, { mode } = {}) {
    if (!model) return; const t = theme();
    const state = p => p == null ? 'none' : p.p <= 0.001 ? 'todo' : p.p >= 0.999 ? 'done' : 'doing';
    const finish = prog.finish ? prog.finish.p : 0;
    Object.values(model.parts).forEach(parts => {
      const e = parts.e, p = prog[e.key], s = state(p);
      const visible = s === 'none' ? (e.kind !== 'finish') : s !== 'todo';
      const late = mode === 'actual' && p && p.late && s !== 'done';
      const tint = late ? t.crit : s === 'doing' ? t.accent : null;
      if (e.kind === 'struct') {
        const k = s === 'none' ? 1 : Math.min(1, p.p * 2), ks = s === 'none' ? 1 : Math.max(0, (p.p - 0.5) * 2);
        parts.cols.forEach(c => { c.visible = visible && k > 0; c.scale.y = Math.max(0.001, k); c.position.y = e.z + (e.h - 0.2) * k / 2; });
        parts.slab.visible = visible && ks > 0; parts.slab.scale.z = Math.max(0.001, ks); parts.slab.position.z = e.depth * ks / 2;
        parts.solids.forEach(m => { m.material = tint ? mat('tint:' + tint, tint) : mat('struct', '#b9bcc2'); });
      } else if (e.kind === 'wall') {
        const k = s === 'none' ? 1 : p.p;
        parts.walls.forEach((w, i) => { w.visible = visible; const full = i === 3 ? 1.0 : e.h - 0.2; w.scale.y = Math.max(0.001, k); w.position.y = e.z + full * k / 2; });
        const base = new THREE.Color('#c8764f').lerp(new THREE.Color(t.dark ? '#d9d2c4' : '#efe7d8'), Math.min(1, finish));
        parts.wm.color.copy(late ? col(t.crit) : s === 'doing' ? col(t.accent) : base);
        parts.walls.forEach(w => { w.material = parts.wm; });
        parts.glass.visible = visible && k >= 0.999 && finish >= 0.6;
      } else if (e.kind !== 'finish') {
        parts.solids.forEach(m => { m.visible = visible; m.material = tint ? mat('tint:' + tint, tint) : mat(e.kind, e.kind === 'found' ? '#8d8a83' : '#b9bcc2'); });
      }
    });
    dirty = true;
  }

  const ray = new THREE.Raycaster(), ptr = new THREE.Vector2(); let down = null;
  renderer.domElement.addEventListener('pointerdown', e => { down = [e.clientX, e.clientY]; });
  renderer.domElement.addEventListener('pointerup', e => {
    if (!down || Math.hypot(e.clientX - down[0], e.clientY - down[1]) > 4) return;
    const r = renderer.domElement.getBoundingClientRect();
    ptr.set(((e.clientX - r.left) / r.width) * 2 - 1, -((e.clientY - r.top) / r.height) * 2 + 1);
    ray.setFromCamera(ptr, camera);
    const hit = ray.intersectObjects(root.children, true).find(h => h.object.visible && h.object.isMesh && h.object.userData.key);
    if (onPick) onPick(hit ? hit.object.userData.key : null, e);
  });
  controls.addEventListener('change', () => { dirty = true; });
  const resize = () => { const w = container.clientWidth || 600, h = container.clientHeight || 400; renderer.setSize(w, h, false); renderer.domElement.style.width = '100%'; renderer.domElement.style.height = '100%'; camera.aspect = w / h; camera.updateProjectionMatrix(); dirty = true; };
  const ro = new ResizeObserver(resize); ro.observe(container); resize();
  const loop = () => { raf = requestAnimationFrame(loop); if (controls.update() || dirty) { renderer.render(scene, camera); dirty = false; } };
  loop();
  return {
    build, apply,
    snapshot: () => { renderer.render(scene, camera); return renderer.domElement.toDataURL('image/png'); },
    resetView: () => { if (model) build(model.B, model.els); },
    dispose: () => { cancelAnimationFrame(raf); ro.disconnect(); controls.dispose(); renderer.dispose(); renderer.domElement.remove(); }
  };
}
