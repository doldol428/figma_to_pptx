/**
 * 원본 PPTX 와 "가져오기 → 내보내기" 를 거친 PPTX 를 대조한다.
 *
 *   npx esbuild tools/compare-roundtrip.ts --bundle --platform=node --format=esm \
 *     --outfile=dist/compare.mjs --external:jszip
 *   node dist/compare.mjs "<원본>" "<왕복 결과>"
 *
 * 원본은 도형이 그룹 안에 들어 있고 그 좌표가 자식 좌표계(chOff/chExt)라, 그룹 변환을 풀어야
 * 왕복 결과(평평하게 펴진 도형들)와 같은 자로 잴 수 있다.
 */
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { deep, num, one, parseXml, type XNode } from '../src/import/xml';

const EMU_MM = 25.4 / 914400;
const mm = (v: number): string => (v * EMU_MM).toFixed(2);

interface Shape {
  name: string;
  kind: string;
  x: number;
  y: number;
  w: number;
  h: number;
  text: string;
  rot: number;
  flipH: boolean;
  flipV: boolean;
}

/**
 * 선의 양 끝점.
 *
 * PPTX 의 선은 상자의 대각선이고 (flip 으로 방향을 뒤집는다), Figma 의 선은 길이와 회전으로
 * 표현된다. 그래서 같은 선이라도 상자가 전혀 다르게 저장된다 — 상자를 맞대면 어긋나 보이지만
 * 실제로 그어지는 자리는 같을 수 있다. 그래서 끝점으로 비교한다.
 */
function ends(s: Shape): [number, number, number, number] {
  const x1 = s.flipH ? s.x + s.w : s.x;
  const x2 = s.flipH ? s.x : s.x + s.w;
  const y1 = s.flipV ? s.y + s.h : s.y;
  const y2 = s.flipV ? s.y : s.y + s.h;
  if (!s.rot) return [x1, y1, x2, y2];

  const rad = (s.rot / 60000) * (Math.PI / 180);
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  const cx = s.x + s.w / 2;
  const cy = s.y + s.h / 2;
  const spin = (px: number, py: number): [number, number] =>
    [cx + (px - cx) * cos - (py - cy) * sin, cy + (px - cx) * sin + (py - cy) * cos];
  const [ax, ay] = spin(x1, y1);
  const [bx, by] = spin(x2, y2);
  return [ax, ay, bx, by];
}

/** 그룹 좌표계를 실제 배치로 옮기는 선형 변환 */
interface Frame {
  sx: number;
  sy: number;
  dx: number;
  dy: number;
}

const IDENT: Frame = { sx: 1, sy: 1, dx: 0, dy: 0 };

function childFrame(xfrm: XNode | null, outer: Frame): Frame {
  if (!xfrm) return outer;
  const off = one(xfrm, 'off');
  const ext = one(xfrm, 'ext');
  const chOff = one(xfrm, 'chOff');
  const chExt = one(xfrm, 'chExt');
  if (!off || !ext || !chOff || !chExt) return outer;

  const cw = num(chExt.attrs.cx);
  const ch = num(chExt.attrs.cy);
  const sx = cw === 0 ? 1 : num(ext.attrs.cx) / cw;
  const sy = ch === 0 ? 1 : num(ext.attrs.cy) / ch;
  // 자식 좌표 c → 바깥 좌표: off + (c - chOff) * s, 그 위에 바깥 변환을 다시 얹는다
  return {
    sx: outer.sx * sx,
    sy: outer.sy * sy,
    dx: outer.dx + outer.sx * (num(off.attrs.x) - num(chOff.attrs.x) * sx),
    dy: outer.dy + outer.sy * (num(off.attrs.y) - num(chOff.attrs.y) * sy),
  };
}

function textOf(node: XNode | null): string {
  if (!node) return '';
  let s = '';
  const walk = (n: XNode): void => {
    if (n.tag === 't') s += n.text;
    for (const c of n.children) walk(c);
  };
  walk(node);
  return s;
}

function geomOf(sp: XNode): string {
  const spPr = one(sp, 'spPr');
  if (one(spPr, 'custGeom')) return 'custGeom';
  const prst = one(spPr, 'prstGeom')?.attrs.prst;
  if (prst) return prst;
  if (sp.tag === 'pic') return '(그림)';
  if (sp.tag === 'graphicFrame') return '(표/차트)';
  return '(없음)';
}

/** spTree 를 훑어 잎 도형만 절대 좌표로 모은다 */
function shapesOf(tree: XNode | null, frame: Frame, out: Shape[]): void {
  for (const el of tree?.children ?? []) {
    if (el.tag === 'grpSp') {
      shapesOf(el, childFrame(deep(el, 'xfrm'), frame), out);
      continue;
    }
    if (el.tag !== 'sp' && el.tag !== 'pic' && el.tag !== 'cxnSp' && el.tag !== 'graphicFrame') continue;

    const xfrm = deep(el, 'xfrm');
    const off = one(xfrm, 'off');
    const ext = one(xfrm, 'ext');
    const name = deep(el, 'cNvPr')?.attrs.name ?? '(이름 없음)';
    out.push({
      name,
      kind: geomOf(el),
      x: off ? frame.dx + frame.sx * num(off.attrs.x) : NaN,
      y: off ? frame.dy + frame.sy * num(off.attrs.y) : NaN,
      w: ext ? frame.sx * num(ext.attrs.cx) : NaN,
      h: ext ? frame.sy * num(ext.attrs.cy) : NaN,
      // 표는 글자가 <a:tbl> 안에 있어 txBody 만 보면 놓친다. 요소 전체를 훑는다.
      text: textOf(el),
      rot: num(xfrm?.attrs.rot),
      flipH: xfrm?.attrs.flipH === '1',
      flipV: xfrm?.attrs.flipV === '1',
    });
  }
}

async function readDeck(path: string): Promise<{
  cx: number; cy: number; slides: Shape[][]; media: number; parts: number; bytes: number;
}> {
  const buf = await readFile(path);
  const zip = await JSZip.loadAsync(buf);
  const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir);

  const pres = parseXml(await zip.file('ppt/presentation.xml')!.async('string'));
  const sz = deep(pres, 'sldSz');

  const rels = await zip.file('ppt/_rels/presentation.xml.rels')!.async('string');
  const relMap = new Map([...rels.matchAll(/Id="([^"]+)"[^>]*Target="([^"]+)"/g)].map((m) => [m[1], m[2]]));
  const order = [...(await zip.file('ppt/presentation.xml')!.async('string'))
    .matchAll(/<p:sldId [^>]*r:id="([^"]+)"/g)].map((m) => relMap.get(m[1]) ?? '');

  const slides: Shape[][] = [];
  for (const target of order) {
    const p = 'ppt/' + target.replace(/^\.\.\//, '');
    const xml = await zip.file(p)?.async('string');
    if (!xml) continue;
    const shapes: Shape[] = [];
    shapesOf(deep(parseXml(xml), 'spTree'), IDENT, shapes);
    slides.push(shapes);
  }

  return {
    cx: num(sz?.attrs.cx),
    cy: num(sz?.attrs.cy),
    slides,
    media: names.filter((n) => /^ppt\/media\/.+\.\w+$/.test(n)).length,
    parts: names.length,
    bytes: buf.length,
  };
}

/* ── 실행 ─────────────────────────────────────────────────────── */

const [srcPath, outPath] = process.argv.slice(2);
if (!srcPath || !outPath) {
  console.error('사용법: node dist/compare.mjs "<원본>" "<왕복 결과>"');
  process.exit(1);
}

const a = await readDeck(srcPath);
const b = await readDeck(outPath);
const mbOf = (n: number): string => (n / 1024 / 1024).toFixed(1);

console.log('■ 문서');
console.log(`  슬라이드 크기   원본 ${a.cx} × ${a.cy} EMU (${mm(a.cx)} × ${mm(a.cy)} mm)`);
console.log(`                왕복 ${b.cx} × ${b.cy} EMU (${mm(b.cx)} × ${mm(b.cy)} mm)`);
console.log(`                차이 ${a.cx - b.cx} × ${a.cy - b.cy} EMU`);
console.log(`  슬라이드 수     ${a.slides.length} → ${b.slides.length}`);
console.log(`  미디어 파일     ${a.media} → ${b.media}`);
console.log(`  파트 수         ${a.parts} → ${b.parts}`);
console.log(`  파일 크기       ${mbOf(a.bytes)} MB → ${mbOf(b.bytes)} MB`);

const countA = a.slides.reduce((s, x) => s + x.length, 0);
const countB = b.slides.reduce((s, x) => s + x.length, 0);
const charsA = a.slides.reduce((s, x) => s + x.reduce((t, y) => t + y.text.length, 0), 0);
const charsB = b.slides.reduce((s, x) => s + x.reduce((t, y) => t + y.text.length, 0), 0);
console.log(`\n■ 도형 ${countA} → ${countB}  (${countB - countA >= 0 ? '+' : ''}${countB - countA})`);
console.log(`■ 글자 ${charsA.toLocaleString()} → ${charsB.toLocaleString()}자  (${charsB - charsA >= 0 ? '+' : ''}${charsB - charsA})`);

/* 도형 종류 분포 */
const kinds = (slides: Shape[][]): Map<string, number> => {
  const m = new Map<string, number>();
  for (const s of slides) for (const x of s) m.set(x.kind, (m.get(x.kind) ?? 0) + 1);
  return m;
};
const ka = kinds(a.slides);
const kb = kinds(b.slides);
const allKinds = [...new Set([...ka.keys(), ...kb.keys()])]
  .sort((x, y) => (kb.get(y) ?? 0) + (ka.get(y) ?? 0) - (kb.get(x) ?? 0) - (ka.get(x) ?? 0));
console.log('\n■ 도형 종류 (원본 → 왕복)');
for (const k of allKinds.slice(0, 18)) {
  const x = ka.get(k) ?? 0;
  const y = kb.get(k) ?? 0;
  if (x === 0 && y === 0) continue;
  console.log(`  ${String(x).padStart(6)} → ${String(y).padStart(6)}  ${k}${x !== y ? `   (${y - x >= 0 ? '+' : ''}${y - x})` : ''}`);
}

/* 이름으로 짝을 지어 위치·크기 비교 */
let paired = 0; let lost = 0; let added = 0;
const offsets: Array<{ d: number; dx: number; dy: number; dw: number; dh: number; label: string }> = [];
const missing: string[] = [];
const lineGaps: Array<{ d: number; label: string }> = [];

for (let i = 0; i < Math.min(a.slides.length, b.slides.length); i++) {
  const byName = new Map<string, Shape[]>();
  for (const s of b.slides[i]) {
    const list = byName.get(s.name);
    if (list) list.push(s);
    else byName.set(s.name, [s]);
  }
  for (const s of a.slides[i]) {
    /*
     * 같은 이름이 여러 개인 도형이 흔하다. 순서대로 짝지으면 엉뚱한 짝이 나서
     * 실제로는 제자리에 있는 도형이 수백 mm 어긋난 것처럼 보인다. 가장 가까운 것과 짝짓는다.
     */
    const pool = byName.get(s.name);
    let hit: Shape | undefined;
    if (pool?.length) {
      let best = 0;
      let bestD = Infinity;
      for (let k = 0; k < pool.length; k++) {
        const d = Number.isFinite(s.x) && Number.isFinite(pool[k].x)
          ? Math.hypot(s.x - pool[k].x, s.y - pool[k].y)
          : 0;
        if (d < bestD) { bestD = d; best = k; }
      }
      hit = pool.splice(best, 1)[0];
    }
    if (!hit) {
      lost++;
      if (missing.length < 10) missing.push(`  ${i + 1}장  ${s.kind.padEnd(20)} "${s.name}"`);
      continue;
    }
    paired++;
    if (!Number.isFinite(s.x) || !Number.isFinite(hit.x)) continue;

    // 선은 상자가 아니라 그어지는 자리로 판단한다
    if (s.kind === 'line' || hit.kind === 'line') {
      const [ax1, ay1, ax2, ay2] = ends(s);
      const [bx1, by1, bx2, by2] = ends(hit);
      // 어느 끝이 어느 끝에 대응하는지는 모르므로 두 대응을 다 보고 가까운 쪽을 쓴다
      const same = Math.max(Math.hypot(ax1 - bx1, ay1 - by1), Math.hypot(ax2 - bx2, ay2 - by2));
      const flipped = Math.max(Math.hypot(ax1 - bx2, ay1 - by2), Math.hypot(ax2 - bx1, ay2 - by1));
      lineGaps.push({ d: Math.min(same, flipped), label: `${i + 1}장 "${s.name}"` });
      continue;
    }

    const dx = Math.abs(s.x - hit.x);
    const dy = Math.abs(s.y - hit.y);
    offsets.push({
      d: Math.hypot(dx, dy),
      dx, dy,
      dw: Math.abs(s.w - hit.w),
      dh: Math.abs(s.h - hit.h),
      label: `${i + 1}장 ${s.kind} "${s.name}"`,
    });
  }
  added += [...byName.values()].reduce((t, v) => t + v.length, 0);
}

console.log('\n■ 이름으로 짝지은 도형');
console.log(`  짝 지어짐 ${paired} · 원본에만 ${lost} · 왕복에만 ${added}`);

const sorted = [...offsets].sort((x, y) => x.d - y.d);
const at = (p: number): string => (sorted.length ? mm(sorted[Math.floor((sorted.length - 1) * p)].d) : '-');
console.log(`  위치 어긋남 분포   중앙값 ${at(0.5)} · 90% ${at(0.9)} · 99% ${at(0.99)} · 최대 ${at(1)} mm`);
const within = (limit: number): string => {
  const n = offsets.filter((o) => o.d * EMU_MM <= limit).length;
  return `${((n / offsets.length) * 100).toFixed(1)}%`;
};
console.log(`  0.1mm 이내 ${within(0.1)} · 0.5mm 이내 ${within(0.5)} · 1mm 이내 ${within(1)} · 5mm 이내 ${within(5)}`);

console.log('  가장 많이 어긋난 것');
for (const o of sorted.slice(-8).reverse()) {
  console.log(`    ${mm(o.d).padStart(8)} mm   ${o.label}  (크기차 ${mm(o.dw)} × ${mm(o.dh)} mm)`);
}
if (lineGaps.length) {
  const ls = [...lineGaps].sort((x, y) => x.d - y.d);
  const lat = (p: number): string => mm(ls[Math.floor((ls.length - 1) * p)].d);
  console.log(`\n■ 선 ${lineGaps.length}개 — 그어지는 자리(양 끝점)로 비교`);
  console.log(`  중앙값 ${lat(0.5)} · 90% ${lat(0.9)} · 99% ${lat(0.99)} · 최대 ${lat(1)} mm`);
  const near = lineGaps.filter((o) => o.d * EMU_MM <= 0.5).length;
  console.log(`  0.5mm 이내 ${((near / lineGaps.length) * 100).toFixed(1)}%`);
  console.log('  가장 많이 어긋난 선');
  for (const o of ls.slice(-4).reverse()) console.log(`    ${mm(o.d).padStart(8)} mm   ${o.label}`);
}

if (missing.length) {
  console.log('\n■ 원본에는 있는데 왕복 결과에서 같은 이름을 못 찾은 것');
  for (const m of missing) console.log(m);
}

/* 슬라이드별 도형 수 차이가 큰 곳 */
const gaps: Array<[number, number, number]> = [];
for (let i = 0; i < Math.min(a.slides.length, b.slides.length); i++) {
  gaps.push([i + 1, a.slides[i].length, b.slides[i].length]);
}
gaps.sort((x, y) => Math.abs(y[2] - y[1]) - Math.abs(x[2] - x[1]));
console.log('\n■ 도형 수 차이가 큰 슬라이드');
for (const [n, x, y] of gaps.slice(0, 8)) {
  console.log(`  ${String(n).padStart(3)}장  ${String(x).padStart(4)} → ${String(y).padStart(4)}  (${y - x >= 0 ? '+' : ''}${y - x})`);
}
