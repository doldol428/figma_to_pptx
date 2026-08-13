/**
 * 가져오기 파서를 실제 PPTX 로 검증한다.
 *
 *   npm run verify:import -- "<pptx 경로>"
 *
 * Figma 노드 생성은 여기서 못 돌리지만, 파서는 순수 함수라 Node 에서 그대로 실행된다.
 * 슬라이드 크기·노드 종류·좌표 범위·폰트·경고를 뽑아 눈으로 확인할 수 있게 한다.
 */
import { readFile } from 'node:fs/promises';
import { setLocale } from '../src/shared/i18n';
import type { ImportNode } from '../src/shared/importir';
import { readPptx } from '../src/import/reader';

setLocale('ko');

const file = process.argv[2];
if (!file) {
  console.error('사용법: npm run verify:import -- "<pptx 경로>"');
  process.exit(1);
}

const buf = await readFile(file);
const started = process.hrtime.bigint();
const doc = await readPptx(
  buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer,
  file.split(/[\\/]/).pop() ?? 'deck.pptx',
);
const ms = Number(process.hrtime.bigint() - started) / 1e6;

const EMU_MM = 25.4 / 72;
const kinds = new Map<string, number>();
const geoms = new Map<string, number>();
let deepest = 0;
let outOfBounds = 0;
let textChars = 0;
let imageBytes = 0;

function bump(m: Map<string, number>, k: string): void {
  m.set(k, (m.get(k) ?? 0) + 1);
}

function walk(nodes: ImportNode[], depth: number): void {
  deepest = Math.max(deepest, depth);
  for (const n of nodes) {
    bump(kinds, n.type);
    if (n.type === 'shape') bump(geoms, n.geometry.kind);
    if (n.type === 'text') {
      for (const p of n.paragraphs) for (const r of p.runs) textChars += r.text.length;
    }
    if (n.type === 'image') imageBytes += (n.data.length * 3) / 4;
    if (n.type === 'group') walk(n.children, depth + 1);

    // 좌표가 슬라이드 밖으로 크게 벗어나면 그룹 좌표계 매핑이 틀어졌다는 신호다.
    const { x, y, w, h } = n.place;
    if (x < -doc.widthPt || y < -doc.heightPt
      || x > doc.widthPt * 2 || y > doc.heightPt * 2
      || w > doc.widthPt * 3 || h > doc.heightPt * 3) {
      outOfBounds++;
    }
  }
}

for (const s of doc.slides) walk(s.nodes, 1);

console.log(`■ ${doc.fileName}`);
console.log(`  슬라이드  ${doc.slides.length}장 · ${(doc.widthPt * EMU_MM).toFixed(1)} × ${(doc.heightPt * EMU_MM).toFixed(1)} mm`);
console.log(`  파싱      ${ms.toFixed(0)} ms`);

console.log('\n■ 노드');
for (const [k, v] of [...kinds].sort((a, b) => b[1] - a[1])) {
  console.log(`  ${String(v).padStart(6)}  ${k}`);
}
console.log(`  중첩 깊이 ${deepest} · 글자 ${textChars.toLocaleString()}자 · 이미지 ${(imageBytes / 1024 / 1024).toFixed(1)} MB`);

console.log('\n■ 도형 기하');
for (const [k, v] of [...geoms].sort((a, b) => b[1] - a[1])) {
  console.log(`  ${String(v).padStart(6)}  ${k}`);
}

console.log(`\n■ 좌표 범위 밖 노드: ${outOfBounds}`);
if (outOfBounds > 0) {
  console.log('  0 이 아니면 그룹 좌표계(chOff/chExt) 매핑을 의심할 것');
}

console.log(`\n■ 폰트 ${doc.fonts.length}종`);
console.log(`  ${doc.fonts.join(', ')}`);

// 문서에 포함된 글꼴의 이름표. Figma 는 영문 이름으로만 등록하므로 이게 매칭의 정답지다.
const aliases = Object.keys(doc.fontAliases);
console.log(`\n■ 포함된 글꼴 이름 ${aliases.length}종 (한글 → 영문)`);
for (const k of aliases.sort()) console.log(`  ${k}  →  ${doc.fontAliases[k]}`);

const byMessage = new Map<string, number>();
for (const w of doc.warnings) bump(byMessage, w.message);
console.log(`\n■ 경고 ${doc.warnings.length}건 (${byMessage.size}종)`);
for (const [k, v] of [...byMessage].sort((a, b) => b[1] - a[1]).slice(0, 12)) {
  console.log(`  ${String(v).padStart(4)}  ${k}`);
}
