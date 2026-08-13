/**
 * 슬라이드 한 장을 파서 결과 그대로 펼쳐 본다.
 *
 *   npx esbuild tools/dump-slide.ts --bundle --platform=node --format=esm \
 *     --outfile=dist/dump-slide.mjs --external:jszip
 *   node dist/dump-slide.mjs "<pptx>" <1부터>
 *
 * 가져온 결과가 원본과 다를 때, 파서가 잘못 읽은 것인지 노드 생성이 잘못한 것인지
 * 가르는 데 쓴다. 여기 값이 맞으면 문제는 create.ts 에 있다.
 */
import { readFile } from 'node:fs/promises';
import { readPptx } from '../src/import/reader';
import { setLocale } from '../src/shared/i18n';
import type { ImportNode } from '../src/shared/importir';

setLocale('ko');

const file = process.argv[2];
const index = Number(process.argv[3] ?? '1') - 1;

const buf = await readFile(file);
const doc = await readPptx(
  buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer,
  'deck.pptx',
);

const slide = doc.slides[index];
if (!slide) {
  console.error(`슬라이드 ${index + 1} 없음 (총 ${doc.slides.length}장)`);
  process.exit(1);
}

console.log(`■ ${slide.name}   슬라이드 ${doc.widthPt.toFixed(1)} × ${doc.heightPt.toFixed(1)} pt\n`);

const f = (n: number): string => n.toFixed(0).padStart(5);

function show(nodes: ImportNode[], depth: number): void {
  for (const n of nodes) {
    const p = n.place;
    const box = `${f(p.x)},${f(p.y)} ${f(p.w)}×${f(p.h)}`;
    const rot = p.rotation ? ` rot${p.rotation.toFixed(0)}` : '';
    const pad = '  '.repeat(depth);

    if (n.type === 'text') {
      const runs = n.paragraphs.flatMap((q) => q.runs);
      const text = runs.map((r) => r.text).join('').replace(/\s+/g, ' ').slice(0, 40);
      const sizes = Array.from(new Set(runs.map((r) => r.size))).join('/');
      const fonts = Array.from(new Set(runs.map((r) => r.fontFamily))).join(',').slice(0, 22);
      const colors = Array.from(new Set(runs.map((r) => `#${r.color}`))).join(',');
      console.log(`${pad}text  ${box}${rot}  ${sizes}pt ${colors} [${fonts}] "${text}"`);
      console.log(`${pad}      여백 L${n.insets.left.toFixed(0)}/T${n.insets.top.toFixed(0)} · ${n.vertical} · 문단 ${n.paragraphs.length} · autoWidth=${n.autoWidth}`);
    } else if (n.type === 'group') {
      console.log(`${pad}group ${box}${rot} "${n.name.slice(0, 26)}" (${n.children.length})`);
      show(n.children, depth + 1);
    } else if (n.type === 'shape') {
      console.log(`${pad}${n.geometry.kind.padEnd(5)} ${box}${rot} "${n.name.slice(0, 26)}"`);
    } else {
      console.log(`${pad}${n.type.padEnd(5)} ${box}${rot} "${n.name.slice(0, 26)}"`);
    }
  }
}

show(slide.nodes, 0);
