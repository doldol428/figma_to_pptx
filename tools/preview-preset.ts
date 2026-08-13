/**
 * preset 도형을 SVG 로 그려본다 — 기하가 맞는지 눈으로 확인하는 용도.
 *
 *   npx esbuild tools/preview-preset.ts --bundle --platform=node --format=esm \
 *     --outfile=dist/preview-preset.mjs
 *   node dist/preview-preset.mjs
 *
 * OOXML 의 조절값은 대부분 짧은 변(ss) 기준이라, 너비로 계산하면 납작한 도형에서
 * 몇 배로 어긋난다. 숫자만 봐서는 티가 안 나므로 그려서 본다.
 */
import { writeFileSync } from 'node:fs';
import { presetPath } from '../src/import/preset';

interface Sample {
  name: string;
  prst: string;
  x: number;
  y: number;
  w: number;
  h: number;
  color: string;
}

// 실측 파일 슬라이드 1 의 Step 행 (좌표·크기 그대로)
const SAMPLES: Sample[] = [
  { name: 'homePlate', prst: 'homePlate', x: 34, y: 545, w: 147, h: 43, color: '#BFD3D6' },
  { name: 'chevron', prst: 'chevron', x: 164, y: 545, w: 144, h: 43, color: '#A9C4CC' },
  { name: 'chevron', prst: 'chevron', x: 290, y: 545, w: 144, h: 43, color: '#7C93C8' },
  { name: 'chevron', prst: 'chevron', x: 417, y: 545, w: 144, h: 43, color: '#5B6FB5' },
];

let body = '';
for (const s of SAMPLES) {
  const path = presetPath(s.prst, s.w, s.h, {});
  if (!path) {
    console.error(`${s.prst}: 경로 없음`);
    continue;
  }
  body += `<g transform="translate(${s.x} ${s.y})"><path d="${path.data}" fill="${s.color}"/></g>`;
}

const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="1120" height="110" viewBox="20 535 560 55">`
  + `<rect x="20" y="535" width="560" height="55" fill="#FFFFFF"/>${body}</svg>`;

const out = new URL('../dist/preview-preset.svg', import.meta.url);
writeFileSync(out, svg);
console.log('dist/preview-preset.svg');
