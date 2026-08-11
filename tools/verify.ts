/**
 * 회귀 검증 — 이 플러그인이 존재하는 이유를 자동으로 지킨다.
 *
 * A4(210×297mm) 프레임을 IR 로 만들어 pptx 를 뽑고, 실제 XML 의 sldSz 가
 * 7560000 × 10692000 EMU 인지 확인한다. 16:9 로 뭉개는 회귀가 생기면 여기서 걸린다.
 *
 *   npm run verify
 */
import { writeFile } from 'node:fs/promises';
import JSZip from 'jszip';
import type { Doc } from '../src/shared/ir';
import { composePptx } from '../src/ui/build';

const EMU_PER_MM = 36000;
const PT_PER_MM = 72 / 25.4;

/** mm → Figma px (1px = 1pt 규약) */
const mm = (v: number): number => v * PT_PER_MM;

const A4_W = mm(210);
const A4_H = mm(297);

const doc: Doc = {
  widthPx: A4_W,
  heightPx: A4_H,
  paper: 'A4 세로',
  warnings: [],
  slides: [
    {
      name: 'A4 테스트',
      fill: { kind: 'solid', color: 'FFFFFF', transparency: 0 },
      items: [
        {
          type: 'shape',
          name: '전면 사각형',
          box: { x: 0, y: 0, w: A4_W, h: A4_H, rot: 0, flipH: false, flipV: false },
          geom: { kind: 'rect' },
          fill: { kind: 'solid', color: 'F5F5F5', transparency: 0 },
        },
        {
          type: 'shape',
          name: '20mm 여백 사각형',
          box: { x: mm(20), y: mm(20), w: mm(170), h: mm(50), rot: 0, flipH: false, flipV: false },
          geom: { kind: 'roundRect', radius: mm(4) },
          fill: { kind: 'solid', color: '0D99FF', transparency: 0 },
          stroke: { color: '003366', transparency: 0, width: 2, dashType: 'solid' },
        },
        {
          type: 'shape',
          name: '회전 45도 타원',
          box: { x: mm(20), y: mm(90), w: mm(40), h: mm(40), rot: 45, flipH: false, flipV: false },
          geom: { kind: 'ellipse' },
          fill: { kind: 'solid', color: 'FF6B00', transparency: 20 },
        },
        {
          type: 'shape',
          name: '커스텀 경로(삼각형+베지어)',
          box: { x: mm(80), y: mm(90), w: mm(40), h: mm(40), rot: 0, flipH: false, flipV: false },
          geom: {
            kind: 'custom',
            points: [
              { x: 0, y: mm(40), moveTo: true },
              { x: mm(20), y: 0 },
              { x: mm(40), y: mm(40), c: [mm(34), mm(6), mm(40), mm(24)] },
              { close: true },
            ],
          },
          fill: { kind: 'solid', color: '22C55E', transparency: 0 },
        },
        {
          type: 'text',
          name: '제목',
          box: { x: mm(20), y: mm(150), w: mm(170), h: mm(30), rot: 0, flipH: false, flipV: false },
          align: 'left',
          valign: 'top',
          wrap: true,
          runs: [
            {
              text: '한글 제목 ', fontFace: 'Pretendard', fontSize: 24, bold: true,
              italic: false, underline: false, strike: false, color: '111111', transparency: 0,
            },
            {
              text: '두 번째 줄', fontFace: 'Pretendard', fontSize: 14, bold: false,
              italic: false, underline: false, strike: false, color: '666666', transparency: 0,
              breakLine: true,
            },
            {
              text: '세 번째 줄', fontFace: 'Pretendard', fontSize: 14, bold: false,
              italic: false, underline: false, strike: false, color: '666666', transparency: 0,
            },
          ],
        },
      ],
    },
  ],
};

const failures: string[] = [];

function check(label: string, actual: unknown, expected: unknown): void {
  const ok = actual === expected;
  console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${label}: ${actual}${ok ? '' : ` (기대값 ${expected})`}`);
  if (!ok) failures.push(label);
}

function checkTruthy(label: string, ok: boolean): void {
  console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${label}`);
  if (!ok) failures.push(label);
}

async function main(): Promise<void> {
  const buffer = await composePptx(doc).write({ outputType: 'nodebuffer' });
  const bytes = buffer as Uint8Array;
  await writeFile(new URL('../dist/verify-a4.pptx', import.meta.url), bytes);

  const zip = await JSZip.loadAsync(bytes);
  const presentation = await zip.file('ppt/presentation.xml')!.async('string');
  const slide = await zip.file('ppt/slides/slide1.xml')!.async('string');

  const sldSz = /<p:sldSz([^/]*)\/>/.exec(presentation)?.[1] ?? '';
  const cx = Number(/cx="(\d+)"/.exec(sldSz)?.[1]);
  const cy = Number(/cy="(\d+)"/.exec(sldSz)?.[1]);

  console.log('\n슬라이드 크기');
  check('가로 EMU (210mm)', cx, 210 * EMU_PER_MM);
  check('세로 EMU (297mm)', cy, 297 * EMU_PER_MM);
  console.log(`  → ${cx / EMU_PER_MM} × ${cy / EMU_PER_MM} mm`);

  console.log('\n도형 배치');
  // 20mm 여백 사각형: off = 20mm, ext = 170×50mm
  checkTruthy(
    '20mm 여백이 정확히 720000 EMU 로 들어감',
    slide.includes(`<a:off x="${20 * EMU_PER_MM}" y="${20 * EMU_PER_MM}"/>`),
  );
  checkTruthy(
    '170×50mm 크기가 그대로 유지됨',
    slide.includes(`<a:ext cx="${170 * EMU_PER_MM}" cy="${50 * EMU_PER_MM}"/>`),
  );
  checkTruthy('회전 45도가 rot="2700000" 으로 기록됨', slide.includes('rot="2700000"'));
  checkTruthy('custGeom 경로가 생성됨', slide.includes('<a:custGeom>'));
  checkTruthy('베지어 세그먼트가 포함됨', slide.includes('<a:cubicBezTo>'));
  checkTruthy('roundRect adj 값이 기록됨', slide.includes('<a:gd name="adj"'));

  console.log('\n텍스트');
  checkTruthy('한글 런이 살아 있음', slide.includes('한글 제목'));
  checkTruthy('문단이 2개 이상으로 나뉨', (slide.match(/<a:p>/g) ?? []).length >= 2);
  checkTruthy(
    '텍스트 박스 내부 여백이 0 (PowerPoint 기본 0.1인치가 제거됨)',
    /lIns="0"/.test(slide) && /tIns="0"/.test(slide),
  );

  console.log('\n패키지 구조');
  for (const part of [
    '[Content_Types].xml', '_rels/.rels', 'ppt/presentation.xml',
    'ppt/slides/slide1.xml', 'ppt/slideMasters/slideMaster1.xml', 'ppt/theme/theme1.xml',
  ]) {
    checkTruthy(`${part} 존재`, zip.file(part) !== null);
  }

  console.log(`\ndist/verify-a4.pptx 생성 (${(bytes.length / 1024).toFixed(1)} KB)`);

  if (failures.length > 0) {
    console.error(`\n실패 ${failures.length}건: ${failures.join(', ')}`);
    process.exit(1);
  }
  console.log('\n전체 통과');
}

main();
