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
  slideWPt: A4_W,
  slideHPt: A4_H,
  ptPerPx: 1,
  offsetXPt: 0,
  offsetYPt: 0,
  frameWPx: A4_W,
  frameHPx: A4_H,
  presetLabel: '프레임 실측 (1px = 1pt)',
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

async function render(d: Doc): Promise<{
  cx: number; cy: number; xml: string; bytes: Uint8Array; zip: JSZip;
}> {
  const bytes = (await composePptx(d).write({ outputType: 'nodebuffer' })) as Uint8Array;
  const zip = await JSZip.loadAsync(bytes);
  const presentation = await zip.file('ppt/presentation.xml')!.async('string');
  const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
  const sldSz = /<p:sldSz([^/]*)\/>/.exec(presentation)?.[1] ?? '';
  return {
    cx: Number(/cx="(\d+)"/.exec(sldSz)?.[1]),
    cy: Number(/cy="(\d+)"/.exec(sldSz)?.[1]),
    xml,
    bytes,
    zip,
  };
}

/**
 * 1920×1080 프레임 → PowerPoint 와이드스크린.
 * 균등 배율 0.5 가 좌표·크기·폰트에 모두 걸려야 하고, 종횡비는 그대로여야 한다.
 */
function widescreenDoc(): Doc {
  return {
    slideWPt: 960,
    slideHPt: 540,
    ptPerPx: 0.5,
    offsetXPt: 0,
    offsetYPt: 0,
    frameWPx: 1920,
    frameHPx: 1080,
    presetLabel: 'PowerPoint 와이드스크린 16:9',
    warnings: [],
    slides: [{
      name: '표지',
      fill: { kind: 'solid', color: '101010', transparency: 0 },
      items: [
        {
          type: 'shape',
          name: '가운데 사각형',
          // 프레임 좌표계(px) 기준 — 배율 0.5 가 걸려 240pt 위치에 480×270pt 로 앉아야 한다
          box: { x: 480, y: 270, w: 960, h: 540, rot: 0, flipH: false, flipV: false },
          geom: { kind: 'rect' },
          fill: { kind: 'solid', color: 'FFFFFF', transparency: 0 },
          stroke: { color: 'FF0000', transparency: 0, width: 8, dashType: 'solid' },
        },
        {
          type: 'text',
          name: '제목',
          box: { x: 160, y: 120, w: 1600, h: 200, rot: 0, flipH: false, flipV: false },
          align: 'center', valign: 'middle', wrap: true,
          runs: [{
            text: '96px 제목', fontFace: 'Pretendard', fontSize: 96, bold: true,
            italic: false, underline: false, strike: false, color: 'FFFFFF', transparency: 0,
          }],
        },
      ],
    }],
  };
}

async function main(): Promise<void> {
  const { cx, cy, xml: slide, bytes, zip } = await render(doc);
  await writeFile(new URL('../dist/verify-a4.pptx', import.meta.url), bytes);

  console.log('\nA4 실측 — 슬라이드 크기');
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

  /* ── 1920×1080 → PowerPoint 와이드스크린 ─────────────────────── */

  const wide = widescreenDoc();
  const w = await render(wide);
  await writeFile(new URL('../dist/verify-16x9.pptx', import.meta.url), w.bytes);

  console.log('\n1920×1080 → PowerPoint 와이드스크린 (균등 배율 0.5×)');
  check('가로 EMU (13.333in)', w.cx, 12192000);
  check('세로 EMU (7.5in)', w.cy, 6858000);
  checkTruthy(
    '종횡비가 정확히 16:9 로 유지됨',
    Math.abs(w.cx / w.cy - 16 / 9) < 1e-9,
  );
  checkTruthy(
    '프레임 (480,270) → 슬라이드 240pt (3048000 EMU)',
    w.xml.includes(`<a:off x="${240 * 12700}" y="${135 * 12700}"/>`),
  );
  checkTruthy(
    '960×540px 사각형 → 480×270pt',
    w.xml.includes(`<a:ext cx="${480 * 12700}" cy="${270 * 12700}"/>`),
  );
  checkTruthy('폰트 96px → 48pt (배율이 폰트에도 적용됨)', w.xml.includes('sz="4800"'));
  checkTruthy('선 굵기 8px → 4pt', w.xml.includes(`<a:ln w="${4 * 12700}">`));
  console.log(`  → ${w.cx / EMU_PER_MM} × ${w.cy / EMU_PER_MM} mm`);
  console.log(`dist/verify-16x9.pptx 생성 (${(w.bytes.length / 1024).toFixed(1)} KB)`);

  if (failures.length > 0) {
    console.error(`\n실패 ${failures.length}건: ${failures.join(', ')}`);
    process.exit(1);
  }
  console.log('\n전체 통과');
}

main();
