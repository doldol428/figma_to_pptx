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
import type { Box, Doc, TextItem } from '../src/shared/ir';
import { fitsInMaster } from '../src/shared/ir';
import { resolveLocale, setLocale } from '../src/shared/i18n';
import { resolveSlide } from '../src/shared/slidesize';
import { exportScaleForDpi } from '../src/shared/units';
import { PptxBuilder, composePptx } from '../src/ui/build';
import { aliasesFor, pickFont } from '../src/main/fontalias';
import { linePlacement } from '../src/import/transform';
import { presetPath } from '../src/import/preset';
import { pathBounds } from '../src/import/pathbox';
import { readFill } from '../src/import/color';
import { deep as deepXml, parseXml } from '../src/import/xml';

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
  chip: 'A4 세로',
  warnings: [],
  masters: [],
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

function round(n: number): number {
  return Math.round(n * 100) / 100;
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
    chip: '16:9',
    warnings: [],
    masters: [],
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

  /* ── 슬라이드 크기 자동 결정 ─────────────────────────────────── */

  // 사용자에게 묻지 않고 프레임 크기만으로 결정된다. 판정표를 그대로 못박는다.
  setLocale('ko');
  console.log('\n슬라이드 크기 자동 결정');
  const cases: Array<[string, number, number, number, number, number, string | null]> = [
    // 이름                프레임 W  프레임 H   슬라이드W 슬라이드H  배율        칩
    ['A4 세로',            mm(210), mm(297),  mm(210), mm(297),  1,          'A4 세로'],
    ['A4 가로',            mm(297), mm(210),  mm(297), mm(210),  1,          'A4 가로'],
    ['960×540 (이미 표준)', 960,     540,      960,     540,      1,          '16:9'],
    ['1920×1080',          1920,    1080,     960,     540,      0.5,        '16:9'],
    ['1280×720',           1280,    720,      960,     540,      0.75,       '16:9'],
    ['3840×2160 (4K)',     3840,    2160,     960,     540,      0.25,       '16:9'],
    ['5120×2880 (5K)',     5120,    2880,     960,     540,      0.1875,     '16:9'],
    ['1024×768',           1024,    768,      720,     540,      0.703125,   '4:3'],
    ['1080×1920 (세로)',    1080,    1920,     540,     960,      0.5,        '9:16'],
  ];
  for (const [name, fw, fh, sw, sh, k, chip] of cases) {
    const plan = resolveSlide(fw, fh);
    const got = plan
      ? `${round(plan.wPt)}×${round(plan.hPt)}pt ${plan.ptPerPx}× [${plan.chip ?? '-'}]`
      : 'null';
    const want = `${round(sw)}×${round(sh)}pt ${k}× [${chip ?? '-'}]`;
    check(name, got, want);
  }
  // 균등 배율이므로 종횡비는 어떤 경우에도 보존돼야 한다.
  for (const [name, fw, fh] of cases) {
    const plan = resolveSlide(fw, fh)!;
    checkTruthy(
      `${name} — 종횡비 보존`,
      Math.abs((plan.wPt - plan.offsetXPt * 2) / (plan.hPt - plan.offsetYPt * 2) - fw / fh) < 1e-9,
    );
  }

  /* ── 로케일 ──────────────────────────────────────────────────── */

  // 한국어 / 영어 둘만 지원하고 나머지는 전부 영어로 떨어져야 한다.
  console.log('\n로케일 판정');
  check('ko-KR', resolveLocale(['ko-KR']), 'ko');
  check('ko', resolveLocale(['ko']), 'ko');
  check('en-US', resolveLocale(['en-US']), 'en');
  check('ja-JP → 영어', resolveLocale(['ja-JP']), 'en');
  check('zh-CN → 영어', resolveLocale(['zh-CN']), 'en');
  check('빈 목록 → 영어', resolveLocale([]), 'en');
  check('fr-FR, ko-KR → 한국어', resolveLocale(['fr-FR', 'ko-KR']), 'ko');

  // 용지 칩은 번역되고, 비율 칩은 언어와 무관하게 같아야 한다.
  setLocale('en');
  check('영어 용지 칩', resolveSlide(mm(210), mm(297))?.chip, 'A4 Portrait');
  check('영어 가로 칩', resolveSlide(mm(297), mm(210))?.chip, 'A4 Landscape');
  check('영어 비율 칩', resolveSlide(1920, 1080)?.chip, '16:9');
  setLocale('ko');
  check('한국어 용지 칩', resolveSlide(mm(210), mm(297))?.chip, 'A4 세로');
  check('한국어 비율 칩', resolveSlide(1920, 1080)?.chip, '16:9');

  /* ── 이미지 렌더 해상도 ──────────────────────────────────────── */

  // 목표 DPI 는 슬라이드 배율과 무관하게 같은 실효 해상도로 떨어져야 한다.
  // 배율을 사용자에게 직접 받으면 프리셋마다 해상도가 몇 배씩 널뛴다.
  console.log('\n이미지 렌더 해상도 (220 DPI 고정 요청)');
  for (const [label, ptPerPx] of [
    ['A4 실측 (1×)', 1],
    ['1920 → 와이드스크린 (0.5×)', 0.5],
    ['3840 → 와이드스크린 (0.25×)', 0.25],
  ] as const) {
    const r = exportScaleForDpi(220, ptPerPx);
    check(`${label} → 실효 DPI`, r.actualDpi, 220);
    console.log(`         노드 렌더 배율 ${r.scale}×`);
  }
  // A4 실측에서 300 DPI 는 4.17× 라 Figma 한계에 걸린다. 조용히 넘어가면 안 된다.
  const clampedCase = exportScaleForDpi(300, 1);
  checkTruthy('A4 실측 300 DPI 는 4× 한계에 걸린 것으로 보고됨', clampedCase.clamped);
  check('한계 적용 후 실제 배율', clampedCase.scale, 4);

  /* ── 한글 글꼴 이름 → Figma 등록명 ───────────────────────────── */

  /*
   * Figma 데스크탑이 돌려준 실제 목록(2163종)에는 한글 family 가 하나도 없었다.
   * PPTX 가 적어둔 한글 이름을 그대로 찾으면 설치된 글꼴도 전부 "없음" 이 된다.
   * 아래 오른쪽 이름들은 그 목록에서 확인한 실제 등록명이다.
   */
  console.log('\n한글 글꼴 이름 → Figma 등록명');
  const aliasCases: Array<[string, string]> = [
    ['KoPub돋움체 Bold', 'KoPubDotum Bold'],
    ['KoPub바탕체 Light', 'KoPubBatang Light'],
    ['맑은 고딕', 'Malgun Gothic'],
    ['나눔고딕', 'NanumGothic'],
    ['나눔명조', 'NanumMyeongjo'],
    ['굴림', 'Gulim'],
    ['굴림체', 'GulimChe'],
    ['바탕', 'Batang'],
    ['궁서체', 'GungsuhChe'],
    ['HY견고딕', 'HYGothic-Extra'],
    ['HY헤드라인M', 'HYHeadLine-Medium'],
    ['서울남산체 B', 'SeoulNamsan B'],
    ['페이퍼로지 4 Regular', 'Paperlogy 4 Regular'],
  ];
  for (const [korean, want] of aliasCases) {
    const got = aliasesFor(korean);
    checkTruthy(`${korean} → ${want}`, got.indexOf(want) >= 0);
  }
  // 영문 이름은 바꿀 것이 없으므로 후보를 만들지 않는다 (헛돌면 매칭이 느려지고 오탐이 는다).
  check('영문 이름은 후보 없음', aliasesFor('Pretendard').length, 0);
  check('영문 이름은 후보 없음 (공백 포함)', aliasesFor('Times New Roman').length, 0);
  // 구체적인 규칙이 먼저 걸려야 한다 — KoPub돋움체가 돋움체 규칙에 먼저 잡히면 KoPubDotumChe 가 된다.
  check('KoPub 은 첫 후보가 정확', aliasesFor('KoPub돋움체 Bold')[0], 'KoPubDotum Bold');

  /* ── 폰트 이름 → 실제 등록된 family/style ────────────────────── */

  /*
   * 이름을 바꾸는 것만으로는 부족하다. 실제로 설치 목록에서 집히는지가 관건이다.
   * 아래 목록은 Figma 데스크탑이 돌려준 이름을 그대로 옮긴 것이다 (스타일 표기까지 그대로).
   */
  const installed = new Map<string, Set<string>>([
    ['Paperlogy', new Set(['1 Thin', '2 ExtraLight', '3 Light', '4 Regular', '5 Medium', '6 SemiBold', '7 Bold', '8 ExtraBold', '9 Black'])],
    ['Freesentation', new Set(['1 Thin', '2 ExtraLight', '3 Light', '4 Regular', '5 Medium', '6 SemiBold', '7 Bold', '8 ExtraBold', '9 Black'])],
    ['KoPubDotum', new Set(['Bold', 'Light', 'Medium'])],
    ['KoPubBatang', new Set(['Bold', 'Light', 'Medium'])],
    ['KoPubWorldDotum', new Set(['Bold', 'Light', 'Medium'])],
    ['KoPubWorldDotum_Pro', new Set(['Bold', 'Light', 'Medium'])],
    ['NEXON Football Gothic', new Set(['B', 'L'])],
    ['Gmarket Sans TTF', new Set(['Bold', 'Light', 'Medium'])],
    ['NanumSquare', new Set(['Bold', 'ExtraBold', 'Light', 'Regular'])],
    ['Malgun Gothic', new Set(['Bold', 'Regular', 'Semilight'])],
    ['Gulim', new Set(['Regular'])],
    // Pretendard 는 굵기가 style 로만 있고 이름에는 "Variable" 이 끼어 있다
    ['Pretendard', new Set(['Black', 'Bold', 'ExtraBold', 'ExtraLight', 'Light', 'Medium', 'Regular', 'SemiBold', 'Thin'])],
  ]);

  /** 문서에 포함된 글꼴에서 읽어낸 표 (실제 덱에서 뽑은 값) */
  const docAliases: Record<string, string> = {
    '페이퍼로지 4 Regular': 'Paperlogy 4 Regular',
    '페이퍼로지 8 ExtraBold': 'Paperlogy 8 ExtraBold',
    '프리젠테이션 3 Light': 'Freesentation 3 Light',
    '프리젠테이션 6 SemiBold': 'Freesentation 6 SemiBold',
    'KoPub돋움체 Bold': 'KoPubDotum Bold',
    '맑은 고딕': 'Malgun Gothic',
    '공체 Light': 'Gulim',
  };

  /** 플러그인이 하는 순서 그대로: 문서 표 → 정적 별칭표 → family/style 분리 */
  const resolveName = (typeface: string, wantStyle: string): string => {
    const doc = docAliases[typeface];
    const names = doc ? [typeface, doc, ...aliasesFor(typeface)] : [typeface, ...aliasesFor(typeface)];
    for (const name of names) {
      const hit = pickFont(installed, name, wantStyle);
      if (hit) return `${hit.family} / ${hit.style}`;
    }
    return '없음';
  };

  console.log('\n폰트 이름 → 실제 등록된 family/style');
  const resolveCases: Array<[string, string, string]> = [
    // 문서에 포함된 글꼴에서 이름을 얻는 경우
    ['페이퍼로지 4 Regular', 'Regular', 'Paperlogy / 4 Regular'],
    ['페이퍼로지 8 ExtraBold', 'Bold', 'Paperlogy / 8 ExtraBold'],
    ['프리젠테이션 3 Light', 'Regular', 'Freesentation / 3 Light'],
    ['프리젠테이션 6 SemiBold', 'Bold', 'Freesentation / 6 SemiBold'],
    ['맑은 고딕', 'Bold', 'Malgun Gothic / Bold'],
    // 정적 별칭표로만 풀리는 경우 (문서에 포함돼 있지 않은 글꼴)
    ['KoPub돋움체 Medium', 'Regular', 'KoPubDotum / Medium'],
    ['KoPubWorld돋움체 Bold', 'Bold', 'KoPubWorldDotum / Bold'],
    ['KoPubWorld돋움체_Pro Light', 'Regular', 'KoPubWorldDotum_Pro / Light'],
    ['G마켓 산스 TTF Bold', 'Bold', 'Gmarket Sans TTF / Bold'],
    ['넥슨 풋볼고딕 B', 'Bold', 'NEXON Football Gothic / B'],
    ['나눔스퀘어 ExtraBold', 'Bold', 'NanumSquare / ExtraBold'],
    // 이름 중간에 "Variable" 이 끼어도 굵기를 잃지 않아야 한다
    ['Pretendard Variable ExtraBold', 'Regular', 'Pretendard / ExtraBold'],
    ['Pretendard Variable SemiBold', 'Regular', 'Pretendard / SemiBold'],
    ['Pretendard', 'Bold', 'Pretendard / Bold'],
  ];
  for (const [typeface, wantStyle, want] of resolveCases) {
    check(`${typeface} (${wantStyle})`, resolveName(typeface, wantStyle), want);
  }
  // 설치돼 있지 않으면 조용히 아무거나 집지 말고 없다고 해야 한다.
  check('설치 안 된 글꼴', resolveName('Dinmed', 'Regular'), '없음');

  /* ── 선 배치 ─────────────────────────────────────────────────── */

  /*
   * 값이 아니라 **선이 실제로 놓이는 두 끝점**을 확인한다.
   * create.ts 가 쓰는 것과 같은 행렬을 여기서 다시 세워, 회전까지 반영된 결과를 본다.
   */
  const endpoints = (p: {
    x: number; y: number; w: number; h: number;
  }): [number, number, number, number] => {
    const at = linePlacement({ ...p, rotation: 0, flipH: false, flipV: false });
    const rad = (at.rotation * Math.PI) / 180;
    const cos = Math.cos(rad);
    const sin = Math.sin(rad);
    const cx = at.x + at.w / 2;
    const cy = at.y;
    // 로컬 (0,0) 과 (len,0) 을 중심 기준 회전으로 옮긴다
    const put = (t: number): [number, number] => [
      cx + (t - at.w / 2) * cos,
      cy - (t - at.w / 2) * sin,
    ];
    const [x1, y1] = put(0);
    const [x2, y2] = put(at.w);
    return [x1, y1, x2, y2];
  };
  const r2 = (v: number): number => Math.round(v * 100) / 100;

  console.log('\n선 배치 — 실제로 놓이는 두 끝점');
  const lineCases: Array<[string, { x: number; y: number; w: number; h: number }, string]> = [
    // 실측 파일의 "직선 연결선 35" — 이 값이 (62,-28)→(62,28) 로 나오던 것이 버그였다
    ['세로선 (34,0) 0×56', { x: 34, y: 0, w: 0, h: 56 }, '34,0 → 34,56'],
    ['세로선 (34,41) 0×32', { x: 34, y: 41, w: 0, h: 32 }, '34,41 → 34,73'],
    ['가로선 (10,20) 100×0', { x: 10, y: 20, w: 100, h: 0 }, '10,20 → 110,20'],
    ['대각선 (0,0) 100×50', { x: 0, y: 0, w: 100, h: 50 }, '0,0 → 100,50'],
    ['대각선 (5,7) 30×40', { x: 5, y: 7, w: 30, h: 40 }, '5,7 → 35,47'],
  ];
  for (const [name, box, want] of lineCases) {
    const [x1, y1, x2, y2] = endpoints(box);
    check(name, `${r2(x1)},${r2(y1)} → ${r2(x2)},${r2(y2)}`, want);
  }

  /* ── 호(arc) 각도 ────────────────────────────────────────────── */

  /*
   * 시작점·끝점·베지어 조각 수로 확인한다. 각도 값만 보면 방향이 뒤집힌 것을 못 잡는다.
   * 90° 는 조각 1개, 180° 는 2개, 270° 는 3개다 (arc() 가 사분면마다 나눈다).
   */
  const arcShape = (prst: string, adj: Record<string, number>): string => {
    const path = presetPath(prst, 13, 13, adj)?.data ?? '';
    const nums = path.match(/-?\d+(\.\d+)?/g) ?? [];
    const segs = (path.match(/C/g) ?? []).length;
    const start = `${nums[0]},${nums[1]}`;
    const end = `${nums[nums.length - 2]},${nums[nums.length - 1]}`;
    return `${start} → ${end} (${segs}조각)`;
  };

  console.log('\n호 각도 — 시작점 → 끝점');
  // 실측 파일 슬라이드 2 의 브라켓 모서리: <a:avLst/> 라 기본값이 쓰인다.
  // 12시(6.5,0) 에서 3시(13,6.5) 로 가는 90° 모서리여야 한다 — 270° 소용돌이가 아니라.
  check('arc 기본값 (빈 avLst)', arcShape('arc', {}), '6.5,0 → 13,6.5 (1조각)');
  // 슬라이드 3 의 원호: 180.98° 에서 시작해 끝각이 0 이므로 한 바퀴를 더해 179.02° 를 쓸어간다
  check('arc 끝각 < 시작각', arcShape('arc', { adj1: 10858817, adj2: 0 }), '0,6.39 → 13,6.5 (2조각)');
  // 슬라이드 9~13 의 원호: 185.53° → 354.31° 로 끝각이 더 커서 정규화가 필요 없다 (예전에도 맞던 쪽)
  check('arc 끝각 > 시작각', arcShape('arc', { adj1: 11131912, adj2: 21258779 }), '0.03,5.87 → 12.97,5.86 (2조각)');
  // pie 는 0°→270° 라 정규화가 걸리지 않는다. 마지막 점은 중심으로 닫는 선이다.
  check('pie 기본값', arcShape('pie', {}), '13,6.5 → 6.5,6.5 (3조각)');
  // blockArc 는 180°→0° — 바깥 호 180°(2조각) + 안쪽 되돌아오는 호 180°(2조각)
  check('blockArc 기본값', arcShape('blockArc', {}), '0,6.5 → 3.25,6.5 (4조각)');

  /* ── 공통 서식 → slideLayout ─────────────────────────────────── */

  /*
   * `#레이아웃` 컴포넌트가 진짜 slideLayout 파트로 나가는지, 슬라이드가 그것을 가리키는지,
   * `#슬라이드번호` 가 글자가 아니라 **필드**로 나가는지 — XML 을 열어 확인한다.
   */
  const box = (x: number, y: number, w: number, h: number): Box =>
    ({ x, y, w, h, rot: 0, flipH: false, flipV: false });

  const withMaster: Doc = {
    ...doc,
    masters: [{
      name: '#레이아웃/백지',
      items: [
        {
          type: 'shape', name: '괘선', box: box(mm(20), mm(30), mm(170), 0),
          geom: { kind: 'line' }, fill: { kind: 'none' },
          stroke: { color: 'A9B2BB', transparency: 0, width: 1, dashType: 'solid' },
        },
        {
          type: 'shape', name: '장식', box: box(mm(20), mm(10), 20, 20),
          geom: { kind: 'custom', points: [
            { x: 0, y: 0, moveTo: true }, { x: 20, y: 0 }, { x: 10, y: 20 }, { close: true },
          ] },
          fill: { kind: 'solid', color: 'D5D9DD', transparency: 0 },
        },
        {
          type: 'text', name: '머리글', box: box(mm(100), mm(12), mm(90), mm(8)),
          align: 'right', valign: 'top', wrap: false,
          runs: [{
            text: '제 2 권', fontFace: 'Arial', fontSize: 12, bold: true,
            italic: false, underline: false, strike: false, color: 'A9B2BB', transparency: 0,
          }],
        },
      ],
      slideNumber: {
        box: box(mm(100), mm(285), mm(10), mm(5)),
        fontFace: 'Arial', fontSize: 9, color: '404040', align: 'center',
      },
    }],
    slides: [{ ...doc.slides[0], master: '#레이아웃/백지' }],
  };

  const m = await render(withMaster);
  const layoutNames = Object.keys(m.zip.files).filter((n) => /slideLayout\d+\.xml$/.test(n));
  const layoutXml = await m.zip.file(layoutNames[layoutNames.length - 1])!.async('string');
  const slideRels = await m.zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string');

  console.log('\n공통 서식 → slideLayout 파트');
  checkTruthy('레이아웃 파트가 생성됨', layoutNames.length > 0);
  checkTruthy('슬라이드가 그 레이아웃을 가리킴', slideRels.indexOf('slideLayout') >= 0);
  checkTruthy('머리글 글자가 레이아웃에 들어감', layoutXml.indexOf('제 2 권') >= 0);
  checkTruthy('선이 레이아웃에 들어감', layoutXml.indexOf('prst="line"') >= 0);
  checkTruthy('custGeom 장식이 레이아웃에 들어감', layoutXml.indexOf('<a:custGeom>') >= 0);
  // 페이지 번호는 글자가 아니라 필드여야 PowerPoint 가 장마다 다시 채운다.
  checkTruthy('슬라이드 번호가 slidenum 필드로 나감', layoutXml.indexOf('type="slidenum"') >= 0);
  checkTruthy('슬라이드 번호 자리표시자가 붙음', layoutXml.indexOf('type="sldNum"') >= 0);
  // 공통 서식으로 뺀 것이 슬라이드에도 중복으로 들어가면 안 된다.
  checkTruthy('머리글이 슬라이드에 중복되지 않음', m.xml.indexOf('제 2 권') < 0);

  console.log(`\n■ 마스터 판정`);
  const multiRun: TextItem = {
    type: 'text', name: '2서식', box: box(0, 0, 10, 10), align: 'left', valign: 'top', wrap: true,
    runs: [
      { text: 'A', fontFace: 'Arial', fontSize: 12, bold: true, italic: false, underline: false, strike: false, color: '000000', transparency: 0 },
      { text: 'B', fontFace: 'Arial', fontSize: 10, bold: false, italic: false, underline: false, strike: false, color: '000000', transparency: 0 },
    ],
  };
  checkTruthy('서식 하나짜리 텍스트는 공통 서식에 들어감', fitsInMaster(withMaster.masters[0].items[2]));
  checkTruthy('서식 여러 개인 텍스트는 못 들어감', !fitsInMaster(multiRun));
  checkTruthy('custGeom 도형은 들어감', fitsInMaster(withMaster.masters[0].items[1]));

  /* ── 같은 이미지 공유 ────────────────────────────────────────── */

  /*
   * 실측한 원본 덱은 이미지 참조 322개가 파일 235개를 가리키고, 내용이 같은데 따로 저장된 것이
   * 하나도 없다. PPTX 는 원래 그렇게 쓴다. 우리도 같은 이미지는 파일 하나만 두어야 한다.
   */
  const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
  const OTHER = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
  const shared: Doc = {
    ...doc,
    slides: [1, 2, 3].map((n) => ({
      name: `${n}장`,
      fill: { kind: 'solid', color: 'FFFFFF', transparency: 0 },
      items: [
        // 세 장에 같은 이미지
        { type: 'image', name: '공용', box: box(10, 10, 50, 50), data: PNG, mime: 'image/png', sizing: 'stretch' },
        // 마지막 장에만 다른 이미지 — 이건 따로 저장돼야 한다
        ...(n === 3
          ? [{ type: 'image' as const, name: '단독', box: box(70, 10, 50, 50), data: OTHER, mime: 'image/png' as const, sizing: 'stretch' as const }]
          : []),
      ],
    })),
  };
  const sharedOut = await render(shared);
  const mediaParts = Object.keys(sharedOut.zip.files).filter((n) => /^ppt\/media\/.+\.\w+$/.test(n));
  let imageRefs = 0;
  for (const n of Object.keys(sharedOut.zip.files).filter((x) => /^ppt\/slides\/_rels\/.+\.rels$/.test(x))) {
    const rels = await sharedOut.zip.file(n)!.async('string');
    imageRefs += (rels.match(/Type="[^"]*\/image"/g) ?? []).length;
  }

  console.log('\n같은 이미지 공유 — 내용 비교로');
  check('참조는 4번 (같은 것 3 + 다른 것 1)', imageRefs, 4);
  check('저장된 파일은 2개', mediaParts.length, 2);

  /*
   * Figma 의 imageHash 가 있으면 그것만 보고 판정한다 — 내용 주소라 같은 그림이면 반드시 같다.
   * 여기서는 그 경로가 실제로 동작하는지 보려고, 데이터가 달라도 열쇠가 같으면 합쳐지는지 본다.
   * (실제로는 열쇠가 같은데 데이터가 다를 수 없다. Figma 가 내용으로 주소를 매기기 때문이다.)
   */
  const keyed: Doc = {
    ...doc,
    slides: [1, 2].map((n) => ({
      name: `${n}장`,
      fill: { kind: 'solid', color: 'FFFFFF', transparency: 0 },
      items: [{
        type: 'image' as const, name: '같은 그림', box: box(10, 10, 50, 50),
        key: 'figma-hash-abc', data: n === 1 ? PNG : OTHER,
        mime: 'image/png' as const, sizing: 'stretch' as const,
      }],
    })),
  };
  const keyedOut = await render(keyed);
  const keyedParts = Object.keys(keyedOut.zip.files).filter((n) => /^ppt\/media\/.+\.\w+$/.test(n));
  console.log('\n같은 이미지 공유 — Figma 주소로');
  check('열쇠가 같으면 파일 하나', keyedParts.length, 1);

  /* ── 도형 id 충돌 ────────────────────────────────────────────── */

  /*
   * OOXML 은 슬라이드 안에서 도형 id 가 유일해야 한다. 겹치면 PowerPoint 가 그 도형을
   * "읽을 수 없는 내용"으로 지우고 복구 창을 띄운다.
   *
   * PptxGenJS 는 일반 도형에 `순번 + 2` 를, 슬라이드 번호 자리표시자에는 25 를 준다.
   * 그래서 도형이 24개를 넘는 순간 부딪힌다. 30개를 깔아 그 지점을 지나가게 한다.
   */
  const crowded: Doc = {
    ...doc,
    masters: [{
      name: '#레이아웃/번호',
      items: [],
      slideNumber: {
        box: box(mm(100), mm(285), mm(10), mm(5)),
        fontFace: 'Arial', fontSize: 9, color: '404040', align: 'center',
      },
    }],
    slides: [{
      name: '도형 30개',
      fill: { kind: 'solid', color: 'FFFFFF', transparency: 0 },
      master: '#레이아웃/번호',
      items: Array.from({ length: 30 }, (_, i) => ({
        type: 'shape' as const,
        name: `도형 ${i}`,
        box: box(10 + i, 10, 20, 20),
        geom: { kind: 'rect' as const },
        fill: { kind: 'solid' as const, color: '888888', transparency: 0 },
      })),
    }],
  };
  const crowdOut = await render(crowded);
  const crowdIds = [...crowdOut.xml.matchAll(/<p:cNvPr id="(\d+)"/g)].map((m) => m[1]);
  const crowdLayouts = Object.keys(crowdOut.zip.files).filter((n) => /slideLayout\d+\.xml$/.test(n));
  const numberedLayout = await Promise.all(
    crowdLayouts.map(async (n) => crowdOut.zip.file(n)!.async('string')),
  );

  console.log('\n도형 id 충돌 (도형 30개 + 슬라이드 번호)');
  check('도형 id 가 전부 다름', new Set(crowdIds).size, crowdIds.length);
  checkTruthy('슬라이드에는 번호 자리표시자를 넣지 않음', crowdOut.xml.indexOf('type="sldNum"') < 0);
  // 번호 자체는 레이아웃에 남아 PowerPoint 가 장마다 채운다.
  checkTruthy(
    '레이아웃에는 slidenum 필드가 남음',
    numberedLayout.some((x) => x.indexOf('type="slidenum"') >= 0),
  );

  /* ── 여러 장 흘려보내기 ──────────────────────────────────────── */

  /*
   * 내보내기는 장을 한 장씩 받아 붙인다 (문서를 통째로 넘기면 이미지가 부푼 만큼 실려 터진다).
   * 여러 번 붙여도 장수와 순서가 그대로인지 확인한다.
   */
  const many: Doc = {
    ...doc,
    slides: [1, 2, 3].map((n) => ({
      name: `${n}장`,
      fill: { kind: 'solid', color: 'FFFFFF', transparency: 0 },
      items: [{
        type: 'text', name: '쪽번호', box: box(10, 10, 100, 20),
        align: 'left', valign: 'top', wrap: false,
        runs: [{
          text: `장 ${n}`, fontFace: 'Arial', fontSize: 12, bold: false,
          italic: false, underline: false, strike: false, color: '000000', transparency: 0,
        }],
      }],
    })),
  };
  const streamed = new PptxBuilder(many);
  for (const s of many.slides) streamed.add(s);
  const zipMany = await JSZip.loadAsync(
    await streamed.presentation.write({ outputType: 'nodebuffer' }) as Buffer,
  );
  const parts = Object.keys(zipMany.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n));

  console.log('\n여러 장 흘려보내기');
  check('붙인 만큼 슬라이드 파트가 생김', parts.length, 3);
  for (let n = 1; n <= 3; n++) {
    const xml = await zipMany.file(`ppt/slides/slide${n}.xml`)!.async('string');
    checkTruthy(`${n}번째 장의 내용이 제자리`, xml.indexOf(`장 ${n}`) >= 0);
  }

  /* ── 열린 경로 내보내기 ──────────────────────────────────────── */

  /*
   * 브라켓의 원호는 **열린** 경로다. 닫아서 내보내면 끝점과 시작점이 직선으로 이어져
   * 초승달이 되고, 칠을 안 적으면 PowerPoint 가 기본 서식으로 채워 버린다.
   * 원본과 같이 열린 경로 + 선 + noFill 로 나가야 한다.
   */
  const openArc: Doc = {
    ...doc,
    slides: [{
      name: '열린 호',
      fill: { kind: 'solid', color: 'FFFFFF', transparency: 0 },
      items: [{
        type: 'shape',
        name: '원호',
        box: { x: 100, y: 100, w: 20, h: 20, rot: 0, flipH: false, flipV: false },
        geom: { kind: 'custom', points: [
          { x: 0, y: 0, moveTo: true },
          { x: 20, y: 20, c: [11, 0, 20, 9] },
        ] },
        fill: { kind: 'none' },
        stroke: { color: '808080', transparency: 0, width: 1, dashType: 'solid' },
      }],
    }],
  };
  const arcOut = await render(openArc);

  console.log('\n열린 경로 내보내기');
  checkTruthy('custGeom 으로 나감', arcOut.xml.indexOf('<a:custGeom>') >= 0);
  checkTruthy('베지어가 들어감', arcOut.xml.indexOf('<a:cubicBezTo>') >= 0);
  checkTruthy('경로를 닫지 않음 (a:close 없음)', arcOut.xml.indexOf('<a:close') < 0);
  // 칠을 안 적으면 PowerPoint 가 기본 강조색으로 채운다. 명시적으로 비워야 한다.
  checkTruthy('칠 없음이 <a:noFill/> 로 명시됨', arcOut.xml.indexOf('<a:noFill/>') >= 0);
  checkTruthy('선은 그대로 나감', arcOut.xml.indexOf('<a:ln w=') >= 0);

  /* ── 경로 배치 (벡터 노드) ───────────────────────────────────── */

  /*
   * 브라켓 모서리 실측: 원호는 13×13 상자에 이름 붙어 있지만 실제로는 오른쪽 위 사분면만 쓴다.
   * 이름상의 상자로 변환을 걸면 도형이 밀리고, flipH 의 반전 축까지 어긋난다.
   *
   * create.ts 와 같은 계산을 여기서 다시 세워, 경로의 네 귀퉁이가 어디에 놓이는지 본다.
   */
  const placedBounds = (
    data: string, p: { x: number; y: number; w: number; h: number; flipH?: boolean },
  ): string => {
    const bb = pathBounds(data);
    const w = p.w;
    const h = p.h;
    const fx = p.flipH ? -1 : 1;
    // 회전 0 기준 — applyPlacement 의 a,b,c,d 와 같다
    const a = fx;
    const d = 1;
    const cx = p.x + w / 2;
    const cy = p.y + h / 2;
    const tx = cx - a * (w / 2) + a * bb.x;
    const ty = cy - d * (h / 2) + d * bb.y;
    // 노드 로컬 (0,0)~(bw,bh) 가 놓이는 실제 범위
    const xs = [tx, tx + a * bb.w];
    const ys = [ty, ty + d * bb.h];
    const r2n = (v: number): number => Math.round(v * 100) / 100;
    return `${r2n(Math.min(...xs))},${r2n(Math.min(...ys))} ~ ${r2n(Math.max(...xs))},${r2n(Math.max(...ys))}`;
  };

  console.log('\n경로 배치 — 원호가 실제로 덮는 범위');
  const arc90 = presetPath('arc', 13, 13, {})!.data;
  check('원호 경계는 사분면 하나', (() => {
    const b = pathBounds(arc90);
    return `${Math.round(b.x * 10) / 10},${Math.round(b.y * 10) / 10} ${Math.round(b.w * 10) / 10}×${Math.round(b.h * 10) / 10}`;
  })(), '6.5,0 6.5×6.5');
  // 오른쪽 모서리 (원호 46): 반전 없음 — 상자의 오른쪽 위에 붙어야 한다
  check('오른쪽 모서리', placedBounds(arc90, { x: 548, y: 310, w: 13, h: 13 }), '554.5,310 ~ 561,316.5');
  // 왼쪽 모서리 (원호 43): flipH — 같은 상자의 **왼쪽** 위에 붙어야 한다
  check('왼쪽 모서리 (flipH)', placedBounds(arc90, { x: 36, y: 310, w: 13, h: 13, flipH: true }), '36,310 ~ 42.5,316.5');

  /* ── 패턴 채우기 ─────────────────────────────────────────────── */

  /*
   * Figma 에 무늬 개념이 없어 앞색·뒷색을 밀도만큼 섞은 단색으로 근사한다.
   * 안 읽으면 칠 없는 도형이 되어 노드가 통째로 버려진다 — 실제로 십자가 도형이 그렇게 사라졌었다.
   */
  const pattern = (prst: string, fg: string, bg: string): string => {
    const xml = parseXml(`<a:spPr><a:pattFill prst="${prst}">`
      + `<a:fgClr><a:srgbClr val="${fg}"/></a:fgClr>`
      + `<a:bgClr><a:srgbClr val="${bg}"/></a:bgClr></a:pattFill></a:spPr>`);
    // parseXml 은 #root 로 감싸서 돌려준다. readFill 은 spPr 자체를 받는다.
    const paint = readFill(deepXml(xml, 'spPr'), { scheme: {}, map: {} });
    return paint && paint.kind === 'solid' ? paint.color : '없음';
  };

  console.log('\n패턴 채우기 — 밀도만큼 섞은 단색');
  check('dkUpDiag 검정/흰색 (50%)', pattern('dkUpDiag', '000000', 'FFFFFF'), '808080');
  check('pct25 검정/흰색', pattern('pct25', '000000', 'FFFFFF'), 'BFBFBF');
  check('pct75 검정/흰색', pattern('pct75', '000000', 'FFFFFF'), '404040');
  // 모르는 이름은 50% 로 떨어지되, 도형이 사라지지는 않아야 한다
  check('모르는 무늬', pattern('notAPattern', '000000', 'FFFFFF'), '808080');

  if (failures.length > 0) {
    console.error(`\n실패 ${failures.length}건: ${failures.join(', ')}`);
    process.exit(1);
  }
  console.log('\n전체 통과');
}

main();
