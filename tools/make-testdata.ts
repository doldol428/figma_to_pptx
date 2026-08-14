/**
 * 왕복 정확도를 재기 위한 시험용 PPTX 를 만든다.
 *
 *   npm run testdata -- "<나올 파일 경로>"
 *
 * 실제 제안서로는 무엇이 왜 틀어졌는지 가려내기 어렵다. 도형이 겹치고, 그룹에 파묻히고,
 * 좌표가 어중간해서 0.3mm 어긋난 것이 반올림 때문인지 계산이 틀린 것인지 알 수 없다.
 *
 * 그래서 이 파일은 **좌표를 전부 mm 정수** 로 놓는다. 왕복 후 0.1mm라도 어긋나면 그건
 * 반올림이 아니라 우리 계산이 틀린 것이다. 한 장에 한 갈래씩만 담아 원인이 섞이지 않게 한다.
 *
 * 덮는 범위는 **preset.ts 가 이름으로 다루는 87종 전부** 다. 일부만 넣으면 나머지가
 * 언제 깨졌는지 알 방법이 없다. 지원하지 않는 도형도 몇 개 넣어 대체 동작을 확인한다.
 *
 * adj 값(모서리 반경·화살표 목·호 각도 …)은 PptxGenJS 로는 대부분 넣을 수 없어,
 * 다 만든 뒤 zip 을 열어 직접 심는다. 실제로 틀렸던 버그가 죄다 adj 에서 나왔다.
 */
import { writeFile } from 'node:fs/promises';
import JSZip from 'jszip';
import PptxGenJS from 'pptxgenjs';

const MM_IN = 1 / 25.4;
const mm = (v: number): number => v * MM_IN;

const INK = '18539B';
const FILL = 'D5D9DD';

/** preset.ts 가 이름으로 다루는 도형 (그 파일의 case 목록과 같아야 한다) */
const SUPPORTED = [
  'rect', 'flowChartProcess', 'flowChartPredefinedProcess', 'ellipse', 'flowChartConnector',
  'roundRect', 'flowChartAlternateProcess', 'round1Rect', 'round2SameRect', 'round2DiagRect',
  'snip1Rect', 'snip2SameRect', 'snip2DiagRect', 'snipRoundRect', 'triangle', 'rtTriangle',
  'diamond', 'flowChartDecision', 'parallelogram', 'flowChartInputOutput', 'trapezoid',
  'pentagon', 'hexagon', 'heptagon', 'octagon', 'decagon', 'dodecagon',
  'star4', 'star5', 'star6', 'star8', 'star10', 'star12',
  'rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'leftRightArrow', 'upDownArrow',
  'homePlate', 'chevron', 'donut', 'pie', 'arc', 'chord', 'blockArc', 'teardrop', 'moon',
  'frame', 'halfFrame', 'corner', 'plaque', 'bevel', 'plus',
  'flowChartTerminator', 'flowChartDocument', 'flowChartMagneticDisk', 'flowChartPreparation',
  'flowChartManualOperation',
  'leftBracket', 'rightBracket', 'bracketPair', 'leftBrace', 'rightBrace', 'bracePair',
  'cube', 'can',
  'curvedRightArrow', 'curvedUpArrow', 'curvedDownArrow', 'curvedLeftArrow',
  'mathPlus', 'mathMinus', 'mathEqual', 'mathMultiply',
  'wedgeEllipseCallout', 'wedgeRectCallout',
];

/** 우리가 모르는 도형 — 사각형으로 대체하고 경고를 남기는지 본다 */
const UNSUPPORTED = [
  'smileyFace', 'sun', 'heart', 'cloud', 'lightningBolt', 'gear6',
  'funnel', 'ribbon', 'wave', 'verticalScroll', 'swooshArrow', 'circularArrow',
];

/** 선으로 그려지는 연결선 (도형 격자가 아니라 따로 시험한다) */
const CONNECTORS = [
  'straightConnector1', 'bentConnector2', 'bentConnector3', 'bentConnector4', 'bentConnector5',
  'curvedConnector2', 'curvedConnector3', 'curvedConnector4', 'curvedConnector5',
];

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'A4', width: mm(210), height: mm(297) });
pptx.layout = 'A4';

type Slide = ReturnType<PptxGenJS['addSlide']>;
const shapeOf = (name: string): PptxGenJS.SHAPE_NAME => name as PptxGenJS.SHAPE_NAME;

function title(s: Slide, text: string): void {
  s.addText(text, {
    x: mm(10), y: mm(8), w: mm(190), h: mm(8),
    fontSize: 12, bold: true, color: '000000', margin: 0, isTextBox: true, fit: 'none',
    objectName: `제목 ${text.slice(0, 6)}`,
  });
}

/** 도형 격자 한 장 — 5칸 × 6줄, 아래에 이름표 */
function grid(s: Slide, names: string[], tag: string): void {
  names.forEach((name, i) => {
    const col = i % 5;
    const row = Math.floor(i / 5);
    const x = 20 + col * 34;
    const y = 25 + row * 43;
    s.addShape(shapeOf(name), {
      x: mm(x), y: mm(y), w: mm(26), h: mm(26),
      fill: { color: FILL }, line: { color: INK, width: 1 },
      objectName: `${tag}:${name}`,
    });
    s.addText(name, {
      x: mm(x), y: mm(y + 27), w: mm(30), h: mm(5),
      fontSize: 6, color: '666666', margin: 0, isTextBox: true, fit: 'none',
      objectName: `이름표:${name}`,
    });
  });
}

/* ── 도형 87종 ────────────────────────────────────────────────── */
for (let i = 0; i < SUPPORTED.length; i += 30) {
  const chunk = SUPPORTED.slice(i, i + 30);
  const s = pptx.addSlide();
  title(s, `지원 도형 ${i + 1}–${i + chunk.length} / ${SUPPORTED.length}`);
  grid(s, chunk, '도형');
}

/* ── 모르는 도형 ──────────────────────────────────────────────── */
{
  const s = pptx.addSlide();
  title(s, '지원하지 않는 도형 — 사각형으로 대체되고 경고가 남아야 한다');
  grid(s, UNSUPPORTED, '미지원');
}

/* ── adj 값 변형 ──────────────────────────────────────────────── */
{
  const s = pptx.addSlide();
  title(s, 'adj 값 — 기본값만으로는 안 걸리는 것들 (모서리·목·두께·각도)');

  // 이름에 심을 값을 적어 두면, 다 만든 뒤 zip 에서 찾아 넣는다
  const cases: Array<[string, string]> = [
    ['roundRect', 'adj=5000'], ['roundRect', 'adj=50000'],
    ['chevron', 'adj=10000'], ['chevron', 'adj=45000'],
    ['homePlate', 'adj=10000'], ['homePlate', 'adj=45000'],
    ['donut', 'adj=10000'], ['donut', 'adj=40000'],
    ['arc', 'adj1=0,adj2=10800000'], ['arc', 'adj1=16200000,adj2=5400000'],
    ['pie', 'adj1=0,adj2=5400000'], ['blockArc', 'adj1=10800000,adj2=0,adj3=10000'],
    ['star5', 'adj=20000'], ['frame', 'adj1=5000'],
    ['corner', 'adj1=20000,adj2=20000'], ['bevel', 'adj=30000'],
    ['can', 'adj=10000'], ['cube', 'adj=40000'],
    ['rightArrow', 'adj1=30000,adj2=30000'], ['leftBracket', 'adj=30000'],
  ];

  cases.forEach(([name, adj], i) => {
    const col = i % 5;
    const row = Math.floor(i / 5);
    const x = 20 + col * 34;
    const y = 25 + row * 43;
    s.addShape(shapeOf(name), {
      x: mm(x), y: mm(y), w: mm(26), h: mm(26),
      fill: { color: 'E3F2FD' }, line: { color: INK, width: 1 },
      objectName: `adj|${name}|${adj}`,
    });
    s.addText(`${name}\n${adj}`, {
      x: mm(x), y: mm(y + 27), w: mm(32), h: mm(8),
      fontSize: 6, color: '666666', margin: 0, isTextBox: true, fit: 'none',
      objectName: `이름표:adj ${name} ${i}`,
    });
  });
}

/* ── 선: 방향 × 뒤집기 ───────────────────────────────────────── */
{
  const s = pptx.addSlide();
  title(s, '선 — 방향 × 뒤집기 (끝점이 mm 정수 위에 있어야 한다)');

  const cases: Array<[string, number, number, number, number, boolean, boolean]> = [
    ['가로', 20, 30, 60, 0, false, false],
    ['가로 flipH', 20, 40, 60, 0, true, false],
    ['세로', 20, 50, 0, 40, false, false],
    ['세로 flipV', 30, 50, 0, 40, false, true],
    ['대각 ↘', 50, 50, 40, 40, false, false],
    ['대각 flipH ↙', 100, 50, 40, 40, true, false],
    ['대각 flipV ↗', 150, 50, 40, 40, false, true],
    ['대각 flipHV ↖', 50, 110, 40, 40, true, true],
    ['긴 대각', 100, 110, 90, 60, false, false],
    ['짧은 대각', 20, 180, 10, 10, false, false],
    ['얇고 긴 가로', 20, 200, 170, 0, false, false],
    ['얇고 긴 세로', 190, 30, 0, 170, false, false],
  ];
  for (const [name, x, y, w, h, flipH, flipV] of cases) {
    s.addShape(pptx.ShapeType.line, {
      x: mm(x), y: mm(y), w: mm(w), h: mm(h), flipH, flipV,
      line: { color: INK, width: 1.5 }, objectName: name,
    });
  }
  [15, 30, 45, 90].forEach((deg, i) => {
    s.addShape(pptx.ShapeType.line, {
      x: mm(20 + i * 40), y: mm(230), w: mm(30), h: mm(20), rotate: deg,
      line: { color: '888888', width: 1.5 }, objectName: `대각 rot${deg}`,
    });
  });

  // 연결선 — 선으로 떨어지는지
  CONNECTORS.forEach((name, i) => {
    s.addShape(shapeOf(name), {
      x: mm(20 + (i % 5) * 34), y: mm(265 + Math.floor(i / 5) * 12), w: mm(28), h: mm(8),
      line: { color: 'E33348', width: 1 }, objectName: `연결선:${name}`,
    });
  });
}

/* ── 회전과 뒤집기 ───────────────────────────────────────────── */
{
  const s = pptx.addSlide();
  title(s, '회전 · 뒤집기 — 귀퉁이가 어디로 가는지');

  [0, 15, 30, 45, 90, 135, 180, 270].forEach((deg, i) => {
    s.addShape(pptx.ShapeType.rect, {
      x: mm(20 + (i % 4) * 45), y: mm(30 + Math.floor(i / 4) * 55), w: mm(34), h: mm(18),
      rotate: deg, fill: { color: 'CCE0F5' }, line: { color: INK, width: 1 },
      objectName: `회전 ${deg}도`,
    });
  });

  const flips: Array<[string, boolean, boolean]> = [
    ['그대로', false, false], ['flipH', true, false], ['flipV', false, true], ['flipHV', true, true],
  ];
  flips.forEach(([label, fh, fv], i) => {
    s.addShape(pptx.ShapeType.triangle, {
      x: mm(20 + i * 45), y: mm(150), w: mm(34), h: mm(26), flipH: fh, flipV: fv,
      fill: { color: 'FFE0B2' }, line: { color: '888888', width: 1 },
      objectName: `삼각 ${label}`,
    });
    s.addShape(pptx.ShapeType.homePlate, {
      x: mm(20 + i * 45), y: mm(190), w: mm(34), h: mm(26), flipH: fh, flipV: fv,
      fill: { color: 'FFE0B2' }, line: { color: '888888', width: 1 },
      objectName: `오각 ${label}`,
    });
    // 회전과 뒤집기가 같이 걸린 경우
    s.addShape(pptx.ShapeType.rtTriangle, {
      x: mm(20 + i * 45), y: mm(230), w: mm(34), h: mm(26), flipH: fh, flipV: fv, rotate: 30,
      fill: { color: 'C8E6C9' }, line: { color: '888888', width: 1 },
      objectName: `직각삼각 ${label} rot30`,
    });
  });
}

/* ── 텍스트 ──────────────────────────────────────────────────── */
{
  const s = pptx.addSlide();
  title(s, '텍스트 — 정렬 · 여러 서식 · 도형 안의 글자');

  (['left', 'center', 'right'] as const).forEach((align, i) => {
    s.addText(`${align} 정렬 · 한글 Abc 123`, {
      x: mm(20), y: mm(25 + i * 14), w: mm(170), h: mm(10),
      align, fontSize: 12, margin: 0, isTextBox: true, fit: 'none',
      objectName: `정렬 ${align}`,
    });
  });

  s.addText([
    { text: '큰 굵은 글씨', options: { fontSize: 18, bold: true, color: INK, breakLine: true } },
    { text: '작은 보통 글씨', options: { fontSize: 10, color: '404040' } },
  ], {
    x: mm(20), y: mm(70), w: mm(170), h: mm(20),
    margin: 0, isTextBox: true, fit: 'none', objectName: '여러 서식',
  });

  (['top', 'middle', 'bottom'] as const).forEach((valign, i) => {
    s.addText(valign, {
      x: mm(20 + i * 58), y: mm(100), w: mm(50), h: mm(30),
      valign, fontSize: 11, margin: 0, isTextBox: true, fit: 'none',
      objectName: `세로 ${valign}`,
    });
  });

  s.addText('도형 안의 글자', {
    x: mm(20), y: mm(140), w: mm(80), h: mm(25),
    shape: pptx.ShapeType.roundRect, fill: { color: 'E3F2FD' },
    line: { color: INK, width: 1 }, align: 'center', valign: 'middle',
    fontSize: 12, margin: 0, objectName: '도형+글자',
  });

  s.addText('여백 5mm 인 상자', {
    x: mm(110), y: mm(140), w: mm(80), h: mm(25),
    margin: [mm(5) * 72, mm(5) * 72, mm(5) * 72, mm(5) * 72],
    fontSize: 11, isTextBox: true, fit: 'none',
    line: { color: '888888', width: 0.5 }, objectName: '여백 상자',
  });

  s.addText('줄바꿈 없는 긴 글씨 abcdefghijklmnop', {
    x: mm(20), y: mm(180), w: mm(60), h: mm(10),
    fontSize: 11, wrap: false, margin: 0, isTextBox: true, fit: 'none',
    objectName: '줄바꿈 없음',
  });
}

/* ── 좌표 격자 ───────────────────────────────────────────────── */
{
  const s = pptx.addSlide();
  title(s, '좌표 격자 — 20mm 눈금과 교차점 표식');

  for (let x = 20; x <= 190; x += 20) {
    s.addShape(pptx.ShapeType.line, {
      x: mm(x), y: mm(30), w: 0, h: mm(240),
      line: { color: 'EEEEEE', width: 0.5 }, objectName: `세로눈금 ${x}`,
    });
  }
  for (let y = 30; y <= 270; y += 20) {
    s.addShape(pptx.ShapeType.line, {
      x: mm(20), y: mm(y), w: mm(170), h: 0,
      line: { color: 'EEEEEE', width: 0.5 }, objectName: `가로눈금 ${y}`,
    });
  }
  for (const [x, y] of [[20, 30], [100, 130], [180, 250], [60, 210], [140, 70]]) {
    s.addShape(pptx.ShapeType.ellipse, {
      x: mm(x - 2), y: mm(y - 2), w: mm(4), h: mm(4),
      fill: { color: 'E33348' }, line: { type: 'none' },
      objectName: `표식 ${x},${y}`,
    });
  }
}

/* ── adj 값을 zip 에 직접 심는다 ─────────────────────────────── */

const buf = await pptx.write({ outputType: 'nodebuffer' }) as Buffer;
const zip = await JSZip.loadAsync(buf);
let injected = 0;

for (const name of Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))) {
  let xml = await zip.file(name)!.async('string');
  /*
   * 이름에 `adj|<도형>|adj=…` 를 적어 두었다. 그 도형의 빈 <a:avLst/> 를 실제 값으로 채운다.
   * PptxGenJS 로는 rectRadius 말고는 adj 를 넣을 수 없어서, 다 만든 뒤 여기서 손본다.
   */
  xml = xml.replace(
    /<p:cNvPr id="\d+" name="adj\|[^|]+\|([^"]+)"[\s\S]*?<a:prstGeom prst="[^"]+"><a:avLst>/g,
    (block, spec: string) => {
      injected++;
      const gds = spec.split(',')
        .map((pair) => {
          const [k, v] = pair.split('=');
          return `<a:gd name="${k}" fmla="val ${v}"/>`;
        })
        .join('');
      return block + gds;
    },
  );
  zip.file(name, xml);
}

const path = process.argv[2] ?? 'TestDATA.pptx';
const final = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
await writeFile(path, final);

console.log(`${path}`);
console.log(`  슬라이드 ${Object.keys(zip.files).filter((n) => /slide\d+\.xml$/.test(n)).length}장 · ${(final.length / 1024).toFixed(0)} KB`);
console.log(`  지원 도형 ${SUPPORTED.length}종 · 미지원 ${UNSUPPORTED.length}종 · 연결선 ${CONNECTORS.length}종`);
console.log(`  adj 값 심은 도형 ${injected}개`);
console.log('  좌표는 전부 mm 정수 — 왕복 후 0.1mm 라도 어긋나면 계산이 틀린 것이다.');
