/**
 * 왕복 정확도를 재기 위한 시험용 PPTX 를 만든다.
 *
 *   npm run testdata -- "<나올 파일 경로>"
 *
 * 실제 제안서로는 무엇이 왜 틀어졌는지 가려내기 어렵다. 도형이 서로 겹치고, 그룹에 파묻히고,
 * 좌표가 어중간해서 0.3mm 어긋난 것이 반올림 때문인지 계산이 틀린 것인지 알 수 없다.
 *
 * 그래서 이 파일은 **좌표를 전부 10mm 배수의 정수 EMU** 로 놓는다. 왕복 후 0.1mm라도
 * 어긋나면 그건 반올림이 아니라 우리 계산이 틀린 것이다. 한 장에 한 가지씩만 시험해
 * 원인이 섞이지 않게 한다.
 */
import { writeFile } from 'node:fs/promises';
import PptxGenJS from 'pptxgenjs';

const MM_IN = 1 / 25.4;
/** mm → inch. PptxGenJS 가 inch 로 받으므로 여기서 한 번만 바꾼다. */
const mm = (v: number): number => v * MM_IN;

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'A4', width: mm(210), height: mm(297) });
pptx.layout = 'A4';

const GREY = '888888';
const BLUE = '18539B';

/** 장 제목 — 어느 장이 무엇을 재는지 남긴다 */
function title(s: ReturnType<PptxGenJS['addSlide']>, text: string): void {
  s.addText(text, {
    x: mm(10), y: mm(10), w: mm(190), h: mm(10),
    fontSize: 14, bold: true, color: '000000', margin: 0, isTextBox: true,
  });
}

/* ── 1장: 선 — 방향과 뒤집기의 모든 조합 ──────────────────────── */
{
  const s = pptx.addSlide();
  title(s, '1. 선 — 방향 × 뒤집기 (끝점이 정확히 10mm 격자 위에 있어야 한다)');

  const cases: Array<[string, number, number, number, number, boolean, boolean]> = [
    //  이름          x    y    w    h   flipH  flipV
    ['가로',          20,  40,  60,   0, false, false],
    ['가로 flipH',    20,  50,  60,   0, true,  false],
    ['세로',          20,  60,   0,  40, false, false],
    ['세로 flipV',    30,  60,   0,  40, false, true],
    ['대각 ↘',        50,  60,  40,  40, false, false],
    ['대각 flipH ↙',  100, 60,  40,  40, true,  false],
    ['대각 flipV ↗',  150, 60,  40,  40, false, true],
    ['대각 flipHV ↖', 50, 120,  40,  40, true,  true],
    ['긴 대각',      100, 120,  90,  60, false, false],
    ['짧은 대각',     20, 200,  10,  10, false, false],
  ];

  for (const [name, x, y, w, h, flipH, flipV] of cases) {
    s.addShape(pptx.ShapeType.line, {
      x: mm(x), y: mm(y), w: mm(w), h: mm(h), flipH, flipV,
      line: { color: BLUE, width: 1.5 },
      objectName: name,
    });
  }

  // 회전이 걸린 선
  s.addShape(pptx.ShapeType.line, {
    x: mm(120), y: mm(200), w: mm(60), h: mm(0), rotate: 30,
    line: { color: GREY, width: 1.5 }, objectName: '가로 rot30',
  });
  s.addShape(pptx.ShapeType.line, {
    x: mm(120), y: mm(230), w: mm(60), h: mm(20), rotate: 45, flipV: true,
    line: { color: GREY, width: 1.5 }, objectName: '대각 flipV rot45',
  });
}

/* ── 2장: Figma 에 없는 preset 도형 ───────────────────────────── */
{
  const s = pptx.addSlide();
  title(s, '2. preset 도형 — 경로로 다시 그려지는 것들');

  const shapes: PptxGenJS.SHAPE_NAME[] = [
    pptx.ShapeType.rect, pptx.ShapeType.roundRect, pptx.ShapeType.ellipse,
    pptx.ShapeType.triangle, pptx.ShapeType.hexagon, pptx.ShapeType.donut,
    pptx.ShapeType.arc, pptx.ShapeType.chevron, pptx.ShapeType.homePlate,
    pptx.ShapeType.bracketPair, pptx.ShapeType.snip1Rect, pptx.ShapeType.round1Rect,
    pptx.ShapeType.round2SameRect, pptx.ShapeType.halfFrame, pptx.ShapeType.flowChartDecision,
    pptx.ShapeType.pie, pptx.ShapeType.blockArc, pptx.ShapeType.plus,
  ];

  shapes.forEach((shape, i) => {
    const col = i % 4;
    const row = Math.floor(i / 4);
    s.addShape(shape, {
      x: mm(20 + col * 45), y: mm(30 + row * 45), w: mm(30), h: mm(30),
      fill: { color: 'D5D9DD' },
      line: { color: BLUE, width: 1 },
      objectName: String(shape),
    });
  });
}

/* ── 3장: 회전과 뒤집기 ──────────────────────────────────────── */
{
  const s = pptx.addSlide();
  title(s, '3. 회전 — 사각형의 네 귀퉁이가 어디로 가는지');

  [0, 15, 30, 45, 90, 180].forEach((deg, i) => {
    s.addShape(pptx.ShapeType.rect, {
      x: mm(20 + (i % 3) * 60), y: mm(40 + Math.floor(i / 3) * 70), w: mm(40), h: mm(20),
      rotate: deg, fill: { color: 'CCE0F5' }, line: { color: BLUE, width: 1 },
      objectName: `회전 ${deg}도`,
    });
  });

  // 뒤집힌 삼각형 — 회전만으로는 재현되지 않는다
  s.addShape(pptx.ShapeType.triangle, {
    x: mm(20), y: mm(200), w: mm(40), h: mm(30), flipV: true,
    fill: { color: 'FFE0B2' }, line: { color: GREY, width: 1 }, objectName: '삼각 flipV',
  });
  s.addShape(pptx.ShapeType.homePlate, {
    x: mm(80), y: mm(200), w: mm(40), h: mm(30), flipH: true,
    fill: { color: 'FFE0B2' }, line: { color: GREY, width: 1 }, objectName: '오각 flipH',
  });
}

/* ── 4장: 텍스트 ─────────────────────────────────────────────── */
{
  const s = pptx.addSlide();
  title(s, '4. 텍스트 — 정렬 · 여러 서식 · 자동 크기');

  (['left', 'center', 'right'] as const).forEach((align, i) => {
    s.addText(`${align} 정렬 · 한글 Abc 123`, {
      x: mm(20), y: mm(30 + i * 15), w: mm(170), h: mm(10),
      align, fontSize: 12, margin: 0, isTextBox: true, fit: 'none',
      line: { color: 'DDDDDD', width: 0.5 },
      objectName: `정렬 ${align}`,
    });
  });

  // 한 상자 안에 서식이 여럿 — 공통 서식에 못 들어가는 그 경우
  s.addText([
    { text: '큰 굵은 글씨', options: { fontSize: 18, bold: true, color: BLUE, breakLine: true } },
    { text: '작은 보통 글씨', options: { fontSize: 10, color: '404040' } },
  ], {
    x: mm(20), y: mm(85), w: mm(170), h: mm(20),
    margin: 0, isTextBox: true, fit: 'none',
    line: { color: 'DDDDDD', width: 0.5 },
    objectName: '여러 서식',
  });

  (['top', 'middle', 'bottom'] as const).forEach((valign, i) => {
    s.addText(`${valign}`, {
      x: mm(20 + i * 60), y: mm(120), w: mm(50), h: mm(30),
      valign, fontSize: 11, margin: 0, isTextBox: true, fit: 'none',
      line: { color: 'DDDDDD', width: 0.5 },
      objectName: `세로 ${valign}`,
    });
  });

  // 도형 안의 글자 — Figma 에서 두 노드로 갈라지는 그 경우
  s.addText('도형 안의 글자', {
    x: mm(20), y: mm(165), w: mm(80), h: mm(25),
    shape: pptx.ShapeType.roundRect, fill: { color: 'E3F2FD' },
    line: { color: BLUE, width: 1 }, align: 'center', valign: 'middle',
    fontSize: 12, margin: 0, objectName: '도형+글자',
  });
}

/* ── 5장: 그룹 중첩 ──────────────────────────────────────────── */
{
  const s = pptx.addSlide();
  title(s, '5. 좌표 격자 — 20mm 마다 눈금 (어긋남을 눈으로도 재려고)');

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
  // 격자 교차점에 정확히 놓인 표식
  for (const [x, y] of [[20, 30], [100, 130], [180, 250], [60, 210]]) {
    s.addShape(pptx.ShapeType.ellipse, {
      x: mm(x - 2), y: mm(y - 2), w: mm(4), h: mm(4),
      fill: { color: 'E33348' }, line: { type: 'none' },
      objectName: `표식 ${x},${y}`,
    });
  }
}

const path = process.argv[2] ?? 'TestDATA.pptx';
const buf = await pptx.write({ outputType: 'nodebuffer' }) as Buffer;
await writeFile(path, buf);
console.log(`${path} — 슬라이드 5장, ${(buf.length / 1024).toFixed(0)} KB`);
console.log('좌표는 전부 mm 정수라, 왕복 후 0.1mm 라도 어긋나면 계산이 틀린 것이다.');
