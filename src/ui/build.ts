import PptxGenJS from 'pptxgenjs';
import type { Box, Doc, ImageItem, PathPoint, ShapeItem, TextItem } from '../shared/ir';
import { ptToIn } from '../shared/units';

/**
 * IR → PPTX.
 *
 * 슬라이드 크기는 프레임에서만 나온다. 16:9 같은 고정 레이아웃을 끼워넣지 않는다.
 *
 * 프레임 비율이 PowerPoint 표준 비율과 일치할 때만 표준 크기를 쓸 수 있는데,
 * 그때도 적용되는 건 `doc.ptPerPx` 라는 **단일 균등 배율** 하나뿐이다.
 * 가로·세로에 다른 배율이 걸릴 여지가 구조적으로 없으므로 종횡비는 변하지 않는다.
 *
 * 이 파일의 Scale 이 px → inch 변환이 일어나는 유일한 지점이다.
 */
export async function buildPptx(doc: Doc): Promise<Blob> {
  const blob = await composePptx(doc).write({ outputType: 'blob' });
  return blob as Blob;
}

/** IR 을 PptxGenJS 인스턴스까지만 조립한다. 출력 형식은 호출부가 정한다(브라우저 blob / Node buffer). */
export function composePptx(doc: Doc): PptxGenJS {
  const pptx = new PptxGenJS();
  const scale = new Scale(doc);

  pptx.defineLayout({
    name: 'FIGMA_FRAME',
    width: ptToIn(doc.slideWPt),
    height: ptToIn(doc.slideHPt),
  });
  pptx.layout = 'FIGMA_FRAME';

  /*
   * 공통 서식을 먼저 정의한다 — 슬라이드가 masterName 으로 가리키려면 이미 있어야 한다.
   * 이것이 진짜 slideLayout 파트가 되어, 머리글을 PowerPoint 에서도 한 번만 고치면 된다.
   */
  for (const master of doc.masters) {
    pptx.defineSlideMaster({
      title: master.name,
      objects: doc.masters.length > 0
        ? master.items.map((item) => masterObject(pptx, item, scale))
        : undefined,
      slideNumber: master.slideNumber
        ? {
          ...position(master.slideNumber.box, scale),
          align: master.slideNumber.align,
          fontFace: master.slideNumber.fontFace,
          fontSize: scale.pt(master.slideNumber.fontSize),
          color: master.slideNumber.color,
        }
        : undefined,
    });
  }

  for (const slide of doc.slides) {
    const s = slide.master ? pptx.addSlide({ masterName: slide.master }) : pptx.addSlide();
    if (slide.fill.kind === 'solid') {
      s.background = { color: slide.fill.color, transparency: slide.fill.transparency };
    }

    for (const item of slide.items) {
      if (item.type === 'shape') addShape(pptx, s, item, scale);
      else if (item.type === 'text') addText(s, item, scale);
      else addImage(s, item, scale);
    }
  }

  return pptx;
}

/**
 * Figma px 를 PPTX 단위로 옮기는 변환기.
 *
 * - `x` / `y` : 위치. 배율 + 중앙 정렬 보정까지 적용해 inch 로.
 * - `len`     : 길이(폭·높이·반경·경로 좌표). 보정 없이 배율만, inch 로.
 * - `pt`      : 폰트 크기·선 굵기·자간처럼 PptxGenJS 가 pt 로 받는 값.
 *
 * 폰트 크기와 선 굵기도 반드시 같은 배율을 타야 한다. 좌표만 줄이면 글자만 커진다.
 */
class Scale {
  private readonly k: number;
  private readonly offX: number;
  private readonly offY: number;

  constructor(doc: Doc) {
    this.k = doc.ptPerPx;
    this.offX = doc.offsetXPt;
    this.offY = doc.offsetYPt;
  }

  x(px: number): number {
    return ptToIn(px * this.k + this.offX);
  }

  y(px: number): number {
    return ptToIn(px * this.k + this.offY);
  }

  len(px: number): number {
    return ptToIn(px * this.k);
  }

  pt(px: number): number {
    return px * this.k;
  }
}

type Slide = ReturnType<PptxGenJS['addSlide']>;

/**
 * PptxGenJS 4.0.1 런타임은 custGeom 을 지원하지만 (dist 의 `shape === 'custGeom'` 분기)
 * 타입 정의의 SHAPE_NAME 유니온에는 아직 빠져 있다. 그래서 여기서만 캐스팅한다.
 */
const CUST_GEOM = 'custGeom' as PptxGenJS.SHAPE_NAME;

interface Position {
  x: number;
  y: number;
  w: number;
  h: number;
  rotate?: number;
  flipH?: boolean;
  flipV?: boolean;
}

function position(box: Box, s: Scale): Position {
  const pos: Position = {
    x: s.x(box.x),
    y: s.y(box.y),
    w: s.len(box.w),
    h: s.len(box.h),
  };
  if (box.rot !== 0) pos.rotate = box.rot;
  if (box.flipH) pos.flipH = true;
  if (box.flipV) pos.flipV = true;
  return pos;
}

function addShape(pptx: PptxGenJS, slide: Slide, item: ShapeItem, s: Scale): void {
  const { shape, opts } = shapeSpec(pptx, item, s);
  slide.addShape(shape, opts);
}

/** 도형 하나의 PptxGenJS 인자. 슬라이드와 공통 서식이 같은 값을 쓴다. */
function shapeSpec(
  pptx: PptxGenJS, item: ShapeItem, s: Scale,
): { shape: PptxGenJS.SHAPE_NAME; opts: PptxGenJS.ShapeProps } {
  const opts: PptxGenJS.ShapeProps = {
    ...position(item.box, s),
    objectName: item.name,
  };

  /*
   * 칠이 없으면 아예 넘기지 않는다.
   * `{ type: 'none' }` 을 주면 PptxGenJS 가 그 분기에서 **아무것도 쓰지 않고**, 그러면
   * PowerPoint 가 기본 도형 서식(강조색)으로 채워 버린다. 비워야 `<a:noFill/>` 이 나간다.
   */
  if (item.fill.kind === 'solid') {
    opts.fill = { color: item.fill.color, transparency: item.fill.transparency };
  }

  opts.line = item.stroke
    ? {
      color: item.stroke.color,
      transparency: item.stroke.transparency,
      width: s.pt(item.stroke.width),
      dashType: item.stroke.dashType,
    }
    : { type: 'none' };

  if (item.shadow) opts.shadow = shadowOf(item.shadow, s);

  let shape: PptxGenJS.SHAPE_NAME;
  switch (item.geom.kind) {
    case 'roundRect':
      shape = pptx.ShapeType.roundRect;
      // rectRadius 는 inch. PptxGenJS 가 min(cx,cy) 대비 비율로 환산한다.
      opts.rectRadius = s.len(Math.min(item.geom.radius, Math.min(item.box.w, item.box.h) / 2));
      break;
    case 'ellipse':
      shape = pptx.ShapeType.ellipse;
      break;
    case 'line':
      shape = pptx.ShapeType.line;
      break;
    case 'custom':
      shape = CUST_GEOM;
      opts.points = toPoints(item.geom.points, s);
      break;
    default:
      shape = pptx.ShapeType.rect;
  }

  return { shape, opts };
}

function shadowOf(shadow: NonNullable<ShapeItem['shadow']>, s: Scale): PptxGenJS.ShadowProps {
  return {
    type: shadow.type,
    angle: shadow.angle,
    offset: s.pt(shadow.offset),
    blur: s.pt(shadow.blur),
    color: shadow.color,
    opacity: shadow.opacity,
  };
}

/**
 * custGeom 점은 도형 로컬 좌표다 (경로 공간 = 도형 크기).
 * 위치 보정(offset)이 아니라 길이 변환을 써야 한다.
 */
function toPoints(points: PathPoint[], s: Scale): NonNullable<PptxGenJS.ShapeProps['points']> {
  return points.map((p) => {
    if ('close' in p) return { close: true as const };
    if ('c' in p) {
      return {
        x: s.len(p.x),
        y: s.len(p.y),
        curve: {
          type: 'cubic' as const,
          x1: s.len(p.c[0]),
          y1: s.len(p.c[1]),
          x2: s.len(p.c[2]),
          y2: s.len(p.c[3]),
        },
      };
    }
    return p.moveTo
      ? { x: s.len(p.x), y: s.len(p.y), moveTo: true }
      : { x: s.len(p.x), y: s.len(p.y) };
  });
}

function addText(slide: Slide, item: TextItem, s: Scale): void {
  const { runs, opts } = textSpec(item, s);
  slide.addText(runs, opts);
}

/** 텍스트 하나의 PptxGenJS 인자. 슬라이드와 공통 서식이 같은 값을 쓴다. */
function textSpec(
  item: TextItem, s: Scale,
): { runs: PptxGenJS.TextProps[]; opts: PptxGenJS.TextPropsOptions } {
  const runs: PptxGenJS.TextProps[] = item.runs.map((r) => {
    const options: PptxGenJS.TextPropsOptions = {
      fontFace: r.fontFace,
      fontSize: s.pt(r.fontSize),
      bold: r.bold,
      italic: r.italic,
      color: r.color,
      align: item.align,
    };
    if (r.transparency > 0) options.transparency = r.transparency;
    if (r.underline) options.underline = { style: 'sng' };
    if (r.strike) options.strike = 'sngStrike';
    if (r.charSpacing !== undefined) options.charSpacing = s.pt(r.charSpacing);
    if (r.lineSpacing !== undefined) options.lineSpacing = s.pt(r.lineSpacing);
    if (r.breakLine) options.breakLine = true;
    if (r.hyperlink) options.hyperlink = { url: r.hyperlink };
    return { text: r.text, options };
  });

  const opts: PptxGenJS.TextPropsOptions = {
    ...position(item.box, s),
    objectName: item.name,
    align: item.align,
    valign: item.valign,
    wrap: item.wrap,
    // Figma 텍스트 박스에는 내부 여백이 없다. PowerPoint 기본 여백(0.1")을 지우지 않으면 전부 밀린다.
    margin: 0,
    // 자동 축소/확대는 Figma 에 없는 동작이라 끈다.
    fit: 'none',
    isTextBox: true,
    fill: { type: 'none' },
    line: { type: 'none' },
  };
  if (item.shadow) opts.shadow = shadowOf(item.shadow, s);

  return { runs, opts };
}

function addImage(slide: Slide, item: ImageItem, s: Scale): void {
  slide.addImage(imageSpec(item, s));
}

function imageSpec(item: ImageItem, s: Scale): PptxGenJS.ImageProps {
  const pos = position(item.box, s);
  const opts: PptxGenJS.ImageProps = {
    ...pos,
    objectName: item.name,
    data: `data:${item.mime};base64,${item.data}`,
  };
  if (item.sizing !== 'stretch') {
    opts.sizing = { type: item.sizing, w: pos.w, h: pos.h };
  }
  return opts;
}

/**
 * 항목 하나를 공통 서식(slideLayout)에 넣을 형태로.
 *
 * PptxGenJS 의 마스터는 `text`·`rect`·`line`·`image` 키만 받지만, `text` 키는 내부적으로
 * `addTextDefinition` 을 타고 그것이 `options.shape` 를 존중한다. 그래서 도형도 `text` 키에
 * 실어 보내면 custGeom 까지 그대로 나간다 (`rect` 키는 사각형으로 고정이라 쓸 수 없다).
 *
 * 다만 글자는 문자열 하나만 실린다 — 서식이 여러 개인 텍스트는 애초에 여기까지 오지 않는다
 * (shared/ir.ts 의 fitsInMaster 가 걸러 슬라이드 쪽에 남긴다).
 */
function masterObject(
  pptx: PptxGenJS, item: ShapeItem | TextItem | ImageItem, s: Scale,
): NonNullable<PptxGenJS.SlideMasterProps['objects']>[number] {
  if (item.type === 'image') {
    return { image: imageSpec(item, s) } as never;
  }
  if (item.type === 'shape') {
    const { shape, opts } = shapeSpec(pptx, item, s);
    return { text: { text: '', options: { ...opts, shape } as PptxGenJS.TextPropsOptions } } as never;
  }
  const { runs, opts } = textSpec(item, s);
  const run = runs[0];
  return {
    text: { text: run?.text ?? '', options: { ...opts, ...run?.options } },
  } as never;
}
