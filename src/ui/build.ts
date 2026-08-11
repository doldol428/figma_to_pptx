import PptxGenJS from 'pptxgenjs';
import type { Doc, ImageItem, PathPoint, ShapeItem, TextItem } from '../shared/ir';
import { pxToIn } from '../shared/units';

/**
 * IR → PPTX.
 *
 * 핵심 규칙 하나: 슬라이드 크기는 프레임 크기에서만 나온다.
 * 16:9 같은 고정 레이아웃을 절대 끼워넣지 않으며, 어떤 좌표도 스케일링하지 않는다.
 * pxToIn() 이 유일한 단위 변환 지점이고 그 안에서 1px = 1pt 가 고정된다.
 */
export async function buildPptx(doc: Doc): Promise<Blob> {
  const blob = await composePptx(doc).write({ outputType: 'blob' });
  return blob as Blob;
}

/** IR 을 PptxGenJS 인스턴스까지만 조립한다. 출력 형식은 호출부가 정한다(브라우저 blob / Node buffer). */
export function composePptx(doc: Doc): PptxGenJS {
  const pptx = new PptxGenJS();

  pptx.defineLayout({
    name: 'FIGMA_FRAME',
    width: pxToIn(doc.widthPx),
    height: pxToIn(doc.heightPx),
  });
  pptx.layout = 'FIGMA_FRAME';

  for (const slide of doc.slides) {
    const s = pptx.addSlide();
    if (slide.fill.kind === 'solid') {
      s.background = { color: slide.fill.color, transparency: slide.fill.transparency };
    }

    for (const item of slide.items) {
      if (item.type === 'shape') addShape(pptx, s, item);
      else if (item.type === 'text') addText(s, item);
      else addImage(s, item);
    }
  }

  return pptx;
}

type Slide = ReturnType<PptxGenJS['addSlide']>;

/**
 * PptxGenJS 4.0.1 런타임은 custGeom 을 지원하지만 (dist 의 `shape === 'custGeom'` 분기)
 * 타입 정의의 SHAPE_NAME 유니온에는 아직 빠져 있다. 그래서 여기서만 캐스팅한다.
 */
const CUST_GEOM = 'custGeom' as PptxGenJS.SHAPE_NAME;

function position(box: ShapeItem['box']): {
  x: number; y: number; w: number; h: number; rotate?: number; flipH?: boolean; flipV?: boolean;
} {
  const pos = {
    x: pxToIn(box.x),
    y: pxToIn(box.y),
    w: pxToIn(box.w),
    h: pxToIn(box.h),
  } as ReturnType<typeof position>;
  if (box.rot !== 0) pos.rotate = box.rot;
  if (box.flipH) pos.flipH = true;
  if (box.flipV) pos.flipV = true;
  return pos;
}

function addShape(pptx: PptxGenJS, slide: Slide, item: ShapeItem): void {
  const opts: PptxGenJS.ShapeProps = {
    ...position(item.box),
    objectName: item.name,
  };

  opts.fill = item.fill.kind === 'solid'
    ? { color: item.fill.color, transparency: item.fill.transparency }
    : { type: 'none' };

  opts.line = item.stroke
    ? {
      color: item.stroke.color,
      transparency: item.stroke.transparency,
      width: item.stroke.width,
      dashType: item.stroke.dashType,
    }
    : { type: 'none' };

  if (item.shadow) {
    opts.shadow = {
      type: item.shadow.type,
      angle: item.shadow.angle,
      offset: item.shadow.offset,
      blur: item.shadow.blur,
      color: item.shadow.color,
      opacity: item.shadow.opacity,
    };
  }

  let shape: PptxGenJS.SHAPE_NAME;
  switch (item.geom.kind) {
    case 'roundRect':
      shape = pptx.ShapeType.roundRect;
      // rectRadius 는 inch. PptxGenJS 가 min(cx,cy) 대비 비율로 환산한다.
      opts.rectRadius = pxToIn(Math.min(item.geom.radius, Math.min(item.box.w, item.box.h) / 2));
      break;
    case 'ellipse':
      shape = pptx.ShapeType.ellipse;
      break;
    case 'line':
      shape = pptx.ShapeType.line;
      break;
    case 'custom':
      shape = CUST_GEOM;
      opts.points = toPoints(item.geom.points);
      break;
    default:
      shape = pptx.ShapeType.rect;
  }

  slide.addShape(shape, opts);
}

/** custGeom 점은 도형 로컬 좌표를 inch 로 표현한 값이다 (경로 공간 = 도형 크기). */
function toPoints(points: PathPoint[]): NonNullable<PptxGenJS.ShapeProps['points']> {
  return points.map((p) => {
    if ('close' in p) return { close: true as const };
    if ('c' in p) {
      return {
        x: pxToIn(p.x),
        y: pxToIn(p.y),
        curve: {
          type: 'cubic' as const,
          x1: pxToIn(p.c[0]),
          y1: pxToIn(p.c[1]),
          x2: pxToIn(p.c[2]),
          y2: pxToIn(p.c[3]),
        },
      };
    }
    return p.moveTo
      ? { x: pxToIn(p.x), y: pxToIn(p.y), moveTo: true }
      : { x: pxToIn(p.x), y: pxToIn(p.y) };
  });
}

function addText(slide: Slide, item: TextItem): void {
  const runs: PptxGenJS.TextProps[] = item.runs.map((r) => {
    const options: PptxGenJS.TextPropsOptions = {
      fontFace: r.fontFace,
      fontSize: r.fontSize,
      bold: r.bold,
      italic: r.italic,
      color: r.color,
      align: item.align,
    };
    if (r.transparency > 0) options.transparency = r.transparency;
    if (r.underline) options.underline = { style: 'sng' };
    if (r.strike) options.strike = 'sngStrike';
    if (r.charSpacing !== undefined) options.charSpacing = r.charSpacing;
    if (r.lineSpacing !== undefined) options.lineSpacing = r.lineSpacing;
    if (r.breakLine) options.breakLine = true;
    if (r.hyperlink) options.hyperlink = { url: r.hyperlink };
    return { text: r.text, options };
  });

  const opts: PptxGenJS.TextPropsOptions = {
    ...position(item.box),
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
  if (item.shadow) {
    opts.shadow = {
      type: item.shadow.type,
      angle: item.shadow.angle,
      offset: item.shadow.offset,
      blur: item.shadow.blur,
      color: item.shadow.color,
      opacity: item.shadow.opacity,
    };
  }

  slide.addText(runs, opts);
}

function addImage(slide: Slide, item: ImageItem): void {
  const opts: PptxGenJS.ImageProps = {
    ...position(item.box),
    objectName: item.name,
    data: `data:${item.mime};base64,${item.data}`,
  };
  if (item.sizing !== 'stretch') {
    opts.sizing = {
      type: item.sizing,
      w: pxToIn(item.box.w),
      h: pxToIn(item.box.h),
    };
  }
  slide.addImage(opts);
}
