import PptxGenJS from 'pptxgenjs';
// IR 의 Slide 와 PptxGenJS 의 Slide 가 이름이 겹친다. 이쪽은 "만들 내용"이라 SlideSpec 으로 받는다.
import type {
  Box, Doc, DocMeta, ImageItem, PathPoint, ShapeItem, Slide as SlideSpec, TextItem,
} from '../shared/ir';
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

/**
 * 슬라이드를 한 장씩 받아 쌓는 조립기.
 *
 * 문서를 통째로 받으면 58장 분량의 base64 이미지가 한꺼번에 메모리에 올라가고,
 * main → UI 로 넘길 때도 한 번에 넘겨야 한다. 장수가 늘면 그대로 터진다.
 * 그래서 머리말로 시작하고, 장을 받을 때마다 붙이고, 마지막에 파일로 굳힌다.
 */
export class PptxBuilder {
  private readonly pptx = new PptxGenJS();
  private readonly scale: Scale;
  /** Figma 이미지 주소 → 이미 만들어 둔 미디어 파일. 같은 그림은 그 파일을 함께 가리킨다. */
  private readonly byKey = new Map<string, string>();

  constructor(meta: DocMeta) {
    this.scale = new Scale(meta);
    definePresentation(this.pptx, meta, this.scale);
  }

  add(slide: SlideSpec): void {
    addSlideTo(this.pptx, slide, this.scale, this.byKey);
  }

  async finish(): Promise<Blob> {
    shareRepeatedMedia(this.pptx);
    return await this.pptx.write({ outputType: 'blob' }) as Blob;
  }

  /** 조립 중인 인스턴스. 출력 형식을 직접 정해야 하는 검증 도구용. */
  get presentation(): PptxGenJS {
    return this.pptx;
  }
}

/**
 * IR 을 PptxGenJS 인스턴스까지만 조립한다. 출력 형식은 호출부가 정한다(브라우저 blob / Node buffer).
 * 한 장씩 붙이는 경로를 그대로 쓴다 — 검증이 실제 내보내기와 다른 길을 타면 의미가 없다.
 */
export function composePptx(doc: Doc): PptxGenJS {
  const builder = new PptxBuilder(doc);
  for (const slide of doc.slides) builder.add(slide);
  shareRepeatedMedia(builder.presentation);
  return builder.presentation;
}

/** PptxGenJS 내부의 미디어 관계 항목 — 공개 타입이 아니라 필요한 것만 좁게 적는다. */
interface MediaRel {
  data?: string;
  Target?: string;
  type?: string;
}

/**
 * 같은 이미지는 파일 하나만 두고 여러 슬라이드가 함께 가리키게 한다.
 *
 * PPTX 는 원래 그렇게 쓴다 — 실측한 원본 덱은 이미지 참조 322개가 파일 235개를 가리키고,
 * 내용이 같은데 따로 저장된 것이 하나도 없다. 반면 PptxGenJS 는 참조마다 파일을 하나씩 만든다.
 * 중복 제거 장치가 있긴 한데 **파일 경로**를 열쇠로 쓰기 때문에 data URI 로 넣는 우리에게는
 * 걸리지 않고, 게다가 한 슬라이드 안에서만 비교한다.
 *
 * 파일을 쓰는 쪽은 `zip.file(rel.Target, …)` 이라, 같은 Target 을 주면 한 파일로 합쳐진다.
 * 그래서 Target 만 대표 하나로 맞춰 준다. (실측: 465MB → 이미지 197종 기준 약 90MB)
 */
function shareRepeatedMedia(pptx: PptxGenJS): void {
  const parts = pptx as unknown as {
    slides?: Array<{ _relsMedia?: MediaRel[] }>;
    slideLayouts?: Array<{ _relsMedia?: MediaRel[] }>;
    masterSlide?: { _relsMedia?: MediaRel[] };
  };

  /*
   * base64 문자열 전체를 Map 열쇠로 쓰면 몇 MB 짜리를 매번 해싱한다.
   * 길이와 앞뒤 조각으로 먼저 통을 나눈 뒤, 통 안에서만 전체를 비교한다 — 빠르면서 정확하다.
   */
  const buckets = new Map<string, Array<{ data: string; target: string }>>();

  const share = (rel: MediaRel): void => {
    if (!rel.data || !rel.Target) return;
    const sig = `${rel.type ?? ''}|${rel.data.length}|${rel.data.slice(0, 48)}|${rel.data.slice(-48)}`;
    const bucket = buckets.get(sig);
    if (!bucket) {
      buckets.set(sig, [{ data: rel.data, target: rel.Target }]);
      return;
    }
    for (const seen of bucket) {
      if (seen.data === rel.data) {
        rel.Target = seen.target;
        return;
      }
    }
    bucket.push({ data: rel.data, target: rel.Target });
  };

  for (const slide of parts.slides ?? []) for (const rel of slide._relsMedia ?? []) share(rel);
  for (const layout of parts.slideLayouts ?? []) for (const rel of layout._relsMedia ?? []) share(rel);
  for (const rel of parts.masterSlide?._relsMedia ?? []) share(rel);
}

/** 슬라이드 크기와 공통 서식 — 장을 붙이기 전에 끝나 있어야 한다. */
function definePresentation(pptx: PptxGenJS, doc: DocMeta, scale: Scale): void {
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

}

function addSlideTo(
  pptx: PptxGenJS, slide: SlideSpec, scale: Scale, byKey = new Map<string, string>(),
): void {
  const s = slide.master ? pptx.addSlide({ masterName: slide.master }) : pptx.addSlide();

  /*
   * 슬라이드 번호는 레이아웃에만 둔다.
   *
   * PptxGenJS 는 레이아웃의 slideNumber 를 각 장에도 물려 자리표시자를 하나 더 그리는데,
   * 그 도형 id 가 25 로 박혀 있다. 일반 도형은 `순번 + 2` 라 **24번째 도형이 정확히 25** 가 되어
   * 부딪히고, id 가 겹친 슬라이드를 PowerPoint 는 "읽을 수 없는 내용"으로 지우며 복구 창을 띄운다.
   * (실측: 232장 중 180장에서 충돌. 슬라이드 한 장에 도형이 58개였다.)
   *
   * 원본 덱도 번호를 레이아웃에만 두고 슬라이드에는 두지 않는다 — 그 편이 옳기도 하다.
   */
  s.slideNumber = null as unknown as PptxGenJS.SlideNumberProps;
  if (slide.fill.kind === 'solid') {
    s.background = { color: slide.fill.color, transparency: slide.fill.transparency };
  }

  for (const item of slide.items) {
    if (item.type === 'shape') addShape(pptx, s, item, scale);
    else if (item.type === 'text') addText(s, item, scale);
    else {
      addImage(s, item, scale);
      if (item.key) reuseByKey(s, item.key, byKey);
    }
  }
}

/**
 * 방금 만들어진 미디어 관계를, 같은 그림을 이미 담고 있는 파일로 돌려놓는다.
 *
 * Figma 의 imageHash 는 내용 주소라 같은 그림이면 반드시 같다. 몇 MB 짜리 base64 를
 * 맞대볼 필요 없이 이 40자만 보면 된다. 다시 렌더해서 뽑은 그림에는 이 주소가 없어,
 * 그런 것들은 마지막에 내용 비교로 걸러진다(shareRepeatedMedia).
 */
function reuseByKey(slide: Slide, key: string, byKey: Map<string, string>): void {
  const rels = (slide as unknown as { _relsMedia?: MediaRel[] })._relsMedia;
  const rel = rels?.[rels.length - 1];
  if (!rel?.Target) return;

  const first = byKey.get(key);
  if (first) rel.Target = first;
  else byKey.set(key, rel.Target);
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

  constructor(doc: DocMeta) {
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
