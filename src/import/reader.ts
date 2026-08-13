import JSZip from 'jszip';
import { t } from '../shared/i18n';
import type { Warning } from '../shared/ir';
import type {
  ImportDoc, ImportNode, ImportSlide, Paint, Paragraph, Placement, ShapeGeometry,
  ShadowSpec, ShapeNode, StrokeSpec, TableCell, TextRun,
} from '../shared/importir';
import {
  colorNode, inheritsGroupFill, readFill, readLinePaint, resolveColorNode, sliceFillForChild,
  type ColorContext,
} from './color';
import { connectorPath, nativeFor, presetPath, readAdjust } from './preset';
import {
  IDENTITY, decompose, multiply, placeBox, scale, translate, type Mat,
} from './transform';
import { all, bool, deep, num, one, parseXml, path as xpath, type XNode } from './xml';

const EMU_PT = 12700;
const pt = (emu: number): number => emu / EMU_PT;

/** 텍스트 상자 기본 여백 (EMU) — bodyPr 에 값이 없을 때 PowerPoint 가 쓰는 값 */
const DEFAULT_INSETS = { l: 91440, t: 45720, r: 91440, b: 45720 };

interface Ctx {
  zip: JSZip;
  color: ColorContext;
  /** 테마 폰트 — `+mj-lt` / `+mn-lt` 참조를 푸는 데 쓴다 */
  themeFonts: { major: string; minor: string };
  warn: (slide: string, node: string, message: string) => void;
  fonts: Set<string>;
  slideName: string;
  /** 마스터의 txStyles — 자리표시자 텍스트의 기본 서식 */
  masterStyles: { title: XNode | null; body: XNode | null; other: XNode | null };
  /** presentation.xml 의 defaultTextStyle — 일반 텍스트 상자의 기본 서식 */
  defaultTextStyle: XNode | null;
}

/**
 * 텍스트 서식 상속 사슬.
 *
 * PPTX 의 글자 크기·색·폰트는 대부분 런에 직접 적혀 있지 않다.
 * 도형의 `<a:lstStyle>`, 레이아웃 자리표시자, 마스터의 `txStyles`, 프레젠테이션 기본값이
 * 차례로 얹힌다. 런만 읽으면 값이 없을 때 임의의 기본값으로 떨어져 크기가 통째로 틀어진다.
 *
 * 우선순위가 높은 것부터 담는다.
 */
type StyleChain = Array<XNode | null>;

/** `<a:lvl{n}pPr><a:defRPr>` — 문단 수준(level)에 해당하는 기본 런 속성 */
function levelDefaults(lstStyle: XNode | null, level: number): XNode | null {
  if (!lstStyle) return null;
  const lvl = one(lstStyle, `lvl${Math.min(9, level + 1)}pPr`);
  return one(lvl, 'defRPr');
}

/** 속성이 실제로 적힌 첫 노드를 찾는다 */
function pickAttr(nodes: StyleChain, attr: string): string | undefined {
  for (const n of nodes) {
    const v = n?.attrs[attr];
    if (v !== undefined) return v;
  }
  return undefined;
}

/** 자식 요소가 있는 첫 노드를 찾는다 */
function pickChild(nodes: StyleChain, tag: string): XNode | null {
  for (const n of nodes) {
    const c = one(n, tag);
    if (c) return c;
  }
  return null;
}

export interface ReadOptions {
  onProgress?: (done: number, total: number) => void;
  /** 가져올 슬라이드 번호(1부터). 비우면 전부. 한 장만 확인할 때 파싱까지 건너뛴다. */
  only?: Set<number>;
}

/** "1", "1-5", "1,3,7-9" → {1,3,7,8,9}. 빈 문자열이면 null(전체). */
export function parseSlideRange(input: string): Set<number> | null {
  const text = input.trim();
  if (!text) return null;
  const out = new Set<number>();
  for (const part of text.split(',')) {
    const m = /^\s*(\d+)\s*(?:-\s*(\d+))?\s*$/.exec(part);
    if (!m) continue;
    const from = Number(m[1]);
    const to = m[2] ? Number(m[2]) : from;
    for (let i = Math.min(from, to); i <= Math.max(from, to); i++) out.add(i);
  }
  return out.size > 0 ? out : null;
}

export async function readPptx(
  data: ArrayBuffer,
  fileName: string,
  opts: ReadOptions = {},
): Promise<ImportDoc> {
  const zip = await JSZip.loadAsync(data);
  const warnings: Warning[] = [];
  const seen = new Set<string>();
  const warn = (slide: string, node: string, message: string): void => {
    const key = `${slide}|${node}|${message}`;
    if (seen.has(key)) return;
    seen.add(key);
    warnings.push({ slide, node, message });
  };

  const presentation = await readXml(zip, 'ppt/presentation.xml');
  const sldSz = deep(presentation, 'sldSz');
  const widthPt = pt(num(sldSz?.attrs.cx, 9144000));
  const heightPt = pt(num(sldSz?.attrs.cy, 6858000));

  const slidePaths = await orderedSlidePaths(zip, presentation);

  // 테마와 색 매핑은 마스터를 거쳐야 한다. 슬라이드 XML 만 읽으면 색이 전부 틀어진다.
  const master = await readXml(zip, 'ppt/slideMasters/slideMaster1.xml');
  const clrMapEl = one(master, 'clrMap');
  const map: Record<string, string> = clrMapEl ? { ...clrMapEl.attrs } : {};

  const theme = await readXml(zip, 'ppt/theme/theme1.xml');
  const scheme: Record<string, string> = {};
  for (const c of deep(theme, 'clrScheme')?.children ?? []) {
    const v = one(c, 'srgbClr')?.attrs.val ?? one(c, 'sysClr')?.attrs.lastClr;
    if (v) scheme[c.tag] = v;
  }

  const fontScheme = deep(theme, 'fontScheme');
  const themeFonts = {
    major: xpath(fontScheme, 'majorFont', 'latin')?.attrs.typeface || 'Arial',
    minor: xpath(fontScheme, 'minorFont', 'latin')?.attrs.typeface || 'Arial',
  };

  const txStyles = one(master, 'txStyles');
  const fonts = new Set<string>();
  const ctx: Ctx = {
    zip,
    color: { scheme, map },
    themeFonts,
    warn,
    fonts,
    slideName: '',
    masterStyles: {
      title: one(txStyles, 'titleStyle'),
      body: one(txStyles, 'bodyStyle'),
      other: one(txStyles, 'otherStyle'),
    },
    defaultTextStyle: one(deep(presentation, 'presentation') ?? presentation, 'defaultTextStyle'),
  };

  // 지정된 번호만 남긴다. 원래 번호는 이름에 유지해서 어느 장인지 알 수 있게 한다.
  const wanted = slidePaths
    .map((path, i) => ({ path, number: i + 1 }))
    .filter((s) => !opts.only || opts.only.has(s.number));

  const slides: ImportSlide[] = [];
  for (let i = 0; i < wanted.length; i++) {
    opts.onProgress?.(i, wanted.length);
    const { path: slidePath, number } = wanted[i];
    ctx.slideName = `슬라이드 ${number}`;

    const slideXml = await readXml(zip, slidePath);
    const rels = await readRels(zip, slidePath);
    const layout = await layoutFor(zip, rels);

    const tree = deep(slideXml, 'spTree');
    const nodes: ImportNode[] = [];
    if (tree) {
      for (const child of tree.children) {
        const node = await readNode(child, IDENTITY, ctx, rels, layout, null);
        if (node) nodes.push(node);
      }
    }

    slides.push({
      name: `${number}. ${slideTitle(tree) || ctx.slideName}`,
      fill: readSlideBackground(slideXml, ctx) ?? undefined,
      nodes,
    });
  }
  opts.onProgress?.(wanted.length, wanted.length);

  return {
    widthPt,
    heightPt,
    fileName,
    slides,
    warnings,
    fonts: Array.from(fonts).sort(),
  };
}

/* ── 패키지 탐색 ──────────────────────────────────────────────── */

async function readXml(zip: JSZip, p: string): Promise<XNode | null> {
  const file = zip.file(p);
  if (!file) return null;
  return parseXml(await file.async('string'));
}

async function readRels(zip: JSZip, partPath: string): Promise<Record<string, string>> {
  const slash = partPath.lastIndexOf('/');
  const relPath = `${partPath.slice(0, slash)}/_rels/${partPath.slice(slash + 1)}.rels`;
  const xml = await readXml(zip, relPath);
  const out: Record<string, string> = {};
  for (const r of deep(xml, 'Relationships')?.children ?? []) {
    if (r.attrs.Id && r.attrs.Target) out[r.attrs.Id] = r.attrs.Target;
  }
  return out;
}

/** presentation.xml 의 sldIdLst 순서를 rels 로 실제 경로에 대응시킨다 */
async function orderedSlidePaths(zip: JSZip, presentation: XNode | null): Promise<string[]> {
  const rels = await readRels(zip, 'ppt/presentation.xml');
  const ids = all(deep(presentation, 'sldIdLst'), 'sldId');
  const paths: string[] = [];
  for (const id of ids) {
    const target = rels[id.attrs['r:id'] ?? ''];
    if (target) paths.push(`ppt/${target.replace(/^\.\.\//, '')}`);
  }
  if (paths.length > 0) return paths;

  // sldIdLst 를 못 읽으면 파일명 순서로 떨어진다.
  return Object.keys(zip.files)
    .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => num(/\d+/.exec(a)?.[0]) - num(/\d+/.exec(b)?.[0]));
}

async function layoutFor(
  zip: JSZip, rels: Record<string, string>,
): Promise<XNode | null> {
  const target = Object.values(rels).find((t) => t.indexOf('slideLayout') >= 0);
  if (!target) return null;
  return readXml(zip, `ppt/${target.replace(/^\.\.\//, '')}`);
}

function slideTitle(tree: XNode | null): string {
  for (const sp of all(tree, 'sp')) {
    const ph = deep(xpath(sp, 'nvSpPr', 'nvPr'), 'ph');
    if (ph?.attrs.type !== 'title' && ph?.attrs.type !== 'ctrTitle') continue;
    const text = collectText(one(sp, 'txBody'));
    if (text) return text.slice(0, 40);
  }
  return '';
}

function collectText(txBody: XNode | null): string {
  if (!txBody) return '';
  let out = '';
  for (const p of all(txBody, 'p')) {
    for (const r of all(p, 'r')) out += one(r, 't')?.text ?? '';
  }
  return out.trim();
}

function readSlideBackground(slide: XNode | null, ctx: Ctx): Paint | null {
  const bgPr = xpath(deep(slide, 'bg'), 'bgPr');
  if (!bgPr) return null;
  const fill = readFill(bgPr, ctx.color);
  return fill ?? null;
}

/* ── 노드 ────────────────────────────────────────────────────── */

/** 부모 그룹이 물려주는 채우기와 그 그룹의 범위 (grpFill 해석용) */
interface Inherited {
  paint: Paint;
  box: { x: number; y: number; w: number; h: number };
}

async function readNode(
  el: XNode,
  parent: Mat,
  ctx: Ctx,
  rels: Record<string, string>,
  layout: XNode | null,
  inherited: Inherited | null,
): Promise<ImportNode | null> {
  switch (el.tag) {
    case 'sp':
      return readShape(el, parent, ctx, layout, inherited);
    case 'cxnSp':
      return readConnector(el, parent, ctx);
    case 'pic':
      return readPicture(el, parent, ctx, rels);
    case 'grpSp':
      return readGroup(el, parent, ctx, rels, layout, inherited);
    case 'graphicFrame':
      return readGraphicFrame(el, parent, ctx);
    default:
      return null;
  }
}

/** spPr 의 xfrm 을 읽어 누적 행렬과 로컬 크기를 만든다 */
function localMatrix(xfrm: XNode | null): { m: Mat; w: number; h: number } | null {
  if (!xfrm) return null;
  const off = one(xfrm, 'off');
  const ext = one(xfrm, 'ext');
  if (!off || !ext) return null;
  const w = num(ext.attrs.cx);
  const h = num(ext.attrs.cy);
  return {
    m: placeBox(
      num(off.attrs.x), num(off.attrs.y), w, h,
      num(xfrm.attrs.rot) / 60000,
      bool(xfrm.attrs.flipH),
      bool(xfrm.attrs.flipV),
    ),
    w,
    h,
  };
}

function placementOf(parent: Mat, local: { m: Mat; w: number; h: number }): Placement {
  const box = decompose(multiply(parent, local.m), local.w, local.h);
  return { ...box, x: pt(box.x), y: pt(box.y), w: pt(box.w), h: pt(box.h) };
}

/**
 * 도형의 xfrm 이 없으면 레이아웃의 같은 placeholder 에서 가져온다.
 * 제목·본문 자리표시자는 슬라이드에 위치를 안 적는 경우가 많다.
 */
function layoutPlaceholder(sp: XNode, layout: XNode | null): XNode | null {
  const ph = deep(xpath(sp, 'nvSpPr', 'nvPr'), 'ph');
  if (!ph || !layout) return null;
  const idx = ph.attrs.idx;
  const type = ph.attrs.type;
  for (const cand of all(deep(layout, 'spTree'), 'sp')) {
    const cph = deep(xpath(cand, 'nvSpPr', 'nvPr'), 'ph');
    if (!cph) continue;
    if ((idx !== undefined && cph.attrs.idx === idx)
      || (type !== undefined && cph.attrs.type === type)) {
      return cand;
    }
  }
  return null;
}

function inheritedXfrm(sp: XNode, layout: XNode | null): XNode | null {
  return xpath(layoutPlaceholder(sp, layout), 'spPr', 'xfrm');
}

/** 이 도형의 텍스트에 적용되는 서식 상속 사슬 (우선순위 높은 것부터) */
function styleChainFor(sp: XNode, layout: XNode | null, ctx: Ctx): StyleChain {
  const ph = deep(xpath(sp, 'nvSpPr', 'nvPr'), 'ph');
  const type = ph?.attrs.type;
  const masterStyle = !ph
    ? ctx.defaultTextStyle
    : type === 'title' || type === 'ctrTitle'
      ? ctx.masterStyles.title
      : type === 'body' || type === 'subTitle' || type === undefined
        ? ctx.masterStyles.body
        : ctx.masterStyles.other;

  return [
    xpath(sp, 'txBody', 'lstStyle'),
    xpath(layoutPlaceholder(sp, layout), 'txBody', 'lstStyle'),
    masterStyle,
    ctx.defaultTextStyle,
  ];
}

function readStroke(ln: XNode | null, ctx: Ctx): StrokeSpec | undefined {
  if (!ln) return undefined;
  const paint = readLinePaint(ln, ctx.color);
  if (!paint) return undefined;
  const width = ln.attrs.w ? pt(num(ln.attrs.w)) : 1;
  const dashVal = one(ln, 'prstDash')?.attrs.val;
  const stroke: StrokeSpec = { paint, width };
  if (dashVal && dashVal !== 'solid') {
    stroke.dash = dashVal.indexOf('dot') >= 0 ? [width, width * 2] : [width * 4, width * 3];
  }
  const cap = ln.attrs.cap;
  if (cap === 'rnd') stroke.cap = 'ROUND';
  else if (cap === 'sq') stroke.cap = 'SQUARE';
  return stroke;
}

function readShadow(spPr: XNode | null, ctx: Ctx): ShadowSpec | undefined {
  const outer = deep(one(spPr, 'effectLst'), 'outerShdw');
  if (!outer) return undefined;
  const c = resolveColorNode(colorNode(outer), ctx.color);
  const dist = pt(num(outer.attrs.dist));
  const dir = num(outer.attrs.dir) / 60000;
  const rad = (dir * Math.PI) / 180;
  return {
    offsetX: Math.round(dist * Math.cos(rad) * 100) / 100,
    offsetY: Math.round(dist * Math.sin(rad) * 100) / 100,
    blur: pt(num(outer.attrs.blurRad)),
    color: c?.color ?? '000000',
    opacity: c?.opacity ?? 0.3,
  };
}

async function readShape(
  sp: XNode, parent: Mat, ctx: Ctx, layout: XNode | null, inherited: Inherited | null,
  /** 주어지면 grpFill 도형의 기하를 여기 모으고 개별 도형은 만들지 않는다 (텍스트는 그대로) */
  collect?: (part: MergePart) => void,
): Promise<ImportNode | null> {
  const spPr = one(sp, 'spPr');
  const local = localMatrix(one(spPr, 'xfrm') ?? inheritedXfrm(sp, layout));
  if (!local) return null;

  const name = xpath(sp, 'nvSpPr', 'cNvPr')?.attrs.name ?? '도형';
  const place = placementOf(parent, local);
  const opacity = 1;

  const prstGeom = one(spPr, 'prstGeom');
  const custGeom = one(spPr, 'custGeom');
  const prst = prstGeom?.attrs.prst ?? '';
  const adj = readAdjust(prstGeom);

  // grpFill 이면 부모 그룹의 채우기를 이 도형이 차지하는 구간만큼 잘라 쓴다.
  const own = readFill(spPr, ctx.color);
  const fill = own ?? (inheritsGroupFill(spPr) && inherited
    ? sliceFillForChild(inherited.paint, inherited.box, place)
    : undefined);

  const stroke = readStroke(one(spPr, 'ln'), ctx);
  const shadow = readShadow(spPr, ctx);

  const geometry = resolveGeometry(prst, custGeom, place, adj, name, ctx);
  const children: ImportNode[] = [];

  const usesGroupFill = !own && inheritsGroupFill(spPr);
  const mergePath = collect && geometry && usesGroupFill && !stroke
    && place.rotation === 0 && !place.flipH
    ? geometryToPath(geometry, place.w, place.h)
    : null;

  if (mergePath) {
    // 기하는 형제들과 합쳐서 한 도형으로 낸다. 여기서는 만들지 않는다.
    collect?.({ path: mergePath, place });
  } else {
    const hasPaint = !!fill || !!stroke;
    if (geometry && hasPaint) {
      const shape: ImportNode = {
        type: 'shape', name, place, opacity, geometry,
        ...(fill ? { fill } : {}),
        ...(stroke ? { stroke } : {}),
        ...(shadow ? { shadow } : {}),
      };
      children.push(shape);
    }
  }

  const text = readTextBody(one(sp, 'txBody'), place, name, ctx, styleChainFor(sp, layout, ctx));
  if (text) children.push(text);

  if (children.length === 0) return null;
  if (children.length === 1) return children[0];
  // 면과 글자를 함께 가진 도형 — Figma 에는 도형 안의 텍스트가 없으므로 묶어서 낸다.
  return { type: 'group', name, place, opacity, children };
}

function resolveGeometry(
  prst: string,
  custGeom: XNode | null,
  place: Placement,
  adj: Record<string, number>,
  name: string,
  ctx: Ctx,
): ShapeGeometry | null {
  if (custGeom) {
    const data = custGeomPath(custGeom, place.w, place.h);
    if (data) return { kind: 'path', data, evenOdd: false };
    return { kind: 'rect' };
  }
  if (!prst) return { kind: 'rect' };

  const native = nativeFor(prst, place.w, place.h, adj);
  if (native) return native;

  const preset = presetPath(prst, place.w, place.h, adj);
  if (preset) return { kind: 'path', data: preset.data, evenOdd: preset.evenOdd };

  const conn = connectorPath(prst, place.w, place.h);
  if (conn) return { kind: 'path', data: conn.data, evenOdd: false };

  ctx.warn(ctx.slideName, name, `preset 도형 ${prst} 은 아직 지원하지 않아 사각형으로 대체했습니다.`);
  return { kind: 'rect' };
}

function custGeomPath(custGeom: XNode, w: number, h: number): string {
  const pathLst = one(custGeom, 'pathLst');
  let d = '';
  for (const p of all(pathLst, 'path')) {
    const pw = num(p.attrs.w);
    const ph = num(p.attrs.h);
    const sx = pw ? w / pw : 1 / EMU_PT;
    const sy = ph ? h / ph : 1 / EMU_PT;
    const X = (v: string | undefined): string => (num(v) * sx).toFixed(2);
    const Y = (v: string | undefined): string => (num(v) * sy).toFixed(2);
    for (const seg of p.children) {
      const ps = all(seg, 'pt');
      switch (seg.tag) {
        case 'moveTo':
          d += `M${X(ps[0]?.attrs.x)} ${Y(ps[0]?.attrs.y)} `;
          break;
        case 'lnTo':
          d += `L${X(ps[0]?.attrs.x)} ${Y(ps[0]?.attrs.y)} `;
          break;
        case 'cubicBezTo':
          d += `C${X(ps[0]?.attrs.x)} ${Y(ps[0]?.attrs.y)} ${X(ps[1]?.attrs.x)} ${Y(ps[1]?.attrs.y)} ${X(ps[2]?.attrs.x)} ${Y(ps[2]?.attrs.y)} `;
          break;
        case 'quadBezTo':
          d += `Q${X(ps[0]?.attrs.x)} ${Y(ps[0]?.attrs.y)} ${X(ps[1]?.attrs.x)} ${Y(ps[1]?.attrs.y)} `;
          break;
        case 'close':
          d += 'Z ';
          break;
        default:
          break;
      }
    }
  }
  return d.trim();
}

async function readConnector(el: XNode, parent: Mat, ctx: Ctx): Promise<ImportNode | null> {
  const spPr = one(el, 'spPr');
  const local = localMatrix(one(spPr, 'xfrm'));
  if (!local) return null;

  const name = xpath(el, 'nvCxnSpPr', 'cNvPr')?.attrs.name ?? '연결선';
  const place = placementOf(parent, local);
  const prst = one(spPr, 'prstGeom')?.attrs.prst ?? 'line';

  const stroke = readStroke(one(spPr, 'ln'), ctx);
  if (!stroke) return null;

  const conn = connectorPath(prst, place.w, place.h);
  const geometry: ShapeGeometry = conn
    ? { kind: 'path', data: conn.data, evenOdd: false }
    : { kind: 'line' };

  return { type: 'shape', name, place, opacity: 1, geometry, stroke };
}

async function readGroup(
  el: XNode, parent: Mat, ctx: Ctx, rels: Record<string, string>, layout: XNode | null,
  inherited: Inherited | null,
): Promise<ImportNode | null> {
  const grpSpPr = one(el, 'grpSpPr');
  const xfrm = one(grpSpPr, 'xfrm');
  if (!xfrm) return null;

  const off = one(xfrm, 'off');
  const ext = one(xfrm, 'ext');
  const chOff = one(xfrm, 'chOff');
  const chExt = one(xfrm, 'chExt');
  if (!off || !ext) return null;

  const w = num(ext.attrs.cx);
  const h = num(ext.attrs.cy);

  // 그룹의 배치(회전 포함)를 먼저 세우고, 그 안에서 자식 좌표계를 매핑한다.
  const placed = placeBox(
    num(off.attrs.x), num(off.attrs.y), w, h,
    num(xfrm.attrs.rot) / 60000,
    bool(xfrm.attrs.flipH),
    bool(xfrm.attrs.flipV),
  );

  let inner: Mat = IDENTITY;
  if (chOff && chExt) {
    const cw = num(chExt.attrs.cx);
    const ch = num(chExt.attrs.cy);
    inner = multiply(
      scale(cw ? w / cw : 1, ch ? h / ch : 1),
      translate(-num(chOff.attrs.x), -num(chOff.attrs.y)),
    );
  }

  const childMat = multiply(multiply(parent, placed), inner);

  const name = xpath(el, 'nvGrpSpPr', 'cNvPr')?.attrs.name ?? '그룹';
  const place = placementOf(parent, { m: placed, w, h });

  // 이 그룹이 채우기를 들고 있으면 grpFill 을 쓰는 자식들이 그것을 나눠 쓴다.
  // 그룹 자신에게 채우기가 없으면 위에서 물려받은 것을 그대로 넘긴다.
  const ownFill = readFill(grpSpPr, ctx.color);
  const passDown: Inherited | null = ownFill
    ? { paint: ownFill, box: place }
    : inheritsGroupFill(grpSpPr) ? inherited : inherited;

  // 그룹 채우기를 쓰는 자식들은 경로를 모아 하나로 합친다 (mergeParts 주석 참고).
  const parts: MergePart[] = [];
  const collect = ownFill ? (part: MergePart): void => { parts.push(part); } : undefined;

  const children: ImportNode[] = [];
  for (const child of el.children) {
    const node = child.tag === 'sp'
      ? await readShape(child, childMat, ctx, layout, passDown, collect)
      : await readNode(child, childMat, ctx, rels, layout, passDown);
    if (node) children.push(node);
  }

  if (ownFill && parts.length > 0) {
    const merged = mergeParts(parts, ownFill, `${name} 배경`);
    // 합친 배경은 맨 아래에 둔다 — 나머지 요소가 그 위에 올라앉는 구조다.
    if (merged) children.unshift(merged);
  }

  if (children.length === 0) return null;
  return { type: 'group', name, place, opacity: 1, children };
}

const MIME: Record<string, string> = {
  png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif',
  svg: 'image/svg+xml', bmp: 'image/bmp', tiff: 'image/tiff',
};

async function readPicture(
  el: XNode, parent: Mat, ctx: Ctx, rels: Record<string, string>,
): Promise<ImportNode | null> {
  const spPr = one(el, 'spPr');
  const local = localMatrix(one(spPr, 'xfrm'));
  if (!local) return null;

  const name = xpath(el, 'nvPicPr', 'cNvPr')?.attrs.name ?? '그림';
  const place = placementOf(parent, local);

  const blip = deep(one(el, 'blipFill'), 'blip');
  const id = blip?.attrs['r:embed'] ?? blip?.attrs['r:link'];
  const target = id ? rels[id] : undefined;
  if (!target) return null;

  const filePath = `ppt/${target.replace(/^\.\.\//, '')}`;
  const file = ctx.zip.file(filePath);
  if (!file) return null;

  const ext = (filePath.split('.').pop() ?? '').toLowerCase();
  const mime = MIME[ext];
  if (!mime) {
    // wdp(JPEG XR) · emf/wmf(메타파일) 는 Figma 가 읽지 못한다.
    ctx.warn(ctx.slideName, name, `.${ext} 이미지는 Figma 가 읽지 못해 건너뛰었습니다.`);
    return null;
  }

  const prstGeom = one(spPr, 'prstGeom');
  const nativeShape = nativeFor(prstGeom?.attrs.prst ?? 'rect', place.w, place.h,
    readAdjust(prstGeom));

  return {
    type: 'image',
    name,
    place,
    opacity: 1,
    data: await file.async('base64'),
    mime,
    isSvg: ext === 'svg',
    ...(nativeShape && nativeShape.kind === 'roundRect' ? { radii: nativeShape.radii } : {}),
  };
}

async function readGraphicFrame(el: XNode, parent: Mat, ctx: Ctx): Promise<ImportNode | null> {
  const local = localMatrix(one(el, 'xfrm'));
  if (!local) return null;

  const name = xpath(el, 'nvGraphicFramePr', 'cNvPr')?.attrs.name ?? '개체';
  const place = placementOf(parent, local);

  const tbl = deep(el, 'tbl');
  if (tbl) return readTable(tbl, name, place, ctx);

  if (deep(el, 'chart')) {
    ctx.warn(ctx.slideName, name, '차트는 Figma 에 대응 개체가 없어 건너뛰었습니다.');
  } else if (deep(el, 'relIds')) {
    ctx.warn(ctx.slideName, name, 'SmartArt 는 Figma 에 대응 개체가 없어 건너뛰었습니다.');
  }
  return null;
}

function readTable(
  tbl: XNode, name: string, place: Placement, ctx: Ctx,
): ImportNode {
  const colWidths = all(one(tbl, 'tblGrid'), 'gridCol').map((g) => pt(num(g.attrs.w)));
  const rowsXml = all(tbl, 'tr');
  const rowHeights = rowsXml.map((r) => pt(num(r.attrs.h)));

  const rows: TableCell[][] = rowsXml.map((tr) => all(tr, 'tc').map((tc) => {
    const tcPr = one(tc, 'tcPr');
    const fill = readFill(tcPr, ctx.color);
    const anchor = tcPr?.attrs.anchor;
    const cell: TableCell = {
      merged: bool(tc.attrs.hMerge) || bool(tc.attrs.vMerge),
      rowSpan: num(tc.attrs.rowSpan, 1),
      colSpan: num(tc.attrs.gridSpan, 1),
      ...(fill ? { fill } : {}),
      paragraphs: readParagraphs(one(tc, 'txBody'), ctx, [ctx.masterStyles.other]),
      insets: {
        left: pt(num(tcPr?.attrs.marL, 91440)),
        top: pt(num(tcPr?.attrs.marT, 45720)),
        right: pt(num(tcPr?.attrs.marR, 91440)),
        bottom: pt(num(tcPr?.attrs.marB, 45720)),
      },
      vertical: anchor === 'ctr' ? 'CENTER' : anchor === 'b' ? 'BOTTOM' : 'TOP',
      borders: {
        top: readStroke(one(tcPr, 'lnT'), ctx),
        right: readStroke(one(tcPr, 'lnR'), ctx),
        bottom: readStroke(one(tcPr, 'lnB'), ctx),
        left: readStroke(one(tcPr, 'lnL'), ctx),
      },
    };
    return cell;
  }));

  return { type: 'table', name, place, opacity: 1, colWidths, rowHeights, rows };
}

/* ── grpFill 병합 ────────────────────────────────────────────── */

/**
 * 그룹 채우기를 쓰는 형제 도형들을 하나로 합친다.
 *
 * PowerPoint 는 그룹 채우기를 자식들의 **합집합에 한 번** 칠한다. Figma 는 노드가 여러 개면
 * 여러 번 칠하고 겹친 곳을 합성하므로, 반투명 그라디언트에서는 이음매가 진하게 드러난다.
 * 경로를 합쳐 도형 하나로 만들면 원본과 같아진다.
 *
 * 회전·반전이 걸린 도형은 단순 평행이동으로 합칠 수 없어 제외한다.
 */
interface MergePart {
  path: string;
  place: Placement;
}

function mergeParts(parts: MergePart[], fill: Paint, name: string): ShapeNode | null {
  if (parts.length === 0) return null;

  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (const p of parts) {
    minX = Math.min(minX, p.place.x);
    minY = Math.min(minY, p.place.y);
    maxX = Math.max(maxX, p.place.x + p.place.w);
    maxY = Math.max(maxY, p.place.y + p.place.h);
  }

  const data = parts
    .map((p) => translatePath(p.path, p.place.x - minX, p.place.y - minY))
    .join(' ');

  return {
    type: 'shape',
    name,
    place: {
      x: minX,
      y: minY,
      w: Math.max(0.01, maxX - minX),
      h: Math.max(0.01, maxY - minY),
      rotation: 0,
      flipH: false,
      flipV: false,
    },
    opacity: 1,
    geometry: { kind: 'path', data, evenOdd: false },
    fill,
  };
}

/** 절대 좌표 경로를 평행이동한다. 모든 명령의 인자가 (x, y) 쌍이라 번갈아 더하면 된다. */
function translatePath(data: string, dx: number, dy: number): string {
  let i = 0;
  return data.replace(/-?\d*\.?\d+/g, (token) => {
    const v = Number(token) + (i++ % 2 === 0 ? dx : dy);
    return (Math.round(v * 100) / 100).toString();
  });
}

/** 원시 도형을 경로로 편다. 합치려면 전부 경로여야 한다. */
function geometryToPath(geom: ShapeGeometry, w: number, h: number): string | null {
  switch (geom.kind) {
    case 'path':
      return geom.data;
    case 'rect':
      return `M0 0 L${w} 0 L${w} ${h} L0 ${h} Z`;
    case 'line':
      return `M0 0 L${w} ${h}`;
    case 'ellipse': {
      const k = 0.5523;
      const rx = w / 2;
      const ry = h / 2;
      return `M${rx} 0 C${rx + rx * k} 0 ${w} ${ry - ry * k} ${w} ${ry}`
        + ` C${w} ${ry + ry * k} ${rx + rx * k} ${h} ${rx} ${h}`
        + ` C${rx - rx * k} ${h} 0 ${ry + ry * k} 0 ${ry}`
        + ` C0 ${ry - ry * k} ${rx - rx * k} 0 ${rx} 0 Z`;
    }
    case 'roundRect': {
      const k = 0.5523;
      const [tl, tr, br, bl] = geom.radii;
      return `M${tl} 0 L${w - tr} 0 C${w - tr + tr * k} 0 ${w} ${tr - tr * k} ${w} ${tr}`
        + ` L${w} ${h - br} C${w} ${h - br + br * k} ${w - br + br * k} ${h} ${w - br} ${h}`
        + ` L${bl} ${h} C${bl - bl * k} ${h} 0 ${h - bl + bl * k} 0 ${h - bl}`
        + ` L0 ${tl} C0 ${tl - tl * k} ${tl - tl * k} 0 ${tl} 0 Z`;
    }
    default:
      return null;
  }
}

/* ── 텍스트 ──────────────────────────────────────────────────── */

function readTextBody(
  txBody: XNode | null, place: Placement, name: string, ctx: Ctx, styles: StyleChain,
): ImportNode | null {
  if (!txBody) return null;
  const paragraphs = readParagraphs(txBody, ctx, styles);
  if (paragraphs.length === 0) return null;

  // Figma 텍스트는 가로 정렬이 노드 단위다. 문단마다 다르면 첫 문단 기준으로 눌린다.
  if (paragraphs.some((p) => p.align !== paragraphs[0].align)) {
    ctx.warn(ctx.slideName, name, t().mixedAlign);
  }

  const bodyPr = one(txBody, 'bodyPr');
  const anchor = bodyPr?.attrs.anchor;
  const spAutoFit = !!one(bodyPr, 'spAutoFit');

  return {
    type: 'text',
    name,
    place,
    opacity: 1,
    paragraphs,
    vertical: anchor === 'ctr' ? 'CENTER' : anchor === 'b' ? 'BOTTOM' : 'TOP',
    insets: {
      left: pt(num(bodyPr?.attrs.lIns, DEFAULT_INSETS.l)),
      top: pt(num(bodyPr?.attrs.tIns, DEFAULT_INSETS.t)),
      right: pt(num(bodyPr?.attrs.rIns, DEFAULT_INSETS.r)),
      bottom: pt(num(bodyPr?.attrs.bIns, DEFAULT_INSETS.b)),
    },
    autoWidth: bodyPr?.attrs.wrap === 'none' && spAutoFit,
  };
}

function readParagraphs(txBody: XNode | null, ctx: Ctx, styles: StyleChain): Paragraph[] {
  if (!txBody) return [];
  const out: Paragraph[] = [];

  for (const p of all(txBody, 'p')) {
    const pPr = one(p, 'pPr');
    const level = num(pPr?.attrs.lvl);
    const runs: TextRun[] = [];

    for (const child of p.children) {
      if (child.tag === 'r') {
        const run = readRun(child, pPr, ctx, styles, level);
        if (run.text) runs.push(run);
      } else if (child.tag === 'br') {
        runs.push({ ...readRun(null, pPr, ctx, styles, level), text: '\n' });
      } else if (child.tag === 'fld') {
        // 슬라이드 번호 등 필드 — 마지막으로 계산된 값이 그대로 들어 있다
        const run = readRun(child, pPr, ctx, styles, level);
        if (run.text) runs.push(run);
      }
    }

    if (runs.length === 0 && out.length === 0) continue;

    const algn = pPr?.attrs.algn;
    const para: Paragraph = {
      runs,
      align: algn === 'ctr' ? 'CENTER' : algn === 'r' ? 'RIGHT'
        : algn === 'just' ? 'JUSTIFIED' : 'LEFT',
      level,
    };

    const spcPct = xpath(pPr, 'lnSpc', 'spcPct');
    if (spcPct) para.lineHeightPct = num(spcPct.attrs.val) / 1000;
    const before = xpath(pPr, 'spcBef', 'spcPts');
    if (before) para.spaceBefore = num(before.attrs.val) / 100;
    const after = xpath(pPr, 'spcAft', 'spcPts');
    if (after) para.spaceAfter = num(after.attrs.val) / 100;

    const buChar = one(pPr, 'buChar');
    const buNum = one(pPr, 'buAutoNum');
    if (buChar) para.bullet = { kind: 'char', char: buChar.attrs.char ?? '•' };
    else if (buNum) para.bullet = { kind: 'number' };

    out.push(para);
  }

  // 뒤쪽 빈 문단은 버린다
  while (out.length > 0 && out[out.length - 1].runs.length === 0) out.pop();
  return out;
}

/**
 * 런 속성을 상속 사슬을 따라 해석한다.
 *
 * 우선순위: 런의 rPr → 문단의 defRPr → 도형 lstStyle → 레이아웃 자리표시자 → 마스터 → 기본값.
 * 크기·색·폰트 모두 이 사슬을 타므로, 런만 보면 값이 없을 때 임의의 기본값으로 떨어진다.
 */
function readRun(
  r: XNode | null, pPr: XNode | null, ctx: Ctx, styles: StyleChain, level: number,
): TextRun {
  const chain: StyleChain = [
    r ? one(r, 'rPr') : null,
    xpath(pPr, 'defRPr'),
    ...styles.map((s) => levelDefaults(s, level)),
  ];

  const size = num(pickAttr(chain, 'sz'), 1800) / 100;
  const family = resolveFont(chain, ctx);
  ctx.fonts.add(family);

  let color = '000000';
  let opacity = 1;
  const solid = pickChild(chain, 'solidFill');
  if (solid) {
    const c = resolveColorNode(colorNode(solid), ctx.color);
    if (c) {
      color = c.color;
      opacity = c.opacity;
    }
  }

  const bold = bool(pickAttr(chain, 'b'));
  const italic = bool(pickAttr(chain, 'i'));
  const run: TextRun = {
    text: r ? (one(r, 't')?.text ?? '') : '',
    size,
    fontFamily: family,
    fontStyle: bold && italic ? 'Bold Italic' : bold ? 'Bold' : italic ? 'Italic' : 'Regular',
    bold,
    italic,
    underline: (pickAttr(chain, 'u') ?? 'none') !== 'none',
    strike: (pickAttr(chain, 'strike') ?? 'noStrike') !== 'noStrike',
    color,
    opacity,
  };

  const spc = pickAttr(chain, 'spc');
  if (spc) run.letterSpacing = num(spc) / 100;

  const link = pickChild(chain, 'hlinkClick');
  if (link?.attrs['r:id']) run.link = link.attrs['r:id'];

  return run;
}

/** `+mj-lt` / `+mn-lt` 는 테마 폰트 참조다. 한글 문서는 ea(동아시아) 쪽이 실제 폰트인 경우가 많다. */
function resolveFont(chain: StyleChain, ctx: Ctx): string {
  for (const tag of ['ea', 'latin', 'cs']) {
    const node = pickChild(chain, tag);
    const face = node?.attrs.typeface;
    if (!face) continue;
    if (face.startsWith('+mj')) return ctx.themeFonts.major;
    if (face.startsWith('+mn')) return ctx.themeFonts.minor;
    return face;
  }
  return ctx.themeFonts.minor;
}
