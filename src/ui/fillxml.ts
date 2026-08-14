import JSZip from 'jszip';
import type { Fill, Gradient, Pattern, Stroke } from '../shared/ir';

/**
 * PptxGenJS 가 쓰지 못하는 채우기를 XML 에 덧쓴다.
 *
 * PptxGenJS 4.x 는 칠을 'solid' | 'none' 으로만 받는다 — 소스에 `gradFill` 이나 `pattFill`
 * 을 쓰는 코드가 아예 없다. 그래서 색은 평균 단색으로 정상 생성해 두고, 파일을 굳히기 직전에
 * 그 도형의 `<a:solidFill>` 을 제 것으로 바꿔 끼운다.
 *
 * 도형을 찾는 열쇠는 이름이다. PptxGenJS 는 `objectName` 을 `<p:cNvPr name="…">` 에 그대로
 * 쓰므로, 표식을 붙여 내보내고 덧쓰면서 떼어낸다. 이름이 겹쳐도 표식은 고유해서 헷갈리지 않는다.
 *
 * 덧쓰기가 실패해도 평균 단색이 그대로 남는다. 나빠질 수는 없고 맞을 때만 좋아진다.
 */

/** 도형 하나에 덧쓸 XML 조각들 */
export interface Mark {
  id: number;
  /** 도형 자체의 칠 */
  fill?: string;
  /** 테두리 */
  line?: string;
  /** 글자 (그 도형 안의 모든 런) */
  text?: string;
}

/** `ppt/slides/slide3.xml` → 그 장에 덧쓸 표식들 */
export type Marks = Map<string, Mark[]>;

/**
 * 표식을 발급하고 어느 부품에 속하는지 적어 둔다.
 * 이름에 붙일 꼬리표를 돌려주며, 덧쓸 것이 없으면 아무것도 돌려주지 않는다.
 */
export class Marker {
  private next = 1;
  readonly parts: Marks = new Map();

  claim(part: string, frags: Omit<Mark, 'id'>): string | null {
    if (!frags.fill && !frags.line && !frags.text) return null;
    const mark: Mark = { id: this.next++, ...frags };
    const list = this.parts.get(part);
    if (list) list.push(mark);
    else this.parts.set(part, [mark]);
    return tag(mark.id);
  }
}

/**
 * 이름 뒤에 붙는 꼬리표.
 *
 * `objectName` 은 `encodeXmlEntities` 를 거치므로 `& < > " '` 는 쓸 수 없다.
 * 물결과 숫자만 쓰면 원문 그대로 XML 에 실린다.
 */
function tag(id: number): string {
  return `~g${id}~`;
}

/** 남은 꼬리표를 지운다. 덧쓰기가 걸리지 않은 것도 사용자 눈에 보이면 안 된다. */
const LEFTOVER = / ~g\d+~/g;

/** 덧쓸 XML 부품인가 — 슬라이드와 레이아웃에만 도형이 있다. */
export function isDrawingPart(path: string): boolean {
  return /^ppt\/(slides\/slide|slideLayouts\/slideLayout)\d+\.xml$/.test(path);
}

const FULL_TURN = 21600000;

/** 0..100 투명도 → `<a:alpha>`. 불투명하면 아무것도 쓰지 않는다 (기본값이 100%). */
function alphaOf(transparency: number): string {
  return transparency > 0 ? `<a:alpha val="${Math.round((100 - transparency) * 1000)}"/>` : '';
}

function clr(color: string, transparency: number): string {
  const inner = alphaOf(transparency);
  const val = color.toUpperCase();
  return inner ? `<a:srgbClr val="${val}">${inner}</a:srgbClr>` : `<a:srgbClr val="${val}"/>`;
}

export function gradFillXml(g: Gradient): string {
  const stops = g.stops.map((s) => {
    const pos = Math.round(Math.max(0, Math.min(1, s.position)) * 100000);
    return `<a:gs pos="${pos}">${clr(s.color, s.transparency)}</a:gs>`;
  }).join('');

  /*
   * 방사형은 중심에서 퍼진다. `fillToRect` 가 마지막 정지점이 닿는 사각형이고,
   * 네 변을 다 50% 로 주면 도형 한가운데의 한 점이 된다 — 그게 원형 그라디언트다.
   */
  const shape = g.type === 'radial'
    ? '<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>'
    : `<a:lin ang="${(Math.round(g.angle * 60000) % FULL_TURN + FULL_TURN) % FULL_TURN}" scaled="0"/>`;

  return `<a:gradFill rotWithShape="1"><a:gsLst>${stops}</a:gsLst>${shape}</a:gradFill>`;
}

export function pattFillXml(p: Pattern): string {
  return `<a:pattFill prst="${p.preset}">`
    + `<a:fgClr>${clr(p.fg, p.fgTransparency)}</a:fgClr>`
    + `<a:bgClr>${clr(p.bg, p.bgTransparency)}</a:bgClr>`
    + '</a:pattFill>';
}

/**
 * 이 칠을 덧써야 하는가. 덧쓸 것이 없으면 아무것도 돌려주지 않는다.
 *
 * 무늬는 가져올 때 눌러 둔 단색이 그대로일 때만 되돌린다 — 색을 바꿨다면 그쪽이 사용자의
 * 최신 의사이고, 원본 무늬 색으로 되돌리면 그 편집을 뭉갠다.
 */
export function fillXmlOf(fill: Fill | undefined): string | undefined {
  if (fill?.kind !== 'solid') return undefined;
  if (fill.gradient) return gradFillXml(fill.gradient);
  if (fill.pattern && fill.pattern.approximated.toUpperCase() === fill.color.toUpperCase()) {
    return pattFillXml(fill.pattern);
  }
  return undefined;
}

export function lineXmlOf(stroke: Stroke | undefined): string | undefined {
  return stroke?.gradient ? gradFillXml(stroke.gradient) : undefined;
}

/** 첫 `<a:solidFill>…</a:solidFill>` 을 바꾼다. 없으면 원본 그대로. */
function swapFirstFill(xml: string, replacement: string): string {
  const at = xml.indexOf('<a:solidFill>');
  if (at < 0) return xml;
  const end = xml.indexOf('</a:solidFill>', at);
  if (end < 0) return xml;
  return xml.slice(0, at) + replacement + xml.slice(end + '</a:solidFill>'.length);
}

/** 모든 `<a:solidFill>` 을 바꾼다 — 런마다 하나씩 있는 글자 색용. */
function swapEveryFill(xml: string, replacement: string): string {
  return xml.replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/g, replacement);
}

/**
 * 테두리가 시작하는 자리.
 *
 * `<a:ln` 으로 찾으면 안 된다 — custGeom 경로의 `<a:lnTo>` 가 먼저 걸린다. 그러면 경계가
 * 기하 한복판에 잡혀 칠을 엉뚱한 쪽에서 찾게 되고, **직선 구간이 있는 벡터만** 조용히
 * 덧쓰기에 실패한다 (곡선만 있는 벡터는 `<a:cubicBezTo>` 뿐이라 멀쩡했다).
 */
const LINE_START = /<a:ln[ >]/;

/**
 * 도형 하나(`<p:sp>…</p:sp>`) 안에서 자리를 갈라 덧쓴다.
 *
 * PptxGenJS 가 내는 `<p:spPr>` 는 xfrm → 기하 → 칠 → 선 순서다. 칠과 선의 `<a:solidFill>`
 * 은 생김새가 같아서, 선이 시작하는 자리를 경계로 앞뒤를 나눠야 서로를 침범하지 않는다.
 */
function rewriteShape(sp: string, mark: Mark): string {
  const bodyAt = sp.indexOf('<p:txBody>');
  let head = bodyAt < 0 ? sp : sp.slice(0, bodyAt);
  const body = bodyAt < 0 ? '' : sp.slice(bodyAt);

  if (mark.fill || mark.line) {
    const lnAt = head.search(LINE_START);
    const beforeLn = lnAt < 0 ? head : head.slice(0, lnAt);
    const fromLn = lnAt < 0 ? '' : head.slice(lnAt);
    head = (mark.fill ? swapFirstFill(beforeLn, mark.fill) : beforeLn)
      + (mark.line ? swapFirstFill(fromLn, mark.line) : fromLn);
  }

  return head + (mark.text && body ? swapEveryFill(body, mark.text) : body);
}

/** 부품 하나에 그 부품 몫의 표식을 전부 덧쓰고, 남은 꼬리표를 지운다. */
export function injectPart(xml: string, marks: readonly Mark[]): string {
  let out = xml;
  for (const mark of marks) {
    // 자리는 덧쓸 때마다 밀리므로 매번 처음부터 찾는다.
    const at = out.indexOf(`${tag(mark.id)}"`);
    if (at < 0) continue;
    const open = out.lastIndexOf('<p:sp>', at);
    const close = out.indexOf('</p:sp>', at);
    if (open < 0 || close < 0) continue;
    out = out.slice(0, open) + rewriteShape(out.slice(open, close), mark) + out.slice(close);
  }
  return out.replace(LEFTOVER, '');
}

/**
 * PptxGenJS 가 zip 에 넣는 슬라이드 XML 을 가로채 덧쓴다.
 *
 * 다 만든 pptx 를 다시 풀었다 담는 방법도 있지만, 100MB 짜리 파일을 통째로 두 벌 들고 있어야
 * 한다. 여기서 가로채면 몇 KB 짜리 문자열 하나만 만지고 끝난다.
 *
 * 되돌리기 함수를 반드시 부를 것 — 다른 코드가 쓰는 JSZip 을 계속 바꿔 둘 수는 없다.
 */
export function hookSlideXml(marks: Marks): () => void {
  const proto = JSZip.prototype as unknown as {
    file: (name: unknown, data?: unknown, opts?: unknown) => unknown;
  };
  const original = proto.file;

  proto.file = function patched(name: unknown, data?: unknown, opts?: unknown): unknown {
    if (typeof name === 'string' && typeof data === 'string' && isDrawingPart(name)) {
      return original.call(this, name, injectPart(data, marks.get(name) ?? []), opts);
    }
    return original.call(this, name, data, opts);
  };

  return () => {
    proto.file = original;
  };
}
