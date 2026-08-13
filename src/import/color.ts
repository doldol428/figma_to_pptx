import type { Hex } from '../shared/ir';
import type { GradientPaint, Paint, SolidPaint } from '../shared/importir';
import { all, num, one, type XNode } from './xml';

/**
 * OOXML 색 해석.
 *
 * 슬라이드 XML 의 색은 대부분 최종 값이 아니다. `<a:schemeClr val="accent1">` 은 테마를
 * 가리키고, 그 위에 lumMod/lumOff/shade/tint/satMod 변형이 얹힌다. 이 변형을 무시하면
 * 문서 전체 색조가 통째로 틀어진다 — 실제 제안서 파일에서 테마 참조가 1,650회 나왔다.
 *
 * lumMod/lumOff/satMod 는 RGB 가 아니라 HSL 공간에서 동작한다. RGB 곱셈으로 근사하면
 * 회색조 계열에서 특히 크게 어긋난다.
 */

export interface ColorContext {
  /** 테마 clrScheme — dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink */
  scheme: Record<string, string>;
  /** 마스터의 clrMap — bg1/tx1/bg2/tx2 가 어느 슬롯을 가리키는지 */
  map: Record<string, string>;
  /** placeholder 색(phClr) — 도형 스타일 해석 시 채워진다 */
  phClr?: string;
}

const DEFAULT_MAP: Record<string, string> = {
  bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2',
};

const PRESET: Record<string, string> = {
  black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF',
  yellow: 'FFFF00', cyan: '00FFFF', magenta: 'FF00FF', gray: '808080', grey: '808080',
  darkGray: 'A9A9A9', lightGray: 'D3D3D3', orange: 'FFA500', purple: '800080',
};

/* ── HSL 변환 ─────────────────────────────────────────────────── */

interface Rgb { r: number; g: number; b: number }

function toHsl({ r, g, b }: Rgb): { h: number; s: number; l: number } {
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return { h: 0, s: 0, l };
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === r) h = ((g - b) / d + (g < b ? 6 : 0)) / 6;
  else if (max === g) h = ((b - r) / d + 2) / 6;
  else h = ((r - g) / d + 4) / 6;
  return { h, s, l };
}

function hueToRgb(p: number, q: number, t: number): number {
  let x = t;
  if (x < 0) x += 1;
  if (x > 1) x -= 1;
  if (x < 1 / 6) return p + (q - p) * 6 * x;
  if (x < 1 / 2) return q;
  if (x < 2 / 3) return p + (q - p) * (2 / 3 - x) * 6;
  return p;
}

function toRgb(h: number, s: number, l: number): Rgb {
  if (s === 0) return { r: l, g: l, b: l };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return {
    r: hueToRgb(p, q, h + 1 / 3),
    g: hueToRgb(p, q, h),
    b: hueToRgb(p, q, h - 1 / 3),
  };
}

const clamp01 = (n: number): number => Math.max(0, Math.min(1, n));

function hexOf({ r, g, b }: Rgb): Hex {
  const ch = (v: number): string =>
    Math.round(clamp01(v) * 255).toString(16).toUpperCase().padStart(2, '0');
  return ch(r) + ch(g) + ch(b);
}

function parseHex(hex: string): Rgb {
  const h = hex.replace('#', '').padStart(6, '0');
  return {
    r: parseInt(h.slice(0, 2), 16) / 255,
    g: parseInt(h.slice(2, 4), 16) / 255,
    b: parseInt(h.slice(4, 6), 16) / 255,
  };
}

/* ── 색 노드 해석 ─────────────────────────────────────────────── */

const COLOR_TAGS = ['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'scrgbClr', 'hslClr'];

/** 색을 담고 있는 컨테이너(solidFill, ln 등)에서 첫 색 노드를 찾는다 */
export function colorNode(container: XNode | null): XNode | null {
  if (!container) return null;
  for (const c of container.children) {
    if (COLOR_TAGS.indexOf(c.tag) >= 0) return c;
  }
  return null;
}

export interface ResolvedColor {
  color: Hex;
  /** 0.0 - 1.0 */
  opacity: number;
}

export function resolveColorNode(node: XNode | null, ctx: ColorContext): ResolvedColor | null {
  if (!node) return null;

  let base: Rgb;
  switch (node.tag) {
    case 'srgbClr':
      base = parseHex(node.attrs.val ?? '000000');
      break;
    case 'schemeClr': {
      const val = node.attrs.val ?? 'tx1';
      if (val === 'phClr') {
        base = parseHex(ctx.phClr ?? '808080');
      } else {
        const slot = ctx.map[val] ?? DEFAULT_MAP[val] ?? val;
        base = parseHex(ctx.scheme[slot] ?? ctx.scheme[val] ?? '808080');
      }
      break;
    }
    case 'sysClr':
      base = parseHex(node.attrs.lastClr ?? '000000');
      break;
    case 'prstClr':
      base = parseHex(PRESET[node.attrs.val ?? 'black'] ?? '808080');
      break;
    case 'scrgbClr':
      base = {
        r: num(node.attrs.r) / 100000,
        g: num(node.attrs.g) / 100000,
        b: num(node.attrs.b) / 100000,
      };
      break;
    case 'hslClr':
      base = toRgb(num(node.attrs.hue) / 21600000, num(node.attrs.sat) / 100000,
        num(node.attrs.lum) / 100000);
      break;
    default:
      return null;
  }

  let opacity = 1;
  let hsl: { h: number; s: number; l: number } | null = null;
  const ensureHsl = (): { h: number; s: number; l: number } => {
    if (!hsl) hsl = toHsl(base);
    return hsl;
  };
  const flush = (): void => {
    if (hsl) {
      base = toRgb(hsl.h, clamp01(hsl.s), clamp01(hsl.l));
      hsl = null;
    }
  };

  for (const mod of node.children) {
    const v = num(mod.attrs.val) / 100000;
    switch (mod.tag) {
      case 'alpha':
        opacity *= v;
        break;
      case 'lumMod':
        ensureHsl().l *= v;
        break;
      case 'lumOff':
        ensureHsl().l += v;
        break;
      case 'satMod':
        ensureHsl().s *= v;
        break;
      case 'satOff':
        ensureHsl().s += v;
        break;
      case 'hueMod':
        ensureHsl().h = (ensureHsl().h * v) % 1;
        break;
      case 'shade':
        // shade/tint 는 선형 RGB 에서의 혼합이다. HSL 보정과 순서가 엉키지 않게 먼저 반영한다.
        flush();
        base = { r: base.r * v, g: base.g * v, b: base.b * v };
        break;
      case 'tint':
        flush();
        base = {
          r: base.r * v + (1 - v),
          g: base.g * v + (1 - v),
          b: base.b * v + (1 - v),
        };
        break;
      case 'gray':
        flush();
        base = { r: 0.5, g: 0.5, b: 0.5 };
        break;
      default:
        break;
    }
  }
  flush();

  return { color: hexOf(base), opacity: clamp01(opacity) };
}

/* ── 채우기 ───────────────────────────────────────────────────── */

/**
 * 도형 속성에서 채우기를 읽는다.
 * `noFill` 이면 명시적으로 null 을 돌려주고, 채우기 요소가 아예 없으면 undefined 를 준다
 * (상속을 계속 따라가야 한다는 뜻).
 */
export function readFill(spPr: XNode | null, ctx: ColorContext): Paint | null | undefined {
  if (!spPr) return undefined;
  if (one(spPr, 'noFill')) return null;

  const solid = one(spPr, 'solidFill');
  if (solid) {
    const c = resolveColorNode(colorNode(solid), ctx);
    if (c) {
      const paint: SolidPaint = { kind: 'solid', color: c.color, opacity: c.opacity };
      return paint;
    }
  }

  const grad = one(spPr, 'gradFill');
  if (grad) return readGradient(grad, ctx);

  // blipFill(이미지 채우기)은 호출부가 그림 노드로 따로 처리한다.
  return undefined;
}

export function readGradient(grad: XNode, ctx: ColorContext): GradientPaint | null {
  const list = one(grad, 'gsLst');
  const stops = all(list, 'gs')
    .map((gs) => {
      const c = resolveColorNode(colorNode(gs), ctx);
      if (!c) return null;
      return {
        position: clamp01(num(gs.attrs.pos) / 100000),
        color: c.color,
        opacity: c.opacity,
      };
    })
    .filter((s): s is NonNullable<typeof s> => s !== null)
    .sort((a, b) => a.position - b.position);

  if (stops.length === 0) return null;
  if (stops.length === 1) stops.push({ ...stops[0], position: 1 });

  const lin = one(grad, 'lin');
  const pathEl = one(grad, 'path');

  if (pathEl && pathEl.attrs.path !== 'linear') {
    return { kind: 'gradient', type: 'radial', angle: 0, stops };
  }

  // ang 은 1/60000 도, 화면 기준 시계방향. Figma 도 y 가 아래로 증가하므로 그대로 대응된다.
  const angle = lin ? (num(lin.attrs.ang) / 60000) % 360 : 0;
  return { kind: 'gradient', type: 'linear', angle, stops };
}

/**
 * `<a:grpFill/>` — "채우기는 부모 그룹 것을 쓴다".
 * 이걸 모르면 도형에 채우기가 없다고 판단해 통째로 버리게 된다.
 */
export function inheritsGroupFill(spPr: XNode | null): boolean {
  return !!one(spPr, 'grpFill');
}

/** 그라디언트의 t 지점 색 (선형 보간) */
function sampleGradient(g: GradientPaint, t: number): { color: Hex; opacity: number } {
  const stops = g.stops;
  if (t <= stops[0].position) return { color: stops[0].color, opacity: stops[0].opacity };
  const last = stops[stops.length - 1];
  if (t >= last.position) return { color: last.color, opacity: last.opacity };

  for (let i = 1; i < stops.length; i++) {
    const a = stops[i - 1];
    const b = stops[i];
    if (t > b.position) continue;
    const span = b.position - a.position;
    const k = span <= 0 ? 0 : (t - a.position) / span;
    const ca = parseHex(a.color);
    const cb = parseHex(b.color);
    return {
      color: hexOf({
        r: ca.r + (cb.r - ca.r) * k,
        g: ca.g + (cb.g - ca.g) * k,
        b: ca.b + (cb.b - ca.b) * k,
      }),
      opacity: a.opacity + (b.opacity - a.opacity) * k,
    };
  }
  return { color: last.color, opacity: last.opacity };
}

interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** 상자를 그라디언트 축에 투영한 구간 */
function projectOnAxis(box: Box, dx: number, dy: number): { lo: number; hi: number } {
  let lo = Infinity;
  let hi = -Infinity;
  for (const x of [box.x, box.x + box.w]) {
    for (const y of [box.y, box.y + box.h]) {
      const p = x * dx + y * dy;
      lo = Math.min(lo, p);
      hi = Math.max(hi, p);
    }
  }
  return { lo, hi };
}

/**
 * 그룹에 걸린 채우기를 자식 하나에 맞춰 잘라낸다.
 *
 * PowerPoint 는 그룹의 그라디언트를 그룹 전체 범위에 한 번 그리고 자식들이 그 일부를 보여준다.
 * 자식마다 원본을 통째로 복사하면 같은 그라디언트가 자식 수만큼 반복돼 완전히 다른 그림이 된다.
 * 그래서 자식이 차지하는 구간만 0~1 로 다시 펴서 넘긴다.
 */
export function sliceFillForChild(paint: Paint, group: Box, child: Box): Paint {
  if (paint.kind !== 'gradient' || paint.type !== 'linear') return paint;

  const rad = (paint.angle * Math.PI) / 180;
  const dx = Math.cos(rad);
  const dy = Math.sin(rad);
  const g = projectOnAxis(group, dx, dy);
  const c = projectOnAxis(child, dx, dy);

  const span = g.hi - g.lo;
  if (span <= 1e-6) return paint;

  const t0 = (c.lo - g.lo) / span;
  const t1 = (c.hi - g.lo) / span;
  if (t1 - t0 <= 1e-6) {
    const at = sampleGradient(paint, t0);
    return { kind: 'solid', color: at.color, opacity: at.opacity };
  }

  const start = sampleGradient(paint, t0);
  const end = sampleGradient(paint, t1);
  const stops = [{ position: 0, color: start.color, opacity: start.opacity }];
  for (const s of paint.stops) {
    const p = (s.position - t0) / (t1 - t0);
    if (p > 0.0001 && p < 0.9999) {
      stops.push({ position: p, color: s.color, opacity: s.opacity });
    }
  }
  stops.push({ position: 1, color: end.color, opacity: end.opacity });

  return { ...paint, stops };
}

/** 선(ln) 요소에서 색을 읽는다 */
export function readLinePaint(ln: XNode | null, ctx: ColorContext): Paint | null {
  if (!ln) return null;
  if (one(ln, 'noFill')) return null;
  const solid = one(ln, 'solidFill');
  if (solid) {
    const c = resolveColorNode(colorNode(solid), ctx);
    if (c) return { kind: 'solid', color: c.color, opacity: c.opacity };
  }
  const grad = one(ln, 'gradFill');
  if (grad) return readGradient(grad, ctx);
  return null;
}
