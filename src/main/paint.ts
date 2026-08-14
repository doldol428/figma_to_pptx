import { t } from '../shared/i18n';
import type { Fill, Gradient, Hex, Shadow, Stroke } from '../shared/ir';

export function toHex(c: RGB | RGBA): Hex {
  const ch = (v: number) => {
    const n = Math.max(0, Math.min(255, Math.round(v * 255)));
    return n.toString(16).toUpperCase().padStart(2, '0');
  };
  return ch(c.r) + ch(c.g) + ch(c.b);
}

/** 0..1 불투명도 → PPTX 투명도 퍼센트(0..100) */
function transparency(opacity: number): number {
  return Math.round((1 - clamp01(opacity)) * 100 * 10) / 10;
}

function clamp01(n: number): number {
  return Math.max(0, Math.min(1, n));
}

export function topVisiblePaint(paints: readonly Paint[] | typeof figma.mixed): Paint | null {
  if (paints === figma.mixed || !Array.isArray(paints)) return null;
  for (let i = paints.length - 1; i >= 0; i--) {
    const p = paints[i];
    if (p.visible !== false && (p.opacity ?? 1) > 0) return p;
  }
  return null;
}

export function countVisiblePaints(paints: readonly Paint[] | typeof figma.mixed): number {
  if (paints === figma.mixed || !Array.isArray(paints)) return 0;
  return paints.filter((p) => p.visible !== false && (p.opacity ?? 1) > 0).length;
}

/**
 * PPTX 가 그대로 담을 수 있는 그라디언트인가.
 *
 * 선형과 방사형만 `<a:gradFill>` 로 옮겨진다. 원뿔형(ANGULAR)·마름모형(DIAMOND)은
 * DrawingML 에 대응이 없어 단색 평균으로 남는다.
 */
function gradientOf(paint: GradientPaint): Gradient | undefined {
  if (paint.type !== 'GRADIENT_LINEAR' && paint.type !== 'GRADIENT_RADIAL') return undefined;

  /*
   * Figma 는 단위 사각형에서 왼→오른쪽으로 흐르는 그라디언트를 gradientTransform 으로 돌린다.
   * 첫 행이 곧 그 방향벡터라 각도는 거기서 바로 나온다 (가져오기의 역산 — import/…/create.ts).
   */
  const m = paint.gradientTransform;
  const angle = (Math.atan2(m[0][1], m[0][0]) * 180) / Math.PI;

  return {
    type: paint.type === 'GRADIENT_RADIAL' ? 'radial' : 'linear',
    angle: ((angle % 360) + 360) % 360,
    stops: paint.gradientStops.map((s) => ({
      position: clamp01(s.position),
      color: toHex(s.color),
      transparency: transparency(s.color.a ?? 1),
    })),
  };
}

/**
 * 그라디언트 → 단색 근사.
 *
 * PptxGenJS 4.x 는 도형 fill 을 'solid' | 'none' 으로만 노출하므로 정지점들을 위치 가중
 * 평균해서 가장 덜 튀는 단색을 만든다. 선형·방사형은 이 단색 위에 원본을 얹어 두었다가
 * (`Fill.gradient`) 파일을 굳히기 직전 XML 에 덧써 넣는다 — ui/gradient.ts.
 */
function averageGradient(paint: GradientPaint): { color: Hex; opacity: number } {
  const stops = paint.gradientStops;
  if (stops.length === 0) return { color: '808080', opacity: 1 };

  let r = 0;
  let g = 0;
  let b = 0;
  let a = 0;
  let wSum = 0;
  for (let i = 0; i < stops.length; i++) {
    const prev = i === 0 ? stops[0].position : stops[i - 1].position;
    const next = i === stops.length - 1 ? stops[i].position : stops[i + 1].position;
    const w = Math.max((next - prev) / 2, 1e-3);
    const c = stops[i].color;
    r += c.r * w;
    g += c.g * w;
    b += c.b * w;
    a += (c.a ?? 1) * w;
    wSum += w;
  }
  return {
    color: toHex({ r: r / wSum, g: g / wSum, b: b / wSum }),
    opacity: a / wSum,
  };
}

export type FillResult =
  | { kind: 'fill'; fill: Fill }
  /** 이미지 페인트 — 노드를 렌더해서 그림으로 배치해야 한다 */
  | { kind: 'image' }
  /** 근사 처리됨 — 경고 문구 포함 */
  | { kind: 'fill'; fill: Fill; note: string };

/** 노드의 최상단 보이는 페인트 하나를 PPTX fill 로 변환한다. nodeOpacity 는 곱해서 반영. */
export function resolveFill(
  paints: readonly Paint[] | typeof figma.mixed,
  nodeOpacity: number,
): FillResult {
  const paint = topVisiblePaint(paints);
  if (!paint) return { kind: 'fill', fill: { kind: 'none' } };

  const pOpacity = paint.opacity ?? 1;

  if (paint.type === 'SOLID') {
    return {
      kind: 'fill',
      fill: {
        kind: 'solid',
        color: toHex(paint.color),
        transparency: transparency(pOpacity * nodeOpacity),
      },
    };
  }

  if (paint.type === 'IMAGE' || paint.type === 'VIDEO') {
    return { kind: 'image' };
  }

  if (
    paint.type === 'GRADIENT_LINEAR' ||
    paint.type === 'GRADIENT_RADIAL' ||
    paint.type === 'GRADIENT_ANGULAR' ||
    paint.type === 'GRADIENT_DIAMOND'
  ) {
    const avg = averageGradient(paint);
    const fill: Fill = {
      kind: 'solid',
      color: avg.color,
      transparency: transparency(avg.opacity * pOpacity * nodeOpacity),
    };
    const gradient = gradientOf(paint);
    if (gradient) {
      fill.gradient = gradient;
      return { kind: 'fill', fill };
    }
    return { kind: 'fill', fill, note: t().gradientApproximated(paint.type, avg.color) };
  }

  return { kind: 'fill', fill: { kind: 'none' } };
}

export interface SideWeights {
  top: number;
  right: number;
  bottom: number;
  left: number;
}

/**
 * 변마다 두께가 다르면 네 변의 값을, 균일하면 null 을 돌려준다.
 *
 * Figma 는 변별 두께가 설정되면 `strokeWeight` 로 `figma.mixed` 를 준다.
 * 표는 보통 행 프레임에 아래쪽 선만 주는 식으로 만들기 때문에 이 경우가 흔하다.
 * 균일한 값으로 뭉개면 없던 격자가 생기고 두께도 틀어진다.
 */
export function perSideWeights(node: SceneNode): SideWeights | null {
  if (!('strokeWeight' in node)) return null;
  const n = node as SceneNode & MinimalStrokesMixin & Partial<IndividualStrokesMixin>;
  if (typeof n.strokeWeight === 'number') return null;
  if (typeof n.strokeTopWeight !== 'number') return null;
  return {
    top: n.strokeTopWeight,
    right: n.strokeRightWeight ?? 0,
    bottom: n.strokeBottomWeight ?? 0,
    left: n.strokeLeftWeight ?? 0,
  };
}

export function resolveStroke(
  node: SceneNode & MinimalStrokesMixin,
  nodeOpacity: number,
): Stroke | undefined {
  const paint = topVisiblePaint(node.strokes);
  if (!paint) return undefined;

  // 변별 두께면 대표값으로 가장 두꺼운 변을 쓴다. 실제 그릴 때는 변마다 제 값으로 덮어쓴다.
  let weight: number;
  if (typeof node.strokeWeight === 'number') {
    weight = node.strokeWeight;
  } else {
    const sides = perSideWeights(node);
    weight = sides ? Math.max(sides.top, sides.right, sides.bottom, sides.left) : 1;
  }
  if (weight <= 0) return undefined;

  let color: Hex;
  let gradient: Gradient | undefined;
  let opacity = paint.opacity ?? 1;
  if (paint.type === 'SOLID') {
    color = toHex(paint.color);
  } else if (paint.type.startsWith('GRADIENT_')) {
    const avg = averageGradient(paint as GradientPaint);
    color = avg.color;
    opacity *= avg.opacity;
    gradient = gradientOf(paint as GradientPaint);
  } else {
    return undefined;
  }

  const stroke: Stroke = {
    color,
    transparency: transparency(opacity * nodeOpacity),
    width: weight,
    dashType: dashOf(node.dashPattern),
  };
  if (gradient) stroke.gradient = gradient;
  return stroke;
}

/** 텍스트 노드 전체에 걸린 그라디언트. 런의 색은 그 평균값이라 별도로 얹는다. */
export function textGradient(paints: readonly Paint[] | typeof figma.mixed): Gradient | undefined {
  const paint = topVisiblePaint(paints);
  if (!paint || !paint.type.startsWith('GRADIENT_')) return undefined;
  return gradientOf(paint as GradientPaint);
}

function dashOf(pattern: readonly number[] | undefined): Stroke['dashType'] {
  if (!pattern || pattern.length === 0) return 'solid';
  const [on, off = on] = pattern;
  if (on <= 1.5 && off >= on) return 'sysDot';
  if (pattern.length >= 4) return 'dashDot';
  return 'dash';
}

export function resolveShadow(effects: readonly Effect[] | typeof figma.mixed): Shadow | undefined {
  if (effects === figma.mixed || !Array.isArray(effects)) return undefined;
  for (const e of effects) {
    if (e.visible === false) continue;
    if (e.type !== 'DROP_SHADOW' && e.type !== 'INNER_SHADOW') continue;
    const { x, y } = e.offset;
    // PPT 그림자 각도는 동쪽 0°, 시계방향. Figma 는 y 가 아래로 증가하므로 그대로 대응된다.
    let angle = (Math.atan2(y, x) * 180) / Math.PI;
    if (angle < 0) angle += 360;
    return {
      type: e.type === 'DROP_SHADOW' ? 'outer' : 'inner',
      angle: Math.round(angle),
      offset: Math.round(Math.hypot(x, y) * 100) / 100,
      blur: Math.round(e.radius * 100) / 100,
      color: toHex(e.color),
      opacity: Math.round(clamp01(e.color.a ?? 1) * 100) / 100,
    };
  }
  return undefined;
}

/** 그림자 외 지원 불가 효과가 있으면 이름을 돌려준다 (경고용) */
export function unsupportedEffects(
  effects: readonly Effect[] | typeof figma.mixed,
): string[] {
  if (effects === figma.mixed || !Array.isArray(effects)) return [];
  const names: string[] = [];
  for (const e of effects) {
    if (e.visible === false) continue;
    if (e.type === 'LAYER_BLUR') names.push(t().effectLayerBlur);
    else if (e.type === 'BACKGROUND_BLUR') names.push(t().effectBackgroundBlur);
  }
  return names;
}
