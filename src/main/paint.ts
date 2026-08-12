import { t } from '../shared/i18n';
import type { Fill, Hex, Shadow, Stroke } from '../shared/ir';

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
 * 그라디언트 → 단색 근사.
 *
 * PptxGenJS 4.x 는 도형 fill 을 'solid' | 'none' 으로만 노출한다.
 * 정지점들을 위치 가중 평균해서 가장 덜 튀는 단색을 만들고, 호출부에서 경고를 남긴다.
 * (실제 gradFill 주입은 pptx 재패키징이 필요 — README 의 다음 단계 참고)
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
    return {
      kind: 'fill',
      fill: {
        kind: 'solid',
        color: avg.color,
        transparency: transparency(avg.opacity * pOpacity * nodeOpacity),
      },
      note: t().gradientApproximated(paint.type, avg.color),
    };
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
  let opacity = paint.opacity ?? 1;
  if (paint.type === 'SOLID') {
    color = toHex(paint.color);
  } else if (paint.type.startsWith('GRADIENT_')) {
    const avg = averageGradient(paint as GradientPaint);
    color = avg.color;
    opacity *= avg.opacity;
  } else {
    return undefined;
  }

  return {
    color,
    transparency: transparency(opacity * nodeOpacity),
    width: weight,
    dashType: dashOf(node.dashPattern),
  };
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
