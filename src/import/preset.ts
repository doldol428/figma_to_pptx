import { num, type XNode } from './xml';

/**
 * prstGeom → 경로.
 *
 * OOXML 의 preset 도형은 187종이고, 각각 조절점(adjust value)에 따라 모양이 달라지는
 * 파라메트릭 정의를 갖는다. 전부 옮기는 것은 현실적이지 않으므로 실제 문서에 나오는 것부터
 * 채운다 — 실측 제안서 58장에서 쓰인 종류가 여기 전부 들어 있다.
 *
 * rect · roundRect · ellipse · line 은 여기서 경로를 만들지 않는다.
 * Figma 원시 노드로 만들어야 나중에 크기 조절과 편집이 자연스럽기 때문이다 (`nativeKind`).
 *
 * Figma 의 vectorPaths 는 SVG 경로 문법을 받지만 호(A) 명령은 쓰지 않는다.
 * 원호가 필요한 도형은 3차 베지어로 근사한다.
 */

/** Figma 원시 노드로 만들 수 있는 preset */
export type NativeKind =
  | { kind: 'rect' }
  | { kind: 'roundRect'; radii: [number, number, number, number] }
  | { kind: 'ellipse' }
  | { kind: 'line' };

export interface PresetPath {
  data: string;
  evenOdd: boolean;
}

/** avLst 의 조절값을 이름 → 값(1/100000) 으로 읽는다 */
export function readAdjust(prstGeom: XNode | null): Record<string, number> {
  const out: Record<string, number> = {};
  const avLst = prstGeom?.children.find((c) => c.tag === 'avLst');
  for (const gd of avLst?.children ?? []) {
    if (gd.tag !== 'gd') continue;
    const name = gd.attrs.name;
    const fmla = gd.attrs.fmla ?? '';
    const m = /^val\s+(-?\d+)/.exec(fmla);
    if (name && m) out[name] = num(m[1]);
  }
  return out;
}

const p = (n: number): string => (Math.round(n * 100) / 100).toString();
const M = (x: number, y: number): string => `M${p(x)} ${p(y)} `;
const L = (x: number, y: number): string => `L${p(x)} ${p(y)} `;
const C = (
  x1: number, y1: number, x2: number, y2: number, x: number, y: number,
): string => `C${p(x1)} ${p(y1)} ${p(x2)} ${p(y2)} ${p(x)} ${p(y)} `;
const Z = 'Z ';

/** 타원 호를 3차 베지어로. 각도는 라디안, y 아래 방향 기준. */
function arc(
  cx: number, cy: number, rx: number, ry: number, from: number, to: number, move: boolean,
): string {
  const sweep = to - from;
  const steps = Math.max(1, Math.ceil(Math.abs(sweep) / (Math.PI / 2)));
  const step = sweep / steps;
  const k = (4 / 3) * Math.tan(step / 4);
  let out = '';
  let t = from;
  if (move) out += M(cx + rx * Math.cos(t), cy + ry * Math.sin(t));
  for (let i = 0; i < steps; i++) {
    const t2 = t + step;
    const x0 = cx + rx * Math.cos(t);
    const y0 = cy + ry * Math.sin(t);
    const x1 = cx + rx * Math.cos(t2);
    const y1 = cy + ry * Math.sin(t2);
    out += C(
      x0 - k * rx * Math.sin(t), y0 + k * ry * Math.cos(t),
      x1 + k * rx * Math.sin(t2), y1 - k * ry * Math.cos(t2),
      x1, y1,
    );
    t = t2;
  }
  return out;
}

/** 정n각형 (위쪽 꼭짓점부터 시계방향) */
function polygon(w: number, h: number, sides: number, rotate = -Math.PI / 2): string {
  let d = '';
  for (let i = 0; i < sides; i++) {
    const a = rotate + (i * 2 * Math.PI) / sides;
    const x = w / 2 + (w / 2) * Math.cos(a);
    const y = h / 2 + (h / 2) * Math.sin(a);
    d += i === 0 ? M(x, y) : L(x, y);
  }
  return d + Z;
}

function star(w: number, h: number, points: number, inner: number): string {
  let d = '';
  for (let i = 0; i < points * 2; i++) {
    const r = i % 2 === 0 ? 1 : inner;
    const a = -Math.PI / 2 + (i * Math.PI) / points;
    const x = w / 2 + (w / 2) * r * Math.cos(a);
    const y = h / 2 + (h / 2) * r * Math.sin(a);
    d += i === 0 ? M(x, y) : L(x, y);
  }
  return d + Z;
}

/** 모서리가 잘린/둥근 사각형을 한 번에 그린다. r>0 둥글게, r<0 잘라내기. */
function cornerRect(w: number, h: number, c: [number, number, number, number]): string {
  const [tl, tr, br, bl] = c;
  const A = Math.abs;
  const K = 0.5523;
  let d = M(A(tl), 0);
  d += L(w - A(tr), 0);
  if (tr > 0) d += C(w - tr + tr * K, 0, w, tr - tr * K, w, tr);
  else if (tr < 0) d += L(w, A(tr));
  d += L(w, h - A(br));
  if (br > 0) d += C(w, h - br + br * K, w - br + br * K, h, w - br, h);
  else if (br < 0) d += L(w - A(br), h);
  d += L(A(bl), h);
  if (bl > 0) d += C(bl - bl * K, h, 0, h - bl + bl * K, 0, h - bl);
  else if (bl < 0) d += L(0, h - A(bl));
  d += L(0, A(tl));
  if (tl > 0) d += C(0, tl - tl * K, tl - tl * K, 0, tl, 0);
  else if (tl < 0) d += L(A(tl), 0);
  return d + Z;
}

/**
 * Figma 원시 노드로 만들 수 있으면 그 종류를, 아니면 null 을 돌려준다.
 * 반환된 반경은 pt 단위 실제 값이다.
 */
export function nativeFor(
  prst: string, w: number, h: number, adj: Record<string, number>,
): NativeKind | null {
  const mn = Math.min(w, h);
  const a = (name: string, def: number): number => (adj[name] ?? def) / 100000;

  switch (prst) {
    case 'rect':
    case 'flowChartProcess':
    case 'flowChartPredefinedProcess':
      return { kind: 'rect' };
    case 'ellipse':
    case 'flowChartConnector':
      return { kind: 'ellipse' };
    case 'line':
    case 'straightConnector1':
      return { kind: 'line' };
    case 'roundRect':
    case 'flowChartAlternateProcess': {
      const r = mn * a('adj', 16667);
      return { kind: 'roundRect', radii: [r, r, r, r] };
    }
    case 'round1Rect': {
      const r = mn * a('adj', 16667);
      return { kind: 'roundRect', radii: [0, r, 0, 0] };
    }
    case 'round2SameRect': {
      const r1 = mn * a('adj1', 16667);
      const r2 = mn * a('adj2', 0);
      return { kind: 'roundRect', radii: [r1, r1, r2, r2] };
    }
    case 'round2DiagRect': {
      const r1 = mn * a('adj1', 16667);
      const r2 = mn * a('adj2', 0);
      return { kind: 'roundRect', radii: [r1, r2, r1, r2] };
    }
    default:
      return null;
  }
}

/**
 * 경로로 그려야 하는 preset. 모르는 종류면 null — 호출부가 사각형으로 대체하고 경고한다.
 */
export function presetPath(
  prst: string, w: number, h: number, adj: Record<string, number>,
): PresetPath | null {
  const mn = Math.min(w, h);
  const a = (name: string, def: number): number => (adj[name] ?? def) / 100000;
  const solid = (data: string): PresetPath => ({ data, evenOdd: false });
  const hollow = (data: string): PresetPath => ({ data, evenOdd: true });

  switch (prst) {
    /* ── 잘린 모서리 ─────────────────────────────────────────── */
    case 'snip1Rect':
      return solid(cornerRect(w, h, [0, -mn * a('adj', 16667), 0, 0]));
    case 'snip2SameRect': {
      const s1 = -mn * a('adj1', 16667);
      const s2 = -mn * a('adj2', 0);
      return solid(cornerRect(w, h, [s1, s1, s2, s2]));
    }
    case 'snip2DiagRect': {
      const s1 = -mn * a('adj1', 0);
      const s2 = -mn * a('adj2', 16667);
      return solid(cornerRect(w, h, [s1, s2, s1, s2]));
    }
    case 'snipRoundRect': {
      const s = -mn * a('adj1', 16667);
      const r = mn * a('adj2', 16667);
      return solid(cornerRect(w, h, [s, 0, 0, r]));
    }

    /* ── 다각형 ──────────────────────────────────────────────── */
    case 'triangle': {
      const t = w * a('adj', 50000);
      return solid(M(t, 0) + L(w, h) + L(0, h) + Z);
    }
    case 'rtTriangle':
      return solid(M(0, 0) + L(w, h) + L(0, h) + Z);
    case 'diamond':
    case 'flowChartDecision':
      return solid(M(w / 2, 0) + L(w, h / 2) + L(w / 2, h) + L(0, h / 2) + Z);
    case 'parallelogram':
    case 'flowChartInputOutput': {
      const t = w * a('adj', 25000);
      return solid(M(t, 0) + L(w, 0) + L(w - t, h) + L(0, h) + Z);
    }
    case 'trapezoid': {
      const t = w * a('adj', 25000);
      return solid(M(t, 0) + L(w - t, 0) + L(w, h) + L(0, h) + Z);
    }
    case 'pentagon':
      return solid(polygon(w, h, 5));
    case 'hexagon': {
      // 좌우 꼭짓점이 안쪽으로 들어간 가로형 육각형 — PowerPoint 기본 모양
      const t = Math.min(w / 2, mn * a('adj', 25000));
      return solid(M(t, 0) + L(w - t, 0) + L(w, h / 2) + L(w - t, h) + L(t, h) + L(0, h / 2) + Z);
    }
    case 'heptagon':
      return solid(polygon(w, h, 7));
    case 'octagon':
      return solid(polygon(w, h, 8, -Math.PI / 2 + Math.PI / 8));
    case 'decagon':
      return solid(polygon(w, h, 10));
    case 'dodecagon':
      return solid(polygon(w, h, 12));
    case 'star4':
      return solid(star(w, h, 4, a('adj', 12500) * 8));
    case 'star5':
      return solid(star(w, h, 5, 0.38));
    case 'star6':
      return solid(star(w, h, 6, 0.577));
    case 'star8':
      return solid(star(w, h, 8, 0.7));
    case 'star10':
      return solid(star(w, h, 10, 0.75));
    case 'star12':
      return solid(star(w, h, 12, 0.79));

    /* ── 화살표 ──────────────────────────────────────────────── */
    case 'rightArrow': {
      const t = h * a('adj1', 50000);
      const head = Math.min(w, w * a('adj2', 50000));
      const y0 = (h - t) / 2;
      return solid(M(0, y0) + L(w - head, y0) + L(w - head, 0) + L(w, h / 2)
        + L(w - head, h) + L(w - head, y0 + t) + L(0, y0 + t) + Z);
    }
    case 'leftArrow': {
      const t = h * a('adj1', 50000);
      const head = Math.min(w, w * a('adj2', 50000));
      const y0 = (h - t) / 2;
      return solid(M(w, y0) + L(head, y0) + L(head, 0) + L(0, h / 2)
        + L(head, h) + L(head, y0 + t) + L(w, y0 + t) + Z);
    }
    case 'upArrow': {
      const t = w * a('adj1', 50000);
      const head = Math.min(h, h * a('adj2', 50000));
      const x0 = (w - t) / 2;
      return solid(M(x0, h) + L(x0, head) + L(0, head) + L(w / 2, 0)
        + L(w, head) + L(x0 + t, head) + L(x0 + t, h) + Z);
    }
    case 'downArrow': {
      const t = w * a('adj1', 50000);
      const head = Math.min(h, h * a('adj2', 50000));
      const x0 = (w - t) / 2;
      return solid(M(x0, 0) + L(x0, h - head) + L(0, h - head) + L(w / 2, h)
        + L(w, h - head) + L(x0 + t, h - head) + L(x0 + t, 0) + Z);
    }
    case 'leftRightArrow': {
      const t = h * a('adj1', 50000);
      const head = Math.min(w / 2, w * a('adj2', 25000));
      const y0 = (h - t) / 2;
      return solid(M(0, h / 2) + L(head, 0) + L(head, y0) + L(w - head, y0) + L(w - head, 0)
        + L(w, h / 2) + L(w - head, h) + L(w - head, y0 + t) + L(head, y0 + t) + L(head, h) + Z);
    }
    case 'chevron':
    case 'homePlate': {
      const t = Math.min(w, w * a('adj', 50000)) * (prst === 'chevron' ? 1 : 1);
      if (prst === 'homePlate') {
        return solid(M(0, 0) + L(w - t, 0) + L(w, h / 2) + L(w - t, h) + L(0, h) + Z);
      }
      return solid(M(0, 0) + L(w - t, 0) + L(w, h / 2) + L(w - t, h) + L(0, h) + L(t, h / 2) + Z);
    }

    /* ── 원형 파생 ───────────────────────────────────────────── */
    case 'donut': {
      const t = mn * a('adj', 25000);
      const outer = arc(w / 2, h / 2, w / 2, h / 2, 0, Math.PI * 2, true);
      const inner = arc(w / 2, h / 2, w / 2 - t, h / 2 - t, Math.PI * 2, 0, true);
      return hollow(`${outer}${Z}${inner}${Z}`);
    }
    case 'pie':
    case 'arc':
    case 'chord': {
      const start = (adj.adj1 ?? 0) / 60000 * (Math.PI / 180);
      const end = (adj.adj2 ?? 16200000) / 60000 * (Math.PI / 180);
      const sweep = arc(w / 2, h / 2, w / 2, h / 2, start, end, true);
      if (prst === 'arc') return { data: sweep, evenOdd: false };
      if (prst === 'chord') return solid(sweep + Z);
      return solid(sweep + L(w / 2, h / 2) + Z);
    }
    case 'blockArc': {
      const start = (adj.adj1 ?? 10800000) / 60000 * (Math.PI / 180);
      const end = (adj.adj2 ?? 0) / 60000 * (Math.PI / 180);
      const t = mn * a('adj3', 25000);
      return solid(arc(w / 2, h / 2, w / 2, h / 2, start, end, true)
        + arc(w / 2, h / 2, w / 2 - t, h / 2 - t, end, start, false) + Z);
    }
    case 'teardrop': {
      const t = mn * a('adj', 100000) / 2;
      return solid(arc(w / 2, h / 2, w / 2, h / 2, -Math.PI / 2, Math.PI * 1.5, true).replace(/^M[^C]*/, `M${p(w / 2)} 0 `)
        + L(w / 2 + t, h / 2 - t) + Z);
    }
    case 'moon': {
      const t = w * a('adj', 50000);
      return solid(M(w, 0) + C(w - t * 1.33, 0, w - t * 1.33, h, w, h)
        + C(w - t * 2.66, h, w - t * 2.66, 0, w, 0) + Z);
    }

    /* ── 테두리·프레임 ───────────────────────────────────────── */
    case 'frame': {
      const t = mn * a('adj1', 12500);
      return hollow(M(0, 0) + L(w, 0) + L(w, h) + L(0, h) + Z
        + M(t, t) + L(t, h - t) + L(w - t, h - t) + L(w - t, t) + Z);
    }
    case 'halfFrame': {
      const t1 = mn * a('adj1', 33333);
      const t2 = mn * a('adj2', 33333);
      return solid(M(0, 0) + L(w, 0) + L(w - t1, t2) + L(t1, t2) + L(t1, h) + L(0, h) + Z);
    }
    case 'corner': {
      const t1 = h * a('adj1', 50000);
      const t2 = w * a('adj2', 50000);
      return solid(M(0, 0) + L(t2, 0) + L(t2, h - t1) + L(w, h - t1) + L(w, h) + L(0, h) + Z);
    }
    case 'plaque': {
      const t = mn * a('adj', 16667);
      const K = 0.5523;
      return solid(M(t, 0) + L(w - t, 0) + C(w - t + t * K, 0, w, t - t * K, w, t)
        + L(w, h - t) + C(w, h - t + t * K, w - t + t * K, h, w - t, h)
        + L(t, h) + C(t - t * K, h, 0, h - t + t * K, 0, h - t)
        + L(0, t) + C(0, t - t * K, t - t * K, 0, t, 0) + Z);
    }
    case 'bevel':
    case 'plus': {
      if (prst === 'bevel') {
        const t = mn * a('adj', 12500);
        return hollow(M(0, 0) + L(w, 0) + L(w, h) + L(0, h) + Z
          + M(t, t) + L(t, h - t) + L(w - t, h - t) + L(w - t, t) + Z);
      }
      const t = mn * a('adj', 25000);
      return solid(M(t, 0) + L(w - t, 0) + L(w - t, t) + L(w, t) + L(w, h - t) + L(w - t, h - t)
        + L(w - t, h) + L(t, h) + L(t, h - t) + L(0, h - t) + L(0, t) + L(t, t) + Z);
    }

    /* ── 순서도 ──────────────────────────────────────────────── */
    case 'flowChartTerminator': {
      const r = h / 2;
      return solid(M(r, 0) + L(w - r, 0) + arc(w - r, h / 2, r, r, -Math.PI / 2, Math.PI / 2, false)
        + L(r, h) + arc(r, h / 2, r, r, Math.PI / 2, Math.PI * 1.5, false) + Z);
    }
    case 'flowChartDocument': {
      const t = h * 0.17;
      return solid(M(0, 0) + L(w, 0) + L(w, h - t)
        + C(w * 0.75, h - t * 2.2, w * 0.25, h + t * 0.6, 0, h - t) + Z);
    }
    case 'flowChartMagneticDisk': {
      const ry = h * 0.17;
      return solid(M(0, ry) + arc(w / 2, ry, w / 2, ry, Math.PI, Math.PI * 2, false)
        + L(w, h - ry) + arc(w / 2, h - ry, w / 2, ry, 0, Math.PI, false) + Z);
    }
    case 'flowChartPreparation': {
      const t = w * 0.2;
      return solid(M(t, 0) + L(w - t, 0) + L(w, h / 2) + L(w - t, h) + L(t, h) + L(0, h / 2) + Z);
    }
    case 'flowChartManualOperation': {
      const t = w * 0.2;
      return solid(M(0, 0) + L(w, 0) + L(w - t, h) + L(t, h) + Z);
    }

    /* ── 괄호·중괄호 (선으로만 그려지는 열린 경로) ───────────── */
    case 'leftBracket': {
      const r = Math.min(w, h / 2) * a('adj', 8333) * 2;
      return { data: M(w, 0) + L(r, 0) + C(0, 0, 0, 0, 0, r) + L(0, h - r)
        + C(0, h, 0, h, r, h) + L(w, h), evenOdd: false };
    }
    case 'rightBracket': {
      const r = Math.min(w, h / 2) * a('adj', 8333) * 2;
      return { data: M(0, 0) + L(w - r, 0) + C(w, 0, w, 0, w, r) + L(w, h - r)
        + C(w, h, w, h, w - r, h) + L(0, h), evenOdd: false };
    }
    case 'bracketPair': {
      const r = Math.min(w / 2, h / 2) * a('adj', 16667) * 2;
      const left = M(r, 0) + C(0, 0, 0, 0, 0, r) + L(0, h - r) + C(0, h, 0, h, r, h);
      const right = M(w - r, 0) + C(w, 0, w, 0, w, r) + L(w, h - r) + C(w, h, w, h, w - r, h);
      return { data: left + right, evenOdd: false };
    }
    case 'leftBrace': {
      const t = h * a('adj2', 50000);
      const r = Math.min(w, h / 4);
      return { data: M(w, 0) + C(w / 2, 0, w / 2, 0, w / 2, r)
        + L(w / 2, t - r) + C(w / 2, t, w / 2, t, 0, t)
        + C(w / 2, t, w / 2, t, w / 2, t + r)
        + L(w / 2, h - r) + C(w / 2, h, w / 2, h, w, h), evenOdd: false };
    }
    case 'rightBrace': {
      const t = h * a('adj2', 50000);
      const r = Math.min(w, h / 4);
      return { data: M(0, 0) + C(w / 2, 0, w / 2, 0, w / 2, r)
        + L(w / 2, t - r) + C(w / 2, t, w / 2, t, w, t)
        + C(w / 2, t, w / 2, t, w / 2, t + r)
        + L(w / 2, h - r) + C(w / 2, h, w / 2, h, 0, h), evenOdd: false };
    }
    case 'bracePair': {
      const r = Math.min(w / 4, h / 4);
      const mid = h / 2;
      const left = M(w / 4, 0) + C(w / 8, 0, w / 8, 0, w / 8, r)
        + L(w / 8, mid - r) + C(w / 8, mid, w / 8, mid, 0, mid)
        + C(w / 8, mid, w / 8, mid, w / 8, mid + r)
        + L(w / 8, h - r) + C(w / 8, h, w / 8, h, w / 4, h);
      const right = M((w * 3) / 4, 0) + C(w - w / 8, 0, w - w / 8, 0, w - w / 8, r)
        + L(w - w / 8, mid - r) + C(w - w / 8, mid, w - w / 8, mid, w, mid)
        + C(w - w / 8, mid, w - w / 8, mid, w - w / 8, mid + r)
        + L(w - w / 8, h - r) + C(w - w / 8, h, w - w / 8, h, (w * 3) / 4, h);
      return { data: left + right, evenOdd: false };
    }

    /* ── 입체 ────────────────────────────────────────────────── */
    case 'cube': {
      const t = mn * a('adj', 25000);
      // 앞면 · 윗면 · 오른쪽면을 한 경로에 담는다. 면 사이 경계는 선으로 남는다.
      return {
        data: M(0, t) + L(t, 0) + L(w, 0) + L(w, h - t) + L(w - t, h) + L(0, h) + Z
          + M(0, t) + L(w - t, t) + L(w, 0)
          + M(w - t, t) + L(w - t, h),
        evenOdd: false,
      };
    }
    case 'can': {
      const ry = Math.min(h / 2, h * a('adj', 25000) / 2);
      return {
        data: M(0, ry) + arc(w / 2, ry, w / 2, ry, Math.PI, Math.PI * 2, false)
          + L(w, h - ry) + arc(w / 2, h - ry, w / 2, ry, 0, Math.PI, false) + Z
          + arc(w / 2, ry, w / 2, ry, 0, Math.PI * 2, true) + Z,
        evenOdd: false,
      };
    }

    /* ── 상하 화살표 · 굽은 화살표 ───────────────────────────── */
    case 'upDownArrow': {
      const t = w * a('adj1', 50000);
      const head = Math.min(h / 2, h * a('adj2', 25000));
      const x0 = (w - t) / 2;
      return solid(M(w / 2, 0) + L(w, head) + L(x0 + t, head) + L(x0 + t, h - head) + L(w, h - head)
        + L(w / 2, h) + L(0, h - head) + L(x0, h - head) + L(x0, head) + L(0, head) + Z);
    }
    case 'curvedRightArrow':
    case 'curvedUpArrow':
    case 'curvedDownArrow':
    case 'curvedLeftArrow': {
      // 굽은 화살표는 정확한 곡률보다 방향이 읽히는 게 중요하다. 사분원 + 화살촉으로 근사한다.
      const t = Math.min(w, h) * 0.25;
      if (prst === 'curvedRightArrow' || prst === 'curvedLeftArrow') {
        const dir = prst === 'curvedRightArrow' ? 1 : -1;
        const x0 = dir > 0 ? 0 : w;
        const x1 = dir > 0 ? w : 0;
        return solid(M(x0, 0) + C(x1, 0, x1, h, x0, h) + L(x0, h - t)
          + C(x1 - dir * t, h - t, x1 - dir * t, t, x0, t) + Z);
      }
      const dir = prst === 'curvedUpArrow' ? -1 : 1;
      const y0 = dir > 0 ? 0 : h;
      const y1 = dir > 0 ? h : 0;
      return solid(M(0, y0) + C(0, y1, w, y1, w, y0) + L(w - t, y0)
        + C(w - t, y1 - dir * t, t, y1 - dir * t, t, y0) + Z);
    }

    /* ── 수식 기호 ───────────────────────────────────────────── */
    case 'mathPlus': {
      const t = mn * a('adj1', 23520);
      const y0 = (h - t) / 2;
      const x0 = (w - t) / 2;
      return solid(M(x0, 0) + L(x0 + t, 0) + L(x0 + t, y0) + L(w, y0) + L(w, y0 + t)
        + L(x0 + t, y0 + t) + L(x0 + t, h) + L(x0, h) + L(x0, y0 + t) + L(0, y0 + t)
        + L(0, y0) + L(x0, y0) + Z);
    }
    case 'mathMinus': {
      const t = h * a('adj1', 23520);
      const y0 = (h - t) / 2;
      return solid(M(0, y0) + L(w, y0) + L(w, y0 + t) + L(0, y0 + t) + Z);
    }
    case 'mathEqual': {
      const t = h * a('adj1', 23520);
      const gap = h * a('adj2', 11760);
      const y0 = h / 2 - gap / 2 - t;
      const y1 = h / 2 + gap / 2;
      return solid(M(0, y0) + L(w, y0) + L(w, y0 + t) + L(0, y0 + t) + Z
        + M(0, y1) + L(w, y1) + L(w, y1 + t) + L(0, y1 + t) + Z);
    }
    case 'mathMultiply': {
      const t = mn * a('adj1', 23520) * 0.7;
      const d = Math.atan2(h, w);
      const dx = (t / 2) * Math.cos(d + Math.PI / 2);
      const dy = (t / 2) * Math.sin(d + Math.PI / 2);
      return solid(M(dx, dy) + L(w + dx, h + dy) + L(w - dx, h - dy) + L(-dx, -dy) + Z
        + M(w - dx, dy) + L(-dx, h + dy) + L(dx, h - dy) + L(w + dx, -dy) + Z);
    }

    /* ── 말풍선 ──────────────────────────────────────────────── */
    case 'wedgeEllipseCallout': {
      const tx = w / 2 + w * a('adj1', -20833);
      const ty = h / 2 + h * a('adj2', 62500);
      return solid(arc(w / 2, h / 2, w / 2, h / 2, 0, Math.PI * 2, true) + Z
        + M(w * 0.4, h * 0.93) + L(tx, ty) + L(w * 0.6, h * 0.98) + Z);
    }
    case 'wedgeRectCallout': {
      const tx = w / 2 + w * a('adj1', -20833);
      const ty = h / 2 + h * a('adj2', 62500);
      return solid(M(0, 0) + L(w, 0) + L(w, h) + L(w * 0.35, h)
        + L(tx, ty) + L(w * 0.25, h) + L(0, h) + Z);
    }

    default:
      return null;
  }
}

/**
 * 꺾인 연결선. PowerPoint 는 두 도형 사이의 경로를 계산해 두지만 파일에는 남기지 않으므로
 * 도형 상자만 보고 다시 만든다. 실제 파일에서 800개 넘게 나오는 요소라 빼놓을 수 없다.
 */
export function connectorPath(prst: string, w: number, h: number): PresetPath | null {
  const straight = { data: M(0, 0) + L(w, h), evenOdd: false };
  switch (prst) {
    case 'straightConnector1':
      return straight;
    case 'bentConnector2':
      return { data: M(0, 0) + L(w, 0) + L(w, h), evenOdd: false };
    case 'bentConnector3':
      return { data: M(0, 0) + L(w / 2, 0) + L(w / 2, h) + L(w, h), evenOdd: false };
    case 'bentConnector4':
      return { data: M(0, 0) + L(w / 2, 0) + L(w / 2, h / 2) + L(w, h / 2) + L(w, h), evenOdd: false };
    case 'bentConnector5':
      return {
        data: M(0, 0) + L(w / 4, 0) + L(w / 4, h / 2) + L((w * 3) / 4, h / 2)
          + L((w * 3) / 4, h) + L(w, h),
        evenOdd: false,
      };
    case 'curvedConnector2':
      return { data: M(0, 0) + C(w / 2, 0, w, h / 2, w, h), evenOdd: false };
    case 'curvedConnector3':
      return {
        data: M(0, 0) + C(w / 4, 0, w / 2, 0, w / 2, h / 2) + C(w / 2, h, (w * 3) / 4, h, w, h),
        evenOdd: false,
      };
    case 'curvedConnector4':
    case 'curvedConnector5':
      return { data: M(0, 0) + C(w / 2, 0, w / 2, h, w, h), evenOdd: false };
    default:
      return null;
  }
}
