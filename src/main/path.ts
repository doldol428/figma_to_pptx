import type { PathPoint } from '../shared/ir';

/**
 * SVG path data → PPTX custGeom 점 목록.
 *
 * Figma `fillGeometry` / `strokeGeometry` 가 돌려주는 경로를 그대로 먹인다.
 * PPTX custGeom 이 아는 세그먼트는 moveTo / lnTo / cubicBezTo / quadBezTo / close 뿐이라
 * 호(A)는 3차 베지어로 근사하고 S·T 는 전개한다. 좌표는 도형 로컬 px 그대로 둔다.
 */

const NUM = /[+-]?(?:\d*\.\d+|\d+\.?)(?:[eE][+-]?\d+)?/g;

function tokenize(d: string): Array<{ cmd: string; args: number[] }> {
  const out: Array<{ cmd: string; args: number[] }> = [];
  const re = /([MmLlHhVvCcSsQqTtAaZz])([^MmLlHhVvCcSsQqTtAaZz]*)/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(d)) !== null) {
    const cmd = m[1];
    const args = (m[2].match(NUM) || []).map(Number);
    out.push({ cmd, args });
  }
  return out;
}

/** 각 커맨드가 한 번에 소비하는 인자 개수 */
const ARITY: Record<string, number> = {
  M: 2, L: 2, H: 1, V: 1, C: 6, S: 4, Q: 4, T: 2, A: 7, Z: 0,
};

export function parsePath(d: string): PathPoint[] {
  const pts: PathPoint[] = [];
  let cx = 0;
  let cy = 0;
  let startX = 0;
  let startY = 0;
  // 직전 3차/2차 제어점 (S/T 반사용)
  let lastC: [number, number] | null = null;
  let lastQ: [number, number] | null = null;

  for (const { cmd, args } of tokenize(d)) {
    const upper = cmd.toUpperCase();
    const rel = cmd !== upper;
    const arity = ARITY[upper];

    if (upper === 'Z') {
      pts.push({ close: true });
      cx = startX;
      cy = startY;
      lastC = lastQ = null;
      continue;
    }

    // 인자가 반복되면 같은 커맨드를 여러 번 적용한다 (M 반복은 L 로 강등되는 SVG 규칙 포함).
    for (let i = 0; i + arity <= args.length; i += arity) {
      const a = args.slice(i, i + arity);
      let eff = upper;
      if (upper === 'M' && i > 0) eff = 'L';

      switch (eff) {
        case 'M': {
          cx = rel ? cx + a[0] : a[0];
          cy = rel ? cy + a[1] : a[1];
          startX = cx;
          startY = cy;
          pts.push({ x: cx, y: cy, moveTo: true });
          lastC = lastQ = null;
          break;
        }
        case 'L': {
          cx = rel ? cx + a[0] : a[0];
          cy = rel ? cy + a[1] : a[1];
          pts.push({ x: cx, y: cy });
          lastC = lastQ = null;
          break;
        }
        case 'H': {
          cx = rel ? cx + a[0] : a[0];
          pts.push({ x: cx, y: cy });
          lastC = lastQ = null;
          break;
        }
        case 'V': {
          cy = rel ? cy + a[0] : a[0];
          pts.push({ x: cx, y: cy });
          lastC = lastQ = null;
          break;
        }
        case 'C': {
          const x1 = rel ? cx + a[0] : a[0];
          const y1 = rel ? cy + a[1] : a[1];
          const x2 = rel ? cx + a[2] : a[2];
          const y2 = rel ? cy + a[3] : a[3];
          cx = rel ? cx + a[4] : a[4];
          cy = rel ? cy + a[5] : a[5];
          pts.push({ x: cx, y: cy, c: [x1, y1, x2, y2] });
          lastC = [x2, y2];
          lastQ = null;
          break;
        }
        case 'S': {
          const [px, py]: [number, number] = lastC ? reflect(cx, cy, lastC) : [cx, cy];
          const x2 = rel ? cx + a[0] : a[0];
          const y2 = rel ? cy + a[1] : a[1];
          cx = rel ? cx + a[2] : a[2];
          cy = rel ? cy + a[3] : a[3];
          pts.push({ x: cx, y: cy, c: [px, py, x2, y2] });
          lastC = [x2, y2];
          lastQ = null;
          break;
        }
        case 'Q': {
          const qx = rel ? cx + a[0] : a[0];
          const qy = rel ? cy + a[1] : a[1];
          const ex = rel ? cx + a[2] : a[2];
          const ey = rel ? cy + a[3] : a[3];
          pts.push(quadToCubic(cx, cy, qx, qy, ex, ey));
          cx = ex;
          cy = ey;
          lastQ = [qx, qy];
          lastC = null;
          break;
        }
        case 'T': {
          // 직전 2차 제어점을 현재 점 기준으로 반사한다.
          // 아래에서 lastQ 에 [qx, qy] 를 되돌려 넣기 때문에 타입 추론이 순환한다 → 명시 annotation 필요.
          const qx: number = lastQ ? 2 * cx - lastQ[0] : cx;
          const qy: number = lastQ ? 2 * cy - lastQ[1] : cy;
          const ex = rel ? cx + a[0] : a[0];
          const ey = rel ? cy + a[1] : a[1];
          pts.push(quadToCubic(cx, cy, qx, qy, ex, ey));
          cx = ex;
          cy = ey;
          lastQ = [qx, qy];
          lastC = null;
          break;
        }
        case 'A': {
          const ex = rel ? cx + a[5] : a[5];
          const ey = rel ? cy + a[6] : a[6];
          for (const seg of arcToCubics(cx, cy, a[0], a[1], a[2], a[3] !== 0, a[4] !== 0, ex, ey)) {
            pts.push(seg);
          }
          cx = ex;
          cy = ey;
          lastC = lastQ = null;
          break;
        }
      }
    }
  }

  return pts;
}

function reflect(cx: number, cy: number, prev: [number, number]): [number, number] {
  return [2 * cx - prev[0], 2 * cy - prev[1]];
}

function quadToCubic(
  x0: number, y0: number, qx: number, qy: number, x: number, y: number,
): PathPoint {
  return {
    x,
    y,
    c: [
      x0 + (2 / 3) * (qx - x0),
      y0 + (2 / 3) * (qy - y0),
      x + (2 / 3) * (qx - x),
      y + (2 / 3) * (qy - y),
    ],
  };
}

/** SVG 타원호 → 3차 베지어 근사 (W3C SVG 부록 F.6 의 endpoint→center 변환) */
function arcToCubics(
  x0: number, y0: number,
  rxIn: number, ryIn: number, xAxisRotDeg: number,
  largeArc: boolean, sweep: boolean,
  x: number, y: number,
): PathPoint[] {
  if (x0 === x && y0 === y) return [];
  let rx = Math.abs(rxIn);
  let ry = Math.abs(ryIn);
  if (rx === 0 || ry === 0) return [{ x, y }];

  const phi = (xAxisRotDeg * Math.PI) / 180;
  const cosP = Math.cos(phi);
  const sinP = Math.sin(phi);

  const dx2 = (x0 - x) / 2;
  const dy2 = (y0 - y) / 2;
  const x1p = cosP * dx2 + sinP * dy2;
  const y1p = -sinP * dx2 + cosP * dy2;

  // 반지름이 부족하면 SVG 규격대로 균등 확대
  const lambda = (x1p * x1p) / (rx * rx) + (y1p * y1p) / (ry * ry);
  if (lambda > 1) {
    const s = Math.sqrt(lambda);
    rx *= s;
    ry *= s;
  }

  const sign = largeArc === sweep ? -1 : 1;
  const num = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p;
  const den = rx * rx * y1p * y1p + ry * ry * x1p * x1p;
  const co = sign * Math.sqrt(Math.max(0, num / den));
  const cxp = (co * rx * y1p) / ry;
  const cyp = (-co * ry * x1p) / rx;

  const cx = cosP * cxp - sinP * cyp + (x0 + x) / 2;
  const cy = sinP * cxp + cosP * cyp + (y0 + y) / 2;

  const theta1 = angle(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry);
  let delta = angle((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry);
  if (!sweep && delta > 0) delta -= 2 * Math.PI;
  if (sweep && delta < 0) delta += 2 * Math.PI;

  const segCount = Math.ceil(Math.abs(delta / (Math.PI / 2)));
  const step = delta / segCount;
  const k = (4 / 3) * Math.tan(step / 4);

  const out: PathPoint[] = [];
  let t = theta1;
  let px = x0;
  let py = y0;
  for (let i = 0; i < segCount; i++) {
    const t2 = t + step;
    const [ex, ey] = onArc(cx, cy, rx, ry, cosP, sinP, t2);
    const [dx1, dy1] = arcDeriv(rx, ry, cosP, sinP, t);
    const [dx3, dy3] = arcDeriv(rx, ry, cosP, sinP, t2);
    out.push({
      x: ex,
      y: ey,
      c: [px + k * dx1, py + k * dy1, ex - k * dx3, ey - k * dy3],
    });
    t = t2;
    px = ex;
    py = ey;
  }
  return out;
}

function onArc(
  cx: number, cy: number, rx: number, ry: number, cosP: number, sinP: number, t: number,
): [number, number] {
  const ct = Math.cos(t);
  const st = Math.sin(t);
  return [cx + rx * ct * cosP - ry * st * sinP, cy + rx * ct * sinP + ry * st * cosP];
}

function arcDeriv(
  rx: number, ry: number, cosP: number, sinP: number, t: number,
): [number, number] {
  const ct = Math.cos(t);
  const st = Math.sin(t);
  return [-rx * st * cosP - ry * ct * sinP, -rx * st * sinP + ry * ct * cosP];
}

function angle(ux: number, uy: number, vx: number, vy: number): number {
  const dot = ux * vx + uy * vy;
  const len = Math.sqrt((ux * ux + uy * uy) * (vx * vx + vy * vy));
  let a = Math.acos(Math.min(1, Math.max(-1, dot / len)));
  if (ux * vy - uy * vx < 0) a = -a;
  return a;
}
