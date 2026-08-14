/**
 * 경로가 실제로 차지하는 범위.
 *
 * preset 경로는 **이름상의 상자**(도형의 w×h) 안에서 만들어지지만, 그 상자를 다 채우지는 않는다.
 * 90° 원호는 13×13 상자 안에서 오른쪽 위 사분면만 쓴다. Figma 벡터 노드의 크기는 경로가 정하므로,
 * 이름상의 상자로 변환을 걸면 도형이 그만큼 밀리고 좌우 반전 축도 함께 어긋난다 —
 * 가로 브라켓의 모서리가 그래서 제자리를 벗어나 있었다.
 *
 * 그래서 경로를 원점으로 옮겨 노드의 로컬 상자와 경로를 일치시키고, 얼마나 옮겼는지를
 * 배치 계산에 되돌려 준다.
 */

export interface PathBox {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** 절대 좌표 경로를 평행이동한다. 모든 명령의 인자가 (x, y) 쌍이라 번갈아 더하면 된다. */
export function translatePath(data: string, dx: number, dy: number): string {
  let i = 0;
  return data.replace(/-?\d*\.?\d+/g, (token) => {
    const v = Number(token) + (i++ % 2 === 0 ? dx : dy);
    return (Math.round(v * 100) / 100).toString();
  });
}

const bezier = (p0: number, p1: number, p2: number, p3: number, t: number): number => {
  const u = 1 - t;
  return u * u * u * p0 + 3 * u * u * t * p1 + 3 * u * t * t * p2 + t * t * t * p3;
};

/**
 * 3차 베지어가 방향을 바꾸는 지점 (0 < t < 1).
 * 끝점만 보면 바깥으로 부푼 곡선의 범위를 놓친다 — 원호가 정확히 그런 모양이다.
 */
function extrema(p0: number, p1: number, p2: number, p3: number): number[] {
  const a = p3 - 3 * p2 + 3 * p1 - p0;
  const b = p2 - 2 * p1 + p0;
  const c = p1 - p0;
  const out: number[] = [];
  const push = (t: number): void => {
    if (t > 0 && t < 1) out.push(t);
  };

  if (Math.abs(a) < 1e-12) {
    if (Math.abs(b) > 1e-12) push(-c / (2 * b));
    return out;
  }
  const disc = b * b - a * c;
  if (disc < 0) return out;
  const root = Math.sqrt(disc);
  push((-b + root) / a);
  push((-b - root) / a);
  return out;
}

/** M · L · C · Z 로 이루어진 절대 좌표 경로의 정확한 경계. */
export function pathBounds(data: string): PathBox {
  const tokens = data.match(/[A-Za-z]|-?\d*\.?\d+/g) ?? [];
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  const hit = (x: number, y: number): void => {
    if (x < minX) minX = x;
    if (y < minY) minY = y;
    if (x > maxX) maxX = x;
    if (y > maxY) maxY = y;
  };

  let cx = 0;
  let cy = 0;
  let startX = 0;
  let startY = 0;
  let i = 0;
  const next = (): number => Number(tokens[i++]);
  const isNum = (): boolean => i < tokens.length && !/[A-Za-z]/.test(tokens[i]);

  while (i < tokens.length) {
    const cmd = tokens[i++];
    if (cmd === 'M' || cmd === 'm') {
      cx = next(); cy = next();
      startX = cx; startY = cy;
      hit(cx, cy);
    } else if (cmd === 'L' || cmd === 'l') {
      cx = next(); cy = next();
      hit(cx, cy);
    } else if (cmd === 'C' || cmd === 'c') {
      const x1 = next(); const y1 = next();
      const x2 = next(); const y2 = next();
      const x = next(); const y = next();
      hit(x, y);
      for (const t of extrema(cx, x1, x2, x)) hit(bezier(cx, x1, x2, x, t), bezier(cy, y1, y2, y, t));
      for (const t of extrema(cy, y1, y2, y)) hit(bezier(cx, x1, x2, x, t), bezier(cy, y1, y2, y, t));
      cx = x; cy = y;
    } else if (cmd === 'Z' || cmd === 'z') {
      cx = startX; cy = startY;
    } else {
      // 모르는 명령은 좌표 쌍으로만 훑는다. 범위를 놓치지 않는 쪽으로 넉넉하게.
      while (isNum()) {
        const x = next();
        if (!isNum()) break;
        hit(x, next());
      }
    }
  }

  if (!Number.isFinite(minX)) return { x: 0, y: 0, w: 0.01, h: 0.01 };
  return {
    x: minX,
    y: minY,
    w: Math.max(0.01, maxX - minX),
    h: Math.max(0.01, maxY - minY),
  };
}
