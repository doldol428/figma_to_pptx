import type { Placement } from '../shared/importir';

/**
 * PPTX 좌표계 → Figma 배치.
 *
 * 두 가지를 동시에 풀어야 한다.
 *
 * 1) 그룹은 자식 좌표계(chOff/chExt)를 실제 배치(off/ext)로 매핑한다. 이 배율·이동을
 *    빼먹으면 그룹 안의 모든 요소가 어긋난다 — 실측 파일에 그룹이 1,149개 들어 있다.
 * 2) 회전은 도형 중심 기준이고 PPTX 는 시계 양수, Figma 는 반시계 양수다.
 *
 * 그래서 중첩 구조를 따라 아핀 행렬을 누적한 뒤, 도형마다 한 번 분해해서 배치를 얻는다.
 */

/** [[a, c, tx], [b, d, ty]] */
export type Mat = [[number, number, number], [number, number, number]];

export const IDENTITY: Mat = [[1, 0, 0], [0, 1, 0]];

export function multiply(m: Mat, n: Mat): Mat {
  return [
    [
      m[0][0] * n[0][0] + m[0][1] * n[1][0],
      m[0][0] * n[0][1] + m[0][1] * n[1][1],
      m[0][0] * n[0][2] + m[0][1] * n[1][2] + m[0][2],
    ],
    [
      m[1][0] * n[0][0] + m[1][1] * n[1][0],
      m[1][0] * n[0][1] + m[1][1] * n[1][1],
      m[1][0] * n[0][2] + m[1][1] * n[1][2] + m[1][2],
    ],
  ];
}

export function translate(tx: number, ty: number): Mat {
  return [[1, 0, tx], [0, 1, ty]];
}

export function scale(sx: number, sy: number): Mat {
  return [[sx, 0, 0], [0, sy, 0]];
}

/** PPTX 회전 (도, 시계 양수). 화면 좌표계가 y 아래라 부호가 그대로 들어간다. */
export function rotate(deg: number): Mat {
  const r = (deg * Math.PI) / 180;
  const c = Math.cos(r);
  const s = Math.sin(r);
  return [[c, -s, 0], [s, c, 0]];
}

/**
 * 로컬 (0,0,w,h) 사각형을 놓는 행렬을 만든다.
 * 회전과 반전은 상자 중심을 축으로 한다 — OOXML `<a:xfrm rot flipH flipV>` 의 정의.
 */
export function placeBox(
  x: number, y: number, w: number, h: number,
  rotDeg: number, flipH: boolean, flipV: boolean,
): Mat {
  const cx = x + w / 2;
  const cy = y + h / 2;
  let m = translate(cx, cy);
  if (rotDeg) m = multiply(m, rotate(rotDeg));
  if (flipH || flipV) m = multiply(m, scale(flipH ? -1 : 1, flipV ? -1 : 1));
  return multiply(m, translate(-w / 2, -h / 2));
}

/**
 * 누적 행렬 + 로컬 크기 → Figma 배치.
 *
 * Figma 도 회전을 중심 기준으로 다루므로, 중심을 행렬로 구한 뒤 w/2, h/2 를 빼서
 * 회전 전 좌상단을 복원한다. 반전이 섞이면(det < 0) x축 반전으로 분리한 뒤 각을 잰다.
 */
export function decompose(m: Mat, w: number, h: number): Placement {
  const [[a, c, tx], [b, d, ty]] = m;

  const cx = a * (w / 2) + c * (h / 2) + tx;
  const cy = b * (w / 2) + d * (h / 2) + ty;

  // 배율이 섞여 있으면 크기에 흡수시킨다 (그룹의 chExt→ext 배율이 여기로 들어온다).
  const sx = Math.hypot(a, b) || 1;
  const sy = Math.hypot(c, d) || 1;
  const det = a * d - b * c;
  const flipH = det < 0;

  const sign = flipH ? -1 : 1;
  // PPTX 시계 양수 → Figma 반시계 양수
  const rotation = -(Math.atan2((b * sign) / sx, (a * sign) / sx) * 180) / Math.PI;

  const outW = w * sx;
  const outH = h * sy;

  return {
    x: round(cx - outW / 2),
    y: round(cy - outH / 2),
    w: round(Math.max(0, outW)),
    h: round(Math.max(0, outH)),
    rotation: round(rotation),
    flipH,
    flipV: false,
  };
}

function round(n: number): number {
  const r = Math.round(n * 1000) / 1000;
  return Object.is(r, -0) ? 0 : r;
}
