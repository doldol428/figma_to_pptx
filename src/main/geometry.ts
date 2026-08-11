import type { Box } from '../shared/ir';

/** Figma 의 2x3 아핀 행렬. [[a, c, tx], [b, d, ty]] */
export type Mat = [[number, number, number], [number, number, number]];

export function toMat(t: Transform): Mat {
  return [
    [t[0][0], t[0][1], t[0][2]],
    [t[1][0], t[1][1], t[1][2]],
  ];
}

export function mul(m: Mat, n: Mat): Mat {
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

export function invert(m: Mat): Mat {
  const [[a, c, tx], [b, d, ty]] = m;
  const det = a * d - b * c;
  if (Math.abs(det) < 1e-12) {
    // 특이 행렬 — 실무에서 나올 일이 거의 없지만 터지진 않게 항등행렬로 대체
    return [[1, 0, 0], [0, 1, 0]];
  }
  const ia = d / det;
  const ic = -c / det;
  const ib = -b / det;
  const id = a / det;
  return [
    [ia, ic, -(ia * tx + ic * ty)],
    [ib, id, -(ib * tx + id * ty)],
  ];
}

/** 프레임 로컬 공간 기준의 노드 변환행렬 */
export function relativeTo(frame: SceneNode, node: SceneNode): Mat {
  return mul(invert(toMat(frame.absoluteTransform)), toMat(node.absoluteTransform));
}

/**
 * 아핀 행렬 + 로컬 크기 → PPTX 배치값.
 *
 * PowerPoint 의 `<a:off>` 는 **회전 전** 좌상단이고 회전은 도형 중심을 축으로 한다.
 * 반면 Figma 행렬은 로컬 (0,0) 을 축으로 회전이 이미 반영돼 있다.
 * 그래서 중심점을 행렬로 구한 뒤 w/2, h/2 를 빼서 회전 전 좌상단을 복원한다.
 *
 * 부호: Figma 회전은 반시계 양수, PPTX 는 시계 양수라 뒤집는다.
 */
export function place(m: Mat, w: number, h: number): Box {
  const [[a, c, tx], [b, d, ty]] = m;

  const cx = a * (w / 2) + c * (h / 2) + tx;
  const cy = b * (w / 2) + d * (h / 2) + ty;

  // det < 0 이면 반사(뒤집기)가 섞여 있다. x축 반전으로 분리한 뒤 회전각을 잰다.
  const det = a * d - b * c;
  const flipH = det < 0;
  const s = flipH ? -1 : 1;
  const rotCcw = (Math.atan2(-b * s, a * s) * 180) / Math.PI;

  return {
    x: cx - w / 2,
    y: cy - h / 2,
    w,
    h,
    rot: round(-rotCcw),
    flipH,
    flipV: false,
  };
}

function round(n: number): number {
  const r = Math.round(n * 1000) / 1000;
  return Object.is(r, -0) ? 0 : r;
}

/** stroke 정렬 보정 — PPT 선은 항상 윤곽선 중앙에 그려진다. */
export function insetBox(box: Box, amount: number): Box {
  return {
    ...box,
    x: box.x + amount,
    y: box.y + amount,
    w: Math.max(0, box.w - amount * 2),
    h: Math.max(0, box.h - amount * 2),
  };
}
