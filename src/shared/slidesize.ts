import { MAX_SLIDE_IN, MIN_SLIDE_IN, PT_PER_INCH, detectPaper } from './units';

/**
 * 프레임 크기 → 슬라이드 크기 결정.
 *
 * 기본은 실측(1px = 1pt)이다. A4 처럼 물리 크기가 의미를 갖는 문서는 이 경로로 나간다.
 *
 * 다만 프레젠테이션용 프레임은 보통 1920×1080 처럼 화면 해상도로 만든다.
 * 실측을 그대로 적용하면 67.7×38.1cm 짜리 슬라이드가 나오는데, 비율은 맞지만
 * PowerPoint 표준이 아니다. 그래서 **프레임 비율이 표준 슬라이드 비율과 일치하면**
 * 표준 크기로 자동 환산한다. 사용자에게 물어볼 문제가 아니다 — 비율이 맞으면 표준이 정답이다.
 *
 * 어느 경로든 적용되는 건 단일 균등 배율 하나뿐이라 종횡비는 변하지 않는다.
 */

interface StandardSize {
  /** 가로 방향 기준 (pt) */
  wPt: number;
  hPt: number;
}

/** PowerPoint 의 "슬라이드 크기" 기본 제공 항목 */
const STANDARDS: StandardSize[] = [
  { wPt: 960, hPt: 540 }, // 와이드스크린 13.333 × 7.5 in
  { wPt: 720, hPt: 540 }, // 표준 4:3      10 × 7.5 in
  { wPt: 720, hPt: 450 }, // 16:10         10 × 6.25 in
];

export interface SlidePlan {
  /** 확정된 슬라이드 크기 (pt) */
  wPt: number;
  hPt: number;
  /** Figma px → pt 균등 배율. 실측이면 1. */
  ptPerPx: number;
  /** 중앙 정렬 보정 (pt). 비율이 정확히 일치하면 0. */
  offsetXPt: number;
  offsetYPt: number;
  /** 크기 옆에 붙는 짧은 칩 — "16:9", "A4 세로" 등 */
  chip: string | null;
  /** 프레임 크기를 그대로 쓰는지 */
  native: boolean;
}

/** 비율 일치 허용 오차 (상대값). 1920×1080 같은 정수 해상도를 넉넉히 포괄한다. */
const RATIO_TOL = 0.005;

const EXACT = 0.01;

function fits(wPt: number, hPt: number): boolean {
  const w = wPt / PT_PER_INCH;
  const h = hPt / PT_PER_INCH;
  return w >= MIN_SLIDE_IN && h >= MIN_SLIDE_IN && w <= MAX_SLIDE_IN && h <= MAX_SLIDE_IN;
}

/**
 * 비율이 맞는 표준 크기를 찾는다. 프레임이 이미 그 크기면(환산할 게 없으면) null.
 * 가로/세로 두 방향을 모두 본다.
 */
function matchStandard(wPx: number, hPx: number): StandardSize | null {
  const ratio = wPx / hPx;
  for (const s of STANDARDS) {
    for (const v of [{ w: s.wPt, h: s.hPt }, { w: s.hPt, h: s.wPt }]) {
      const target = v.w / v.h;
      if (Math.abs(ratio - target) / target > RATIO_TOL) continue;
      if (Math.abs(v.w - wPx) < EXACT && Math.abs(v.h - hPx) < EXACT) return null;
      if (!fits(v.w, v.h)) continue;
      return { wPt: v.w, hPt: v.h };
    }
  }
  return null;
}

/**
 * 슬라이드 크기를 확정한다. PowerPoint 가 받아주지 못하는 크기뿐이면 null.
 *
 * 칩 문구는 비율에서 나온다. 배율을 걸어도 비율은 변하지 않으므로
 * 표준 환산 여부와 무관하게 프레임 크기 하나로 판정하면 된다.
 */
export function resolveSlide(wPx: number, hPx: number): SlidePlan | null {
  if (wPx <= 0 || hPx <= 0) return null;

  const chip = detectPaper(wPx, hPx);
  const standard = matchStandard(wPx, hPx);

  if (standard) {
    const ptPerPx = Math.min(standard.wPt / wPx, standard.hPt / hPx);
    return {
      wPt: standard.wPt,
      hPt: standard.hPt,
      ptPerPx,
      // 비율이 미세하게 다른 경우(1000×562 등)에만 0 이 아니다. 늘리는 대신 여백으로 흡수한다.
      offsetXPt: (standard.wPt - wPx * ptPerPx) / 2,
      offsetYPt: (standard.hPt - hPx * ptPerPx) / 2,
      chip,
      native: false,
    };
  }

  if (!fits(wPx, hPx)) return null;

  return {
    wPt: wPx,
    hPt: hPx,
    ptPerPx: 1,
    offsetXPt: 0,
    offsetYPt: 0,
    chip,
    native: true,
  };
}
