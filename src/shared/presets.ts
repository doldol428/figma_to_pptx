import { MAX_SLIDE_IN, MIN_SLIDE_IN, PT_PER_INCH } from './units';

/**
 * 슬라이드 크기 프리셋.
 *
 * 기본은 언제나 "프레임 실측"(1px = 1pt) 이다. A4 처럼 물리 크기가 의미를 갖는 문서는
 * 이 경로로만 나가야 한다.
 *
 * 다만 프레젠테이션용 프레임은 보통 1920×1080 처럼 화면 해상도로 만든다.
 * 실측으로 내보내면 67.7×38.1cm 짜리 슬라이드가 나오는데, 비율은 맞지만 PowerPoint 표준이 아니다.
 * 그래서 **프레임 비율이 표준 슬라이드 비율과 일치할 때만** 표준 크기 옵션을 함께 제시한다.
 *
 * 핵심: 가로/세로를 따로 맞추지 않는다. 언제나 단일 균등 배율(min) 을 쓰고
 * 남는 여백은 중앙 정렬로 흡수한다. 비율이 정확히 같으면 여백은 0 이 된다.
 * 어떤 경우에도 종횡비가 변하지 않는다.
 */

export interface StandardSize {
  id: string;
  label: string;
  /** 가로 방향 기준 (pt) */
  wPt: number;
  hPt: number;
}

/** PowerPoint 의 "슬라이드 크기" 기본 제공 항목 */
const STANDARDS: StandardSize[] = [
  // 디자인 > 슬라이드 크기 > 와이드스크린 (13.333 × 7.5 in)
  { id: 'ppt-16x9', label: 'PowerPoint 와이드스크린 16:9', wPt: 960, hPt: 540 },
  // 표준 (10 × 7.5 in)
  { id: 'ppt-4x3', label: 'PowerPoint 표준 4:3', wPt: 720, hPt: 540 },
  // 화면 슬라이드 쇼 16:10 (10 × 6.25 in)
  { id: 'ppt-16x10', label: 'PowerPoint 16:10', wPt: 720, hPt: 450 },
];

export const NATIVE_PRESET_ID = 'native';

export interface PresetOption {
  id: string;
  label: string;
  /** 결과 슬라이드 크기 (pt) */
  wPt: number;
  hPt: number;
  /** Figma px → pt 균등 배율. 실측이면 1. */
  ptPerPx: number;
  /** 중앙 정렬 보정 (pt). 비율이 정확히 일치하면 0. */
  offsetXPt: number;
  offsetYPt: number;
  /** 프레임 크기를 그대로 쓰는 옵션인지 */
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

/** 프레임 크기가 표준 슬라이드 크기와 정확히 같으면 그 이름을 돌려준다. */
export function exactStandardName(wPx: number, hPx: number): string | null {
  for (const s of STANDARDS) {
    if (Math.abs(s.wPt - wPx) < EXACT && Math.abs(s.hPt - hPx) < EXACT) return s.label;
    if (Math.abs(s.hPt - wPx) < EXACT && Math.abs(s.wPt - hPx) < EXACT) return `${s.label} 세로`;
  }
  return null;
}

/**
 * 이 프레임 크기로 만들 수 있는 슬라이드 크기 목록.
 * 항상 실측이 먼저 오고, 비율이 맞는 표준 크기가 뒤따른다.
 * PowerPoint 가 허용하지 않는 크기(1~56인치 밖)는 걸러낸다.
 */
export function presetOptions(wPx: number, hPx: number): PresetOption[] {
  const options: PresetOption[] = [];

  if (wPx <= 0 || hPx <= 0) return options;

  if (fits(wPx, hPx)) {
    options.push({
      id: NATIVE_PRESET_ID,
      label: '프레임 실측 (1px = 1pt)',
      wPt: wPx,
      hPt: hPx,
      ptPerPx: 1,
      offsetXPt: 0,
      offsetYPt: 0,
      native: true,
    });
  }

  const ratio = wPx / hPx;

  for (const s of STANDARDS) {
    const variants: Array<{ w: number; h: number; suffix: string }> = [
      { w: s.wPt, h: s.hPt, suffix: '' },
      { w: s.hPt, h: s.wPt, suffix: ' 세로' },
    ];

    for (const v of variants) {
      const target = v.w / v.h;
      if (Math.abs(ratio - target) / target > RATIO_TOL) continue;
      // 프레임이 이미 그 크기면 실측 항목과 완전히 겹친다.
      if (Math.abs(v.w - wPx) < EXACT && Math.abs(v.h - hPx) < EXACT) continue;
      if (!fits(v.w, v.h)) continue;

      const ptPerPx = Math.min(v.w / wPx, v.h / hPx);
      options.push({
        id: `${s.id}${v.suffix ? '-portrait' : ''}`,
        label: `${s.label}${v.suffix}`,
        wPt: v.w,
        hPt: v.h,
        ptPerPx,
        offsetXPt: (v.w - wPx * ptPerPx) / 2,
        offsetYPt: (v.h - hPx * ptPerPx) / 2,
        native: false,
      });
    }
  }

  return options;
}

/**
 * 기본 선택.
 *
 * 비율이 표준과 맞는데 크기가 다르면(1920×1080 등) 표준 쪽을 고른다 — 화면 해상도로 만든
 * 프레임에서 67cm 짜리 슬라이드가 나오는 게 놀랍기 때문. A4 처럼 표준과 비율이 안 맞는
 * 문서는 매칭되는 표준이 없으므로 자연히 실측이 남는다.
 */
export function defaultPresetId(options: PresetOption[]): string {
  const standard = options.find((o) => !o.native);
  return standard ? standard.id : (options[0]?.id ?? NATIVE_PRESET_ID);
}

export function findPreset(options: PresetOption[], id: string): PresetOption | undefined {
  return options.find((o) => o.id === id);
}
