import { t } from './i18n';

/**
 * 단위 규약 — 이 플러그인의 존재 이유.
 *
 * Figma 1px = 1pt (72dpi) 로 **고정** 해석한다.
 * 따라서 A4 세로 문서는 Figma 에서 595.28 x 841.89 px 프레임이며,
 * 결과 PPTX 슬라이드는 정확히 210 x 297 mm 가 된다.
 *
 * 기존 도구들이 프레임을 16:9 슬라이드에 욱여넣으면서 비율을 뭉개는 문제를
 * 여기서 원천 차단한다. 슬라이드 크기는 언제나 프레임 크기에서 유도된다.
 *
 * 예외는 presets.ts 의 표준 슬라이드 크기뿐인데, 그것도 프레임 비율과 일치할 때만
 * 제시되고 적용될 때는 단일 균등 배율만 쓴다. 종횡비는 어느 경로로도 변하지 않는다.
 */

export const PT_PER_INCH = 72;
export const EMU_PER_INCH = 914400;
export const EMU_PER_PT = EMU_PER_INCH / PT_PER_INCH; // 12700
export const MM_PER_INCH = 25.4;

/** Figma px → pt (1:1 고정) */
export function pxToPt(px: number): number {
  return px;
}

/** Figma px → inch (PptxGenJS 의 기본 좌표 단위) */
export function pxToIn(px: number): number {
  return pxToPt(px) / PT_PER_INCH;
}

/** Figma px → mm (UI 표시용) */
export function pxToMm(px: number): number {
  return (pxToPt(px) / PT_PER_INCH) * MM_PER_INCH;
}

/** pt → inch (PptxGenJS 좌표 단위) */
export function ptToIn(pt: number): number {
  return pt / PT_PER_INCH;
}

/** pt → mm (UI 표시용) */
export function ptToMm(pt: number): number {
  return (pt / PT_PER_INCH) * MM_PER_INCH;
}

/** PowerPoint 슬라이드 한 변의 허용 범위 (inch). 벗어나면 PowerPoint 가 파일을 거부한다. */
export const MIN_SLIDE_IN = 1;
export const MAX_SLIDE_IN = 56;

/** Figma exportAsync 의 SCALE 제약 한계. UI 내보내기와 동일하게 4배까지만 건다. */
const MIN_EXPORT_SCALE = 0.05;
const MAX_EXPORT_SCALE = 4;

/**
 * 목표 DPI → Figma 노드 렌더 배율.
 *
 * 노드 렌더 배율을 사용자에게 직접 받으면 안 된다. 같은 "2x" 라도 슬라이드 배율(ptPerPx)에
 * 따라 실효 해상도가 몇 배씩 달라지기 때문이다.
 *
 *   폭 W px 인 노드를 배율 S 로 렌더 → S·W 픽셀
 *   슬라이드에서의 물리 폭          → W·ptPerPx / 72 인치
 *   실효 해상도                     → 72·S / ptPerPx  DPI
 *
 * 그래서 DPI 를 고정값으로 받고 배율을 역산한다. 어떤 프리셋을 골라도 결과 해상도가 같다.
 */
export function exportScaleForDpi(
  dpi: number,
  ptPerPx: number,
): { scale: number; clamped: boolean; actualDpi: number } {
  const ideal = (dpi * ptPerPx) / PT_PER_INCH;
  const scale = Math.min(MAX_EXPORT_SCALE, Math.max(MIN_EXPORT_SCALE, ideal));
  const clamped = Math.abs(scale - ideal) > 1e-6;
  return {
    scale: Math.round(scale * 1000) / 1000,
    clamped,
    actualDpi: Math.round((scale * PT_PER_INCH) / ptPerPx),
  };
}

interface Paper {
  name: string;
  /** 세로 방향 기준 mm */
  wMm: number;
  hMm: number;
}

const PAPERS: Paper[] = [
  { name: 'A0', wMm: 841, hMm: 1189 },
  { name: 'A1', wMm: 594, hMm: 841 },
  { name: 'A2', wMm: 420, hMm: 594 },
  { name: 'A3', wMm: 297, hMm: 420 },
  { name: 'A4', wMm: 210, hMm: 297 },
  { name: 'A5', wMm: 148, hMm: 210 },
  { name: 'A6', wMm: 105, hMm: 148 },
  { name: 'B4(ISO)', wMm: 250, hMm: 353 },
  { name: 'B5(ISO)', wMm: 176, hMm: 250 },
  { name: 'Letter', wMm: 215.9, hMm: 279.4 },
  { name: 'Legal', wMm: 215.9, hMm: 355.6 },
  { name: 'Tabloid', wMm: 279.4, hMm: 431.8 },
];

/**
 * 프레임 크기가 알려진 용지 규격에 해당하면 라벨을 돌려준다. (예: "A4 세로")
 * 순수 표시용 — 변환 계산에는 전혀 관여하지 않는다.
 */
export function detectPaper(wPx: number, hPx: number): string | null {
  const wMm = pxToMm(wPx);
  const hMm = pxToMm(hPx);
  const tol = 1.5; // mm
  for (const p of PAPERS) {
    if (Math.abs(wMm - p.wMm) <= tol && Math.abs(hMm - p.hMm) <= tol) {
      return t().paperChip(p.name, true);
    }
    if (Math.abs(wMm - p.hMm) <= tol && Math.abs(hMm - p.wMm) <= tol) {
      return t().paperChip(p.name, false);
    }
  }
  return detectRatio(wPx / hPx);
}

/** 용지 규격에 안 맞을 때 쓰는 비율 칩. 배율을 걸어도 비율은 변하지 않으므로 항상 유효하다. */
const RATIOS: Array<[string, number]> = [
  ['16:9', 16 / 9], ['9:16', 9 / 16],
  ['4:3', 4 / 3], ['3:4', 3 / 4],
  ['16:10', 16 / 10], ['10:16', 10 / 16],
  ['3:2', 3 / 2], ['2:3', 2 / 3],
  ['1:1', 1],
];

function detectRatio(ratio: number): string | null {
  for (const [label, target] of RATIOS) {
    if (Math.abs(ratio - target) / target <= 0.005) return label;
  }
  return null;
}
