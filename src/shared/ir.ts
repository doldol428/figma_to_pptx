/**
 * 중간 표현(IR).
 *
 * main 스레드(Figma 샌드박스)가 씬 그래프를 읽어 IR 로 만들고,
 * UI 스레드(iframe)가 IR 을 PptxGenJS 호출로 옮긴다.
 * 두 스레드 사이는 postMessage 로만 통신하므로 IR 은 반드시 순수 JSON 이어야 한다.
 *
 * 좌표·크기·폰트 크기·선 굵기까지 **모든 수치는 Figma px**, 원점은 각 프레임의 좌상단.
 * pt/inch/EMU 변환과 `ptPerPx` 배율 적용은 UI 쪽 빌더(build.ts)에서 한 번만 수행한다.
 */

/** "RRGGBB" (샵 없음) */
export type Hex = string;

export interface Warning {
  slide: string;
  node: string;
  message: string;
}

export interface SolidFill {
  kind: 'solid';
  color: Hex;
  /** 0(불투명) ~ 100(투명) */
  transparency: number;
}

export interface NoFill {
  kind: 'none';
}

export type Fill = SolidFill | NoFill;

export interface Stroke {
  color: Hex;
  transparency: number;
  /** px */
  width: number;
  dashType?: 'solid' | 'dash' | 'sysDot' | 'dashDot';
}

export interface Shadow {
  type: 'outer' | 'inner';
  /** 0-359, 동쪽 기준 시계방향 */
  angle: number;
  /** px */
  offset: number;
  /** px */
  blur: number;
  color: Hex;
  /** 0.0 - 1.0 */
  opacity: number;
}

/**
 * 회전/반전이 제거된 축 정렬 박스 + 회전각.
 * PowerPoint 는 도형 중심을 기준으로 회전하므로 x/y 는 **회전 전** 좌상단이다.
 */
export interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
  /** degrees, 시계방향 양수 (PPTX 규약). Figma 의 반시계 양수와 부호가 반대다. */
  rot: number;
  flipH: boolean;
  flipV: boolean;
}

/** custGeom 경로 점 — 좌표는 도형 로컬(0,0 = 도형 좌상단) px */
export type PathPoint =
  | { x: number; y: number; moveTo?: boolean }
  | { x: number; y: number; c: [number, number, number, number] }
  | { close: true };

export type Geom =
  | { kind: 'rect' }
  | { kind: 'roundRect'; radius: number }
  | { kind: 'ellipse' }
  | { kind: 'line' }
  | { kind: 'custom'; points: PathPoint[] };

export interface ShapeItem {
  type: 'shape';
  name: string;
  box: Box;
  geom: Geom;
  fill: Fill;
  stroke?: Stroke;
  shadow?: Shadow;
}

export interface Run {
  text: string;
  fontFace: string;
  /** px */
  fontSize: number;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  color: Hex;
  transparency: number;
  /** px */
  charSpacing?: number;
  /** px — 고정 행간 */
  lineSpacing?: number;
  /** 배수 행간 */
  lineSpacingMultiple?: number;
  /** true 면 이 런에서 문단이 끊긴다 */
  breakLine?: boolean;
  hyperlink?: string;
}

export interface TextItem {
  type: 'text';
  name: string;
  box: Box;
  runs: Run[];
  align: 'left' | 'center' | 'right' | 'justify';
  valign: 'top' | 'middle' | 'bottom';
  /** 자동 줄바꿈 (Figma textAutoResize 가 WIDTH_AND_HEIGHT 면 false) */
  wrap: boolean;
  shadow?: Shadow;
}

export interface ImageItem {
  type: 'image';
  name: string;
  box: Box;
  /** base64 (data URI 접두사 없음) */
  data: string;
  mime: 'image/png' | 'image/jpeg' | 'image/gif';
  /** Figma scaleMode 대응. 'stretch' 는 박스에 그대로 늘린다. */
  sizing: 'cover' | 'contain' | 'stretch';
}

export type Item = ShapeItem | TextItem | ImageItem;

export interface Slide {
  name: string;
  fill: Fill;
  items: Item[];
}

export interface Doc {
  /** 최종 슬라이드 크기 (pt) */
  slideWPt: number;
  slideHPt: number;
  /** Figma px → pt 균등 배율. 실측이면 1. 모든 길이·폰트 크기에 동일하게 곱한다. */
  ptPerPx: number;
  /** 콘텐츠 중앙 정렬 보정 (pt). 프레임 비율이 슬라이드 비율과 정확히 같으면 0. */
  offsetXPt: number;
  offsetYPt: number;
  /** 원본 프레임 크기 (표시·디버깅용) */
  frameWPx: number;
  frameHPx: number;
  presetLabel: string;
  slides: Slide[];
  warnings: Warning[];
}

/* ── main ↔ ui 메시지 ─────────────────────────────────────────── */

export interface SizeInfo {
  w: number;
  h: number;
  count: number;
}

/** presets.ts 의 PresetOption 과 같은 모양 — 메시지로 넘기기 위한 구조적 사본 */
export interface PresetInfo {
  id: string;
  label: string;
  wPt: number;
  hPt: number;
  ptPerPx: number;
  offsetXPt: number;
  offsetYPt: number;
  native: boolean;
}

export interface SelectionState {
  /** export 가능 여부 — 프레임이 1개 이상이고 크기가 전부 같을 때만 true */
  ok: boolean;
  /** ok=false 일 때 사용자에게 보여줄 사유 */
  reason: string;
  frameCount: number;
  widthPx: number;
  heightPx: number;
  /** 감지된 규격 이름 (A4 세로, PowerPoint 와이드스크린 16:9 …) */
  paper: string | null;
  /** 크기가 섞여 있을 때 각 크기별 개수 */
  sizes: SizeInfo[];
  /** 선택 가능한 슬라이드 크기 — 항상 실측이 먼저, 비율이 맞는 표준이 뒤따른다 */
  presets: PresetInfo[];
  defaultPresetId: string;
}

export type MainToUi =
  | { type: 'selection'; state: SelectionState }
  | { type: 'progress'; done: number; total: number; label: string }
  | { type: 'doc'; doc: Doc; fileName: string }
  | { type: 'error'; message: string };

export type UiToMain =
  | { type: 'ready' }
  | { type: 'export'; imageDpi: number; presetId: string }
  | { type: 'notify'; message: string; error?: boolean };
