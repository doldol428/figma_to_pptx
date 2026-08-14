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

import type { Locale } from './i18n';
import type { ImportDoc } from './importir';

/** "RRGGBB" (샵 없음) */
export type Hex = string;

export interface Warning {
  slide: string;
  node: string;
  message: string;
}

export interface GradientStop {
  /** 0 ~ 1 */
  position: number;
  color: Hex;
  /** 0(불투명) ~ 100(투명) */
  transparency: number;
}

/**
 * 그라디언트 — **단색 위에 얹히는 덤**이다.
 *
 * PptxGenJS 4.x 는 칠을 'solid' | 'none' 으로만 받는다. 그래서 색은 늘 평균 단색으로 넘기고,
 * 이 값은 파일을 굳히기 직전 XML 에 `<a:gradFill>` 로 덧써 넣는다(ui/gradient.ts).
 * 덧쓰기가 실패해도 평균 단색이 남으므로, 이 필드가 무시되면 딱 예전 동작이 된다.
 */
export interface Gradient {
  type: 'linear' | 'radial';
  /** 도. 화면 기준 시계방향 — PPTX `<a:lin ang>` 과 같은 규약 */
  angle: number;
  stops: GradientStop[];
}

/**
 * 무늬 채우기 — Figma 에 대응 개념이 **없다**.
 *
 * 가져올 때 밀도만큼 섞은 단색으로 눌러 두고, 원본은 노드에 적어 둔다. Figma 에서 무늬를
 * 편집할 방법 자체가 없으므로 그 값이 낡을 일이 없다 — 확인 없이 되돌려도 안전한 드문 경우다.
 * 다만 **색은 바꿀 수 있으니**, 눌러 둔 단색이 그대로일 때만 되돌린다.
 */
export interface Pattern {
  /** DrawingML 무늬 이름 (`ltUpDiag`, `pct25` …) */
  preset: string;
  fg: Hex;
  fgTransparency: number;
  bg: Hex;
  bgTransparency: number;
  /** 가져올 때 눌러 둔 단색. 지금 색이 이것과 다르면 사람이 손댄 것이라 되돌리지 않는다. */
  approximated: Hex;
}

export interface SolidFill {
  kind: 'solid';
  color: Hex;
  /** 0(불투명) ~ 100(투명) */
  transparency: number;
  /** 이 단색이 그라디언트의 평균값이면 원본이 여기 남는다 */
  gradient?: Gradient;
  /** 이 단색이 무늬를 뭉갠 값이면 원본이 여기 남는다 */
  pattern?: Pattern;
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
  /** 이 단색이 그라디언트의 평균값이면 원본이 여기 남는다 */
  gradient?: Gradient;
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
  /** 글자 전체에 걸린 그라디언트. 런의 색은 그 평균값이다. */
  gradient?: Gradient;
  /**
   * 이 글자를 담는 도형.
   *
   * PPTX 는 "도형 안의 글자" 를 `<p:sp>` 하나로 쓴다 — 도형과 글자가 한 몸이다.
   * Figma 에는 그 개념이 없어 가져올 때 둘로 갈리는데, 여기 담아 두면 다시 하나로 나간다.
   */
  shape?: { geom: Geom; fill: Fill; stroke?: Stroke };
}

export interface ImageItem {
  type: 'image';
  name: string;
  box: Box;
  /**
   * Figma 의 imageHash — 이미지 **내용**의 주소다.
   * 같은 바이트를 몇 번 넣든 같은 해시가 나오므로, 같은 그림인지 여기서 바로 알 수 있다.
   * 노드를 다시 렌더해서 뽑은 경우에는 원본 그림이 아니라 없다.
   */
  key?: string;
  /** base64 (data URI 접두사 없음) */
  data: string;
  mime: 'image/png' | 'image/jpeg' | 'image/gif';
  /** Figma scaleMode 대응. 'stretch' 는 박스에 그대로 늘린다. */
  sizing: 'cover' | 'contain' | 'stretch';
  shadow?: Shadow;
}

export type Item = ShapeItem | TextItem | ImageItem;

export interface Slide {
  name: string;
  fill: Fill;
  items: Item[];
  /** 이 장이 쓰는 공통 서식 (Master.name). 없으면 독립 슬라이드. */
  master?: string;
}

/**
 * 여러 장이 공유하는 공통 서식 — PPTX 의 slideLayout 파트가 된다.
 * `#레이아웃/…` 컴포넌트의 인스턴스를 놓은 슬라이드들이 이것을 가리킨다.
 */
export interface Master {
  name: string;
  items: Item[];
  /** 진짜 `<a:fld type="slidenum">` 로 나갈 자리. `#슬라이드번호` 에서 온다. */
  slideNumber?: {
    box: Box;
    fontFace: string;
    /** px */
    fontSize: number;
    color: Hex;
    align: TextItem['align'];
  };
}

/**
 * 이 항목을 공통 서식에 넣을 수 있는가.
 *
 * PptxGenJS 의 마스터는 항목 하나에 문자열 하나만 실을 수 있다 (`[{ text: <값> }]` 로 감싼다).
 * 서식이 여러 개인 텍스트를 넘기면 `[object Object]` 로 찍혀 나온다 — 실제로 확인했다.
 * 도형·이미지·단일 서식 텍스트는 전부 나간다 (custGeom 포함).
 */
export function fitsInMaster(item: Item): boolean {
  return item.type !== 'text' || item.runs.length <= 1;
}

/**
 * 슬라이드를 뺀 문서 정보. 내보내기는 이것을 먼저 보내고 슬라이드를 한 장씩 뒤따라 보낸다 —
 * 이미지가 base64 로 부푼 문서를 통째로 넘기면 다리가 막힌다 (가져오기와 같은 이유).
 */
export interface DocMeta {
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
  /** "16:9", "A4 세로" 같은 짧은 규격 칩 */
  chip: string | null;
  /** 공통 서식. 슬라이드보다 먼저 정의해야 한다. */
  masters: Master[];
}

export interface Doc extends DocMeta {
  slides: Slide[];
  warnings: Warning[];
}

/* ── main ↔ ui 메시지 ─────────────────────────────────────────── */

/**
 * 플러그인 창 높이 상한.
 * UI 는 이 값을 넘을 때만 스크롤을 허용한다 — 그 아래에서는 창이 내용에 맞춰지므로
 * 스크롤바가 필요 없다. main 의 resize 와 UI 의 판단이 어긋나면 안 되므로 한 곳에 둔다.
 */
export const UI_MAX_HEIGHT = 720;

export interface SelectionState {
  /** export 가능 여부 — 프레임이 1개 이상이고 크기가 전부 같을 때만 true */
  ok: boolean;
  /** ok=false 일 때 사용자에게 보여줄 사유 */
  reason: string;
  frameCount: number;
  widthPx: number;
  heightPx: number;
  /** 짧은 규격 칩 ("16:9", "A4 세로"). 해당 없으면 null */
  chip: string | null;
  /** 자동 확정된 슬라이드 크기 (pt) — 사용자가 고르지 않는다 */
  slideWPt: number;
  slideHPt: number;
  /** 적용될 균등 배율. 1 이면 실측. */
  ptPerPx: number;
}

export type MainToUi =
  | { type: 'selection'; state: SelectionState }
  | { type: 'progress'; done: number; total: number }
  /*
   * 내보내기도 가져오기처럼 한 장씩 흘려보낸다.
   * 문서를 통째로 넘기면 이미지가 base64 로 부푼 만큼 한 번에 실려 장수가 늘수록 터진다.
   */
  | { type: 'exportBegin'; meta: DocMeta; total: number; fileName: string; warnings: Warning[] }
  | { type: 'exportSlide'; index: number; slide: Slide; warnings: Warning[] }
  /**
   * 가져오기 준비 완료. UI 는 이 신호를 받고서야 첫 슬라이드를 보낸다.
   * Figma 는 async 핸들러를 기다려주지 않으므로, 확인 없이 연달아 보내면
   * 세션이 만들어지기 전에 도착해 전부 버려진다.
   */
  | { type: 'importReady' }
  /** 슬라이드 한 장 처리 완료 — UI 는 이걸 받고 다음 장을 보낸다 */
  | { type: 'createProgress'; done: number; total: number }
  | { type: 'imported'; slides: number; missingFonts: string[]; failures: string[] }
  | { type: 'error'; message: string };

export type UiToMain =
  /** locale 은 UI 만 알 수 있어(navigator.languages) 첫 메시지에 실어 보낸다 */
  | { type: 'ready'; locale: Locale }
  | { type: 'export'; imageDpi: number }
  /** 다음 장을 달라 — 한 장 받아 붙인 뒤에 보낸다 (앞질러 보내면 쌓인다) */
  | { type: 'exportNext'; index: number }
  /*
   * 가져오기는 슬라이드를 한 장씩 보낸다.
   * 100MB 짜리 문서는 이미지가 base64 로 부풀어 한 번에 넘기면 postMessage 가 감당하지 못한다.
   */
  | {
    type: 'importBegin';
    widthPt: number;
    heightPt: number;
    fileName: string;
    total: number;
    fonts: string[];
    fontAliases: Record<string, string>;
  }
  /** 레이아웃 컴포넌트를 먼저 만든다 — 슬라이드가 인스턴스를 놓으려면 이미 있어야 한다 */
  | { type: 'importLayout'; layout: ImportDoc['layouts'][number] }
  | { type: 'importSlide'; index: number; slide: ImportDoc['slides'][number] }
  | { type: 'importEnd' }
  /** 내용 높이에 맞춰 플러그인 창을 줄인다 — 경고 목록 유무에 따라 높이가 크게 달라진다 */
  | { type: 'resize'; height: number }
  | { type: 'notify'; message: string; error?: boolean };
