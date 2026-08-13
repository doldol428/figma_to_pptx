import type { Hex, Warning } from './ir';

/**
 * 가져오기(PPTX → Figma) 중간 표현.
 *
 * 내보내기 IR 과 일부러 분리했다. 방향이 반대일 뿐 아니라 다루는 모델의 크기가 다르다.
 * 내보내기는 Figma 의 좁은 모델에서 출발해 필요한 것만 만들면 되지만, 가져오기는
 * PPTX 가 표현할 수 있는 것을 받아내야 한다 — 그라디언트, 표, 연결선, 187종 preset 도형.
 *
 * 좌표 단위는 **pt** (EMU / 12700). Figma 는 1px = 1pt 규약을 쓰므로 그대로 px 가 된다.
 * 파싱은 UI 스레드(JSZip · XML)에서, 노드 생성은 main 스레드에서 한다.
 */

export interface SolidPaint {
  kind: 'solid';
  color: Hex;
  /** 0.0 - 1.0 */
  opacity: number;
}

export interface GradientStop {
  /** 0.0 - 1.0 */
  position: number;
  color: Hex;
  opacity: number;
}

/** PPTX 의 gradFill. Figma 는 그라디언트를 지원하므로 근사 없이 그대로 옮긴다. */
export interface GradientPaint {
  kind: 'gradient';
  type: 'linear' | 'radial';
  /** degrees, 0 = 왼→오른쪽 */
  angle: number;
  stops: GradientStop[];
}

export type Paint = SolidPaint | GradientPaint;

export interface StrokeSpec {
  paint: Paint;
  /** pt */
  width: number;
  dash?: number[];
  cap?: 'NONE' | 'ROUND' | 'SQUARE';
}

export interface ShadowSpec {
  offsetX: number;
  offsetY: number;
  blur: number;
  color: Hex;
  opacity: number;
}

/** 회전·반전이 제거된 배치 정보. 좌표는 슬라이드 좌상단 기준 pt. */
export interface Placement {
  x: number;
  y: number;
  w: number;
  h: number;
  /** degrees, 반시계 양수 (Figma 규약). PPTX 의 시계 양수와 부호가 반대다. */
  rotation: number;
  flipH: boolean;
  flipV: boolean;
}

interface NodeBase {
  name: string;
  place: Placement;
  /** 0.0 - 1.0 */
  opacity: number;
}

/**
 * 도형. 네이티브로 만들 수 있는 것은 종류를 밝히고, 나머지는 경로로 넘긴다.
 * 사각형·원은 Figma 원시 노드로 만들어야 나중에 편집이 자연스럽다.
 */
export type ShapeGeometry =
  | { kind: 'rect' }
  /** 모서리별 반경 (좌상, 우상, 우하, 좌하) */
  | { kind: 'roundRect'; radii: [number, number, number, number] }
  | { kind: 'ellipse' }
  | { kind: 'line' }
  /** SVG path — preset 도형·custGeom·연결선이 모두 여기로 온다 */
  | { kind: 'path'; data: string; evenOdd: boolean };

export interface ShapeNode extends NodeBase {
  type: 'shape';
  geometry: ShapeGeometry;
  fill?: Paint;
  stroke?: StrokeSpec;
  shadow?: ShadowSpec;
}

export interface TextRun {
  text: string;
  /** pt */
  size: number;
  fontFamily: string;
  /** Figma FontName.style 후보. 없으면 Regular 로 떨어진다. */
  fontStyle: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  color: Hex;
  opacity: number;
  /** pt */
  letterSpacing?: number;
  link?: string;
}

export interface Paragraph {
  runs: TextRun[];
  align: 'LEFT' | 'CENTER' | 'RIGHT' | 'JUSTIFIED';
  /** 배수 행간. PPTX lnSpc spcPct. */
  lineHeightPct?: number;
  /** pt, 문단 앞 여백 */
  spaceBefore?: number;
  spaceAfter?: number;
  /** 들여쓰기 단계 */
  level: number;
  bullet?: { kind: 'char'; char: string } | { kind: 'number' };
}

export interface TextNodeSpec extends NodeBase {
  type: 'text';
  paragraphs: Paragraph[];
  vertical: 'TOP' | 'CENTER' | 'BOTTOM';
  /** 텍스트 상자 내부 여백 (pt) */
  insets: { left: number; top: number; right: number; bottom: number };
  autoWidth: boolean;
}

export interface ImageNodeSpec extends NodeBase {
  type: 'image';
  /** base64 (data URI 접두사 없음) */
  data: string;
  mime: string;
  /** SVG 는 벡터로 심을 수 있어 따로 표시한다 */
  isSvg: boolean;
  /** 도형에 이미지 채우기로 들어간 경우의 모서리 반경 */
  radii?: [number, number, number, number];
}

export interface TableCell {
  /** 병합으로 가려진 칸 */
  merged: boolean;
  rowSpan: number;
  colSpan: number;
  fill?: Paint;
  paragraphs: Paragraph[];
  insets: { left: number; top: number; right: number; bottom: number };
  vertical: 'TOP' | 'CENTER' | 'BOTTOM';
  borders: {
    top?: StrokeSpec;
    right?: StrokeSpec;
    bottom?: StrokeSpec;
    left?: StrokeSpec;
  };
}

/** PPTX 표. Figma 에는 표 원시 타입이 없어 프레임 격자로 재구성한다. */
export interface TableNodeSpec extends NodeBase {
  type: 'table';
  /** pt */
  colWidths: number[];
  rowHeights: number[];
  rows: TableCell[][];
}

export interface GroupNodeSpec extends NodeBase {
  type: 'group';
  children: ImportNode[];
}

export type ImportNode =
  | ShapeNode
  | TextNodeSpec
  | ImageNodeSpec
  | TableNodeSpec
  | GroupNodeSpec;

export interface ImportSlide {
  name: string;
  fill?: Paint;
  nodes: ImportNode[];
}

export interface ImportDoc {
  /** 슬라이드 크기 (pt) */
  widthPt: number;
  heightPt: number;
  fileName: string;
  slides: ImportSlide[];
  warnings: Warning[];
  /** 파일에서 쓰인 폰트 이름 — main 스레드가 로드 가능 여부를 확인한다 */
  fonts: string[];
}
