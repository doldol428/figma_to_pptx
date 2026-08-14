import type { Hex, Pattern, Warning } from './ir';

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
  /**
   * 이 단색이 무늬(`pattFill`)를 뭉갠 값이면 원본이 여기 남는다.
   * Figma 에 무늬가 없어 단색으로 눌러 두지만, 노드에 적어 두면 내보낼 때 되살릴 수 있다.
   */
  pattern?: Pattern;
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
  /**
   * 장마다 값이 달라 레이아웃 컴포넌트에 넣을 수 없는 노드 (슬라이드 번호).
   * 레이아웃에서 왔더라도 이것만은 슬라이드에 따로 놓는다.
   */
  perSlide?: boolean;
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
  /** 같은 글꼴의 다른 이름 (한글/영문). 첫 이름이 설치돼 있지 않을 때 차례로 시도한다. */
  fontAlternates?: string[];
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
  shadow?: ShadowSpec;
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
  /**
   * 도형 하나에 글자가 얹힌 것을 갈라 담은 그룹.
   * PPTX 는 이것을 `<p:sp>` 하나로 쓴다 — 표시해 두어야 내보낼 때 다시 하나로 합칠 수 있다.
   */
  shapeText?: boolean;
}

export type ImportNode =
  | ShapeNode
  | TextNodeSpec
  | ImageNodeSpec
  | TableNodeSpec
  | GroupNodeSpec;

/**
 * 여러 슬라이드가 공유하는 페이지 공통 요소. PPTX 의 slideLayout 파트 하나에 해당한다.
 * Figma 에서는 컴포넌트 하나가 되고, 각 슬라이드에는 인스턴스만 놓인다.
 */
export interface ImportLayout {
  /** 레이아웃 파트 파일명 — 슬라이드가 이 값으로 자기 레이아웃을 가리킨다 */
  key: string;
  name: string;
  nodes: ImportNode[];
}

export interface ImportSlide {
  name: string;
  fill?: Paint;
  /** 이 장이 쓰는 레이아웃. 없으면 공통 요소가 없는 장이다. */
  layoutKey?: string;
  /** 레이아웃에서 왔지만 장마다 달라 공유할 수 없는 것 (슬라이드 번호) */
  perSlideNodes: ImportNode[];
  nodes: ImportNode[];
}

export interface ImportDoc {
  /** 슬라이드 크기 (pt) */
  widthPt: number;
  heightPt: number;
  fileName: string;
  /** 쓰인 레이아웃들 — 슬라이드보다 먼저 만들어야 인스턴스를 놓을 수 있다 */
  layouts: ImportLayout[];
  slides: ImportSlide[];
  warnings: Warning[];
  /** 파일에서 쓰인 폰트 이름 — main 스레드가 로드 가능 여부를 확인한다 */
  fonts: string[];
  /**
   * 문서에 포함된 글꼴의 "한글 이름 → 영문 이름" 표.
   * Figma 는 영문 이름으로만 글꼴을 등록하므로 이게 있으면 이름 맞추기가 정확해진다.
   */
  fontAliases: Record<string, string>;
}
