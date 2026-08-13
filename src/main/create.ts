import type {
  GradientPaint, ImageNodeSpec, ImportNode, ImportSlide, Paint, Placement,
  ShapeNode, StrokeSpec, TableNodeSpec, TextNodeSpec,
} from '../shared/importir';

/**
 * ImportDoc → Figma 노드.
 *
 * 슬라이드 하나가 프레임 하나가 되고, 캔버스에 가로로 늘어놓는다.
 * 좌표는 이미 pt 로 정규화돼 있고 Figma 는 1px = 1pt 규약을 쓰므로 그대로 넣는다.
 */

/** 슬라이드 사이 간격 (px) */
const GAP = 80;

export interface ImportMeta {
  widthPt: number;
  heightPt: number;
  fileName: string;
  total: number;
  fonts: string[];
}

/**
 * 가져오기 세션.
 *
 * 슬라이드가 한 장씩 도착하므로 상태를 들고 있어야 한다. 폰트는 시작할 때 한 번만 훑는다 —
 * 슬라이드마다 loadFontAsync 를 반복하면 58장짜리 문서에서 눈에 띄게 느려진다.
 */
export class ImportSession {
  private readonly fonts = new FontBook();
  private readonly origin = nextFreeSpot();
  readonly frames: FrameNode[] = [];

  private constructor(private readonly meta: ImportMeta) {}

  get total(): number {
    return this.meta.total;
  }

  static async begin(meta: ImportMeta): Promise<ImportSession> {
    const session = new ImportSession(meta);
    await session.fonts.load(meta.fonts);
    return session;
  }

  /** 노드 생성에 실패한 항목 — 한 개가 터져도 나머지는 살린다 */
  readonly failures: string[] = [];

  async addSlide(slide: ImportSlide, index: number): Promise<void> {
    const frame = figma.createFrame();
    frame.name = slide.name;
    frame.resizeWithoutConstraints(
      Math.max(1, this.meta.widthPt),
      Math.max(1, this.meta.heightPt),
    );
    frame.x = this.origin.x + index * (this.meta.widthPt + GAP);
    frame.y = this.origin.y;
    frame.clipsContent = true;
    frame.fills = slide.fill
      ? [toFigmaPaint(slide.fill) as Paint2]
      : [{ type: 'SOLID', color: { r: 1, g: 1, b: 1 } }];
    figma.currentPage.appendChild(frame);

    /*
     * 노드 하나가 터져도 슬라이드 전체를 잃지 않는다.
     * 실패는 삼키지 않고 어느 노드였는지 이름과 함께 보고한다 — 조용히 빠지면 원인을 못 찾는다.
     */
    for (const node of slide.nodes) {
      try {
        const created = await createNode(node, this.fonts);
        for (const n of created) {
          const { x, y } = n;
          frame.appendChild(n);
          n.x = x;
          n.y = y;
        }
      } catch (err) {
        this.failures.push(`${slide.name} › ${node.name} (${node.type}): ${String(err)}`);
      }
    }

    this.frames.push(frame);
  }

  missingFonts(): string[] {
    return this.fonts.missing();
  }
}

/** 기존 내용과 겹치지 않는 위치를 찾는다 */
function nextFreeSpot(): { x: number; y: number } {
  const nodes = figma.currentPage.children;
  if (nodes.length === 0) return { x: 0, y: 0 };
  let right = -Infinity;
  let top = Infinity;
  for (const n of nodes) {
    const b = n.absoluteBoundingBox;
    if (!b) continue;
    right = Math.max(right, b.x + b.width);
    top = Math.min(top, b.y);
  }
  if (!Number.isFinite(right)) return { x: 0, y: 0 };
  return { x: right + 200, y: Number.isFinite(top) ? top : 0 };
}

/* ── 폰트 ────────────────────────────────────────────────────── */

/**
 * 폰트는 가져오기에서 가장 자주 어긋나는 지점이다.
 * PPTX 는 폰트 이름만 적어두고 파일을 담지 않으므로, 그 폰트가 이 컴퓨터에 없으면 실패한다.
 * 실패한 폰트는 대체하고 목록으로 보고한다 — 조용히 다른 폰트로 바꿔치면 나중에 더 큰 혼란이 된다.
 */
class FontBook {
  private readonly ok = new Set<string>();
  private readonly failed = new Set<string>();
  private fallback: FontName = { family: 'Inter', style: 'Regular' };

  /**
   * 쓸 폰트를 미리 로드한다.
   *
   * 후보마다 loadFontAsync 를 시도하고 실패를 잡아내는 방식은 쓰지 않는다.
   * 폰트 25종 × 4스타일이면 왕복이 100번이라 플러그인이 멈춘 것처럼 보인다.
   * 설치된 목록을 한 번만 받아 존재하는 것만 실제로 로드한다.
   */
  async load(families: string[]): Promise<void> {
    const available = new Set<string>();
    try {
      for (const f of await figma.listAvailableFontsAsync()) {
        available.add(`${f.fontName.family}|${f.fontName.style}`);
      }
    } catch {
      // 목록을 못 받으면 아래 로드가 실패하며 대체 폰트로 떨어진다.
    }

    for (const candidate of [
      { family: 'Inter', style: 'Regular' },
      { family: 'Roboto', style: 'Regular' },
    ]) {
      if (available.size > 0 && !available.has(`${candidate.family}|${candidate.style}`)) continue;
      try {
        await figma.loadFontAsync(candidate);
        this.fallback = candidate;
        break;
      } catch {
        // 다음 후보로
      }
    }

    const styles = ['Regular', 'Bold', 'Italic', 'Bold Italic'];
    const wanted: FontName[] = [];
    for (const family of families) {
      for (const style of styles) {
        const key = `${family}|${style}`;
        if (available.size > 0 && !available.has(key)) {
          this.failed.add(key);
          continue;
        }
        wanted.push({ family, style });
      }
    }

    // 존재가 확인된 것만 로드한다. 병렬로 보내면 58장짜리도 체감 지연이 없다.
    await Promise.all(wanted.map(async (font) => {
      const key = `${font.family}|${font.style}`;
      try {
        await figma.loadFontAsync(font);
        this.ok.add(key);
      } catch {
        this.failed.add(key);
      }
    }));
  }

  /** 쓸 수 있는 가장 가까운 폰트를 돌려준다 */
  resolve(family: string, style: string): FontName {
    if (this.ok.has(`${family}|${style}`)) return { family, style };
    if (this.ok.has(`${family}|Regular`)) return { family, style: 'Regular' };
    return this.fallback;
  }

  missing(): string[] {
    const families = new Set<string>();
    for (const key of this.failed) {
      const family = key.slice(0, key.indexOf('|'));
      if (!this.ok.has(`${family}|Regular`)) families.add(family);
    }
    return Array.from(families).sort();
  }
}

/* ── 배치 ────────────────────────────────────────────────────── */

/**
 * 회전·반전을 relativeTransform 으로 직접 넣는다.
 * x/y 와 rotation 을 따로 설정하면 Figma 가 중심을 다시 잡아 위치가 밀린다.
 */
function applyPlacement(node: SceneNode, place: Placement): void {
  const w = Math.max(0.01, place.w);
  const h = Math.max(0.01, place.h);
  const rad = (place.rotation * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  const fx = place.flipH ? -1 : 1;
  const fy = place.flipV ? -1 : 1;

  // M = T(center) · R(θ) · S(flip) · T(-w/2, -h/2)
  const a = cos * fx;
  const c = sin * fy;
  const b = -sin * fx;
  const d = cos * fy;
  const cx = place.x + w / 2;
  const cy = place.y + h / 2;

  node.relativeTransform = [
    [a, c, cx - (a * (w / 2) + c * (h / 2))],
    [b, d, cy - (b * (w / 2) + d * (h / 2))],
  ];
}

function resize(node: SceneNode & { resizeWithoutConstraints: (w: number, h: number) => void },
  place: Placement): void {
  node.resizeWithoutConstraints(Math.max(0.01, place.w), Math.max(0.01, place.h));
}

/* ── 색 ──────────────────────────────────────────────────────── */

function rgbOf(hex: string): RGB {
  const h = hex.replace('#', '').padStart(6, '0');
  return {
    r: parseInt(h.slice(0, 2), 16) / 255,
    g: parseInt(h.slice(2, 4), 16) / 255,
    b: parseInt(h.slice(4, 6), 16) / 255,
  };
}

function toFigmaPaint(paint: Paint): SolidPaint | GradientPaint2 {
  if (paint.kind === 'solid') {
    return { type: 'SOLID', color: rgbOf(paint.color), opacity: paint.opacity };
  }
  return gradientPaint(paint);
}

type GradientPaint2 = {
  type: 'GRADIENT_LINEAR' | 'GRADIENT_RADIAL';
  gradientTransform: Transform;
  gradientStops: ColorStop[];
};

/**
 * PPTX 그라디언트 → Figma.
 *
 * Figma 는 단위 사각형에서 왼→오른쪽으로 흐르는 그라디언트를 gradientTransform 으로 돌린다.
 * PPTX 의 ang 은 화면 기준 시계방향이고 Figma 도 y 가 아래라 각도가 그대로 대응된다.
 */
function gradientPaint(g: GradientPaint): GradientPaint2 {
  const rad = (g.angle * Math.PI) / 180;
  const c = Math.cos(rad);
  const s = Math.sin(rad);
  return {
    type: g.type === 'radial' ? 'GRADIENT_RADIAL' : 'GRADIENT_LINEAR',
    gradientTransform: [
      [c, s, 0.5 - 0.5 * c - 0.5 * s],
      [-s, c, 0.5 + 0.5 * s - 0.5 * c],
    ],
    gradientStops: g.stops.map((stop) => ({
      position: stop.position,
      color: { ...rgbOf(stop.color), a: stop.opacity },
    })),
  };
}

function applyStroke(node: SceneNode & MinimalStrokesMixin, stroke: StrokeSpec): void {
  node.strokes = [toFigmaPaint(stroke.paint) as Paint2];
  node.strokeWeight = Math.max(0.01, stroke.width);
  node.strokeAlign = 'CENTER';
  if (stroke.dash) node.dashPattern = stroke.dash;
}

type Paint2 = SolidPaint | GradientPaint2;

/* ── 노드 생성 ───────────────────────────────────────────────── */

async function createNode(spec: ImportNode, fonts: FontBook): Promise<SceneNode[]> {
  switch (spec.type) {
    case 'shape':
      return [createShape(spec)];
    case 'text': {
      const node = await createText(spec, fonts);
      return node ? [node] : [];
    }
    case 'image':
      return createImage(spec);
    case 'table':
      return [await createTable(spec, fonts)];
    case 'group': {
      const children: SceneNode[] = [];
      for (const child of spec.children) {
        try {
          children.push(...await createNode(child, fonts));
        } catch (err) {
          // 형제 하나가 실패해도 그룹은 만든다. 바깥에서 잡아 보고한다.
          throw new Error(`${spec.name} › ${child.name}: ${String(err)}`);
        }
      }
      if (children.length === 0) return [];
      // figma.group 은 부모가 있어야 하므로 호출부에서 프레임에 붙인 뒤 묶는다.
      return [wrapGroup(children, spec.name)];
    }
    default:
      return [];
  }
}

/**
 * 그룹은 프레임으로 만든다. figma.group() 은 노드가 이미 부모에 붙어 있어야 동작하는데,
 * 여기서는 아직 어디에도 붙지 않은 상태라 순서가 꼬인다.
 * 크기 제약을 끄고 자르기를 끈 프레임이면 그룹과 동일하게 보인다.
 */
function wrapGroup(children: SceneNode[], name: string): FrameNode {
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (const c of children) {
    minX = Math.min(minX, c.x);
    minY = Math.min(minY, c.y);
    maxX = Math.max(maxX, c.x + c.width);
    maxY = Math.max(maxY, c.y + c.height);
  }
  if (!Number.isFinite(minX)) {
    minX = 0;
    minY = 0;
    maxX = 1;
    maxY = 1;
  }

  const frame = figma.createFrame();
  frame.name = name;
  frame.fills = [];
  frame.clipsContent = false;
  frame.resizeWithoutConstraints(Math.max(0.01, maxX - minX), Math.max(0.01, maxY - minY));
  frame.x = minX;
  frame.y = minY;
  for (const c of children) {
    const { x, y } = c;
    frame.appendChild(c);
    c.x = x - minX;
    c.y = y - minY;
  }
  return frame;
}

function createShape(spec: ShapeNode): SceneNode {
  const geom = spec.geometry;
  let node: SceneNode & MinimalStrokesMixin & MinimalFillsMixin;

  if (geom.kind === 'ellipse') {
    const e = figma.createEllipse();
    resize(e, spec.place);
    node = e;
  } else if (geom.kind === 'line') {
    const line = figma.createLine();
    // PPTX 선은 상자의 좌상단에서 우하단으로 간다. Figma LineNode 는 항상 수평이라
    // 길이와 각도로 바꿔 넣는다.
    const len = Math.hypot(spec.place.w, spec.place.h);
    line.resizeWithoutConstraints(Math.max(0.01, len), 0);
    const angle = Math.atan2(spec.place.h, spec.place.w) * (180 / Math.PI);
    applyPlacement(line, {
      ...spec.place,
      w: len,
      h: 0,
      rotation: spec.place.rotation - angle,
    });
    if (spec.stroke) applyStroke(line, spec.stroke);
    else line.strokes = [];
    line.name = spec.name;
    line.opacity = spec.opacity;
    return line;
  } else if (geom.kind === 'path') {
    const v = figma.createVector();
    v.vectorPaths = [{
      windingRule: geom.evenOdd ? 'EVENODD' : 'NONZERO',
      data: geom.data,
    }];
    node = v;
  } else {
    const r = figma.createRectangle();
    resize(r, spec.place);
    if (geom.kind === 'roundRect') {
      const [tl, tr, br, bl] = geom.radii;
      r.topLeftRadius = tl;
      r.topRightRadius = tr;
      r.bottomRightRadius = br;
      r.bottomLeftRadius = bl;
    }
    node = r;
  }

  node.name = spec.name;
  node.fills = spec.fill ? [toFigmaPaint(spec.fill) as Paint2] : [];
  // 선이 없으면 명시적으로 비운다. createVector() 로 만든 노드는 기본 선을 달고 나온다.
  if (spec.stroke) applyStroke(node, spec.stroke);
  else node.strokes = [];
  if (spec.shadow) {
    node.effects = [{
      type: 'DROP_SHADOW',
      color: { ...rgbOf(spec.shadow.color), a: spec.shadow.opacity },
      offset: { x: spec.shadow.offsetX, y: spec.shadow.offsetY },
      radius: spec.shadow.blur,
      spread: 0,
      visible: true,
      blendMode: 'NORMAL',
    }];
  }
  node.opacity = spec.opacity;
  applyPlacement(node, spec.place);
  return node;
}

function createImage(spec: ImageNodeSpec): SceneNode[] {
  if (spec.isSvg) {
    try {
      const svg = figma.createNodeFromSvg(decodeUtf8(figma.base64Decode(spec.data)));
      svg.name = spec.name;
      svg.resizeWithoutConstraints(Math.max(0.01, spec.place.w), Math.max(0.01, spec.place.h));
      applyPlacement(svg, spec.place);
      return [svg];
    } catch {
      return [];
    }
  }

  try {
    const image = figma.createImage(figma.base64Decode(spec.data));
    const rect = figma.createRectangle();
    rect.name = spec.name;
    resize(rect, spec.place);
    rect.fills = [{ type: 'IMAGE', imageHash: image.hash, scaleMode: 'FILL' }];
    if (spec.radii) {
      const [tl, tr, br, bl] = spec.radii;
      rect.topLeftRadius = tl;
      rect.topRightRadius = tr;
      rect.bottomRightRadius = br;
      rect.bottomLeftRadius = bl;
    }
    rect.opacity = spec.opacity;
    applyPlacement(rect, spec.place);
    return [rect];
  } catch {
    return [];
  }
}

function decodeUtf8(bytes: Uint8Array): string {
  let out = '';
  for (let i = 0; i < bytes.length; i++) {
    const b = bytes[i];
    if (b < 0x80) out += String.fromCharCode(b);
    else if (b < 0xe0) out += String.fromCharCode(((b & 0x1f) << 6) | (bytes[++i] & 0x3f));
    else if (b < 0xf0) {
      out += String.fromCharCode(((b & 0x0f) << 12) | ((bytes[++i] & 0x3f) << 6) | (bytes[++i] & 0x3f));
    } else {
      const cp = ((b & 0x07) << 18) | ((bytes[++i] & 0x3f) << 12)
        | ((bytes[++i] & 0x3f) << 6) | (bytes[++i] & 0x3f);
      out += String.fromCodePoint(cp);
    }
  }
  return out;
}

/* ── 텍스트 ──────────────────────────────────────────────────── */

async function createText(spec: TextNodeSpec, fonts: FontBook): Promise<SceneNode | null> {
  const text = figma.createText();
  text.name = spec.name;

  // 첫 폰트를 먼저 세워야 characters 를 넣을 수 있다.
  const first = spec.paragraphs[0]?.runs[0];
  text.fontName = fonts.resolve(first?.fontFamily ?? 'Inter', first?.fontStyle ?? 'Regular');

  const lines: string[] = [];
  for (const para of spec.paragraphs) {
    lines.push(para.runs.map((r) => r.text).join(''));
  }
  const content = lines.join('\n');
  if (!content) {
    text.remove();
    return null;
  }
  text.characters = content;

  // 문자 범위별로 서식을 다시 입힌다.
  let cursor = 0;
  for (let pi = 0; pi < spec.paragraphs.length; pi++) {
    const para = spec.paragraphs[pi];
    const paraStart = cursor;
    for (const run of para.runs) {
      const len = run.text.length;
      if (len > 0) {
        const end = cursor + len;
        text.setRangeFontName(cursor, end, fonts.resolve(run.fontFamily, run.fontStyle));
        text.setRangeFontSize(cursor, end, Math.max(1, run.size));
        text.setRangeFills(cursor, end,
          [{ type: 'SOLID', color: rgbOf(run.color), opacity: run.opacity }]);
        if (run.underline) text.setRangeTextDecoration(cursor, end, 'UNDERLINE');
        else if (run.strike) text.setRangeTextDecoration(cursor, end, 'STRIKETHROUGH');
        if (run.letterSpacing !== undefined) {
          text.setRangeLetterSpacing(cursor, end, { unit: 'PIXELS', value: run.letterSpacing });
        }
        if (para.lineHeightPct) {
          text.setRangeLineHeight(cursor, end, { unit: 'PERCENT', value: para.lineHeightPct });
        }
        cursor = end;
      }
    }
    /*
     * 문단 서식. 글머리 기호는 Figma 의 목록으로 옮긴다 — 글리프는 고를 수 없지만
     * 내어쓰기(줄바꿈된 줄이 기호가 아니라 글자 아래로 붙는 것)가 자동으로 맞는다.
     * 이 문서에서 내어쓰기가 518곳이라 이게 정렬 차이의 대부분이었다.
     */
    const paraEnd = cursor;
    if (paraEnd > paraStart) {
      if (para.bullet) {
        text.setRangeListOptions(paraStart, paraEnd, {
          type: para.bullet.kind === 'number' ? 'ORDERED' : 'UNORDERED',
        });
      }
      if (para.level > 0) text.setRangeIndentation(paraStart, paraEnd, para.level);
      const spacing = para.spaceBefore ?? para.spaceAfter;
      if (spacing) text.setRangeParagraphSpacing(paraStart, paraEnd, spacing);
    }

    if (pi < spec.paragraphs.length - 1) cursor += 1; // 줄바꿈 문자
  }

  // Figma 는 문단별 가로 정렬을 지원하지 않는다 (setRange* 에 해당 항목이 없다).
  // 첫 문단 기준으로 노드 전체에 건다. 문단마다 다른 경우는 reader 가 경고로 남긴다.
  text.textAlignHorizontal = spec.paragraphs[0]?.align ?? 'LEFT';
  text.textAlignVertical = spec.vertical;
  /*
   * 상자 크기는 언제나 원본 값을 쓴다.
   *
   * PPTX 의 wrap="none" + spAutoFit 상자는 PowerPoint 가 **진짜 폰트로** 계산해 저장해 둔
   * 크기다. Figma 자동 크기에 맡기면 대체 폰트 폭으로 다시 재면서 상자가 커지고,
   * 오른쪽 정렬 머리글이 페이지 여백을 넘어간다. 폰트가 달라 줄바꿈은 어차피 어긋나지만,
   * 위치와 여백만은 원본을 지키는 편이 낫다.
   */
  text.textAutoResize = 'NONE';

  // 상자 여백만큼 안쪽으로 들여 배치한다. Figma 텍스트에는 내부 여백이 없다.
  const place: Placement = {
    ...spec.place,
    x: spec.place.x + spec.insets.left,
    y: spec.place.y + spec.insets.top,
    w: Math.max(1, spec.place.w - spec.insets.left - spec.insets.right),
    h: Math.max(1, spec.place.h - spec.insets.top - spec.insets.bottom),
  };
  resize(text, place);
  applyPlacement(text, place);
  text.opacity = spec.opacity;
  return text;
}

/* ── 표 ──────────────────────────────────────────────────────── */

/**
 * Figma 에는 표 원시 타입이 없다. 행 프레임 안에 셀 프레임을 넣는 격자로 재구성한다.
 * 셀 테두리는 PPTX 가 변마다 따로 주므로, 프레임의 변별 선 두께로 그대로 옮긴다.
 */
async function createTable(spec: TableNodeSpec, fonts: FontBook): Promise<SceneNode> {
  const table = figma.createFrame();
  table.name = spec.name;
  table.fills = [];
  table.clipsContent = false;
  table.layoutMode = 'VERTICAL';
  table.primaryAxisSizingMode = 'AUTO';
  table.counterAxisSizingMode = 'AUTO';
  table.itemSpacing = 0;

  for (let r = 0; r < spec.rows.length; r++) {
    const rowSpec = spec.rows[r];
    const row = figma.createFrame();
    row.name = `행 ${r + 1}`;
    row.fills = [];
    row.clipsContent = false;
    row.layoutMode = 'HORIZONTAL';
    row.primaryAxisSizingMode = 'AUTO';
    row.counterAxisSizingMode = 'FIXED';
    row.itemSpacing = 0;
    table.appendChild(row);
    row.resizeWithoutConstraints(
      Math.max(1, spec.colWidths.reduce((a, b) => a + b, 0)),
      Math.max(1, spec.rowHeights[r] ?? 20),
    );

    for (let c = 0; c < rowSpec.length; c++) {
      const cellSpec = rowSpec[c];
      const cell = figma.createFrame();
      cell.name = `셀 ${r + 1}-${c + 1}`;
      cell.clipsContent = false;
      cell.fills = cellSpec.fill ? [toFigmaPaint(cellSpec.fill) as Paint2] : [];
      row.appendChild(cell);

      const width = (spec.colWidths[c] ?? 60) * (cellSpec.colSpan || 1);
      const height = spec.rowHeights[r] ?? 20;
      cell.resizeWithoutConstraints(Math.max(1, width), Math.max(1, height));

      applyCellBorders(cell, cellSpec.borders);

      if (cellSpec.paragraphs.length > 0) {
        const text = await createText({
          type: 'text',
          name: '내용',
          place: { x: 0, y: 0, w: width, h: height, rotation: 0, flipH: false, flipV: false },
          opacity: 1,
          paragraphs: cellSpec.paragraphs,
          vertical: cellSpec.vertical,
          insets: cellSpec.insets,
          autoWidth: false,
        }, fonts);
        if (text) cell.appendChild(text);
      }
    }
  }

  applyPlacement(table, spec.place);
  return table;
}

function applyCellBorders(cell: FrameNode, borders: TableCell2['borders']): void {
  const sides = [borders.top, borders.right, borders.bottom, borders.left];
  const present = sides.filter((s): s is StrokeSpec => !!s);
  if (present.length === 0) return;

  // Figma 프레임은 변별 두께를 지원한다 — PPTX 표의 가로선만 있는 구조가 그대로 옮겨진다.
  cell.strokes = [toFigmaPaint(present[0].paint) as Paint2];
  cell.strokeAlign = 'INSIDE';
  cell.strokeTopWeight = borders.top?.width ?? 0;
  cell.strokeRightWeight = borders.right?.width ?? 0;
  cell.strokeBottomWeight = borders.bottom?.width ?? 0;
  cell.strokeLeftWeight = borders.left?.width ?? 0;
}

type TableCell2 = TableNodeSpec['rows'][number][number];
