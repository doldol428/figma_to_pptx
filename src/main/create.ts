import type {
  GradientPaint, ImageNodeSpec, ImportLayout, ImportNode, ImportSlide, Paint, Placement,
  ShadowSpec, ShapeNode, StrokeSpec, TableNodeSpec, TextNodeSpec,
} from '../shared/importir';
import { type PathBox, pathBounds, translatePath } from '../import/pathbox';
import { linePlacement } from '../import/transform';
import { KEY, NAME, layoutName } from '../shared/roles';
import { aliasesFor, latinHead, pickFont } from './fontalias';

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
  /** 문서에 포함된 글꼴에서 읽어낸 한글 → 영문 이름 대응 */
  fontAliases: Record<string, string>;
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
  /** 레이아웃 파트 → 컴포넌트. 슬라이드는 여기에 인스턴스만 놓는다. */
  private readonly layouts = new Map<string, ComponentNode>();

  private constructor(private readonly meta: ImportMeta) {}

  get total(): number {
    return this.meta.total;
  }

  static async begin(meta: ImportMeta): Promise<ImportSession> {
    const session = new ImportSession(meta);
    await session.fonts.load(meta.fonts, meta.fontAliases);
    return session;
  }

  /** 노드 생성에 실패한 항목 — 한 개가 터져도 나머지는 살린다 */
  readonly failures: string[] = [];

  /**
   * 페이지 공통 요소를 컴포넌트 하나로 만든다.
   *
   * PPTX 의 slideLayout 과 1:1 이다. 같은 레이아웃을 쓰는 장이 인스턴스를 놓으므로
   * 머리글을 한 번 고치면 그 레이아웃을 쓰는 모든 장에 반영되고, 다시 내보낼 때도
   * 어느 것이 공통 서식인지 판단할 필요 없이 레이아웃 파트로 되돌릴 수 있다.
   */
  async addLayout(layout: ImportLayout): Promise<void> {
    const component = figma.createComponent();
    component.name = layoutName(layout.name);
    component.resizeWithoutConstraints(
      Math.max(1, this.meta.widthPt),
      Math.max(1, this.meta.heightPt),
    );
    // 슬라이드 줄 왼쪽에 세로로 쌓는다 — 캔버스에서 원본과 섞이지 않게.
    component.x = this.origin.x - (this.meta.widthPt + GAP) * 2;
    component.y = this.origin.y + this.layouts.size * (this.meta.heightPt + GAP);
    component.fills = [];
    component.clipsContent = false;
    component.setPluginData(KEY.role, 'layout');
    component.setPluginData(KEY.layout, layout.key);
    figma.currentPage.appendChild(component);

    await this.fill(component, layout.nodes, component.name);
    this.layouts.set(layout.key, component);
  }

  /** 노드들을 부모 안에 만들어 넣는다. 하나가 터져도 나머지는 살린다. */
  private async fill(parent: FrameNode | ComponentNode, nodes: ImportNode[], where: string): Promise<void> {
    for (const node of nodes) {
      try {
        const created = await createNode(node, this.fonts);
        for (const n of created) {
          const { x, y } = n;
          parent.appendChild(n);
          n.x = x;
          n.y = y;
        }
      } catch (err) {
        this.failures.push(`${where} › ${node.name} (${node.type}): ${String(err)}`);
      }
    }
  }

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
    frame.setPluginData(KEY.slide, String(index + 1));
    figma.currentPage.appendChild(frame);

    /*
     * 공통 요소는 인스턴스 하나로 깔고 시작한다. 맨 아래에 놓아야 슬라이드 내용이 그 위에 앉는다.
     * (레이아웃 컴포넌트가 아직 없으면 그냥 넘어간다 — 공통 요소가 없는 문서다.)
     */
    const component = slide.layoutKey ? this.layouts.get(slide.layoutKey) : undefined;
    if (component) {
      const instance = component.createInstance();
      instance.setPluginData(KEY.role, 'layout');
      instance.setPluginData(KEY.layout, slide.layoutKey ?? '');
      frame.appendChild(instance);
      instance.x = 0;
      instance.y = 0;
    }

    /*
     * 슬라이드 번호처럼 장마다 값이 다른 것은 컴포넌트에 넣을 수 없어 여기서 따로 만든다.
     * 이름을 규칙대로 붙여 두면 다시 내보낼 때 필드로 되돌릴 수 있다.
     */
    for (const node of slide.perSlideNodes) {
      const before = frame.children.length;
      await this.fill(frame, [node], slide.name);
      for (let i = before; i < frame.children.length; i++) {
        const n = frame.children[i];
        n.setPluginData(KEY.pptxName, n.name);
        n.setPluginData(KEY.role, 'slideNumber');
        n.name = NAME.slideNumber;
      }
    }

    /*
     * 노드 하나가 터져도 슬라이드 전체를 잃지 않는다.
     * 실패는 삼키지 않고 어느 노드였는지 이름과 함께 보고한다 — 조용히 빠지면 원인을 못 찾는다.
     */
    await this.fill(frame, slide.nodes, slide.name);

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
const STYLES = ['Regular', 'Bold', 'Italic', 'Bold Italic'];

class FontBook {
  /** 설치된 family → 그 family 가 가진 style 들 */
  private readonly installed = new Map<string, Set<string>>();
  /** "타이프페이스|요청스타일" → 실제로 쓸 FontName */
  private readonly resolved = new Map<string, FontName>();
  private readonly unresolved = new Set<string>();
  /** 문서가 포함한 글꼴에서 읽어낸 "한글 이름 → 영문 이름". 손으로 만든 표보다 정확하다. */
  private docAliases: Record<string, string> = {};
  private fallback: FontName = { family: 'Inter', style: 'Regular' };

  /**
   * 쓸 폰트를 미리 로드한다.
   *
   * 후보마다 loadFontAsync 를 시도하고 실패를 잡는 방식은 쓰지 않는다.
   * 25종 × 4스타일이면 왕복이 100번이라 플러그인이 멈춘 것처럼 보인다.
   * 설치 목록을 한 번만 받아 이름을 맞춘 뒤, 실재하는 것만 병렬로 로드한다.
   */
  async load(typefaces: string[], docAliases: Record<string, string> = {}): Promise<void> {
    this.docAliases = docAliases;
    try {
      for (const f of await figma.listAvailableFontsAsync()) {
        const { family, style } = f.fontName;
        let styles = this.installed.get(family);
        if (!styles) {
          styles = new Set<string>();
          this.installed.set(family, styles);
        }
        styles.add(style);
      }
    } catch {
      // 목록을 못 받으면 아래 로드가 실패하며 대체 폰트로 떨어진다.
    }

    for (const candidate of [
      { family: 'Inter', style: 'Regular' },
      { family: 'Roboto', style: 'Regular' },
    ]) {
      if (!this.installed.get(candidate.family)?.has(candidate.style)) continue;
      try {
        await figma.loadFontAsync(candidate);
        this.fallback = candidate;
        break;
      } catch {
        // 다음 후보로
      }
    }

    const wanted = new Map<string, FontName>();
    for (const typeface of typefaces) {
      let anyMatch = false;
      for (const style of STYLES) {
        const font = this.match(typeface, style);
        if (!font) continue;
        anyMatch = true;
        this.resolved.set(`${typeface}|${style}`, font);
        wanted.set(`${font.family}|${font.style}`, font);
      }
      if (!anyMatch) this.unresolved.add(typeface);
    }

    await Promise.all(Array.from(wanted.values()).map(async (font) => {
      try {
        await figma.loadFontAsync(font);
      } catch {
        for (const [key, value] of this.resolved) {
          if (value.family === font.family && value.style === font.style) this.resolved.delete(key);
        }
        this.unresolved.add(font.family);
      }
    }));
  }

  /**
   * PPTX 의 타이프페이스 이름을 설치된 family/style 로 맞춘다.
   *
   * 두 가지가 동시에 어긋난다.
   * 1. PPTX 는 한글 이름("KoPub돋움체")을 적지만 Figma 는 영문 이름("KoPubDotum")으로만 등록한다.
   * 2. PPTX 는 "페이퍼로지 6 SemiBold" 처럼 굵기까지 붙은 전체 이름을 쓰지만
   *    Figma 는 family 와 style 을 나눠서 갖는다.
   * 그래서 이름을 영문으로 바꾼 후보들까지 포함해 각각 family/style 분리를 시도한다.
   */
  private match(typeface: string, wantStyle: string): FontName | null {
    // 문서가 알려준 이름을 먼저 본다 — 파일에 박힌 실제 이름표라 추측이 아니다.
    const fromDoc = this.docAliases[typeface];
    const candidates = fromDoc
      ? [typeface, fromDoc, ...aliasesFor(typeface)]
      : [typeface, ...aliasesFor(typeface)];
    for (const name of candidates) {
      const hit = this.matchName(name, wantStyle);
      if (hit) return hit;
    }
    return null;
  }

  /** 이름 하나를 family/style 로 쪼개 본다. 규칙은 순수 함수로 빼 두어 목록만 있으면 검증할 수 있다. */
  private matchName(typeface: string, wantStyle: string): FontName | null {
    return pickFont(this.installed, typeface, wantStyle);
  }

  /**
   * 쓸 수 있는 가장 가까운 폰트를 돌려준다.
   * 같은 글꼴의 다른 이름(한글/영문)을 차례로 시도한다 — 어느 쪽이 등록돼 있는지는 설치 방식에 달렸다.
   */
  resolve(typeface: string, style: string, alternates?: string[]): FontName {
    for (const name of [typeface, ...(alternates ?? [])]) {
      const hit = this.resolved.get(`${name}|${style}`) ?? this.resolved.get(`${name}|Regular`);
      if (hit) return hit;
    }
    return this.fallback;
  }

  /**
   * 끝내 못 찾은 이름들. 설치 목록에서 비슷한 이름을 함께 붙인다 —
   * 같은 글꼴이 한글/영문 두 이름을 갖는 경우가 많아, 실제 등록명을 봐야 원인이 잡힌다.
   */
  missing(): string[] {
    return Array.from(this.unresolved).sort().map((name) => {
      // 문서에 박힌 영문 이름을 같이 보여준다 — 설치할 때 찾아야 할 이름이 이것이다.
      const english = this.docAliases[name];
      const label = english ? `${name} (${english})` : name;
      const hint = this.similar(english ?? name);
      return hint ? `${label} → 설치 목록의 비슷한 이름: ${hint}` : label;
    });
  }

  private similar(typeface: string): string {
    /*
     * 앞머리 영문 부분으로도 찾는다. "KoPub돋움체" 를 통째로 대조하면 아무것도 안 걸리지만
     * "KoPub" 으로는 KoPubDotum · KoPubBatang 이 잡힌다 — 등록명을 눈으로 확인할 수 있어야 한다.
     */
    const heads = [typeface.split(' ')[0].toLowerCase()];
    const latin = latinHead(typeface).toLowerCase();
    if (latin.length >= 2 && heads.indexOf(latin) < 0) heads.push(latin);

    const hits: string[] = [];
    for (const family of this.installed.keys()) {
      const lower = family.toLowerCase();
      for (const head of heads) {
        if (head.length < 2) continue;
        if (lower.indexOf(head) < 0 && head.indexOf(lower) < 0) continue;
        if (hits.indexOf(family) < 0) hits.push(family);
        break;
      }
      if (hits.length >= 4) break;
    }
    return hits.join(', ');
  }
}

/* ── 배치 ────────────────────────────────────────────────────── */

/**
 * 회전·반전을 relativeTransform 으로 직접 넣는다.
 * x/y 와 rotation 을 따로 설정하면 Figma 가 중심을 다시 잡아 위치가 밀린다.
 */
function applyPlacement(node: SceneNode, place: Placement, content?: { x: number; y: number }): void {
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

  /*
   * 노드의 로컬 원점이 이름상의 상자 좌상단이 아닐 수 있다 (벡터는 경로가 차지하는 만큼만 크다).
   * 그 어긋남을 **로컬 공간에서** 더한다 — 회전·반전을 이미 거친 축을 따라야 하므로
   * 결과 좌표에 그냥 더하면 안 된다.
   */
  const ox = content ? content.x : 0;
  const oy = content ? content.y : 0;

  node.relativeTransform = [
    [a, c, cx - (a * (w / 2) + c * (h / 2)) + a * ox + c * oy],
    [b, d, cy - (b * (w / 2) + d * (h / 2)) + b * ox + d * oy],
  ];
}

function resize(node: SceneNode & { resizeWithoutConstraints: (w: number, h: number) => void },
  place: Placement): void {
  node.resizeWithoutConstraints(Math.max(0.01, place.w), Math.max(0.01, place.h));
}

/**
 * 이미 제 크기를 잡은 텍스트를 원래 상자의 정렬 기준점에 맞춰 놓는다.
 *
 * 자동 크기 상자는 폰트가 바뀌면 폭이 달라진다. 좌상단을 고정하면 가운데 정렬 글이 오른쪽으로,
 * 오른쪽 정렬 머리글이 여백 밖으로 밀린다. 정렬 방향에 해당하는 변을 붙잡아야 제자리에 남는다.
 */
function anchor(
  text: TextNode,
  place: Placement,
  align: TextNode['textAlignHorizontal'],
  vertical: 'TOP' | 'CENTER' | 'BOTTOM',
): void {
  const dx = align === 'RIGHT' ? place.w - text.width
    : align === 'CENTER' ? (place.w - text.width) / 2
      : 0;
  const dy = vertical === 'BOTTOM' ? place.h - text.height
    : vertical === 'CENTER' ? (place.h - text.height) / 2
      : 0;
  applyPlacement(text, {
    ...place,
    x: place.x + dx,
    y: place.y + dy,
    w: text.width,
    h: text.height,
  });
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
  /*
   * 끝 모양을 빠뜨리면 선이 양끝에서 굵기의 절반씩 짧아진다.
   * 브라켓처럼 선과 호를 이어 붙인 그림은 이음매마다 그만큼 파여 보인다 — 원본은 cap="rnd" 다.
   */
  if (stroke.cap && 'strokeCap' in node) node.strokeCap = stroke.cap;
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
      const group = wrapGroup(
        children,
        // 도형+글자는 이름으로도 알아볼 수 있게 표시를 남긴다 (손으로 만든 것도 통하게).
        spec.shapeText && children.length === 2 ? `${NAME.shapeText}/${spec.name}` : spec.name,
      );
      if (spec.shapeText && children.length === 2) {
        group.setPluginData(KEY.role, 'shapeText');
        group.setPluginData(KEY.pptxName, spec.name);
      }
      return [group];
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
  /** 벡터일 때, 경로가 이름상의 상자 안 어디에서 시작하는지 */
  let pathBox: PathBox | null = null;
  let vector: VectorNode | null = null;

  if (geom.kind === 'ellipse') {
    const e = figma.createEllipse();
    resize(e, spec.place);
    node = e;
  } else if (geom.kind === 'line') {
    const line = figma.createLine();
    const at = linePlacement(spec.place);
    line.resizeWithoutConstraints(Math.max(0.01, at.w), 0);
    applyPlacement(line, at);
    if (spec.stroke) applyStroke(line, spec.stroke);
    else line.strokes = [];
    line.name = spec.name;
    line.opacity = spec.opacity;
    // 선도 그림자를 진다. 여기서 일찍 돌아가느라 빠뜨리고 있었다 — 연결선 163개가 그랬다.
    if (spec.shadow) line.effects = [dropShadow(spec.shadow)];
    return line;
  } else if (geom.kind === 'path') {
    const v = figma.createVector();
    /*
     * 경로를 원점으로 당겨 놓는다. 그래야 벡터 노드의 로컬 상자가 경로와 정확히 일치하고,
     * 회전·반전이 도형 자신의 축을 따른다. 당긴 만큼은 배치에서 되돌린다.
     */
    pathBox = pathBounds(geom.data);
    v.vectorPaths = [{
      windingRule: geom.evenOdd ? 'EVENODD' : 'NONZERO',
      data: translatePath(geom.data, -pathBox.x, -pathBox.y),
    }];
    vector = v;
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
  /*
   * 무늬는 Figma 에 없어 단색으로 눌러 뒀다. 원본을 적어 두지 않으면 되돌릴 길이 없다 —
   * 눌러 둔 색만 남아 내보낼 때도 그대로 단색으로 나간다.
   */
  if (spec.fill?.kind === 'solid' && spec.fill.pattern) {
    node.setPluginData(KEY.pattern, JSON.stringify(spec.fill.pattern));
  }
  // 선이 없으면 명시적으로 비운다. createVector() 로 만든 노드는 기본 선을 달고 나온다.
  if (spec.stroke) applyStroke(node, spec.stroke);
  else node.strokes = [];
  if (spec.shadow) node.effects = [dropShadow(spec.shadow)];
  node.opacity = spec.opacity;

  /*
   * 경로를 원점으로 옮겨 놨으니 노드 크기는 경로 크기와 같아야 한다.
   * 다르면 Figma 가 무언가를 덧대며 원점을 옮긴 것이므로(선 굵기 등) 그만큼 되돌린다.
   * 규칙을 가정하지 않고 실제 크기를 재서 맞춘다 — 여기서는 Figma 를 돌려 확인할 수 없다.
   */
  let content: { x: number; y: number } | undefined;
  if (pathBox && vector) {
    content = {
      x: pathBox.x - (vector.width - pathBox.w) / 2,
      y: pathBox.y - (vector.height - pathBox.h) / 2,
    };
  }
  applyPlacement(node, spec.place, content);
  return node;
}

function dropShadow(shadow: ShadowSpec): DropShadowEffect {
  return {
    type: 'DROP_SHADOW',
    color: { ...rgbOf(shadow.color), a: shadow.opacity },
    offset: { x: shadow.offsetX, y: shadow.offsetY },
    radius: shadow.blur,
    spread: 0,
    visible: true,
    blendMode: 'NORMAL',
  };
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
    if (spec.shadow) rect.effects = [dropShadow(spec.shadow)];
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
  text.fontName = fonts.resolve(
    first?.fontFamily ?? 'Inter', first?.fontStyle ?? 'Regular', first?.fontAlternates,
  );

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
        text.setRangeFontName(cursor, end,
          fonts.resolve(run.fontFamily, run.fontStyle, run.fontAlternates));
        text.setRangeFontSize(cursor, end, Math.max(1, run.size));
        text.setRangeFills(cursor, end, [
          run.gradient
            ? gradientPaint(run.gradient) as Paint2
            : { type: 'SOLID', color: rgbOf(run.color), opacity: run.opacity },
        ]);
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
  // 상자 여백만큼 안쪽으로 들여 배치한다. Figma 텍스트에는 내부 여백이 없다.
  const place: Placement = {
    ...spec.place,
    x: spec.place.x + spec.insets.left,
    y: spec.place.y + spec.insets.top,
    w: Math.max(1, spec.place.w - spec.insets.left - spec.insets.right),
    h: Math.max(1, spec.place.h - spec.insets.top - spec.insets.bottom),
  };

  if (spec.autoWidth) {
    /*
     * wrap="none" 상자는 **줄을 바꾸지 않는다.** PowerPoint 는 글자가 넘치면 상자 밖으로
     * 흘려보낼 뿐 접지 않는다. 저장된 크기에 가둬 두면 "CHAPTER" 가 "CHAP/TER" 로 접힌다.
     *
     * Figma 에는 "줄바꿈 금지 + 크기 고정" 이 없으므로 자동 크기로 둔다. 대신 상자가
     * 좌상단에서 자라면 오른쪽 정렬 머리글이 여백을 넘어가므로, 정렬 기준점을 붙잡아
     * 그쪽을 축으로 자라게 한다 — PowerPoint 가 자동 맞춤에서 하는 것과 같다.
     */
    text.textAutoResize = 'WIDTH_AND_HEIGHT';
    anchor(text, place, text.textAlignHorizontal, spec.vertical);
    text.opacity = spec.opacity;
    return text;
  }

  /*
   * 일반 상자는 원본 크기를 그대로 쓴다. 폰트가 달라 줄바꿈 지점은 어차피 어긋나지만,
   * 위치와 여백만은 원본을 지키는 편이 낫다.
   */
  text.textAutoResize = 'NONE';
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
  table.name = `${NAME.table}/${spec.name}`;
  table.setPluginData(KEY.role, 'table');
  table.setPluginData(KEY.pptxName, spec.name);
  /*
   * 병합만 따로 적어 둔다. 칸 크기·색·글자·테두리는 노드에서 그대로 읽히지만 병합은
   * 프레임 격자에 남지 않고, Figma 에서 칸을 병합할 방법도 없어 이 기록이 낡지 않는다.
   */
  table.setPluginData(KEY.table, JSON.stringify({
    spans: spec.rows.map((row) => row.map((c) => [c.colSpan, c.rowSpan, c.merged ? 1 : 0])),
  }));
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
