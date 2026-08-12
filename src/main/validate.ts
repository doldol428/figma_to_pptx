import type { SelectionState, SizeInfo } from '../shared/ir';
import { t } from '../shared/i18n';
import { resolveSlide } from '../shared/slidesize';
import { pxToMm } from '../shared/units';

const EXPORTABLE = new Set(['FRAME', 'COMPONENT', 'INSTANCE']);

/**
 * 선택에서 슬라이드가 될 프레임들을 뽑는다.
 * 섹션을 고르면 그 안의 최상위 프레임들을 대신 쓴다.
 */
export function collectFrames(selection: readonly SceneNode[]): SceneNode[] {
  const frames: SceneNode[] = [];
  const seen = new Set<string>();

  const add = (node: SceneNode) => {
    if (seen.has(node.id)) return;
    seen.add(node.id);
    frames.push(node);
  };

  for (const node of selection) {
    if (EXPORTABLE.has(node.type)) {
      add(node);
    } else if (node.type === 'SECTION') {
      for (const child of node.children) {
        if (EXPORTABLE.has(child.type)) add(child);
      }
    }
  }

  return sortByCanvasOrder(frames);
}

/** 캔버스 배치 순서(위→아래, 왼→오른쪽)로 정렬. 사람이 프레임을 늘어놓은 순서와 일치한다. */
function sortByCanvasOrder(frames: SceneNode[]): SceneNode[] {
  if (frames.length < 2) return frames;
  const rowTol = Math.max(...frames.map((f) => f.height)) / 2;
  return frames.slice().sort((a, b) => {
    const ay = a.absoluteTransform[1][2];
    const by = b.absoluteTransform[1][2];
    if (Math.abs(ay - by) > rowTol) return ay - by;
    return a.absoluteTransform[0][2] - b.absoluteTransform[0][2];
  });
}

function sizeKey(node: SceneNode): string {
  return `${node.width.toFixed(2)}x${node.height.toFixed(2)}`;
}

export function validate(selection: readonly SceneNode[]): SelectionState {
  const frames = collectFrames(selection);

  const base: SelectionState = {
    ok: false,
    reason: '',
    frameCount: frames.length,
    widthPx: 0,
    heightPx: 0,
    chip: null,
    slideWPt: 0,
    slideHPt: 0,
    ptPerPx: 1,
    sizes: [],
  };

  // 선택이 비어 있는 건 오류가 아니라 시작 상태다.
  // UI 가 슬라이드 크기 영역에서 안내하므로 별도 사유를 만들지 않는다.
  if (frames.length === 0) {
    return base;
  }

  // 크기별로 묶는다 — 하나로 모이지 않으면 내보낼 수 없다.
  const groups = new Map<string, SizeInfo>();
  for (const f of frames) {
    const key = sizeKey(f);
    const g = groups.get(key);
    if (g) g.count += 1;
    else groups.set(key, { w: f.width, h: f.height, count: 1 });
  }
  const sizes = Array.from(groups.values()).sort((a, b) => b.count - a.count);

  const { width, height } = frames[0];
  const plan = resolveSlide(width, height);
  const state: SelectionState = {
    ...base,
    widthPx: width,
    heightPx: height,
    chip: plan?.chip ?? null,
    slideWPt: plan?.wPt ?? width,
    slideHPt: plan?.hPt ?? height,
    ptPerPx: plan?.ptPerPx ?? 1,
    sizes,
  };

  if (sizes.length > 1) {
    // PPTX 는 파일 하나에 슬라이드 크기가 하나뿐이다. 섞이면 어느 쪽도 원본 비율을 지킬 수 없다.
    const list = sizes
      .map((s) => t().sizeEntry(fmt(s.w), fmt(s.h), s.count))
      .join(', ');
    return { ...state, reason: t().mixedSizes(sizes.length, list) };
  }

  const rotated = frames.filter((f) => 'rotation' in f && Math.abs(f.rotation) > 0.01);
  if (rotated.length > 0) {
    return { ...state, reason: t().rotatedFrame(rotated[0].name) };
  }

  // resolveSlide() 가 PowerPoint 한계(1~56인치)를 이미 본다.
  // 실측이 한계를 넘더라도 비율이 맞는 표준 크기가 있으면 그쪽으로 내보낼 수 있다.
  if (!plan) {
    const mm = `${fmt(pxToMm(width))}×${fmt(pxToMm(height))}mm`;
    const shortest = Math.min(width, height);
    return {
      ...state,
      reason: shortest / 72 < 1 ? t().tooSmall(mm) : t().tooLarge(mm),
    };
  }

  return { ...state, ok: true, reason: '' };
}

function fmt(n: number): string {
  return (Math.round(n * 10) / 10).toString();
}
