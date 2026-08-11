import type { SelectionState, SizeInfo } from '../shared/ir';
import { MAX_SLIDE_IN, MIN_SLIDE_IN, detectPaper, pxToIn, pxToMm } from '../shared/units';

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
    paper: null,
    sizes: [],
  };

  if (frames.length === 0) {
    return { ...base, reason: '내보낼 프레임을 선택하세요. (프레임 / 컴포넌트 / 인스턴스)' };
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
  const state: SelectionState = {
    ...base,
    widthPx: width,
    heightPx: height,
    paper: detectPaper(width, height),
    sizes,
  };

  if (sizes.length > 1) {
    // PPTX 는 파일 하나에 슬라이드 크기가 하나뿐이다. 섞이면 어느 쪽도 원본 비율을 지킬 수 없다.
    const list = sizes
      .map((s) => `${fmt(s.w)}×${fmt(s.h)}px (${s.count}개)`)
      .join(', ');
    return {
      ...state,
      reason: `프레임 크기가 ${sizes.length}종류로 섞여 있습니다: ${list}\nPPTX 는 파일당 슬라이드 크기가 하나뿐이라, 같은 크기끼리만 선택해 주세요.`,
    };
  }

  const rotated = frames.filter((f) => 'rotation' in f && Math.abs(f.rotation) > 0.01);
  if (rotated.length > 0) {
    return {
      ...state,
      reason: `회전된 프레임이 있습니다 (${rotated[0].name}). 슬라이드 기준이 되는 프레임은 회전이 0 이어야 합니다.`,
    };
  }

  const wIn = pxToIn(width);
  const hIn = pxToIn(height);
  if (wIn < MIN_SLIDE_IN || hIn < MIN_SLIDE_IN) {
    return {
      ...state,
      reason: `프레임이 너무 작습니다 (${fmt(pxToMm(width))}×${fmt(pxToMm(height))}mm). PowerPoint 슬라이드는 한 변이 최소 1인치(25.4mm)여야 합니다.`,
    };
  }
  if (wIn > MAX_SLIDE_IN || hIn > MAX_SLIDE_IN) {
    return {
      ...state,
      reason: `프레임이 너무 큽니다 (${fmt(pxToMm(width))}×${fmt(pxToMm(height))}mm). PowerPoint 슬라이드는 한 변이 최대 56인치(1422mm)입니다.`,
    };
  }

  return { ...state, ok: true, reason: '' };
}

function fmt(n: number): string {
  return (Math.round(n * 10) / 10).toString();
}
