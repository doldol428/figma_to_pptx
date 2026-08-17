import { pathBounds, translatePath } from '../import/pathbox';
import { connectorPath, nativeFor, presetPath } from '../import/preset';
import { KEY } from '../shared/roles';
import { parsePath } from './path';

/**
 * 벡터를 원래의 preset 도형으로 되돌린다 — **지금 경로가 그것과 같을 때만**.
 *
 * 가져오기가 `hexagon` 같은 preset 을 경로로 펴서 벡터로 만든다. 그대로 내보내면 custGeom
 * 이 되어 파워포인트에서 편집할 수 없고, 상자가 경로에 맞게 조여진 도형(호·파이)은 크기까지
 * 달라진다. 그래서 가져올 때 이름과 조절값을 적어 둔다.
 *
 * 다만 **적어 둔 이름을 그냥 믿으면 안 된다.** Figma 에서 경로를 고쳤는데 preset 으로
 * 되돌리면 그 편집이 통째로 날아간다. 그래서 그 이름·조절값으로 지금 크기에 다시 그려 보고,
 * 경로가 같을 때만 되돌린다. 크기 변경은 다시 그리면 따라오므로 그대로 통과한다.
 */
export interface RestoredPreset {
  prst: string;
  adj: Record<string, number>;
  /**
   * 이름 상자 — 경로가 아니라 preset 이 원래 차지하던 상자. 노드 로컬 px.
   * 호처럼 경로가 상자 일부만 덮는 도형은 이 상자를 되살려야 크기가 맞는다.
   */
  box: { dx: number; dy: number; w: number; h: number };
}

interface Stored {
  prst?: string;
  adj?: Record<string, number>;
  w?: number;
  h?: number;
}

/** 좌표 허용 오차 (px). 우리 베지어 근사 오차보다는 크고, 눈에 띄는 변형보다는 작게. */
const EPS = 0.05;

export function restorePreset(node: SceneNode): RestoredPreset | null {
  const raw = node.getPluginData(KEY.preset);
  if (!raw) return null;

  let stored: Stored;
  try {
    stored = JSON.parse(raw) as Stored;
  } catch {
    return null;
  }
  const { prst, adj, w, h } = stored;
  if (!prst || !adj || !w || !h) return null;

  const paths = (node as SceneNode & Partial<VectorNode>).vectorPaths;
  if (!Array.isArray(paths) || paths.length === 0) {
    return byCorners(node, prst, adj, w, h);
  }

  /*
   * 배수는 **경로가 실제로 덮는 범위**로 잰다. 노드 크기로 재면 안 된다 — Figma 는 선 굵기
   * 만큼 노드를 부풀리므로 테두리가 있는 도형은 노드가 경로보다 크고, 그만큼 어긋난 크기로
   * 다시 그리게 되어 멀쩡한 도형이 전부 대조에서 떨어진다.
   */
  const actualData = paths.map((p) => p.data).join(' ');
  const at = pathBounds(actualData);
  if (!(at.w > 0) && !(at.h > 0)) return null;

  const first = drawPreset(prst, w, h, adj);
  if (!first) return null;
  const base = pathBounds(first.data);
  if (!(base.w > 0) || !(base.h > 0)) return null;

  const scaleX = at.w / base.w;
  const scaleY = at.h / base.h;
  const drawn = drawPreset(prst, w * scaleX, h * scaleY, adj);
  if (!drawn) return null;

  const box = pathBounds(drawn.data);
  if (!(box.w > 0) || !(box.h > 0)) return null;

  // 다시 그린 것을 지금 경로가 놓인 자리로 옮겨 놓고 견준다.
  const expected = parsePath(translatePath(drawn.data, at.x - box.x, at.y - box.y));
  if (!samePath(expected, parsePath(actualData))) return null;

  return {
    prst,
    adj,
    // 이름 상자의 좌상단은 경로 상자보다 그만큼 앞선다.
    box: { dx: at.x - box.x, dy: at.y - box.y, w: w * scaleX, h: h * scaleY },
  };
}

/** preset 이든 연결선이든 같은 자리에서 그린다 — 둘 다 이름으로 경로가 정해진다. */
function drawPreset(
  prst: string, w: number, h: number, adj: Record<string, number>,
): { data: string } | null {
  return presetPath(prst, w, h, adj) ?? connectorPath(prst, w, h);
}

/**
 * 모서리별 반경이 다른 사각형.
 *
 * `round2SameRect` 같은 것은 경로가 아니라 Figma 사각형으로 들어온다. 내보낼 때는 반경이
 * 제각각이라 결국 경로가 되므로, 여기서 이름으로 되돌린다. 대조는 네 모서리 반경으로 한다 —
 * 사람이 반경을 바꿨으면 값이 달라져 걸러진다.
 */
function byCorners(
  node: SceneNode, prst: string, adj: Record<string, number>, w: number, h: number,
): RestoredPreset | null {
  const n = node as SceneNode & Partial<RectangleCornerMixin>;
  if (typeof n.topLeftRadius !== 'number') return null;

  const scaleX = node.width / w;
  const scaleY = node.height / h;
  const want = nativeFor(prst, w * scaleX, h * scaleY, adj);
  if (!want || want.kind !== 'roundRect') return null;

  const have = [n.topLeftRadius, n.topRightRadius ?? 0, n.bottomRightRadius ?? 0,
    n.bottomLeftRadius ?? 0];
  for (let i = 0; i < 4; i++) if (Math.abs(want.radii[i] - have[i]) > EPS) return null;

  return { prst, adj, box: { dx: 0, dy: 0, w: node.width, h: node.height } };
}

type Point = ReturnType<typeof parsePath>[number];

/** 두 경로가 같은가 — 명령 순서와 좌표를 그대로 견준다. */
function samePath(a: Point[], b: Point[]): boolean {
  if (a.length !== b.length || a.length === 0) return false;
  for (let i = 0; i < a.length; i++) {
    const p = a[i];
    const q = b[i];
    if ('close' in p || 'close' in q) {
      if (!('close' in p) || !('close' in q)) return false;
      continue;
    }
    if (('c' in p) !== ('c' in q)) return false;
    if (('moveTo' in p ? p.moveTo : false) !== ('moveTo' in q ? q.moveTo : false)) return false;
    if (Math.abs(p.x - q.x) > EPS || Math.abs(p.y - q.y) > EPS) return false;
    if ('c' in p && 'c' in q) {
      for (let k = 0; k < 4; k++) if (Math.abs(p.c[k] - q.c[k]) > EPS) return false;
    }
  }
  return true;
}
