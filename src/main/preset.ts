import { pathBounds, translatePath } from '../import/pathbox';
import { presetPath } from '../import/preset';
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
   * 이름 상자 — 경로가 아니라 preset 이 원래 차지하던 상자.
   * 벡터 상자에 대한 배수·오프셋으로 적는다 (크기를 바꿔도 따라오게).
   */
  box: { x: number; y: number; w: number; h: number };
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
  if (!Array.isArray(paths) || paths.length === 0) return null;

  /*
   * 가져올 때와 같은 배수로 다시 그린다. 도형을 늘렸으면 이름 상자도 같은 비율로 늘어난다 —
   * 경로 상자와 이름 상자의 비율은 preset 이 정하는 것이라 크기가 변해도 유지된다.
   */
  const first = presetPath(prst, w, h, adj);
  if (!first) return null;
  const nominal = pathBounds(first.data);
  if (!(nominal.w > 0) || !(nominal.h > 0)) return null;

  const scaleX = node.width / nominal.w;
  const scaleY = node.height / nominal.h;
  const drawn = presetPath(prst, w * scaleX, h * scaleY, adj);
  if (!drawn) return null;

  const box = pathBounds(drawn.data);
  if (!(box.w > 0) || !(box.h > 0)) return null;

  const expected = parsePath(translatePath(drawn.data, -box.x, -box.y));
  const actual = parsePath(paths.map((p) => p.data).join(' '));
  if (!samePath(expected, actual)) return null;

  return {
    prst,
    adj,
    box: {
      x: -box.x / box.w,
      y: -box.y / box.h,
      w: (w * scaleX) / box.w,
      h: (h * scaleY) / box.h,
    },
  };
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
