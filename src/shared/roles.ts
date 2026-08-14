/**
 * 왕복(PPTX → Figma → PPTX)을 위한 이름·메타데이터 규칙.
 *
 * 가져오기가 역할을 버리면 내보내기가 그것을 되살릴 방법이 없다. "텍스트 개체 틀 3" 같은
 * PowerPoint 이름은 도움이 되지 않는다 — 실제 이 문서에는 `<p:ph>` 자리표시자가 하나도 없고,
 * 저 이름은 예전에 자리표시자였던 흔적일 뿐이다. 그래서 역할은 **출처**로 정한다.
 * 레이아웃 파트에서 온 것이면 페이지 공통 요소, 슬라이드에서 온 것이면 그 장의 내용이다.
 *
 * 규칙은 두 겹이다.
 *
 * 1. **이름** — 사람이 읽고 직접 쓸 수 있다. Figma 에서 새로 만든 프레임도 이 이름만
 *    지키면 내보낼 때 같은 대접을 받는다.
 * 2. **pluginData** — 기계가 읽는 정확한 기록. 이름을 바꿔도 남고, 원본 PPTX 이름까지 들고 있다.
 *
 * 이름이 규칙이고 pluginData 는 보강이다. 이름을 지웠다고 왕복이 깨지지는 않지만,
 * 이름만 맞춰도 왕복은 된다.
 */

/** 이름 앞에 붙는 표시. 사용자가 지은 이름과 섞이지 않게 한다. */
export const MARK = '#';

export type Role =
  /** 레이아웃에서 온 페이지 공통 요소 — 컴포넌트 하나로 묶고 각 장에는 인스턴스를 놓는다 */
  | 'layout'
  /** 장마다 값이 달라 공통으로 묶을 수 없는 것 (슬라이드 번호) */
  | 'slideNumber'
  /** 그 슬라이드 고유 내용 */
  | 'content';

/** 레이어 이름. `#레이아웃/제목 슬라이드` 처럼 뒤에 레이아웃 이름이 붙는다. */
export const NAME = {
  layout: `${MARK}레이아웃`,
  slideNumber: `${MARK}슬라이드번호`,
} as const;

export function layoutName(name: string): string {
  return name ? `${NAME.layout}/${name}` : NAME.layout;
}

/** 이름만 보고 역할을 되읽는다 — 내보내기가 쓰는 쪽. */
export function roleFromName(name: string): Role | null {
  if (name === NAME.slideNumber) return 'slideNumber';
  if (name === NAME.layout || name.indexOf(`${NAME.layout}/`) === 0) return 'layout';
  return null;
}

/**
 * pluginData 키.
 * Figma 가 플러그인별로 이름 공간을 나눠 주므로 다른 플러그인과 부딪히지 않는다.
 */
export const KEY = {
  /** Role */
  role: 'role',
  /** 레이아웃 식별자 (레이아웃 파트 파일명) */
  layout: 'layout',
  /** 원본 PPTX 도형 이름 — 되돌릴 때 그대로 쓴다 */
  pptxName: 'pptxName',
  /** 슬라이드 번호 (1부터) */
  slide: 'slide',
} as const;
