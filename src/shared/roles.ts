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
  /**
   * 도형 하나에 글자가 얹힌 것.
   *
   * PPTX 는 이것을 `<p:sp>` **하나**로 쓴다 — 도형과 글자가 한 몸이라 같이 움직이고 같이 지워진다.
   * Figma 에는 그런 개념이 없어 가져올 때 도형 노드와 글자 노드로 갈라 그룹에 담는다.
   * 표시를 남겨 두지 않으면 내보낼 때 둘로 갈라진 채 나가, 파워포인트에서 글자만 따로 노는
   * 텍스트 상자가 된다 (실측: 테두리 있는 상자 8개가 16개로).
   */
  | 'shapeText'
  /**
   * 표.
   *
   * Figma 에는 표 원시 타입이 없어 프레임 격자로 재구성한다. 표시가 없으면 내보낼 때
   * 그 격자가 도형과 선 수백 개로 풀려, 파워포인트에서 행을 넣거나 열 너비를 바꿀 수 없게 된다.
   */
  | 'table'
  /** 그 슬라이드 고유 내용 */
  | 'content';

/** 레이어 이름. `#레이아웃/제목 슬라이드` 처럼 뒤에 레이아웃 이름이 붙는다. */
export const NAME = {
  layout: `${MARK}레이아웃`,
  slideNumber: `${MARK}슬라이드번호`,
  /** 도형+글자 그룹. 원래 이름을 뒤에 붙여 `#도형글자/직사각형 5` 처럼 쓴다. */
  shapeText: `${MARK}도형글자`,
  /** 표 격자. `#표/표 1` 처럼 쓴다. */
  table: `${MARK}표`,
} as const;

export function layoutName(name: string): string {
  return name ? `${NAME.layout}/${name}` : NAME.layout;
}

/** 이름만 보고 역할을 되읽는다 — 내보내기가 쓰는 쪽. */
export function roleFromName(name: string): Role | null {
  if (name === NAME.slideNumber) return 'slideNumber';
  if (name === NAME.layout || name.indexOf(`${NAME.layout}/`) === 0) return 'layout';
  if (name === NAME.shapeText || name.indexOf(`${NAME.shapeText}/`) === 0) return 'shapeText';
  if (name === NAME.table || name.indexOf(`${NAME.table}/`) === 0) return 'table';
  return null;
}

/** `#도형글자/직사각형 5` → `직사각형 5` */
export function nameAfterMark(name: string): string {
  const slash = name.indexOf('/');
  return slash >= 0 ? name.slice(slash + 1) : name;
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
  /**
   * 무늬 채우기 원본 (ir.Pattern 의 JSON).
   *
   * Figma 에 무늬가 없어 단색으로 눌러 두는데, 그 사실을 여기 말고는 적어 둘 곳이 없다.
   * 무늬는 Figma 에서 편집할 방법 자체가 없으므로 이 기록이 낡지 않는다 — 확인 없이
   * 되돌려도 안전한 드문 경우다. 다만 색은 바꿀 수 있어 내보낼 때 단색이 그대로인지만 본다.
   */
  pattern: 'pattern',
  /**
   * 표의 칸 병합 (JSON: `{ spans: [[[colSpan, rowSpan, merged]]] }`).
   *
   * 칸 크기와 색·글자·테두리는 노드에서 그대로 읽으면 되지만, 병합은 프레임 격자에 남지
   * 않는다. Figma 에서 칸을 병합할 방법 자체가 없으니 이 기록이 낡을 일도 없다.
   */
  table: 'table',
} as const;
