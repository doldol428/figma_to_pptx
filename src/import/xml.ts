/**
 * 최소 XML 파서.
 *
 * UI 스레드에는 DOMParser 가 있지만 쓰지 않는다. OOXML 은 네임스페이스가 잔뜩 붙어 있어
 * (`a:`, `p:`, `r:`) getElementsByTagNameNS 를 계속 써야 하고, 파트마다 접두사가 달라진다.
 * 접두사를 무시하고 지역명으로만 찾는 편이 훨씬 짧고, 파서 자체도 200줄이 안 된다.
 *
 * OOXML 은 항상 잘 정형화된 XML 이라 오류 복구가 필요 없다. CDATA 도 쓰지 않는다.
 */

export interface XNode {
  /** 접두사를 뗀 지역명 */
  tag: string;
  attrs: Record<string, string>;
  children: XNode[];
  text: string;
}

const TOKEN = /<(\/?)([\w:.-]+)((?:\s+[\w:.-]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>|<\?[^>]*\?>|<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>/g;
const ATTR = /([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

const ENTITY: Record<string, string> = {
  amp: '&', lt: '<', gt: '>', quot: '"', apos: "'",
};

export function decodeEntities(s: string): string {
  if (s.indexOf('&') < 0) return s;
  return s.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, (whole, e: string) => {
    if (e[0] === '#') {
      const code = e[1] === 'x' || e[1] === 'X'
        ? parseInt(e.slice(2), 16)
        : Number(e.slice(1));
      return Number.isFinite(code) ? String.fromCodePoint(code) : whole;
    }
    return ENTITY[e] ?? whole;
  });
}

export function parseXml(source: string): XNode {
  const root: XNode = { tag: '#root', attrs: {}, children: [], text: '' };
  const stack: XNode[] = [root];
  let last = 0;
  let m: RegExpExecArray | null;

  TOKEN.lastIndex = 0;
  while ((m = TOKEN.exec(source)) !== null) {
    const between = source.slice(last, m.index);
    last = TOKEN.lastIndex;

    // 공백만 있는 조각은 버린다. <a:t> 안의 실제 공백은 앞뒤에 태그가 붙어 살아남는다.
    if (between !== '' && stack.length > 1) {
      const top = stack[stack.length - 1];
      top.text += decodeEntities(between);
    }

    if (m[2] === undefined) continue; // 선언 · 주석 · CDATA

    if (m[1] === '/') {
      if (stack.length > 1) stack.pop();
      continue;
    }

    const attrs: Record<string, string> = {};
    if (m[3]) {
      ATTR.lastIndex = 0;
      let a: RegExpExecArray | null;
      while ((a = ATTR.exec(m[3])) !== null) {
        attrs[a[1]] = decodeEntities(a[2] ?? a[3] ?? '');
      }
    }

    const node: XNode = {
      tag: m[2].replace(/^[\w.-]+:/, ''),
      attrs,
      children: [],
      text: '',
    };
    stack[stack.length - 1].children.push(node);
    if (m[4] !== '/') stack.push(node);
  }

  return root;
}

/* ── 탐색 도우미 ──────────────────────────────────────────────── */

/** 직계 자식 중 태그가 일치하는 것 전부 */
export function all(node: XNode | null | undefined, tag: string): XNode[] {
  if (!node) return [];
  return node.children.filter((c) => c.tag === tag);
}

/** 직계 자식 중 첫 번째 */
export function one(node: XNode | null | undefined, tag: string): XNode | null {
  if (!node) return null;
  for (const c of node.children) if (c.tag === tag) return c;
  return null;
}

/** 경로를 따라 내려간다. one(one(n,'a'),'b') 를 path(n,'a','b') 로. */
export function path(node: XNode | null | undefined, ...tags: string[]): XNode | null {
  let cur = node ?? null;
  for (const t of tags) {
    cur = one(cur, t);
    if (!cur) return null;
  }
  return cur;
}

/** 자손 전체에서 첫 일치 (깊이 우선) */
export function deep(node: XNode | null | undefined, tag: string): XNode | null {
  if (!node) return null;
  for (const c of node.children) {
    if (c.tag === tag) return c;
    const found = deep(c, tag);
    if (found) return found;
  }
  return null;
}

export function num(value: string | undefined, fallback = 0): number {
  if (value === undefined) return fallback;
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

export function bool(value: string | undefined, fallback = false): boolean {
  if (value === undefined) return fallback;
  return value === '1' || value === 'true';
}
