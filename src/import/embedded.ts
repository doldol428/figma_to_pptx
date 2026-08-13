import type JSZip from 'jszip';
import { all, deep, type XNode } from './xml';

/**
 * PPTX 에 박혀 있는 글꼴(`ppt/fonts/*.fntdata`)의 **이름표만** 읽는다.
 *
 * 파워포인트는 글꼴을 문서에 포함할 수 있고, 그래서 설치돼 있지 않아도 원본은 멀쩡히 보인다.
 * Figma 는 포함된 글꼴을 쓰지 못하므로 그 글꼴은 결국 대체되지만, 최소한 **이름은 정확히**
 * 알 수 있다 — EOT 헤더에 한글 이름(FamilyName)과 영문 이름(FullName)이 나란히 들어 있다.
 *
 *   "페이퍼로지 4 Regular"  →  "Paperlogy 4 Regular"
 *   "프리젠테이션 4 Regular" →  "Freesentation 4 Regular"
 *   "KoPub돋움체 Bold"      →  "KoPubDotum Bold"
 *
 * Figma 는 영문 이름으로만 글꼴을 등록하므로(설치 목록 2163종에 한글 이름이 하나도 없다)
 * 이 대응표가 있으면 손으로 만든 별칭표보다 정확하다 — 문서가 직접 정답을 들고 있는 셈이다.
 *
 * 글꼴 본체는 MicroType Express 로 압축돼 있어 여기서 꺼내 쓸 수는 없다. 헤더는 압축 대상이
 * 아니라 그대로 읽힌다.
 */

/** EOT 헤더에서 MagicNumber 가 있어야 하는 위치와 값 ('LP') */
const MAGIC_OFFSET = 34;
const MAGIC = 0x504c;

interface EotNames {
  /** 문서가 적어둔 이름 — 보통 한글 */
  family: string;
  /** 영문 전체 이름 */
  fullName: string;
}

/**
 * EOT 헤더를 읽어 이름 두 개를 꺼낸다. 형식이 안 맞으면 null — 추측하지 않는다.
 *
 * 레이아웃은 EOT 명세를 그대로 따른다. 문자열은 전부 (padding u16, 길이 u16, UTF-16LE) 묶음이다.
 */
export function readEotNames(u8: Uint8Array): EotNames | null {
  if (u8.length < 84) return null;
  const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  if (dv.getUint16(MAGIC_OFFSET, true) !== MAGIC) return null;

  let p = 80;
  const str = (): string => {
    p += 2;                                   // padding
    const n = dv.getUint16(p, true);
    p += 2;
    const bytes = u8.subarray(p, p + n);
    p += n;
    let s = '';
    for (let i = 0; i + 1 < bytes.length; i += 2) {
      s += String.fromCharCode(bytes[i] | (bytes[i + 1] << 8));
    }
    return s.replace(/\0/g, '').trim();
  };

  try {
    const family = str();
    str();                                    // StyleName
    str();                                    // VersionName
    const fullName = str();
    if (!family || !fullName) return null;
    return { family, fullName };
  } catch {
    return null;
  }
}

/**
 * presentation.xml 의 `<p:embeddedFont>` 를 훑어 "문서가 쓴 이름 → 영문 이름" 표를 만든다.
 *
 * 한글이 없는 이름(Arial, Wingdings)은 건너뛴다 — 바꿀 것이 없고, 글꼴 본체를 풀어내는 비용만 든다.
 * 같은 typeface 가 굵기별로 여러 번 나오므로 이름 기준으로 한 번만 읽는다.
 */
export async function readEmbeddedFontNames(
  zip: JSZip,
  presentation: XNode | null,
  rels: Record<string, string>,
): Promise<Record<string, string>> {
  const out: Record<string, string> = {};
  if (!presentation) return out;

  for (const entry of all(deep(presentation, 'embeddedFontLst'), 'embeddedFont')) {
    const typeface = deep(entry, 'font')?.attrs.typeface;
    if (!typeface || !/[가-힣]/.test(typeface)) continue;
    if (out[typeface]) continue;

    for (const child of entry.children) {
      const id = child.attrs['r:id'] ?? child.attrs.id;
      const target = id ? rels[id] : undefined;
      if (!target) continue;
      // rels 의 Target 은 ppt/ 기준 상대 경로다 (slide 경로 해석과 같은 규칙).
      const file = zip.file(`ppt/${target.replace(/^\.\.\//, '')}`);
      if (!file) continue;
      const names = readEotNames(await file.async('uint8array'));
      if (!names) continue;
      // 영문 이름이 원래 이름과 같으면 표에 넣을 이유가 없다.
      if (names.fullName !== typeface) out[typeface] = names.fullName;
      break;
    }
  }
  return out;
}
