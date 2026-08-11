import type { Run, TextItem } from '../shared/ir';
import { toHex, topVisiblePaint } from './paint';

const ITALIC_RE = /\b(italic|oblique)\b/i;
const BOLD_RE = /^(bold|black|heavy|extra ?bold|ultra ?bold)$/i;
const REGULAR_RE = /^(regular|normal|book|roman|)$/i;

/**
 * Figma FontName → PPTX 런 속성.
 *
 * PPTX 런에는 굵기 축이 없고 bold 플래그뿐이다. Bold 이상은 플래그로 올리고,
 * SemiBold/Light 처럼 중간 굵기는 "가족명 + 스타일" 을 폰트 이름으로 넘긴다.
 * 한글 폰트(Pretendard, Noto Sans KR 등)는 굵기별로 별도 패밀리로 설치되는 경우가 많아
 * 이 방식이 PowerPoint 에서 실제로 잘 붙는다.
 */
export function mapFont(font: FontName): { fontFace: string; bold: boolean; italic: boolean } {
  const italic = ITALIC_RE.test(font.style);
  const weight = font.style.replace(ITALIC_RE, '').replace(/\s+/g, ' ').trim();
  const bold = BOLD_RE.test(weight);

  let fontFace = font.family;
  if (!bold && !REGULAR_RE.test(weight)) {
    fontFace = `${font.family} ${weight}`;
  }
  return { fontFace, bold, italic };
}

function applyTextCase(text: string, textCase: TextCase): string {
  switch (textCase) {
    case 'UPPER':
      return text.toUpperCase();
    case 'LOWER':
      return text.toLowerCase();
    case 'TITLE':
      return text.replace(/\b\w/g, (c) => c.toUpperCase());
    default:
      return text;
  }
}

/** buildRuns() 가 실제로 요청하는 필드만 골라 쓴다 */
type Segment = Pick<
  StyledTextSegment,
  'characters' | 'start' | 'end' | 'fontSize' | 'fontName' | 'fills'
  | 'textDecoration' | 'textCase' | 'letterSpacing' | 'lineHeight' | 'hyperlink'
>;

function segmentToRunBase(seg: Segment): Omit<Run, 'text' | 'breakLine'> {
  const font = seg.fontName as FontName;
  const { fontFace, bold, italic } = mapFont(font);
  const fontSize = seg.fontSize as number;

  const paint = topVisiblePaint(seg.fills as Paint[]);
  let color = '000000';
  let transparency = 0;
  if (paint && paint.type === 'SOLID') {
    color = toHex(paint.color);
    transparency = Math.round((1 - (paint.opacity ?? 1)) * 100 * 10) / 10;
  }

  const run: Omit<Run, 'text' | 'breakLine'> = {
    fontFace,
    fontSize,
    bold,
    italic,
    underline: seg.textDecoration === 'UNDERLINE',
    strike: seg.textDecoration === 'STRIKETHROUGH',
    color,
    transparency,
  };

  const ls = seg.letterSpacing as LetterSpacing;
  if (ls && ls.value !== 0) {
    run.charSpacing = ls.unit === 'PERCENT' ? (fontSize * ls.value) / 100 : ls.value;
  }

  const lh = seg.lineHeight as LineHeight;
  if (lh && lh.unit === 'PIXELS') {
    run.lineSpacing = lh.value;
  } else if (lh && lh.unit === 'PERCENT') {
    // Figma 의 PERCENT 는 폰트 크기 기준. 고정 pt 로 환산해야 PowerPoint 에서 어긋나지 않는다.
    run.lineSpacing = (fontSize * lh.value) / 100;
  }

  const link = seg.hyperlink as HyperlinkTarget | null;
  if (link && link.type === 'URL') {
    run.hyperlink = link.value;
  }

  return run;
}

// U+2028 = 줄 바꿈(Shift+Enter), U+2029 = 문단 구분. Figma 는 둘 다 쓴다.
const LINE_BREAK = /\r\n|[\r\n\u2028\u2029]/;

/**
 * 스타일 세그먼트 목록 → PPTX 런 배열.
 * 줄바꿈 문자는 런에서 제거하고 직전 런에 breakLine 을 세워 문단을 나눈다.
 */
export function buildRuns(node: TextNode): Run[] {
  const segments = node.getStyledTextSegments([
    'fontSize',
    'fontName',
    'fills',
    'textDecoration',
    'textCase',
    'letterSpacing',
    'lineHeight',
    'hyperlink',
  ]);

  const runs: Run[] = [];
  for (const seg of segments) {
    const base = segmentToRunBase(seg);
    const text = applyTextCase(seg.characters, seg.textCase as TextCase);
    const lines = text.split(LINE_BREAK);
    lines.forEach((line, i) => {
      const isLast = i === lines.length - 1;
      // 마지막 조각이 비어 있고 앞에 줄바꿈이 있었다면, 앞 런의 breakLine 으로 이미 표현된다.
      if (isLast && line === '' && lines.length > 1) {
        return;
      }
      runs.push(isLast ? { ...base, text: line } : { ...base, text: line, breakLine: true });
    });
  }

  if (runs.length === 0) {
    return [];
  }
  return runs;
}

export function textAlign(node: TextNode): TextItem['align'] {
  switch (node.textAlignHorizontal) {
    case 'CENTER':
      return 'center';
    case 'RIGHT':
      return 'right';
    case 'JUSTIFIED':
      return 'justify';
    default:
      return 'left';
  }
}

export function textValign(node: TextNode): TextItem['valign'] {
  switch (node.textAlignVertical) {
    case 'CENTER':
      return 'middle';
    case 'BOTTOM':
      return 'bottom';
    default:
      return 'top';
  }
}
