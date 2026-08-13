/**
 * 한글 글꼴 이름 → Figma 등록명.
 *
 * Figma 데스크탑(Windows)이 돌려주는 2163종의 family 이름에는 **한글이 하나도 없다.**
 * 굴림도 `Gulim`, 나눔고딕도 `NanumGothic`, KoPub 돋움체도 `KoPubDotum` 으로 등록된다.
 * 즉 Figma 는 폰트 파일의 영문 이름(name ID 1, en-US)만 읽는다.
 *
 * 반면 PPTX 의 `<a:ea typeface="...">` 에는 파워포인트가 보여주는 한글 이름이 그대로 들어간다.
 * 그래서 이름을 그대로 찾으면 설치돼 있는 글꼴도 전부 "없음" 이 된다.
 *
 * 여기서는 이름의 한글 부분만 영문으로 바꾼 후보들을 만든다. 굵기가 뒤에 붙어 있어도
 * (`KoPub돋움체 Bold`) 그대로 남으므로 family/style 분리는 기존 로직이 이어서 처리한다.
 * 존재하지 않는 후보가 섞여도 설치 목록에 없으면 그냥 버려지니 넉넉하게 만든다.
 */

/** 긴 이름이 먼저 와야 한다 — `KoPub돋움체` 가 `돋움체` 규칙에 먼저 잡히면 안 된다. */
const ALIASES: Array<[string, string[]]> = [
  // KoPub — 배포판마다 World / _Pro 접미사가 다르다
  ['KoPub돋움체', ['KoPubDotum', 'KoPubWorldDotum', 'KoPubWorldDotum_Pro', 'KoPub Dotum']],
  ['KoPub바탕체', ['KoPubBatang', 'KoPubWorldBatang', 'KoPubWorldBatang_Pro', 'KoPub Batang']],
  ['KoPub돋움', ['KoPubDotum', 'KoPubWorldDotum']],
  ['KoPub바탕', ['KoPubBatang', 'KoPubWorldBatang']],

  // 나눔
  ['나눔고딕코딩', ['NanumGothicCoding']],
  ['나눔스퀘어라운드', ['NanumSquareRound', 'NanumSquareRoundOTF']],
  ['나눔스퀘어', ['NanumSquare', 'NanumSquareOTF', 'NanumSquare_ac']],
  ['나눔바른고딕', ['NanumBarunGothic']],
  ['나눔고딕', ['NanumGothic']],
  ['나눔명조', ['NanumMyeongjo']],
  ['나눔손글씨 펜', ['Nanum Pen']],
  ['나눔손글씨 붓', ['Nanum Brush Script']],

  // Windows 기본
  ['맑은 고딕', ['Malgun Gothic']],
  ['맑은고딕', ['Malgun Gothic']],
  ['새굴림', ['New Gulim']],
  ['굴림체', ['GulimChe']],
  ['굴림', ['Gulim']],
  ['돋움체', ['DotumChe']],
  ['돋움', ['Dotum']],
  ['바탕체', ['BatangChe']],
  ['바탕', ['Batang']],
  ['궁서체', ['GungsuhChe']],
  ['궁서', ['Gungsuh']],

  // HY — 한글 이름과 영문 등록명이 전혀 닮지 않아 표가 없으면 못 찾는다
  ['HY견고딕', ['HYGothic-Extra']],
  ['HY중고딕', ['HYGothic-Medium']],
  ['HY그래픽M', ['HYGraphic-Medium']],
  ['HY그래픽', ['HYGraphic-Medium']],
  ['HY궁서B', ['HYGungSo-Bold']],
  ['HY궁서', ['HYGungSo-Bold']],
  ['HY헤드라인M', ['HYHeadLine-Medium']],
  ['HY견명조', ['HYMyeongJo-Extra']],
  ['HY목각파임B', ['HYPMokGak-Bold']],
  ['HY엽서L', ['HYPost-Light']],
  ['HY엽서M', ['HYPost-Medium']],
  ['HY얕은샘물M', ['HYShortSamul-Medium']],
  ['HY신명조', ['HYSinMyeongJo-Medium']],

  // 지자체·기업 배포 글꼴
  ['서울남산체', ['SeoulNamsan']],
  ['서울남산', ['SeoulNamsan']],
  ['서울한강체', ['SeoulHangang']],
  ['서울한강', ['SeoulHangang']],
  ['제주고딕', ['JejuGothic']],
  ['제주명조', ['JejuMyeongjo']],
  ['제주한라산', ['JejuHallasan']],
  ['배달의민족 한나', ['BM HANNA_TTF']],
  ['지마켓 산스', ['Gmarket Sans TTF']],
  ['에스코어 드림', ['S-Core Dream']],

  // 본문용
  ['페이퍼로지', ['Paperlogy']],
  ['프리텐다드', ['Pretendard']],
  ['본고딕', ['Noto Sans KR', 'Source Han Sans KR']],
  ['본명조', ['Noto Serif KR', 'Source Han Serif KR']],
  ['함초롬돋움', ['HCR Dotum', 'HCR Dotum LVT']],
  ['함초롬바탕', ['HCR Batang', 'HCR Batang LVT']],
];

/** 이름 앞머리의 영문 부분. `KoPub돋움체` → `KoPub` — 힌트를 찾을 때 쓴다. */
export function latinHead(typeface: string): string {
  const m = /^[A-Za-z][A-Za-z0-9 .\-_]*/.exec(typeface);
  return m ? m[0].trim() : '';
}

/**
 * 찾아볼 다른 이름들. 원래 이름은 포함하지 않는다 (부르는 쪽이 먼저 시도한다).
 * 한글이 없으면 바꿀 것도 없으므로 빈 배열이다.
 */
export function aliasesFor(typeface: string): string[] {
  if (!/[가-힣]/.test(typeface)) return [];

  const out: string[] = [];
  for (const [korean, latins] of ALIASES) {
    if (typeface.indexOf(korean) < 0) continue;
    for (const latin of latins) {
      const candidate = typeface.split(korean).join(latin).trim();
      if (candidate && out.indexOf(candidate) < 0) out.push(candidate);
    }
  }
  return out;
}
