/**
 * 한국어 / 영어 2개 언어만 지원한다. 그 외 언어는 전부 영어로 떨어진다.
 *
 * 주의: Figma 플러그인 API 에는 **계정 언어를 읽는 수단이 없다.** PluginAPI 에 locale 속성이
 * 없고 codegen 의 `language` 는 코드 언어라 무관하다. 그래서 UI iframe 의
 * `navigator.languages`(데스크톱 앱이면 Electron 로케일, 웹이면 브라우저 언어)로 판정하고,
 * UI 가 시작할 때 main 스레드에 알려준다.
 *
 * 로케일은 실행 중 바뀌지 않으므로 모듈 전역에 한 번 심고 `t()` 로 꺼내 쓴다.
 * 두 스레드가 각자 자기 사본을 설정한다.
 */

export type Locale = 'ko' | 'en';

export function resolveLocale(tags: readonly string[]): Locale {
  return tags.some((tag) => tag.toLowerCase().startsWith('ko')) ? 'ko' : 'en';
}

export interface Strings {
  numberLocale: string;

  /* UI */
  slideSize: string;
  selectFrame: string;
  frameCount: (n: number) => string;
  uniformScale: (k: number) => string;
  imageResolution: string;
  dpiScreen: string;
  dpiDefault: string;
  dpiPrint: string;
  exportButton: string;
  building: string;
  warningsTitle: (n: number) => string;
  buildFailed: (err: string) => string;
  exported: (slides: number, size: string) => string;
  exportFailed: string;

  /** 규격 칩 — "A4 세로" / "A4 Portrait" */
  paperChip: (name: string, portrait: boolean) => string;

  /* 선택 검증 */
  noFrameSelected: string;
  mixedSizes: (kinds: number) => string;
  rotatedFrame: (name: string) => string;
  sizeOutOfRange: (mm: string) => string;
  noSlideSize: string;
  conversionError: (err: string) => string;

  /* 가져오기 */
  tabExport: string;
  tabImport: string;
  pickFile: string;
  dropHint: string;
  reading: (done: number, total: number) => string;
  parsing: string;
  creating: (done: number, total: number) => string;
  imported: (slides: number) => string;
  importFailed: (err: string) => string;
  notPptx: string;
  fontsMissing: (families: string) => string;
  mixedAlign: string;

  /* 변환 경고 */
  scopeAll: string;
  nodeImageRender: string;
  dpiClamped: (requested: number, actual: number) => string;
  maskSkipped: string;
  blendMode: (mode: string) => string;
  effectUnsupported: (effect: string) => string;
  effectLayerBlur: string;
  effectBackgroundBlur: string;
  clipNotSupported: string;
  multipleFills: string;
  vectorStrokeAlign: (align: string) => string;
  vectorPathFailed: string;
  gradientApproximated: (type: string, hex: string) => string;
  textGradientApproximated: string;
  videoFillSkipped: string;
  imageRoundedWithEffects: string;
  nodeRenderFailed: string;
  imageNotFound: string;
  imageReadFailed: string;
}

const KO: Strings = {
  numberLocale: 'ko-KR',

  slideSize: '슬라이드 크기',
  selectFrame: '프레임을 선택하세요',
  frameCount: (n) => `프레임 ${n}개`,
  uniformScale: (k) => `균등 배율 ${k}×`,
  imageResolution: '이미지 해상도',
  dpiScreen: '화면',
  dpiDefault: '기본',
  dpiPrint: '인쇄',
  exportButton: 'PPTX 내보내기',
  building: '만드는 중…',

  tabExport: 'PPTX 로 내보내기',
  tabImport: 'PPTX 가져오기',
  pickFile: 'PPTX 파일 선택',
  dropHint: '슬라이드가 프레임으로 들어옵니다',
  reading: (done, total) => `읽는 중… ${done}/${total}`,
  parsing: '파일 분석 중…',
  creating: (done, total) => `프레임 만드는 중… ${done}/${total}`,
  imported: (slides) => `슬라이드 ${slides}장을 가져왔습니다`,
  importFailed: (err) => `가져오기 실패: ${err}`,
  notPptx: '.pptx 파일만 가져올 수 있습니다',
  fontsMissing: (families) =>
    `이 컴퓨터에 없는 폰트는 대체했습니다: ${families}`,
  mixedAlign: '문단마다 가로 정렬이 달라 첫 문단 기준으로 맞췄습니다. Figma 텍스트는 정렬이 상자 단위입니다.',
  warningsTitle: (n) => `변환 참고 ${n}건`,
  buildFailed: (err) => `PPTX 생성 실패: ${err}`,
  exported: (slides, size) => `슬라이드 ${slides}장 · ${size}`,
  exportFailed: 'PPTX 생성에 실패했습니다.',

  paperChip: (name, portrait) => `${name} ${portrait ? '세로' : '가로'}`,

  noFrameSelected: '내보낼 프레임을 선택하세요',
  mixedSizes: (kinds) => `프레임 크기가 ${kinds}종류로 섞여 있습니다\n같은 크기끼리만 선택해 주세요`,
  rotatedFrame: (name) => `회전된 프레임은 내보낼 수 없습니다 (${name})`,
  // 허용 범위는 PowerPoint 사정이라 사용자가 알 길이 없다. 숫자는 남기고 설명은 뺀다.
  sizeOutOfRange: (mm) => `슬라이드 한 변은 25.4~1422mm 여야 합니다 (지금 ${mm})`,
  noSlideSize: '이 프레임 크기로 만들 수 있는 슬라이드 크기가 없습니다.',
  conversionError: (err) => `변환 중 오류: ${err}`,

  scopeAll: '전체',
  nodeImageRender: '이미지 렌더',
  dpiClamped: (requested, actual) =>
    `Figma 렌더 배율 한계(4배)에 걸려 ${requested} DPI 대신 약 ${actual} DPI 로 내보냅니다.`,
  maskSkipped: '마스크 레이어는 PPTX 에 대응 도형이 없어 건너뛰었습니다.',
  blendMode: (mode) => `혼합 모드 ${mode} 는 PPTX 에서 재현되지 않습니다.`,
  effectUnsupported: (effect) => `${effect} 효과는 PPTX 에서 재현되지 않습니다.`,
  effectLayerBlur: '레이어 블러',
  effectBackgroundBlur: '배경 블러',
  clipNotSupported: '프레임 밖으로 넘치는 요소가 있습니다. PPTX 에는 내용 자르기가 없어 그대로 보입니다.',
  multipleFills: 'PPTX 도형은 채우기를 하나만 가질 수 있어 맨 위 레이어만 반영했습니다.',
  vectorStrokeAlign: (align) => `벡터 도형의 선 정렬 ${align} 은 PPTX 에서 중앙 정렬로 그려집니다.`,
  vectorPathFailed: '벡터 경로를 읽지 못해 사각형으로 대체했습니다.',
  gradientApproximated: (type, hex) => `그라디언트(${type})는 PPTX 단색 #${hex} 으로 근사했습니다.`,
  textGradientApproximated: '텍스트의 그라디언트/이미지 채우기는 단색으로 근사했습니다.',
  videoFillSkipped: '비디오 채우기는 PPTX 로 옮길 수 없어 건너뛰었습니다.',
  imageRoundedWithEffects:
    '이미지의 둥근 모서리/크롭은 그림자 등 효과와 함께 쓰면 재현되지 않습니다.',
  nodeRenderFailed: '노드 렌더에 실패해 원본 이미지를 그대로 넣었습니다.',
  imageNotFound: '이미지 원본을 찾지 못해 건너뛰었습니다.',
  imageReadFailed: '이미지 데이터를 읽지 못해 건너뛰었습니다.',
};

const EN: Strings = {
  numberLocale: 'en-US',

  slideSize: 'Slide size',
  selectFrame: 'Select a frame',
  frameCount: (n) => (n === 1 ? '1 frame' : `${n} frames`),
  uniformScale: (k) => `Uniform scale ${k}×`,
  imageResolution: 'Image resolution',
  dpiScreen: 'Screen',
  dpiDefault: 'Default',
  dpiPrint: 'Print',
  exportButton: 'Export PPTX',
  building: 'Building…',

  tabExport: 'Export to PPTX',
  tabImport: 'Import PPTX',
  pickFile: 'Choose a PPTX file',
  dropHint: 'Slides come in as frames',
  reading: (done, total) => `Reading… ${done}/${total}`,
  parsing: 'Reading the file…',
  creating: (done, total) => `Creating frames… ${done}/${total}`,
  imported: (slides) => `${slides === 1 ? '1 slide' : `${slides} slides`} imported`,
  importFailed: (err) => `Import failed: ${err}`,
  notPptx: 'Only .pptx files can be imported',
  fontsMissing: (families) => `Fonts not available here were substituted: ${families}`,
  mixedAlign: 'Paragraph alignments differed; the first one was applied to the whole box. Figma aligns text per box, not per paragraph.',
  warningsTitle: (n) => (n === 1 ? '1 conversion note' : `${n} conversion notes`),
  buildFailed: (err) => `Could not build the PPTX: ${err}`,
  exported: (slides, size) => `${slides === 1 ? '1 slide' : `${slides} slides`} · ${size}`,
  exportFailed: 'Could not build the PPTX.',

  paperChip: (name, portrait) => `${name} ${portrait ? 'Portrait' : 'Landscape'}`,

  noFrameSelected: 'Select the frames you want to export',
  mixedSizes: (kinds) =>
    `${kinds} different frame sizes are selected\nSelect frames that are all the same size`,
  rotatedFrame: (name) => `A rotated frame cannot be exported (${name})`,
  sizeOutOfRange: (mm) => `A slide side must be between 25.4mm and 1422mm (currently ${mm})`,
  noSlideSize: 'No slide size can be produced from this frame size.',
  conversionError: (err) => `Conversion failed: ${err}`,

  scopeAll: 'All slides',
  nodeImageRender: 'Image render',
  dpiClamped: (requested, actual) =>
    `Figma caps node rendering at 4×, so this exports at roughly ${actual} DPI instead of ${requested} DPI.`,
  maskSkipped: 'Mask layers have no PPTX equivalent and were skipped.',
  blendMode: (mode) => `Blend mode ${mode} cannot be reproduced in PPTX.`,
  effectUnsupported: (effect) => `${effect} cannot be reproduced in PPTX.`,
  effectLayerBlur: 'Layer blur',
  effectBackgroundBlur: 'Background blur',
  clipNotSupported:
    'Content overflows this frame. PPTX has no clipping, so it stays visible.',
  multipleFills: 'A PPTX shape holds one fill, so only the topmost layer was used.',
  vectorStrokeAlign: (align) =>
    `Stroke alignment ${align} on a vector shape is drawn centred in PPTX.`,
  vectorPathFailed: 'The vector path could not be read, so a rectangle was used instead.',
  gradientApproximated: (type, hex) =>
    `The ${type} gradient was approximated as the solid colour #${hex}.`,
  textGradientApproximated: 'Gradient and image fills on text were approximated as a solid colour.',
  videoFillSkipped: 'Video fills cannot be carried into PPTX and were skipped.',
  imageRoundedWithEffects:
    'Rounded corners and cropping on an image are not reproduced when effects are also applied.',
  nodeRenderFailed: 'Rendering the node failed, so the original image was embedded as-is.',
  imageNotFound: 'The source image could not be found and was skipped.',
  imageReadFailed: 'The image data could not be read and was skipped.',
};

let current: Locale = 'en';

export function setLocale(locale: Locale): void {
  current = locale;
}

export function getLocale(): Locale {
  return current;
}

export function t(): Strings {
  return current === 'ko' ? KO : EN;
}
