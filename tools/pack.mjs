/**
 * 사내 공유용 zip 묶기.
 *
 *   npm run pack
 *
 * Figma 가 개발 플러그인을 돌리는 데 필요한 건 manifest 와 빌드 결과물뿐이다.
 * 소스는 저장소에 있으니 넣지 않는다 — 받는 사람이 압축을 풀고 manifest 를 지정하면 끝이다.
 *
 * `dist/` 는 gitignore 대상이라 저장소에 없다. 반드시 빌드 후에 실행해야 한다.
 */
import JSZip from 'jszip';
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');

const REQUIRED = ['manifest.json', 'dist/code.js', 'dist/ui.html'];

const missing = REQUIRED.filter((f) => !existsSync(resolve(root, f)));
if (missing.length > 0) {
  console.error(`빠진 파일: ${missing.join(', ')}`);
  console.error('먼저 빌드하세요:  npm run build');
  process.exit(1);
}

const pkg = JSON.parse(await readFile(resolve(root, 'package.json'), 'utf8'));
const manifest = JSON.parse(await readFile(resolve(root, 'manifest.json'), 'utf8'));

const folder = 'innodep-export';

// Windows 에서 더블클릭으로 열었을 때 한글이 깨지지 않도록 BOM 을 붙인다.
const guide = `﻿${manifest.name} — Figma 플러그인 설치 안내
${'='.repeat(50)}

■ 설치

1. 이 폴더를 원하는 위치에 압축 해제합니다.
   플러그인을 쓰는 동안 폴더가 계속 있어야 합니다. 옮기거나 지우면 동작하지 않습니다.

2. Figma "데스크톱 앱" 을 엽니다. (웹 브라우저에서는 불러올 수 없습니다)

3. 메뉴 → Plugins → Development → Import plugin from manifest…

4. 압축 해제한 폴더 안의 manifest.json 을 선택합니다.

■ 사용

1. 캔버스에서 내보낼 프레임을 선택합니다. 여러 개 선택하면 여러 장짜리 한 파일이 됩니다.

2. 메뉴 → Plugins → Development → ${manifest.name}

3. [PPTX 내보내기] 를 누르면 파일이 다운로드됩니다.

■ 알아두실 것

· 선택한 프레임의 크기가 모두 같아야 합니다.
  PPTX 는 파일 하나에 슬라이드 크기가 하나뿐이라, 크기가 섞이면 버튼이 비활성화됩니다.
  같은 크기끼리 나눠서 여러 번 내보내면 됩니다.

· 슬라이드 크기는 프레임 크기에서 그대로 나옵니다.
  A4 프레임(595.28 × 841.89 px)은 정확히 210 × 297 mm 로 출력됩니다.
  비율이 PowerPoint 표준(16:9 등)과 맞으면 표준 크기로 자동 환산합니다.

· 도형과 텍스트는 이미지가 아니라 편집 가능한 PPT 개체로 들어갑니다.

· 그라디언트는 단색으로 근사합니다. 내용 자르기·마스크·블러는 PPTX 로 재현되지 않습니다.
  해당하는 항목은 변환 후 목록으로 전부 보여줍니다.

■ 문제가 있으면

Plugins → Development → Show/Hide console 로 콘솔을 열어 메시지를 확인해 주세요.

소스: ${pkg.repository?.url?.replace(/^git\+/, '').replace(/\.git$/, '') ?? ''}
버전: ${pkg.version}
`;

const zip = new JSZip();
zip.file(`${folder}/README.txt`, guide);
zip.file(`${folder}/manifest.json`, await readFile(resolve(root, 'manifest.json')));
zip.file(`${folder}/dist/code.js`, await readFile(resolve(root, 'dist/code.js')));
zip.file(`${folder}/dist/ui.html`, await readFile(resolve(root, 'dist/ui.html')));

const out = resolve(root, `release/innodep-export-v${pkg.version}.zip`);
await mkdir(dirname(out), { recursive: true });

const bytes = await zip.generateAsync({
  type: 'nodebuffer',
  compression: 'DEFLATE',
  compressionOptions: { level: 9 },
});
await writeFile(out, bytes);

console.log(`release/innodep-export-v${pkg.version}.zip  ${(bytes.length / 1024).toFixed(0)} KB`);
console.log('  README.txt / manifest.json / dist/code.js / dist/ui.html');
