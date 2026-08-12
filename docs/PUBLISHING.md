# 게시 가이드 — Figma Community

공개 스토어에 게시한다. 조직 전용(비공개) 게시는 Organization/Enterprise 플랜에서만 가능한데
해당 플랜이 아니므로 공개 경로를 쓴다. 주 사용자는 사내 인원이지만 링크만 있으면 누구나 설치할 수 있다.

---

## 게시 전 반드시 할 것

### 1. 플러그인 ID 교체 — **이걸 안 하면 게시가 막힌다**

현재 `manifest.json` 의 값은 개발용 placeholder 다.

```json
"id": "innodep-export-dev"
```

Figma 가 발급하는 ID 는 `1234567890123456789` 형태의 숫자 문자열이다.
로컬 import 는 아무 문자열이나 통과하지만 게시는 발급된 ID 를 요구한다.

받는 절차:

1. 데스크톱 앱 → **Plugins → Development → New plugin…**
2. **Figma design** + **Empty** 선택 → 아무 임시 폴더에나 생성
3. 생성된 `manifest.json` 의 `id` 값을 복사
4. 이 저장소의 `manifest.json` 에 붙여넣기
5. 임시 폴더 삭제 후 이 저장소를 다시 import

### 2. 빌드

```powershell
npm.cmd run check
```

Figma 는 게시 시점의 **빌드 결과물**(`dist/code.js`, `dist/ui.html`)을 업로드한다.
소스가 아니라 `dist/` 가 올라가므로 빌드가 반드시 선행돼야 한다.
`dist/` 는 `.gitignore` 대상이라 저장소에는 없다 — 게시하는 PC 에서 직접 빌드할 것.

### 3. 에셋

| 항목 | 규격 | 비고 |
|---|---|---|
| 아이콘 | 128×128 PNG | `docs/assets/icon.svg` 를 128×128 로 내보내기 |
| 커버 아트 | 1920×960 PNG (2:1) | `docs/assets/cover.svg` 를 1920×960 으로 내보내기 |

manifest 가 아니라 **게시 모달에서 업로드**한다.
SVG 는 초안이므로 Figma 에서 열어 다듬은 뒤 PNG 로 내보내면 된다.

---

## 게시 절차

데스크톱 앱 → **Plugins → Development → Manage plugins in development**
→ `Innodep Export` → **Publish** → 아래 정보 입력 → 제출.

심사는 보통 며칠 걸린다. 승인 후에는 같은 경로에서 새 버전을 올린다.

---

## 스토어 정보

### 이름

```
Innodep Export
```

### 한 줄 설명 (tagline)

영문:

```
Export Figma frames to PPTX at their true size — A4 stays A4, and every shape stays editable.
```

한국어:

```
Figma 프레임을 실제 크기 그대로 PPTX 로. A4 는 A4 로 나오고, 도형은 편집 가능한 상태로 남습니다.
```

### 상세 설명

영문:

```
Most Figma-to-PPTX tools force every frame into a fixed 16:9 slide. Design an A4 document
and it comes out stretched or letterboxed. Innodep Export derives the slide size from the
frame instead, and never scales one axis independently of the other.

TRUE SIZE
1 Figma px = 1 pt (72dpi). An A4 frame (595.28 × 841.89 px) becomes a slide of exactly
210 × 297 mm. Font sizes carry over unchanged: 12px in Figma is 12pt in PowerPoint.

STANDARD SLIDE SIZES, DETECTED
Presentation frames are usually drawn at screen resolution. A 1920×1080 frame would
otherwise export as a 67.7 × 38.1 cm slide — correct ratio, but not a PowerPoint standard
size. When the frame ratio matches a standard, the frame is converted to that size using a
single uniform factor applied to coordinates, dimensions, font sizes and stroke widths
alike. Aspect ratio never changes. A4 matches no standard ratio, so document work keeps
its true measurements automatically.

EDITABLE, NOT A SCREENSHOT
Nothing is flattened to an image. Rectangles, ellipses and lines become native PPTX
shapes. Vectors, booleans, arcs and per-corner radii are rebuilt as custom geometry from
the real Bézier paths. Text becomes real text boxes with per-character runs — font, size,
colour, letter spacing, line height, underline, strikethrough and links all survive.
Rotation, flips and shadows carry over.

MULTIPLE FRAMES, ONE FILE
Select several same-sized frames and they become one deck, ordered by their layout on the
canvas. A PPTX file holds a single slide size, so mixed sizes are blocked up front rather
than silently distorted.

KNOWN LIMITATIONS
Gradients are approximated as a solid colour. Clip content, masks, blur effects and blend
modes other than Normal are not reproducible in PPTX. Every one of these is reported in a
list after conversion — nothing is dropped silently.

Korean and English. Runs entirely on your machine; no network access.

Source: https://github.com/doldol428/figma_to_pptx
```

한국어:

```
대부분의 Figma → PPTX 도구는 프레임을 16:9 슬라이드에 욱여넣습니다. A4 문서를 만들어도
결과물은 늘어나거나 레터박스가 붙습니다. Innodep Export 는 슬라이드 크기를 프레임에서
유도하며, 가로와 세로에 다른 배율이 걸리는 일이 없습니다.

실측 유지
Figma 1px = 1pt (72dpi). A4 프레임(595.28 × 841.89 px)은 정확히 210 × 297 mm 슬라이드가
됩니다. 폰트 크기도 그대로 넘어갑니다 — Figma 12px 는 PowerPoint 12pt 입니다.

표준 슬라이드 크기 자동 인식
프레젠테이션 프레임은 보통 화면 해상도로 만듭니다. 1920×1080 을 실측으로 내보내면
67.7 × 38.1 cm 짜리 슬라이드가 나오는데, 비율은 맞지만 PowerPoint 표준이 아닙니다.
프레임 비율이 표준과 일치하면 단일 균등 배율로 표준 크기에 맞춥니다. 그 배율이 좌표·크기·
폰트 크기·선 굵기에 똑같이 걸리므로 종횡비는 변하지 않습니다. A4 는 어떤 표준과도 비율이
맞지 않아 문서 작업은 자동으로 실측을 유지합니다.

이미지가 아니라 편집 가능한 도형
통짜 이미지로 굽지 않습니다. 사각형·원·선은 PPTX 네이티브 도형이 되고, 벡터·불린·호·
모서리별 반경은 실제 베지어 경로에서 커스텀 도형으로 다시 만듭니다. 텍스트는 진짜 텍스트
상자가 되며 폰트·크기·색·자간·행간·밑줄·취소선·링크가 문자 단위로 보존됩니다. 회전·반전·
그림자도 함께 넘어갑니다.

여러 장을 한 파일로
같은 크기 프레임을 여러 개 선택하면 캔버스 배치 순서대로 한 덱이 됩니다. PPTX 는 파일당
슬라이드 크기가 하나뿐이라, 크기가 섞이면 조용히 왜곡하는 대신 미리 막습니다.

알려진 한계
그라디언트는 단색으로 근사합니다. 내용 자르기·마스크·블러·Normal 외 혼합 모드는 PPTX 로
재현되지 않습니다. 해당하는 항목은 변환 후 목록으로 전부 보여줍니다 — 조용히 넘어가지
않습니다.

한국어·영어 지원. 모든 처리는 로컬에서 이뤄지며 외부 통신이 없습니다.

소스: https://github.com/doldol428/figma_to_pptx
```

### 태그

이름에 포맷 키워드가 없으므로 검색 노출은 태그와 설명이 담당한다.

```
pptx, powerpoint, export, presentation, slides, deck, a4, print, document, frame, korean
```

### 카테고리

`Import and export` (없으면 `Utilities`)

---

## 다음 버전 계획

설명문에 로드맵을 적지 않는다 — 지키지 못하면 리뷰에 남는다.
구현이 끝난 시점에 릴리스 노트로 알린다.

- 그라디언트를 실제 `gradFill` 로 (pptx zip 후처리, `ShapeItem.name` 이 앵커로 준비돼 있음)
- DOCX 내보내기
- HWPX 내보내기

---

## 릴리스 노트 초안 (v1.0.0)

영문:

```
First release.

- Slide size derived from the frame: an A4 frame exports as exactly 210 × 297 mm
- Standard PowerPoint sizes detected automatically when the frame ratio matches
- Native PPTX shapes and editable text, including custom geometry for vectors and booleans
- Multiple same-sized frames export as one deck
- Image render resolution set by target DPI, consistent across slide sizes
- Korean and English
```

한국어:

```
첫 릴리스.

- 슬라이드 크기를 프레임에서 유도 — A4 프레임은 정확히 210 × 297 mm 로 출력
- 프레임 비율이 표준과 맞으면 PowerPoint 표준 크기 자동 인식
- 네이티브 도형과 편집 가능한 텍스트, 벡터·불린은 커스텀 도형으로 변환
- 같은 크기 프레임 여러 장을 한 덱으로
- 이미지 렌더 해상도를 목표 DPI 로 지정 — 슬라이드 크기가 달라져도 결과 해상도가 같음
- 한국어·영어 지원
```
