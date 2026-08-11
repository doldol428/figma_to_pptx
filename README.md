# figma_to_pptx

Figma 프레임을 **원본 물리 크기 그대로** PPTX 로 내보내는 Figma 플러그인.

기존 Figma → PPTX 도구들은 프레임을 16:9 슬라이드에 욱여넣는다.
A4(210×297mm) 문서를 만들어도 결과물은 4:3 이나 16:9 로 늘어나거나 레터박스가 붙는다.
이 플러그인은 **슬라이드 크기를 프레임 크기에서만 유도**하며, 어떤 좌표도 스케일링하지 않는다.

## 단위 규약

**Figma 1px = 1pt (72dpi)** 로 고정한다. 이 값은 설정이 아니라 계약이다.

| Figma 프레임 | PPTX 슬라이드 |
|---|---|
| 595.28 × 841.89 px | A4 세로 — 210 × 297 mm |
| 841.89 × 595.28 px | A4 가로 — 297 × 210 mm |
| 841.89 × 1190.55 px | A3 세로 — 297 × 420 mm |
| 960 × 540 px | 16:9 — 338.7 × 190.5 mm |

폰트 크기도 그대로 넘어간다. Figma 12px → PPTX 12pt.

`npm run verify` 가 A4 프레임에서 나온 pptx 의 `sldSz` 가 정확히
`7560000 × 10692000` EMU(= 210 × 297 mm)인지 매번 확인한다.

## 프레임 크기가 섞이면 내보낼 수 없다

PPTX 는 **파일 하나에 슬라이드 크기가 하나뿐**이다.
A4 프레임과 16:9 프레임을 같이 고르면 어느 한쪽은 반드시 비율이 깨진다.
그래서 크기가 다른 프레임이 섞여 있으면 내보내기 버튼이 비활성화되고,
어떤 크기가 몇 개씩 섞였는지 알려준다. 같은 크기끼리 나눠서 여러 번 내보내면 된다.

## 변환 방식 — 완전 네이티브

프레임 통째로 PNG 를 박는 방식이 아니라, 모든 노드를 **편집 가능한 PPT 개체**로 옮긴다.

| Figma | PPTX |
|---|---|
| RECTANGLE (모서리 둥글기 균일) | `rect` / `roundRect` prstGeom |
| RECTANGLE (모서리별 반경 다름) | `custGeom` — 실제 경로 |
| ELLIPSE (원형) / 호·도넛 | `ellipse` / `custGeom` |
| VECTOR, STAR, POLYGON, BOOLEAN_OPERATION | `custGeom` — `fillGeometry` 경로를 베지어 그대로 |
| LINE | `line` |
| TEXT | 텍스트 상자 + 문자 단위 런 (폰트·크기·색·자간·행간·밑줄·취소선·링크) |
| 이미지 채우기 | 원본 비트맵 (cover/contain/stretch) |
| 이미지 + 둥근 모서리/크롭 | 해당 노드만 렌더해서 PNG (효과가 없을 때) |
| DROP_SHADOW / INNER_SHADOW | PPT 그림자 |
| 회전 / 좌우 반전 | `xfrm` 의 `rot` / `flipH` |

### 좌표 변환에서 실제로 어려운 부분

- **회전 기준점이 다르다.** Figma 행렬은 노드 로컬 (0,0) 기준으로 회전이 이미 반영돼 있고,
  PowerPoint 는 도형 중심을 축으로 회전하며 `<a:off>` 는 *회전 전* 좌상단이다.
  그래서 행렬로 중심점을 구한 뒤 `w/2, h/2` 를 빼서 복원한다 → `src/main/geometry.ts`
- **회전 부호가 반대다.** Figma 는 반시계 양수, PPTX 는 시계 양수.
- **반전이 행렬에 섞여 있다.** `det < 0` 이면 x축 반전으로 분리한 뒤 회전각을 잰다.
- **선 정렬.** PPT 선은 항상 윤곽선 중앙. Figma 의 INSIDE/OUTSIDE 는 박스를 `strokeWeight/2` 만큼
  물리거나 부풀려서 맞춘다.
- **텍스트 여백.** PowerPoint 텍스트 상자는 기본 좌우 0.1인치 여백이 있다.
  이걸 0 으로 지우지 않으면 모든 텍스트가 밀린다.
- **부모 프레임 기준 좌표.** `inverse(frame.absoluteTransform) × node.absoluteTransform` 으로
  중첩 깊이나 오토레이아웃에 관계없이 프레임 로컬 좌표를 얻는다.

## 아직 안 되는 것

| 항목 | 현재 동작 |
|---|---|
| 그라디언트 채우기 | 정지점 가중 평균 단색으로 근사 + 경고 |
| 내용 자르기(clip) | 재현 안 됨 — 넘친 요소가 그대로 보임, 경고 |
| 마스크 레이어 | 건너뜀 + 경고 |
| 레이어/배경 블러 | 재현 안 됨 + 경고 |
| NORMAL 외 혼합 모드 | 재현 안 됨 + 경고 |
| 채우기 2개 이상 | 맨 위 레이어만 + 경고 |
| 그룹 자체의 불투명도 | 자식들에 곱해서 내림 |

변환하면서 걸린 항목은 전부 UI 하단에 "변환 참고" 목록으로 뜬다. 조용히 넘어가지 않는다.

### 그라디언트를 진짜로 넣으려면

PptxGenJS 4.0.1 의 `ShapeFillProps.type` 은 `'none' | 'solid'` 뿐이라 `gradFill` 을 못 만든다.
다음 단계는 `composePptx()` 결과 zip 을 JSZip 으로 열어
`objectName` 으로 도형을 찾아 `<a:solidFill>` 을 `<a:gradFill>` 로 치환하는 후처리다.
(IR 의 `ShapeItem.name` 이 그대로 `objectName` 으로 들어가 있어 앵커는 이미 준비돼 있다.)

## 구조

```
src/
  shared/
    units.ts      1px = 1pt 고정. 용지 규격 감지(표시용)
    ir.ts         중간 표현 — main ↔ ui 사이 순수 JSON
  main/           Figma 샌드박스 스레드
    code.ts       진입점 / 메시지 라우팅
    validate.ts   선택 검증 — 크기가 섞이면 여기서 막는다
    extract.ts    씬 그래프 순회 → IR
    geometry.ts   아핀 행렬 분해 (회전·반전·중심 보정)
    paint.ts      채우기 / 선 / 그림자
    text.ts       스타일 세그먼트 → 런, 폰트 굵기 매핑
    path.ts       SVG path → custGeom 점 (S/T 전개, A → 베지어 근사)
  ui/             iframe 스레드
    ui.html       UI 템플릿 (빌드 시 스크립트 인라인)
    ui.ts         상태 표시 / 다운로드
    build.ts      IR → PptxGenJS
tools/
  verify.ts       A4 회귀 검증
```

Figma 는 UI 로 단일 HTML 파일만 받기 때문에 빌드 시 JS 번들을 `ui.html` 안에 인라인한다.

## 개발

```bash
npm install
npm run build
```

Figma 데스크톱 앱 → Plugins → Development → Import plugin from manifest… → 이 폴더의 `manifest.json`

```bash
npm run watch    # 변경 감지 재빌드
npm run check    # 타입 검사 + 빌드 + A4 검증
```

`npm run verify` 는 `dist/verify-a4.pptx` 도 남기므로 PowerPoint 로 직접 열어볼 수 있다.
