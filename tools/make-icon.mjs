/**
 * 게시용 아이콘 PNG 생성 — docs/assets/icon.svg 와 같은 도형을 직접 래스터화한다.
 *
 *   node tools/make-icon.mjs [size]
 *
 * SVG 를 PNG 로 바꾸려면 보통 헤드리스 브라우저나 렌더링 라이브러리가 필요한데,
 * 이 아이콘은 둥근 사각형 네 개뿐이라 의존성을 하나 더 들이는 것보다 직접 그리는 편이 낫다.
 * 4×4 슈퍼샘플링으로 안티에일리어싱하고 PNG 는 zlib(내장)으로 인코딩한다.
 */
import { deflateSync } from 'node:zlib';
import { writeFileSync, mkdirSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');

const BLUE = [0x0d, 0x99, 0xff];
const WHITE = [0xff, 0xff, 0xff];

/**
 * 128 단위 좌표계의 도형들. 뒤에 올수록 위에 그려진다.
 *
 * 세로 용지(42/59 = 0.71, A4)와 가로 슬라이드(40/22.5 = 1.78, 16:9)를 대각선으로 배치한다.
 * 둘 다 원래 비율 그대로라는 게 이 플러그인의 전부라, 마크도 그것만 말한다.
 * 겹치면 한 덩어리로 읽히므로 가로·세로 모두 5~6 단위 간격을 띄웠다.
 */
const SHAPES = [
  { x: 0, y: 0, w: 128, h: 128, r: 28, color: BLUE },
  { x: 20, y: 21, w: 42, h: 59, r: 5, color: WHITE },
  { x: 68, y: 85, w: 40, h: 22.5, r: 4, color: WHITE },
];

/** 둥근 사각형 내부 판정 (모서리 반경까지 정확한 거리 계산) */
function inside(px, py, s, k) {
  const x = s.x * k;
  const y = s.y * k;
  const w = s.w * k;
  const h = s.h * k;
  const r = s.r * k;
  const cx = Math.min(Math.max(px, x + r), x + w - r);
  const cy = Math.min(Math.max(py, y + r), y + h - r);
  return Math.hypot(px - cx, py - cy) <= r;
}

const SS = 4; // 픽셀당 4×4 표본

function render(size) {
  const k = size / 128;
  const rgba = Buffer.alloc(size * size * 4);

  for (let py = 0; py < size; py++) {
    for (let px = 0; px < size; px++) {
      let r = 0;
      let g = 0;
      let b = 0;
      let covered = 0;

      for (let sy = 0; sy < SS; sy++) {
        for (let sx = 0; sx < SS; sx++) {
          const fx = px + (sx + 0.5) / SS;
          const fy = py + (sy + 0.5) / SS;
          // 위에 있는 도형부터 찾아 첫 번째 적중을 쓴다
          for (let i = SHAPES.length - 1; i >= 0; i--) {
            if (!inside(fx, fy, SHAPES[i], k)) continue;
            const c = SHAPES[i].color;
            r += c[0];
            g += c[1];
            b += c[2];
            covered++;
            break;
          }
        }
      }

      const total = SS * SS;
      const o = (py * size + px) * 4;
      if (covered === 0) continue; // 완전 투명 — 0 으로 남긴다
      // 비율 곱하지 않은 색으로 저장해야 가장자리에 어두운 테두리가 생기지 않는다
      rgba[o] = Math.round(r / covered);
      rgba[o + 1] = Math.round(g / covered);
      rgba[o + 2] = Math.round(b / covered);
      rgba[o + 3] = Math.round((covered / total) * 255);
    }
  }

  return rgba;
}

/* ── PNG 인코딩 ───────────────────────────────────────────────── */

const CRC_TABLE = (() => {
  const table = new Int32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c;
  }
  return table;
})();

function crc32(buf) {
  let c = -1;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ -1) >>> 0;
}

function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const body = Buffer.concat([Buffer.from(type, 'ascii'), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body));
  return Buffer.concat([len, body, crc]);
}

function encodePng(rgba, size) {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // RGBA
  ihdr[10] = 0; // deflate
  ihdr[11] = 0; // adaptive filtering
  ihdr[12] = 0; // no interlace

  // 각 스캔라인 앞에 필터 바이트 0 을 붙인다
  const raw = Buffer.alloc(size * (size * 4 + 1));
  for (let y = 0; y < size; y++) {
    raw[y * (size * 4 + 1)] = 0;
    rgba.copy(raw, y * (size * 4 + 1) + 1, y * size * 4, (y + 1) * size * 4);
  }

  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', deflateSync(raw, { level: 9 })),
    chunk('IEND', Buffer.alloc(0)),
  ]);
}

const size = Number(process.argv[2]) || 128;
const out = resolve(root, `docs/assets/icon-${size}.png`);
mkdirSync(dirname(out), { recursive: true });
const png = encodePng(render(size), size);
writeFileSync(out, png);
console.log(`${out.replace(root + '\\', '')}  ${size}×${size}  ${(png.length / 1024).toFixed(1)} KB`);
