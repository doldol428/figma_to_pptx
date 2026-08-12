import type { Doc, MainToUi, SelectionState, UiToMain } from '../shared/ir';
import { ptToMm } from '../shared/units';
import { buildPptx } from './build';
import { Dropdown, type DropdownOption } from './dropdown';

const $ = <T extends HTMLElement>(id: string): T => document.getElementById(id) as T;

const elSize = $('size');
const elDetail = $('detail');
const elBlocked = $('blocked');
const elExport = $<HTMLButtonElement>('export');
const elWarnings = $('warnings');

const DPI_OPTIONS: DropdownOption[] = [
  { value: 96, label: '96 DPI', note: '화면' },
  { value: 150, label: '150 DPI' },
  { value: 220, label: '220 DPI', note: '기본' },
  { value: 300, label: '300 DPI', note: '인쇄' },
];

let busy = false;
let state: SelectionState | null = null;

function toMain(msg: UiToMain): void {
  parent.postMessage({ pluginMessage: msg }, '*');
}

function fmt(n: number): string {
  return (Math.round(n * 10) / 10).toLocaleString('ko-KR');
}

function round3(n: number): number {
  return Math.round(n * 1000) / 1000;
}

function renderSelection(s: SelectionState): void {
  state = s;

  if (s.frameCount === 0) {
    elSize.textContent = '—';
    elDetail.textContent = '프레임을 선택하세요';
  } else {
    elSize.innerHTML =
      `${fmt(ptToMm(s.slideWPt))} × ${fmt(ptToMm(s.slideHPt))}<small> mm</small>` +
      (s.chip ? `<span class="badge">${escapeHtml(s.chip)}</span>` : '');

    const parts = [`프레임 ${s.frameCount}개`, `${fmt(s.widthPx)} × ${fmt(s.heightPx)} px`];
    if (s.ptPerPx !== 1) parts.push(`균등 배율 ${round3(s.ptPerPx)}×`);
    elDetail.textContent = parts.join(' · ');
  }

  if (s.reason) {
    elBlocked.textContent = s.reason;
    elBlocked.classList.remove('hidden');
  } else {
    elBlocked.classList.add('hidden');
  }

  if (!busy) elExport.disabled = !s.ok;
}

function renderWarnings(doc: Doc): void {
  if (doc.warnings.length === 0) {
    elWarnings.classList.add('hidden');
    return;
  }
  const items = doc.warnings
    .map((w) => `<li><b>${escapeHtml(w.slide)} › ${escapeHtml(w.node)}</b><br>${escapeHtml(w.message)}</li>`)
    .join('');
  elWarnings.innerHTML = `<h1>변환 참고 ${doc.warnings.length}건</h1><ul>${items}</ul>`;
  elWarnings.classList.remove('hidden');
}

function escapeHtml(s: string): string {
  return s.replace(/[&<>"]/g, (c) =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c] as string));
}

function download(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  // 다운로드가 시작될 시간을 준 뒤 회수한다.
  setTimeout(() => URL.revokeObjectURL(url), 30_000);
}

/**
 * 창 높이를 내용에 맞춘다.
 * 드롭다운 목록은 absolute 라 body 높이에 안 잡히므로 따로 더해줘야 잘리지 않는다.
 */
function syncHeight(): void {
  const body = document.body.getBoundingClientRect().height;
  const menuBottom = dpi?.menuBottom();
  const height = menuBottom === null || menuBottom === undefined
    ? body
    : Math.max(body, menuBottom + 12);
  toMain({ type: 'resize', height });
}

const dpi = new Dropdown(
  $('dpiSelect'),
  $<HTMLButtonElement>('dpiTrigger'),
  $('dpiMenu'),
  DPI_OPTIONS,
  220,
  syncHeight,
);

new ResizeObserver(syncHeight).observe(document.body);

elExport.addEventListener('click', () => {
  if (busy || !state?.ok) return;
  busy = true;
  elExport.disabled = true;
  dpi.setDisabled(true);
  elExport.textContent = '읽는 중…';
  elWarnings.classList.add('hidden');
  toMain({ type: 'export', imageDpi: dpi.value });
});

function finish(): void {
  busy = false;
  elExport.disabled = !state?.ok;
  dpi.setDisabled(false);
  elExport.textContent = 'PPTX 내보내기';
}

window.onmessage = async (event: MessageEvent) => {
  const msg = event.data?.pluginMessage as MainToUi | undefined;
  if (!msg) return;

  if (msg.type === 'selection') {
    renderSelection(msg.state);
    return;
  }

  if (msg.type === 'progress') {
    elExport.textContent = `읽는 중… ${msg.done}/${msg.total}`;
    return;
  }

  if (msg.type === 'error') {
    finish();
    elBlocked.textContent = msg.message;
    elBlocked.classList.remove('hidden');
    toMain({ type: 'notify', message: msg.message, error: true });
    return;
  }

  if (msg.type === 'doc') {
    elExport.textContent = '만드는 중…';
    try {
      const blob = await buildPptx(msg.doc);
      download(blob, msg.fileName);
      renderWarnings(msg.doc);
      const size = `${fmt(ptToMm(msg.doc.slideWPt))} × ${fmt(ptToMm(msg.doc.slideHPt))} mm`;
      toMain({ type: 'notify', message: `슬라이드 ${msg.doc.slides.length}장 · ${size}` });
    } catch (err) {
      elBlocked.textContent = `PPTX 생성 실패: ${String(err)}`;
      elBlocked.classList.remove('hidden');
      toMain({ type: 'notify', message: 'PPTX 생성에 실패했습니다.', error: true });
    } finally {
      finish();
    }
  }
};

toMain({ type: 'ready' });
