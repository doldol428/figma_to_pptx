import type { Doc, MainToUi, SelectionState, UiToMain } from '../shared/ir';
import { ptToMm } from '../shared/units';
import { buildPptx } from './build';

const $ = <T extends HTMLElement>(id: string): T => document.getElementById(id) as T;

const elSize = $('size');
const elDetail = $('detail');
const elBlocked = $('blocked');
const elExport = $<HTMLButtonElement>('export');
const elDpi = $<HTMLSelectElement>('dpi');
const elWarnings = $('warnings');

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

/** 경고 목록이 붙고 빠질 때마다 필요한 높이가 달라지므로 창을 내용에 맞춘다. */
function syncHeight(): void {
  toMain({ type: 'resize', height: document.body.getBoundingClientRect().height });
}

new ResizeObserver(syncHeight).observe(document.body);

elExport.addEventListener('click', () => {
  if (busy || !state?.ok) return;
  busy = true;
  elExport.disabled = true;
  elExport.textContent = '읽는 중…';
  elWarnings.classList.add('hidden');
  toMain({ type: 'export', imageDpi: Number(elDpi.value) });
});

function finish(): void {
  busy = false;
  elExport.disabled = !state?.ok;
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
