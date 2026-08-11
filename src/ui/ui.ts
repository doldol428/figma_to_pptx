import type { Doc, MainToUi, SelectionState, UiToMain } from '../shared/ir';
import { pxToMm } from '../shared/units';
import { buildPptx } from './build';

const $ = <T extends HTMLElement>(id: string): T => document.getElementById(id) as T;

const elSize = $('size');
const elDetail = $('detail');
const elBlocked = $('blocked');
const elExport = $<HTMLButtonElement>('export');
const elScale = $<HTMLSelectElement>('scale');
const elWarnings = $('warnings');

let busy = false;

function toMain(msg: UiToMain): void {
  parent.postMessage({ pluginMessage: msg }, '*');
}

function fmt(n: number): string {
  return (Math.round(n * 10) / 10).toLocaleString('ko-KR');
}

function renderSelection(state: SelectionState): void {
  if (state.frameCount === 0) {
    elSize.textContent = '—';
    elDetail.textContent = '프레임을 선택하세요';
  } else {
    elSize.innerHTML =
      `${fmt(pxToMm(state.widthPx))} × ${fmt(pxToMm(state.heightPx))}<small> mm</small>` +
      (state.paper ? `<span class="badge">${state.paper}</span>` : '');
    elDetail.textContent =
      `프레임 ${state.frameCount}개 · ${fmt(state.widthPx)} × ${fmt(state.heightPx)} px`;
  }

  if (state.reason) {
    elBlocked.textContent = state.reason;
    elBlocked.classList.remove('hidden');
  } else {
    elBlocked.classList.add('hidden');
  }

  elExport.disabled = !state.ok || busy;
  if (!busy) {
    elExport.textContent = state.ok
      ? `PPTX 내보내기 (슬라이드 ${state.frameCount}장)`
      : 'PPTX 내보내기';
  }
}

function renderWarnings(doc: Doc): void {
  if (doc.warnings.length === 0) {
    elWarnings.classList.add('hidden');
    return;
  }
  const items = doc.warnings
    .map((w) => `<li><b>${escapeHtml(w.slide)} › ${escapeHtml(w.node)}</b><br>${escapeHtml(w.message)}</li>`)
    .join('');
  elWarnings.innerHTML =
    `<h1>변환 참고 ${doc.warnings.length}건</h1><ul>${items}</ul>`;
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

elExport.addEventListener('click', () => {
  if (busy) return;
  busy = true;
  elExport.disabled = true;
  elExport.textContent = '프레임 읽는 중…';
  elWarnings.classList.add('hidden');
  toMain({ type: 'export', imageScale: Number(elScale.value) });
});

window.onmessage = async (event: MessageEvent) => {
  const msg = event.data?.pluginMessage as MainToUi | undefined;
  if (!msg) return;

  if (msg.type === 'selection') {
    renderSelection(msg.state);
    return;
  }

  if (msg.type === 'progress') {
    elExport.textContent = `프레임 읽는 중… ${msg.done}/${msg.total}`;
    return;
  }

  if (msg.type === 'error') {
    busy = false;
    elExport.disabled = false;
    elExport.textContent = 'PPTX 내보내기';
    elBlocked.textContent = msg.message;
    elBlocked.classList.remove('hidden');
    toMain({ type: 'notify', message: msg.message, error: true });
    return;
  }

  if (msg.type === 'doc') {
    elExport.textContent = 'PPTX 만드는 중…';
    try {
      const blob = await buildPptx(msg.doc);
      download(blob, msg.fileName);
      renderWarnings(msg.doc);
      const n = msg.doc.slides.length;
      toMain({ type: 'notify', message: `슬라이드 ${n}장을 내보냈습니다.` });
    } catch (err) {
      elBlocked.textContent = `PPTX 생성 실패: ${String(err)}`;
      elBlocked.classList.remove('hidden');
      toMain({ type: 'notify', message: 'PPTX 생성에 실패했습니다.', error: true });
    } finally {
      busy = false;
      elExport.disabled = false;
      elExport.textContent = 'PPTX 내보내기';
      toMain({ type: 'ready' });
    }
  }
};

toMain({ type: 'ready' });
