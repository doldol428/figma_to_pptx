import type { Doc, MainToUi, PresetInfo, SelectionState, UiToMain } from '../shared/ir';
import { ptToMm } from '../shared/units';
import { buildPptx } from './build';

const $ = <T extends HTMLElement>(id: string): T => document.getElementById(id) as T;

const elSize = $('size');
const elDetail = $('detail');
const elBlocked = $('blocked');
const elExport = $<HTMLButtonElement>('export');
const elScale = $<HTMLSelectElement>('scale');
const elPreset = $<HTMLSelectElement>('preset');
const elPresetRow = $('presetRow');
const elNote = $('note');
const elWarnings = $('warnings');

let busy = false;
let state: SelectionState | null = null;
/** 사용자가 직접 고른 프리셋. 선택이 바뀌어도 같은 id 가 남아 있으면 유지한다. */
let chosenPresetId: string | null = null;

function toMain(msg: UiToMain): void {
  parent.postMessage({ pluginMessage: msg }, '*');
}

function fmt(n: number): string {
  return (Math.round(n * 10) / 10).toLocaleString('ko-KR');
}

function currentPreset(): PresetInfo | null {
  if (!state || state.presets.length === 0) return null;
  const id = chosenPresetId ?? state.defaultPresetId;
  return state.presets.find((p) => p.id === id) ?? state.presets[0];
}

function renderPresetOptions(s: SelectionState): void {
  if (s.presets.length <= 1) {
    elPresetRow.classList.add('hidden');
    elPreset.innerHTML = '';
    return;
  }

  const selected = currentPreset();
  elPreset.innerHTML = s.presets
    .map((p) => {
      const size = `${fmt(ptToMm(p.wPt))} × ${fmt(ptToMm(p.hPt))} mm`;
      const sel = p.id === selected?.id ? ' selected' : '';
      return `<option value="${escapeHtml(p.id)}"${sel}>${escapeHtml(p.label)} — ${size}</option>`;
    })
    .join('');
  elPresetRow.classList.remove('hidden');
}

function renderSelection(s: SelectionState): void {
  state = s;
  // 이전에 고른 프리셋이 새 선택에 없으면 기본값으로 되돌린다.
  if (chosenPresetId && !s.presets.some((p) => p.id === chosenPresetId)) {
    chosenPresetId = null;
  }

  renderPresetOptions(s);
  const preset = currentPreset();

  if (s.frameCount === 0 || !preset) {
    elSize.textContent = '—';
    elDetail.textContent = s.frameCount === 0 ? '프레임을 선택하세요' : '';
    elNote.textContent =
      'Figma 1px = 1pt(72dpi) 로 환산합니다. 프레임 크기가 그대로 슬라이드 크기가 되며 비율은 손대지 않습니다.';
  } else {
    const badge = preset.native ? s.paper : preset.label;
    elSize.innerHTML =
      `${fmt(ptToMm(preset.wPt))} × ${fmt(ptToMm(preset.hPt))}<small> mm</small>` +
      (badge ? `<span class="badge">${escapeHtml(badge)}</span>` : '');

    const parts = [
      `프레임 ${s.frameCount}개`,
      `${fmt(s.widthPx)} × ${fmt(s.heightPx)} px`,
    ];
    if (preset.ptPerPx !== 1) parts.push(`균등 배율 ${round3(preset.ptPerPx)}×`);
    elDetail.textContent = parts.join(' · ');

    elNote.textContent = preset.native
      ? 'Figma 1px = 1pt(72dpi). 프레임 크기가 그대로 슬라이드 크기가 됩니다.'
      : `프레임 비율이 이 표준과 일치해서, 좌표·크기·폰트에 균등 배율 ${round3(preset.ptPerPx)}× 를 걸어 표준 크기로 맞춥니다. 비율은 바뀌지 않습니다.`;
  }

  if (s.reason) {
    elBlocked.textContent = s.reason;
    elBlocked.classList.remove('hidden');
  } else {
    elBlocked.classList.add('hidden');
  }

  elExport.disabled = !s.ok || busy;
  if (!busy) {
    elExport.textContent = s.ok
      ? `PPTX 내보내기 (슬라이드 ${s.frameCount}장)`
      : 'PPTX 내보내기';
  }
}

function round3(n: number): number {
  return Math.round(n * 1000) / 1000;
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

elPreset.addEventListener('change', () => {
  chosenPresetId = elPreset.value;
  if (state) renderSelection(state);
});

elExport.addEventListener('click', () => {
  if (busy || !state) return;
  const preset = currentPreset();
  if (!preset) return;
  busy = true;
  elExport.disabled = true;
  elExport.textContent = '프레임 읽는 중…';
  elWarnings.classList.add('hidden');
  toMain({ type: 'export', imageScale: Number(elScale.value), presetId: preset.id });
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
      toMain({ type: 'notify', message: `슬라이드 ${n}장을 내보냈습니다. (${msg.doc.presetLabel})` });
    } catch (err) {
      elBlocked.textContent = `PPTX 생성 실패: ${String(err)}`;
      elBlocked.classList.remove('hidden');
      toMain({ type: 'notify', message: 'PPTX 생성에 실패했습니다.', error: true });
    } finally {
      busy = false;
      elExport.disabled = false;
      toMain({ type: 'ready' });
    }
  }
};

toMain({ type: 'ready' });
