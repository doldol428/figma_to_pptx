import { resolveLocale, setLocale, t } from '../shared/i18n';
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

/*
 * Figma 는 계정 언어를 플러그인에 노출하지 않는다. iframe 로케일(데스크톱 앱이면 Electron,
 * 웹이면 브라우저)이 가장 가까운 신호라 이걸 쓰고, 한국어가 아니면 전부 영어로 떨어진다.
 * main 스레드는 navigator 에 접근할 수 없으므로 ready 메시지로 결과를 전달한다.
 */
const locale = resolveLocale(navigator.languages ?? [navigator.language]);
setLocale(locale);

let busy = false;
let state: SelectionState | null = null;

function toMain(msg: UiToMain): void {
  parent.postMessage({ pluginMessage: msg }, '*');
}

function fmt(n: number): string {
  return (Math.round(n * 10) / 10).toLocaleString(t().numberLocale);
}

function round3(n: number): number {
  return Math.round(n * 1000) / 1000;
}

function renderStaticText(): void {
  $('slideSizeTitle').textContent = t().slideSize;
  $('dpiLabel').textContent = t().imageResolution;
  $('dpiMenu').setAttribute('aria-label', t().imageResolution);
  elDetail.textContent = t().selectFrame;
  elExport.textContent = t().exportButton;
}

function renderSelection(s: SelectionState): void {
  state = s;

  // 선택이 비어 있을 땐 별도 오류 박스를 띄우지 않고 이 줄만 빨갛게 둔다 — 같은 말을 두 번 할 필요가 없다.
  elDetail.classList.toggle('danger', s.frameCount === 0);

  if (s.frameCount === 0) {
    elSize.textContent = '—';
    elDetail.textContent = t().selectFrame;
  } else {
    elSize.innerHTML =
      `${fmt(ptToMm(s.slideWPt))} × ${fmt(ptToMm(s.slideHPt))}<small> mm</small>` +
      (s.chip ? `<span class="badge">${escapeHtml(s.chip)}</span>` : '');

    const parts = [t().frameCount(s.frameCount), `${fmt(s.widthPx)} × ${fmt(s.heightPx)} px`];
    if (s.ptPerPx !== 1) parts.push(t().uniformScale(round3(s.ptPerPx)));
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
  elWarnings.innerHTML = `<h1>${t().warningsTitle(doc.warnings.length)}</h1><ul>${items}</ul>`;
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

const DPI_OPTIONS: DropdownOption[] = [
  { value: 96, label: '96 DPI', note: t().dpiScreen },
  { value: 150, label: '150 DPI' },
  { value: 220, label: '220 DPI', note: t().dpiDefault },
  { value: 300, label: '300 DPI', note: t().dpiPrint },
];

renderStaticText();

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
  elExport.textContent = t().reading(0, 0);
  elWarnings.classList.add('hidden');
  toMain({ type: 'export', imageDpi: dpi.value });
});

function finish(): void {
  busy = false;
  elExport.disabled = !state?.ok;
  dpi.setDisabled(false);
  elExport.textContent = t().exportButton;
}

window.onmessage = async (event: MessageEvent) => {
  const msg = event.data?.pluginMessage as MainToUi | undefined;
  if (!msg) return;

  if (msg.type === 'selection') {
    renderSelection(msg.state);
    return;
  }

  if (msg.type === 'progress') {
    elExport.textContent = t().reading(msg.done, msg.total);
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
    elExport.textContent = t().building;
    try {
      const blob = await buildPptx(msg.doc);
      download(blob, msg.fileName);
      renderWarnings(msg.doc);
      const size = `${fmt(ptToMm(msg.doc.slideWPt))} × ${fmt(ptToMm(msg.doc.slideHPt))} mm`;
      toMain({ type: 'notify', message: t().exported(msg.doc.slides.length, size) });
    } catch (err) {
      elBlocked.textContent = t().buildFailed(String(err));
      elBlocked.classList.remove('hidden');
      toMain({ type: 'notify', message: t().exportFailed, error: true });
    } finally {
      finish();
    }
  }
};

toMain({ type: 'ready', locale });
