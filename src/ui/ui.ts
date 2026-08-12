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

  syncHeight();
}

function renderWarnings(doc: Doc): void {
  if (doc.warnings.length === 0) {
    elWarnings.classList.add('hidden');
  } else {
    const items = doc.warnings
      .map((w) => `<li><b>${escapeHtml(w.slide)} › ${escapeHtml(w.node)}</b><br>${escapeHtml(w.message)}</li>`)
      .join('');
    elWarnings.innerHTML = `<h1>${t().warningsTitle(doc.warnings.length)}</h1><ul>${items}</ul>`;
    elWarnings.classList.remove('hidden');
  }
  syncHeight();
}

function showBlocked(message: string): void {
  elBlocked.textContent = message;
  elBlocked.classList.remove('hidden');
  syncHeight();
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
 *
 * body 박스 높이를 그대로 쓰지 않는다. 오류 박스가 사라져 내용이 줄어도 body 가 따라 줄지 않아
 * 창 아래에 빈 공간이 남는 경우가 있었다. 보이는 자식들의 실제 아래 끝을 재면 늘어날 때도
 * 줄어들 때도 같은 값이 나온다.
 *
 * 드롭다운 목록은 absolute 라 어느 자식의 박스에도 안 잡히므로 따로 더한다.
 */
function contentHeight(): number {
  const padBottom = parseFloat(getComputedStyle(document.body).paddingBottom) || 0;
  let bottom = 0;
  for (const child of Array.from(document.body.children)) {
    const rect = child.getBoundingClientRect();
    // display:none 인 요소와 <script> 는 크기가 0 이라 자연히 걸러진다.
    if (rect.height === 0 && rect.width === 0) continue;
    bottom = Math.max(bottom, rect.bottom);
  }
  const menuBottom = dpi?.menuBottom();
  if (menuBottom != null) bottom = Math.max(bottom, menuBottom);
  return bottom + padBottom;
}

let lastHeight = -1;

function syncHeight(): void {
  const height = contentHeight();
  // 같은 값을 반복해서 보내지 않는다. ResizeObserver 와 명시 호출이 겹칠 수 있다.
  if (Math.abs(height - lastHeight) < 0.5) return;
  lastHeight = height;
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
  // 지난 시도의 오류와 경고는 새로 시작할 때 치운다 — 남겨두면 창만 계속 길어진다.
  elWarnings.classList.add('hidden');
  elBlocked.classList.add('hidden');
  syncHeight();
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
    showBlocked(msg.message);
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
      showBlocked(t().buildFailed(String(err)));
      toMain({ type: 'notify', message: t().exportFailed, error: true });
    } finally {
      finish();
    }
  }
};

toMain({ type: 'ready', locale });
