import { parseSlideRange, readPptx } from '../import/reader';
import type { ImportDoc } from '../shared/importir';
import { resolveLocale, setLocale, t } from '../shared/i18n';
import { UI_MAX_HEIGHT } from '../shared/ir';
import type { MainToUi, SelectionState, UiToMain, Warning } from '../shared/ir';
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
  $('tabExport').textContent = t().tabExport;
  $('tabImport').textContent = t().tabImport;
  $('importTitle').textContent = t().tabImport;
  elPick.textContent = t().pickFile;
  elImport.textContent = t().importButton;
  $('rangeLabel').textContent = t().slideRange;
  elRange.placeholder = t().slideRangeHint;
  renderChosen();
}

/* ── 방향 전환 ───────────────────────────────────────────────── */

const elTabExport = $<HTMLButtonElement>('tabExport');
const elTabImport = $<HTMLButtonElement>('tabImport');
const elPaneExport = $('paneExport');
const elPaneImport = $('paneImport');
const elPick = $<HTMLButtonElement>('pick');
const elImport = $<HTMLButtonElement>('import');
const elFile = $<HTMLInputElement>('file');
const elRange = $<HTMLInputElement>('range');
const elFileName = $('fileName');
const elImportHint = $('importHint');

function showTab(which: 'export' | 'import'): void {
  const isExport = which === 'export';
  elTabExport.setAttribute('aria-selected', String(isExport));
  elTabImport.setAttribute('aria-selected', String(!isExport));
  elPaneExport.classList.toggle('hidden', !isExport);
  elPaneImport.classList.toggle('hidden', isExport);
  elWarnings.classList.add('hidden');
  elBlocked.classList.add('hidden');
  syncHeight();
}

elTabExport.addEventListener('click', () => showTab('export'));
elTabImport.addEventListener('click', () => showTab('import'));

/* ── 가져오기 ────────────────────────────────────────────────── */

/**
 * 고른 파일. 고르는 것과 가져오는 것을 나눈 이유는 내보내기 탭과 짜임을 맞추기 위해서다 —
 * 머리 영역이 "무엇을 고를 것인가"를 보여주고, 아래 주 버튼이 실행한다.
 */
let chosen: File | null = null;

function renderChosen(): void {
  elFileName.textContent = chosen ? chosen.name : '—';
  elImportHint.textContent = chosen
    ? t().fileSize(fmt(chosen.size / 1024 / 1024))
    : t().selectFile;
  // 파일이 없을 때만 빨갛게 — 내보내기 탭에서 프레임 미선택을 알리는 방식과 같다.
  elImportHint.classList.toggle('danger', !chosen);
  elImport.disabled = busy || !chosen;
}

elPick.addEventListener('click', () => {
  if (!busy) elFile.click();
});

elFile.addEventListener('change', () => {
  const file = elFile.files?.[0];
  elFile.value = '';
  if (!file) return;
  if (!/\.pptx$/i.test(file.name)) {
    showBlocked(t().notPptx);
    return;
  }
  elBlocked.classList.add('hidden');
  chosen = file;
  renderChosen();
  syncHeight();
});

elImport.addEventListener('click', () => {
  if (!busy && chosen) void runImport(chosen);
});

async function runImport(file: File): Promise<void> {
  if (busy) return;

  busy = true;
  elPick.disabled = true;
  elImport.disabled = true;
  elWarnings.classList.add('hidden');
  elBlocked.classList.add('hidden');
  elImport.textContent = t().parsing;

  try {
    const buffer = await file.arrayBuffer();
    const doc = await readPptx(buffer, file.name, {
      only: parseSlideRange(elRange.value) ?? undefined,
      onProgress: (done, total) => {
        elImport.textContent = t().reading(done, total);
      },
    });

    pending = doc;
    nextSlide = 0;
    nextLayout = 0;
    pendingWarnings = doc.warnings;
    elImport.textContent = t().creating(0, doc.slides.length);

    toMain({
      type: 'importBegin',
      widthPt: doc.widthPt,
      heightPt: doc.heightPt,
      fileName: doc.fileName,
      total: doc.slides.length,
      fonts: doc.fonts,
      fontAliases: doc.fontAliases,
    });
    // 여기서 슬라이드를 바로 보내면 안 된다 — main 이 세션을 만들기 전에 도착해 버려진다.
    // importReady 를 받고 한 장씩, 처리 확인을 받아가며 보낸다.
  } catch (err) {
    resetImport();
    showBlocked(t().importFailed(String(err)));
  }
}

let pendingWarnings: Warning[] = [];
let pending: ImportDoc | null = null;
let nextSlide = 0;

let nextLayout = 0;

/**
 * 다음 조각 하나를 보낸다 — 레이아웃 먼저, 그 다음 슬라이드.
 *
 * 58장을 한꺼번에 밀어 넣으면 이미지가 base64 로 부푼 만큼(실측 115MB) 다리가 막힌다.
 * 보낸 것은 즉시 비워 메모리도 같이 놓아준다.
 * 레이아웃을 앞세우는 이유는 슬라이드가 인스턴스를 놓을 때 컴포넌트가 이미 있어야 하기 때문이다.
 */
function sendNextSlide(): void {
  if (!pending) return;

  if (nextLayout < pending.layouts.length) {
    const layout = pending.layouts[nextLayout++];
    toMain({ type: 'importLayout', layout });
    pending.layouts[nextLayout - 1] = { key: layout.key, name: layout.name, nodes: [] };
    return;
  }

  if (nextSlide >= pending.slides.length) {
    toMain({ type: 'importEnd' });
    return;
  }
  const index = nextSlide++;
  const slide = pending.slides[index];
  toMain({ type: 'importSlide', index, slide });
  pending.slides[index] = { name: slide.name, nodes: [], perSlideNodes: [] };
}

function resetImport(): void {
  busy = false;
  pending = null;
  nextSlide = 0;
  nextLayout = 0;
  elPick.disabled = false;
  elImport.textContent = t().importButton;
  renderChosen();
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

function renderWarnings(warnings: Warning[]): void {
  if (warnings.length === 0) {
    elWarnings.classList.add('hidden');
  } else {
    const items = warnings
      .map((w) => `<li><b>${escapeHtml(w.slide)} › ${escapeHtml(w.node)}</b><br>${escapeHtml(w.message)}</li>`)
      .join('');
    elWarnings.innerHTML = `<h1>${t().warningsTitle(warnings.length)}</h1><ul>${items}</ul>`;
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

  /*
   * 창 크기는 main 을 거쳐 비동기로 바뀐다. 그래서 탭을 옮기면 새 내용이 먼저 그려지고
   * 창은 한 박자 뒤에 따라오는데, 그 틈에 내용이 창보다 크면 스크롤바가 번쩍였다.
   *
   * 창은 어차피 내용 높이에 맞춰지므로 상한 아래에서는 스크롤이 필요 없다.
   * 상한을 넘을 때만 열어 둔다 — 그때는 진짜로 스크롤해야 한다.
   */
  document.documentElement.style.overflowY = height > UI_MAX_HEIGHT ? 'auto' : 'hidden';

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

  if (msg.type === 'importReady') {
    sendNextSlide();
    return;
  }

  if (msg.type === 'createProgress') {
    elImport.textContent = t().creating(msg.done, msg.total);
    sendNextSlide();
    return;
  }

  if (msg.type === 'imported') {
    resetImport();
    const warnings = pendingWarnings.slice();
    if (msg.missingFonts.length > 0) {
      warnings.unshift({
        slide: t().scopeAll,
        node: t().imageResolution,
        message: t().fontsMissing(msg.missingFonts.join(', ')),
      });
    }
    for (const f of msg.failures) {
      warnings.unshift({ slide: t().scopeAll, node: '노드 생성 실패', message: f });
    }
    pendingWarnings = [];
    renderWarnings(warnings);
    toMain({ type: 'notify', message: t().imported(msg.slides) });
    return;
  }

  if (msg.type === 'error') {
    finish();
    resetImport();
    showBlocked(msg.message);
    toMain({ type: 'notify', message: msg.message, error: true });
    return;
  }

  if (msg.type === 'doc') {
    elExport.textContent = t().building;
    try {
      const blob = await buildPptx(msg.doc);
      download(blob, msg.fileName);
      renderWarnings(msg.doc.warnings);
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
