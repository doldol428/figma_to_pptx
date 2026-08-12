import { setLocale, t } from '../shared/i18n';
import type { MainToUi, UiToMain } from '../shared/ir';
import { resolveSlide } from '../shared/slidesize';
import { extract } from './extract';
import { collectFrames, validate } from './validate';

const UI_WIDTH = 300;

figma.showUI(__html__, { width: UI_WIDTH, height: 180, themeColors: true });

function post(msg: MainToUi): void {
  figma.ui.postMessage(msg);
}

function sendSelection(): void {
  post({ type: 'selection', state: validate(figma.currentPage.selection) });
}

figma.on('selectionchange', sendSelection);

figma.ui.onmessage = async (msg: UiToMain) => {
  if (msg.type === 'ready') {
    // 로케일은 UI 만 알 수 있다 (navigator.languages). 선택 상태를 만들기 전에 먼저 심는다.
    setLocale(msg.locale);
    sendSelection();
    return;
  }

  if (msg.type === 'resize') {
    figma.ui.resize(UI_WIDTH, Math.max(140, Math.min(720, Math.ceil(msg.height))));
    return;
  }

  if (msg.type === 'notify') {
    figma.notify(msg.message, msg.error ? { error: true } : undefined);
    return;
  }

  if (msg.type === 'export') {
    const state = validate(figma.currentPage.selection);
    if (!state.ok) {
      // 선택이 비어 있으면 reason 이 없다 (UI 가 자체 안내한다). 버튼이 막혀 있어 도달할 일은 없지만 대비한다.
      post({ type: 'error', message: state.reason || t().noFrameSelected });
      return;
    }

    const frames = collectFrames(figma.currentPage.selection);
    const plan = resolveSlide(frames[0].width, frames[0].height);
    if (!plan) {
      post({ type: 'error', message: t().noSlideSize });
      return;
    }

    try {
      const doc = await extract(frames, {
        imageDpi: msg.imageDpi,
        plan,
        onProgress: (done, total) => post({ type: 'progress', done, total }),
      });
      post({ type: 'doc', doc, fileName: buildFileName(frames) });
    } catch (err) {
      post({ type: 'error', message: t().conversionError(String(err)) });
    }
  }
};

function buildFileName(frames: SceneNode[]): string {
  const base = frames.length === 1 ? frames[0].name : figma.root.name;
  const safe = base.replace(/[\\/:*?"<>|]/g, '_').trim() || 'figma';
  return `${safe}.pptx`;
}
