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
      post({ type: 'error', message: state.reason });
      return;
    }

    const frames = collectFrames(figma.currentPage.selection);
    const plan = resolveSlide(frames[0].width, frames[0].height);
    if (!plan) {
      post({ type: 'error', message: '이 프레임 크기로 만들 수 있는 슬라이드 크기가 없습니다.' });
      return;
    }

    try {
      const doc = await extract(frames, {
        imageDpi: msg.imageDpi,
        plan,
        onProgress: (done, total, label) => post({ type: 'progress', done, total, label }),
      });
      post({ type: 'doc', doc, fileName: buildFileName(frames) });
    } catch (err) {
      post({ type: 'error', message: `변환 중 오류: ${String(err)}` });
    }
  }
};

function buildFileName(frames: SceneNode[]): string {
  const base = frames.length === 1 ? frames[0].name : figma.root.name;
  const safe = base.replace(/[\\/:*?"<>|]/g, '_').trim() || 'figma';
  return `${safe}.pptx`;
}
