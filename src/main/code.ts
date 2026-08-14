import { setLocale, t } from '../shared/i18n';
import { UI_MAX_HEIGHT } from '../shared/ir';
import type { MainToUi, UiToMain } from '../shared/ir';
import { resolveSlide } from '../shared/slidesize';
import { ImportSession } from './create';
import { type MasterPass, extractMasters, extractSlide } from './extract';
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

/** 진행 중인 가져오기. 슬라이드가 한 장씩 오므로 상태를 들고 있어야 한다. */
let session: ImportSession | null = null;

/** 진행 중인 내보내기. UI 가 한 장씩 달라고 하므로 프레임과 첫 훑기 결과를 들고 있는다. */
let exporting: { frames: SceneNode[]; pass: MasterPass } | null = null;

/*
 * Figma 는 async 핸들러가 끝나기를 기다리지 않고 다음 메시지를 던진다.
 * 처리를 직렬화해 두지 않으면 가져오기처럼 상태를 쌓아가는 흐름이 순서를 잃는다.
 */
let queue: Promise<void> = Promise.resolve();
figma.ui.onmessage = (msg: UiToMain): void => {
  queue = queue.then(() => handle(msg)).catch((err) => {
    post({ type: 'error', message: String(err) });
  });
};

const handle = async (msg: UiToMain): Promise<void> => {
  if (msg.type === 'ready') {
    // 로케일은 UI 만 알 수 있다 (navigator.languages). 선택 상태를 만들기 전에 먼저 심는다.
    setLocale(msg.locale);
    sendSelection();
    return;
  }

  if (msg.type === 'resize') {
    figma.ui.resize(UI_WIDTH, Math.max(140, Math.min(UI_MAX_HEIGHT, Math.ceil(msg.height))));
    return;
  }

  if (msg.type === 'notify') {
    figma.notify(msg.message, msg.error ? { error: true } : undefined);
    return;
  }

  if (msg.type === 'importBegin') {
    try {
      session = await ImportSession.begin(msg);
      post({ type: 'importReady' });
    } catch (err) {
      session = null;
      post({ type: 'error', message: t().importFailed(String(err)) });
    }
    return;
  }

  if (msg.type === 'importLayout') {
    if (!session) return;
    try {
      await session.addLayout(msg.layout);
    } catch (err) {
      post({ type: 'error', message: t().importFailed(String(err)) });
    }
    // 레이아웃도 슬라이드와 같은 확인 신호로 다음 조각을 부른다.
    post({ type: 'createProgress', done: session.frames.length, total: session.total });
    return;
  }

  if (msg.type === 'importSlide') {
    if (!session) return;
    try {
      await session.addSlide(msg.slide, msg.index);
      post({ type: 'createProgress', done: msg.index + 1, total: session.total });
    } catch (err) {
      post({ type: 'error', message: t().importFailed(String(err)) });
    }
    return;
  }

  if (msg.type === 'importEnd') {
    if (!session) return;
    const { frames } = session;
    if (frames.length > 0) {
      figma.currentPage.selection = frames;
      figma.viewport.scrollAndZoomIntoView(frames);
    }
    post({
      type: 'imported',
      slides: frames.length,
      missingFonts: session.missingFonts(),
      failures: session.failures,
    });
    session = null;
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
      /*
       * 공통 서식만 먼저 보내고, 슬라이드는 UI 가 한 장씩 달라고 할 때마다 만든다.
       * 문서를 통째로 만들어 한 번에 넘기면 58장 분량의 base64 이미지가 그대로 실려 터진다.
       */
      const pass = await extractMasters(frames, { imageDpi: msg.imageDpi, plan });
      exporting = { frames, pass };
      post({
        type: 'exportBegin',
        meta: pass.meta,
        total: frames.length,
        fileName: buildFileName(frames),
        warnings: pass.warnings,
      });
    } catch (err) {
      exporting = null;
      post({ type: 'error', message: t().conversionError(String(err)) });
    }
    return;
  }

  if (msg.type === 'exportNext') {
    if (!exporting) return;
    const { frames, pass } = exporting;
    if (msg.index >= frames.length) {
      exporting = null;
      return;
    }
    try {
      const { slide, warnings } = await extractSlide(frames[msg.index], pass);
      post({ type: 'exportSlide', index: msg.index, slide, warnings });
    } catch (err) {
      exporting = null;
      post({ type: 'error', message: t().conversionError(String(err)) });
    }
  }
};

function buildFileName(frames: SceneNode[]): string {
  const base = frames.length === 1 ? frames[0].name : figma.root.name;
  const safe = base.replace(/[\\/:*?"<>|]/g, '_').trim() || 'figma';
  return `${safe}.pptx`;
}
