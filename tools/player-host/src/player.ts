import { createPptxViewer, type PptxViewerInstance } from 'pptx-vanilla-viewer';
import * as pdfjsLib from 'pdfjs-dist';

type PlayerState = {
    slide: number;
    count: number;
};

type HostPlayer = {
    load: (url: string) => Promise<PlayerState>;
    next: () => PlayerState;
    prev: () => PlayerState;
    goTo: (slide: number) => PlayerState;
    state: () => PlayerState;
};

type ViewerHost = PptxViewerInstance & {
    store: {
        get: () => {
            currentSlide: number;
            presenting: boolean;
            canvasSize?: { width: number; height: number };
            zoom?: number | 'fit';
        };
        set: (patch: { presenting?: boolean; zoom?: number | 'fit' }) => void;
    };
    renderer: {
        presentationPlayback: {
            isComplete: () => boolean;
        };
    };
};

type PdfDocument = Awaited<ReturnType<typeof pdfjsLib.getDocument>['promise']>;

const host = document.getElementById('host');
if (!host) {
    throw new Error('player host missing');
}

let viewer: PptxViewerInstance | null = null;
let state: PlayerState = { slide: 1, count: 1 };
let coverObserver: ResizeObserver | null = null;
let rawNext: (() => void) | null = null;
let rawPrev: (() => void) | null = null;
let pdfDoc: PdfDocument | null = null;
let pdfCanvas: HTMLCanvasElement | null = null;
let pdfRenderToken = 0;

/**
 * WebView2 虚拟主机下的 worker 脚本路径。
 */
pdfjsLib.GlobalWorkerOptions.workerSrc = './vendor/pdf.worker.min.mjs';

/**
 * WebView2 里 Fullscreen API 会让后续 next() 失效；本窗已经是无边框铺满，直接禁用。
 */
function disableFullscreenApi(): void {
    const reject = () => Promise.reject(new Error('fullscreen disabled'));
    const proto = Element.prototype as Element & {
        webkitRequestFullscreen?: () => void;
    };
    try {
        proto.requestFullscreen = reject as typeof Element.prototype.requestFullscreen;
    } catch {
        // 环境无此 API
    }

    try {
        proto.webkitRequestFullscreen = () => undefined;
    } catch {
        // 无 webkit 前缀
    }
}

/**
 * 取出带 store / 动画时间线的 viewer。
 */
function asHost(): ViewerHost | null {
    return viewer as ViewerHost | null;
}

/**
 * 地址是否指向 PDF。
 */
function isPdfUrl(url: string): boolean {
    try {
        const path = new URL(url, 'https://player.multippt/').pathname;
        return path.toLowerCase().endsWith('.pdf');
    } catch {
        return url.toLowerCase().includes('.pdf');
    }
}

/**
 * 按屏幕比例覆盖放大幻灯片，裁掉多出的边，避免主屏留黑边。
 */
function applyCoverLayout(): void {
    const instance = asHost();
    const canvas = instance?.store?.get()?.canvasSize;
    if (!instance?.store || !canvas || canvas.width < 1 || canvas.height < 1) {
        return;
    }

    const viewWidth = host.clientWidth;
    const viewHeight = host.clientHeight;
    if (viewWidth < 8 || viewHeight < 8) {
        return;
    }

    const cover = Math.max(viewWidth / canvas.width, viewHeight / canvas.height);
    if (!Number.isFinite(cover) || cover <= 0) {
        return;
    }

    try {
        instance.store.set({ zoom: cover });
    } catch {
        // 缩放失败则保持 fit
    }
}

/**
 * 把页面指令发给 C# 编排器，让副屏一起翻页。
 */
function postHost(command: string): void {
    try {
        const webview = (window as unknown as {
            chrome?: { webview?: { postMessage: (message: unknown) => void } };
        }).chrome?.webview;
        webview?.postMessage({ cmd: command });
    } catch {
        // 非 WebView2 环境忽略
    }
}

/**
 * 播放条 / 单击走宿主同步；真正翻页仍用原始 next/prev。
 */
function hookViewerNavigation(): void {
    if (!viewer || rawNext) {
        return;
    }

    rawNext = viewer.next.bind(viewer);
    rawPrev = viewer.prev.bind(viewer);
    viewer.next = () => postHost('next');
    viewer.prev = () => postHost('prev');
}

/**
 * 开播后按窗口尺寸持续覆盖铺满。
 */
function watchCoverLayout(): void {
    applyCoverLayout();
    if (coverObserver || typeof ResizeObserver === 'undefined') {
        return;
    }

    coverObserver = new ResizeObserver(() => applyCoverLayout());
    coverObserver.observe(host);
}

/**
 * PDF 按窗口完整显示（不裁切）。
 */
function watchPdfLayout(): void {
    if (coverObserver || typeof ResizeObserver === 'undefined') {
        return;
    }

    coverObserver = new ResizeObserver(() => {
        if (pdfDoc && state.slide >= 1) {
            void renderPdfPage(state.slide);
        }
    });
    coverObserver.observe(host);
}

/**
 * 进入放映态但不走 Fullscreen：单击动画才会按 WPS 同拍先播再翻页。
 */
function ensurePresenting(): void {
    const instance = asHost();
    if (!instance?.store) {
        return;
    }

    try {
        if (!instance.store.get().presenting) {
            instance.store.set({ presenting: true });
        }
    } catch {
        // store 未就绪则保持预览
    }
}

/**
 * 从 viewer store 读 1-based 页码；读不到则沿用本地计数。
 */
function storeSlide(): number {
    try {
        const index = asHost()?.store?.get()?.currentSlide;
        if (typeof index === 'number' && Number.isFinite(index)) {
            return index + 1;
        }
    } catch {
        // 沿用本地
    }

    return state.slide;
}

/**
 * 当前页是否还有未播完的单击动画。
 */
function buildsRemain(): boolean {
    try {
        const playback = asHost()?.renderer?.presentationPlayback;
        return !!playback && playback.isComplete() === false;
    } catch {
        return false;
    }
}

/**
 * 返回当前页码状态；页码以本地计数为准，避免滞后的 onSlideChange 把页码打回 1。
 */
function readState(): PlayerState {
    if (pdfDoc) {
        return { slide: state.slide, count: state.count };
    }

    if (viewer) {
        try {
            state.count = Math.max(1, viewer.getSlideCount() || state.count);
        } catch {
            // 保持已有 count
        }
    }

    return { slide: state.slide, count: state.count };
}

/**
 * 拆掉 PPT 播放器实例。
 */
function destroyPptx(): void {
    if (viewer) {
        viewer.destroy();
        viewer = null;
        rawNext = null;
        rawPrev = null;
    }
}

/**
 * 拆掉 PDF 文档与画布。
 */
function destroyPdf(): void {
    pdfDoc = null;
    pdfCanvas = null;
    pdfRenderToken += 1;
}

/**
 * 停掉尺寸监听，避免 PPT / PDF 互相抢布局。
 */
function resetLayoutWatch(): void {
    coverObserver?.disconnect();
    coverObserver = null;
}

/**
 * 把 PDF 当前页画到画布，按窗口完整显示。
 */
async function renderPdfPage(pageNumber: number): Promise<void> {
    if (!pdfDoc || !pdfCanvas) {
        return;
    }

    const token = ++pdfRenderToken;
    const page = await pdfDoc.getPage(pageNumber);
    if (token !== pdfRenderToken) {
        return;
    }

    const viewWidth = Math.max(8, host.clientWidth);
    const viewHeight = Math.max(8, host.clientHeight);
    const base = page.getViewport({ scale: 1 });
    const fit = Math.min(viewWidth / base.width, viewHeight / base.height);
    const dpr = window.devicePixelRatio || 1;
    const viewport = page.getViewport({ scale: fit * dpr });
    pdfCanvas.width = Math.floor(viewport.width);
    pdfCanvas.height = Math.floor(viewport.height);
    pdfCanvas.style.width = `${Math.floor(viewport.width / dpr)}px`;
    pdfCanvas.style.height = `${Math.floor(viewport.height / dpr)}px`;
    const context = pdfCanvas.getContext('2d', { alpha: false });
    if (!context) {
        return;
    }

    await page.render({ canvasContext: context, viewport }).promise;
}

/**
 * 用 pdf.js 打开文稿，第一步只渲染首页。
 */
async function loadPdf(url: string): Promise<PlayerState> {
    destroyPptx();
    destroyPdf();
    resetLayoutWatch();
    host.replaceChildren();

    const stage = document.createElement('div');
    stage.className = 'pdf-stage';
    pdfCanvas = document.createElement('canvas');
    pdfCanvas.className = 'pdf-canvas';
    stage.appendChild(pdfCanvas);
    stage.addEventListener('click', () => postHost('next'));
    host.appendChild(stage);

    const task = pdfjsLib.getDocument({
        url,
        withCredentials: false,
        disableWorker: true
    });
    pdfDoc = await task.promise;
    state = { slide: 1, count: Math.max(1, pdfDoc.numPages || 1) };
    await renderPdfPage(1);
    watchPdfLayout();
    return readState();
}

/**
 * 创建或重建 vanilla viewer，并打开指定文稿。
 */
async function loadPptx(url: string): Promise<PlayerState> {
    destroyPptx();
    destroyPdf();
    resetLayoutWatch();
    host.replaceChildren();

    await new Promise<void>((resolve, reject) => {
        let settled = false;
        const finish = (error?: string) => {
            if (settled) {
                return;
            }

            settled = true;
            if (error) {
                reject(new Error(error));
                return;
            }

            resolve();
        };

        viewer = createPptxViewer(host, {
            source: url,
            editable: false,
            showToolbar: false,
            showThumbnails: false,
            showFormatToolbar: false,
            showInspector: false,
            onLoad: ({ slideCount }) => {
                state = { slide: 1, count: Math.max(1, slideCount || 1) };
                finish();
            },
            onError: (message) => finish(String(message || 'viewer load failed')),
            onSlideChange: () => {
                try {
                    state.count = viewer?.getSlideCount() || state.count;
                } catch {
                    // 只刷新页数
                }
            }
        });

        window.setTimeout(() => {
            if (!settled && viewer && viewer.getSlideCount() > 0) {
                state = {
                    slide: 1,
                    count: viewer.getSlideCount()
                };
                finish();
            }
        }, 8000);
    });

    ensurePresenting();
    hookViewerNavigation();
    watchCoverLayout();
    return readState();
}

/**
 * 按扩展名分流到 PPT 或 PDF 引擎。
 */
async function load(url: string): Promise<PlayerState> {
    if (isPdfUrl(url)) {
        return loadPdf(url);
    }

    return loadPptx(url);
}

/**
 * PDF 翻到指定页；越界则夹紧。
 */
function goToPdf(slide: number): PlayerState {
    if (!pdfDoc) {
        return readState();
    }

    const target = Math.max(1, Math.min(slide, state.count));
    if (target !== state.slide) {
        state.slide = target;
        void renderPdfPage(target);
    }

    return readState();
}

disableFullscreenApi();

const player: HostPlayer = {
    load,
    next() {
        if (pdfDoc) {
            return goToPdf(state.slide + 1);
        }

        if (!viewer) {
            return readState();
        }

        ensurePresenting();
        const before = storeSlide();
        const hadBuilds = buildsRemain();
        try {
            rawNext?.();
        } catch {
            // 单步失败则尝试整页
        }

        const after = storeSlide();
        if (after > before) {
            state.slide = after;
            return readState();
        }

        if (hadBuilds) {
            return readState();
        }

        const nextPage = Math.min(state.slide + 1, state.count);
        if (nextPage !== state.slide) {
            try {
                viewer.goToSlide(nextPage - 1);
            } catch {
                // 忽略
            }

            state.slide = nextPage;
        }

        return readState();
    },
    prev() {
        if (pdfDoc) {
            return goToPdf(state.slide - 1);
        }

        if (!viewer) {
            return readState();
        }

        ensurePresenting();
        const before = storeSlide();
        try {
            rawPrev?.();
        } catch {
            // 忽略
        }

        const after = storeSlide();
        if (after < before) {
            state.slide = after;
            return readState();
        }

        const prevPage = Math.max(1, state.slide - 1);
        try {
            viewer.goToSlide(prevPage - 1);
        } catch {
            // 忽略
        }

        state.slide = prevPage;
        return readState();
    },
    goTo(slide: number) {
        if (pdfDoc) {
            return goToPdf(slide);
        }

        const target = Math.max(1, Math.min(slide, state.count));
        try {
            viewer?.goToSlide(target - 1);
        } catch {
            // 忽略
        }

        state.slide = target;
        return readState();
    },
    state: readState
};

(window as unknown as { player: HostPlayer }).player = player;
