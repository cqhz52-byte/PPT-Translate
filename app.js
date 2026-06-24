const state = {
  file: null,
  fileType: "",
  zip: null,
  segments: [],
  slideCount: 0,
  slideSize: { cx: 12192000, cy: 6858000 },
  slideVisuals: new Map(),
  pdfPageSizes: new Map(),
  pdfTableCells: new Map(),
  pdfBytes: null,
  batchFiles: [],
  batchRunning: false,
  translationRunning: false,
  translationAbortController: null,
  savedFiles: [],
  installPrompt: null,
  wakeLock: null,
  wakeLockNoticeShown: false,
  wakeLockWarningShown: false,
  mobileView: "translate",
};

const els = {
  fileInput: document.querySelector("#fileInput"),
  uploadZone: document.querySelector(".upload-zone"),
  fileMeta: document.querySelector("#fileMeta"),
  segmentTable: document.querySelector("#segmentTable"),
  translateButton: document.querySelector("#translateButton"),
  batchTranslateButton: document.querySelector("#batchTranslateButton"),
  previewButton: document.querySelector("#previewButton"),
  downloadButton: document.querySelector("#downloadButton"),
  shareButton: document.querySelector("#shareButton"),
  resetButton: document.querySelector("#resetButton"),
  helpButton: document.querySelector("#helpButton"),
  helpDialog: document.querySelector("#helpDialog"),
  helpCloseButton: document.querySelector("#helpCloseButton"),
  installButton: document.querySelector("#installButton"),
  workspace: document.querySelector("#workspace"),
  mobileViewButton: document.querySelector("#mobileViewButton"),
  mobileViewMenu: document.querySelector("#mobileViewMenu"),
  mobileViewTargets: [...document.querySelectorAll("[data-mobile-view-target]")],
  previewDialog: document.querySelector("#previewDialog"),
  previewCloseButton: document.querySelector("#previewCloseButton"),
  previewDownloadButton: document.querySelector("#previewDownloadButton"),
  previewShareButton: document.querySelector("#previewShareButton"),
  previewBody: document.querySelector("#previewBody"),
  previewMeta: document.querySelector("#previewMeta"),
  translationDirection: document.querySelector("#translationDirection"),
  pptLayoutMode: document.querySelector("#pptLayoutMode"),
  fontScale: document.querySelector("#fontScale"),
  fontScaleValue: document.querySelector("#fontScaleValue"),
  modelName: document.querySelector("#modelName"),
  apiKey: document.querySelector("#apiKey"),
  slideCount: document.querySelector("#slideCount"),
  segmentCount: document.querySelector("#segmentCount"),
  translatedCount: document.querySelector("#translatedCount"),
  statusText: document.querySelector("#statusText"),
  statusVisual: document.querySelector("#statusVisual"),
  progressFill: document.querySelector("#progressFill"),
  batchQueue: document.querySelector("#batchQueue"),
  savedFileList: document.querySelector("#savedFileList"),
  savedFileCount: document.querySelector("#savedFileCount"),
};

const slidePathPattern = /^ppt\/slides\/slide(\d+)\.xml$/;
const wordPathPattern = /^word\/(?:document|header\d+|footnotes|endnotes|comments)\.xml$/;
const DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const SETTINGS_KEY = "deepseek-document-translator-settings-v1";
const SETTINGS_VERSION = 5;
const SAVED_FILES_DB = "curaway-translated-files-v1";
const SAVED_FILES_STORE = "files";
const CURRENT_DRAFT_DB = "curaway-current-draft-v1";
const CURRENT_DRAFT_STORE = "drafts";
const CURRENT_DRAFT_ID = "current";
const DRAFT_SAVE_DELAY = 600;
const DEEPSEEK_API_BASE = "https://api.deepseek.com";
const TRANSLATE_PROXY = "./api/translate";
const PDFJS_URL = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.min.mjs";
const PDFJS_WORKER_URL = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.worker.min.mjs";
const PDFLIB_URL = "https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/+esm";
const FONTKIT_URL = "https://cdn.jsdelivr.net/npm/@pdf-lib/fontkit@1.1.1/+esm";
const PDF_FONT_URL = "https://cdn.jsdelivr.net/gh/notofonts/noto-cjk@main/Sans/OTF/SimplifiedChinese/NotoSansCJKsc-Regular.otf";
const parser = new DOMParser();
const serializer = new XMLSerializer();
let draftSaveTimer = 0;
let isRestoringDraft = false;

loadSettings();

if ("serviceWorker" in navigator) {
  let isServiceWorkerRefreshing = false;
  navigator.serviceWorker.addEventListener("controllerchange", () => {
    if (isServiceWorkerRefreshing) return;
    isServiceWorkerRefreshing = true;
    window.location.reload();
  });

  navigator.serviceWorker.register("sw.js?v=50").then((registration) => {
    registration.update().catch(() => {});
    registration.addEventListener("updatefound", () => {
      const worker = registration.installing;
      if (!worker) return;
      worker.addEventListener("statechange", () => {
        if (worker.state === "installed" && navigator.serviceWorker.controller) {
          worker.postMessage({ type: "SKIP_WAITING" });
        }
      });
    });
    if (registration.waiting) {
      registration.waiting.postMessage({ type: "SKIP_WAITING" });
    }
  }).catch(() => {
    showToast("PWA 缓存注册失败，应用仍可在浏览器中使用。", true);
  });
}

document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "visible" && document.body.classList.contains("is-busy")) {
    requestScreenWakeLock();
  } else if (document.visibilityState === "hidden") {
    flushCurrentDraftSave();
  }
});

window.addEventListener("pagehide", flushCurrentDraftSave);

window.addEventListener("beforeinstallprompt", (event) => {
  event.preventDefault();
  state.installPrompt = event;
  els.installButton.hidden = false;
});

els.installButton.addEventListener("click", async () => {
  if (!state.installPrompt) {
    showToast("????????????????????????????????????????/???????");
    return;
  }
  state.installPrompt.prompt();
  await state.installPrompt.userChoice;
  state.installPrompt = null;
  els.installButton.hidden = true;
});

els.fileInput.addEventListener("change", async (event) => {
  const files = [...(event.target.files || [])];
  if (!files.length) return;
  if (files.length > 1) {
    queueBatchFiles(files);
    return;
  }
  await loadOfficeFile(files[0]);
});

els.mobileViewButton?.addEventListener("click", () => {
  const expanded = els.mobileViewButton.getAttribute("aria-expanded") === "true";
  setMobileMenuOpen(!expanded);
});

els.mobileViewTargets.forEach((button) => {
  button.addEventListener("click", () => {
    setMobileView(button.dataset.mobileViewTarget || "translate");
    setMobileMenuOpen(false);
  });
});

document.addEventListener("click", (event) => {
  if (!els.mobileViewMenu || !els.mobileViewButton) return;
  if (els.mobileViewMenu.hidden) return;
  if (event.target.closest(".mobile-view-switcher")) return;
  setMobileMenuOpen(false);
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") setMobileMenuOpen(false);
});

setMobileView(state.mobileView);

["dragover", "drop"].forEach((eventName) => {
  document.addEventListener(eventName, (event) => {
    event.preventDefault();
  });
});

["dragenter", "dragover"].forEach((eventName) => {
  els.uploadZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    els.uploadZone.classList.add("drag-over");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  els.uploadZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    if (eventName === "drop") {
      const files = [...(event.dataTransfer?.files || [])];
      if (files.length > 1) {
        queueBatchFiles(files);
      } else if (files[0]) {
        loadOfficeFile(files[0]);
      }
    }
    els.uploadZone.classList.remove("drag-over");
  });
});

els.translateButton.addEventListener("click", handleTranslateButtonClick);
els.batchTranslateButton?.addEventListener("click", processBatchQueue);
els.previewButton.addEventListener("click", openPreview);
els.downloadButton.addEventListener("click", downloadPresentation);
els.shareButton.addEventListener("click", sharePresentation);
els.resetButton.addEventListener("click", () => resetApp());
els.helpButton?.addEventListener("click", openHelp);
els.helpCloseButton?.addEventListener("click", closeHelp);
els.previewCloseButton.addEventListener("click", closePreview);
els.previewDownloadButton.addEventListener("click", downloadPresentation);
els.previewShareButton.addEventListener("click", sharePresentation);
els.previewDialog.addEventListener("click", (event) => {
  if (event.target === els.previewDialog) closePreview();
});
els.helpDialog?.addEventListener("click", (event) => {
  if (event.target === els.helpDialog) closeHelp();
});

[
  els.translationDirection,
  els.pptLayoutMode,
  els.fontScale,
  els.modelName,
  els.apiKey,
].forEach((element) => {
  element?.addEventListener("change", handleSettingsChange);
  element?.addEventListener("input", handleSettingsChange);
});

updateFontScaleLabel();
loadSavedFiles().catch((error) => {
  console.warn("Saved file list unavailable", error);
});
restoreCurrentDraft().catch((error) => {
  console.warn("Draft restore unavailable", error);
});

async function loadOfficeFile(file, options = {}) {
  const silent = Boolean(options.silent);
  try {
    if (/\.ppt$/i.test(file.name) && !/\.pptx$/i.test(file.name)) {
      throw new Error("当前网页版只能可靠解析和回写 PPTX。请先用 PowerPoint/WPS 将 .ppt 另存为 .pptx。");
    }

    if (/\.doc$/i.test(file.name) && !/\.docx$/i.test(file.name)) {
      throw new Error("当前网页版只能可靠解析和回写 DOCX。请先用 Word/WPS 将 .doc 另存为 .docx。");
    }

    if (!/\.(pptx|docx|pdf)$/i.test(file.name)) {
      throw new Error("请选择 .pptx、.docx 或 .pdf 文件。");
    }

    setBusy(true, "正在读取文件...");
    state.file = file;
    state.fileType = getFileTypeFromName(file.name);
    state.zip = null;
    state.segments = [];
    state.slideCount = 0;
    state.slideSize = { cx: 12192000, cy: 6858000 };
    state.slideVisuals = new Map();
    state.pdfPageSizes = new Map();
    state.pdfTableCells = new Map();
    state.pdfBytes = null;

    if (state.fileType === "pdf") {
      await loadPdfDocument(file);
    } else {
      state.zip = await JSZip.loadAsync(file);
    }

    if (state.fileType === "docx") {
      await loadWordDocument();
    } else if (state.fileType === "pptx") {
      await loadPresentation();
    }

    renderSegments();
    updateStats();
    els.fileMeta.textContent = `${file.name} · ${(file.size / 1024 / 1024).toFixed(2)} MB`;
    setStatus(`已解析 ${getFileTypeName()}，共 ${state.segments.length} 段文字。默认使用 DeepSeek 翻译。`);
    if (!options.skipDraftSave) scheduleCurrentDraftSave({ immediate: true });
    if (!silent) showToast(`${getFileTypeName()} 已载入，可以开始翻译或手动编辑。`);
  } catch (error) {
    showToast(error.message || "读取文件失败。", true);
    resetApp(false, { clearDraft: false });
  } finally {
    setBusy(false);
  }
}

async function loadPresentation() {
    state.slideSize = await readSlideSize();
    const slideFiles = Object.keys(state.zip.files)
      .map((path) => ({ path, match: path.match(slidePathPattern) }))
      .filter((item) => item.match)
      .sort((a, b) => Number(a.match[1]) - Number(b.match[1]));

    state.slideCount = slideFiles.length;

    for (const item of slideFiles) {
      const slideNumber = Number(item.match[1]);
      const xmlText = await state.zip.file(item.path).async("text");
      const doc = parser.parseFromString(xmlText, "application/xml");
      const parseError = doc.querySelector("parsererror");
      if (parseError) {
        throw new Error(`第 ${slideNumber} 页 XML 解析失败。`);
      }

      state.slideVisuals.set(item.path, inspectSlideVisuals(doc));

      const paragraphs = [...doc.getElementsByTagNameNS(DRAWING_NS, "p")];
      paragraphs.forEach((paragraph, paragraphIndex) => {
        const textNodes = [...paragraph.getElementsByTagNameNS(DRAWING_NS, "t")];
        const original = textNodes.map((node) => node.textContent || "").join("").trim();
        if (!original) return;
        const layout = inspectPresentationLayout(paragraph, original);
        state.segments.push({
          id: `${slideNumber}-${paragraphIndex}`,
          type: "pptx",
          slideNumber,
          locationLabel: String(slideNumber),
          path: item.path,
          paragraphIndex,
          textNodeCount: textNodes.length,
          layout,
          overrides: createSegmentOverrides(),
          original,
          translation: "",
        });
      });
    }
}

async function loadPdfDocument(file) {
  const pdfjs = await loadPdfJs();
  const pdfBytes = new Uint8Array(await file.arrayBuffer());
  state.pdfBytes = pdfBytes.slice();
  const pdf = await pdfjs.getDocument({ data: pdfBytes }).promise;
  state.slideCount = pdf.numPages;

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    setProgress(pdf.numPages ? pageNumber / pdf.numPages : 0);
    setStatus(`正在解析 PDF 第 ${pageNumber}/${pdf.numPages} 页，提取文字和页面背景...`);
    await waitForUiFrame();
    const page = await pdf.getPage(pageNumber);
    const viewport = page.getViewport({ scale: 1 });
    state.pdfPageSizes.set(`pdf/page-${pageNumber}`, {
      width: viewport.width || 595.28,
      height: viewport.height || 841.89,
    });
    const tableCells = await extractPdfTableCells(page, pdfjs, viewport).catch(() => []);
    state.pdfTableCells.set(`pdf/page-${pageNumber}`, tableCells);
    const pageSample = await renderPdfPageSample(page).catch((error) => {
      console.warn("PDF background sampling failed", error);
      return null;
    });
    const content = await page.getTextContent();
    const lines = extractPdfLineSegments(content.items, viewport.width || 595.28);

    lines.forEach((line, index) => {
      const tableCell = findPdfTableCell(line.bounds, tableCells);
      const backgroundColor = samplePdfBackgroundColor(pageSample, line.bounds, viewport.width || 595.28, viewport.height || 841.89);
      state.segments.push({
        id: `pdf-${pageNumber}-${index}`,
        type: "pdf",
        slideNumber: pageNumber,
        locationLabel: `第 ${pageNumber} 页`,
        path: `pdf/page-${pageNumber}`,
        paragraphIndex: index,
        textNodeCount: 1,
        layout: {
          bounds: line.bounds,
          fontSize: line.fontSize,
          availableWidth: tableCell ? Math.max(line.availableWidth, tableCell.width - 4) : line.availableWidth,
          availableHeight: tableCell ? Math.max(line.availableHeight, tableCell.height - 4) : line.availableHeight,
          rowSegmentCount: tableCell ? Math.max(2, line.rowSegmentCount) : line.rowSegmentCount,
          tableCell,
          cellAlign: tableCell ? inferPdfCellAlign(line.bounds, tableCell, line.text) : "",
          backgroundColor,
          textColor: getReadablePdfTextColor(backgroundColor),
        },
        overrides: createSegmentOverrides(),
        original: line.text,
        translation: "",
      });
    });
  }

  if (!state.segments.length) {
    throw new Error("这个 PDF 没有可提取的文本。若是扫描版 PDF，需要先 OCR 后再翻译。");
  }
}

async function loadPdfJs() {
  if (!window.pdfjsLib) {
    const module = await import(PDFJS_URL);
    window.pdfjsLib = module;
  }

  window.pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER_URL;
  return window.pdfjsLib;
}

async function renderPdfPageSample(page) {
  const scale = 0.72;
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.ceil(viewport.width));
  canvas.height = Math.max(1, Math.ceil(viewport.height));
  const context = canvas.getContext("2d", { willReadFrequently: true });
  if (!context) return null;

  await page.render({ canvasContext: context, viewport }).promise;
  return { canvas, context, scale };
}

async function renderPdfPageBackgroundForExport(page, pdfDoc, pageSize) {
  const width = pageSize.width || 595.28;
  const height = pageSize.height || 841.89;
  const maxPixels = isLikelyMobileDevice() ? 2600000 : 4200000;
  const scale = Math.max(1.35, Math.min(2.2, Math.sqrt(maxPixels / Math.max(1, width * height))));
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.ceil(viewport.width));
  canvas.height = Math.max(1, Math.ceil(viewport.height));
  const context = canvas.getContext("2d", { alpha: false });
  if (!context) throw new Error("Canvas unavailable for PDF export.");

  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, canvas.width, canvas.height);
  await page.render({ canvasContext: context, viewport }).promise;
  const bytes = await canvasToArrayBuffer(canvas, "image/jpeg", 0.94);
  canvas.width = 1;
  canvas.height = 1;
  return { image: await pdfDoc.embedJpg(bytes), width, height };
}

function canvasToArrayBuffer(canvas, type, quality) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (!blob) {
        reject(new Error("PDF page rendering failed."));
        return;
      }
      blob.arrayBuffer().then(resolve, reject);
    }, type, quality);
  });
}

function isLikelyMobileDevice() {
  return /Android|iPhone|iPad|iPod|Mobile/i.test(navigator.userAgent) || Math.min(window.innerWidth, window.innerHeight) < 760;
}

function samplePdfBackgroundColor(sample, bounds, pageWidth, pageHeight) {
  if (!sample || !bounds) return { r: 1, g: 1, b: 1 };

  const { context, canvas, scale } = sample;
  const pad = 3;
  const x = clampNumber((bounds.x - pad) * scale, 0, canvas.width - 1);
  const y = clampNumber((pageHeight - bounds.y - bounds.height - pad) * scale, 0, canvas.height - 1);
  const width = clampNumber((bounds.width + pad * 2) * scale, 1, canvas.width - x);
  const height = clampNumber((bounds.height + pad * 2) * scale, 1, canvas.height - y);

  try {
    const image = context.getImageData(Math.floor(x), Math.floor(y), Math.ceil(width), Math.ceil(height));
    const pixels = image.data;
    const reds = [];
    const greens = [];
    const blues = [];
    const step = Math.max(4, Math.floor((pixels.length / 4) / 180));

    for (let index = 0; index < pixels.length; index += step * 4) {
      reds.push(pixels[index]);
      greens.push(pixels[index + 1]);
      blues.push(pixels[index + 2]);
    }

    return {
      r: medianByte(reds) / 255,
      g: medianByte(greens) / 255,
      b: medianByte(blues) / 255,
    };
  } catch (error) {
    console.warn("PDF background sample unavailable", error);
    return { r: 1, g: 1, b: 1 };
  }
}

function medianByte(values) {
  if (!values.length) return 255;
  values.sort((a, b) => a - b);
  return values[Math.floor(values.length / 2)];
}

function getReadablePdfTextColor(background) {
  const color = background || { r: 1, g: 1, b: 1 };
  const luminance = 0.2126 * color.r + 0.7152 * color.g + 0.0722 * color.b;
  return luminance < 0.48 ? { r: 0.96, g: 0.97, b: 0.98 } : { r: 0.06, g: 0.08, b: 0.09 };
}

function clampNumber(value, min, max) {
  const number = Number(value);
  if (!Number.isFinite(number)) return min;
  return Math.max(min, Math.min(max, number));
}

async function extractPdfTableCells(page, pdfjs, viewport) {
  const operatorList = await page.getOperatorList();
  const segments = extractPdfVectorSegments(operatorList, pdfjs, viewport);
  return buildPdfTableCellsFromLines(segments, viewport.width || 595.28, viewport.height || 841.89);
}

function extractPdfVectorSegments(operatorList, pdfjs, viewport) {
  const OPS = pdfjs.OPS || {};
  const fnArray = operatorList.fnArray || [];
  const argsArray = operatorList.argsArray || [];
  const matrixStack = [[1, 0, 0, 1, 0, 0]];
  const segments = [];

  const op = {
    save: OPS.save ?? 10,
    restore: OPS.restore ?? 11,
    transform: OPS.transform ?? 12,
    constructPath: OPS.constructPath ?? 91,
    moveTo: OPS.moveTo ?? 13,
    lineTo: OPS.lineTo ?? 14,
    curveTo: OPS.curveTo ?? 15,
    curveTo2: OPS.curveTo2 ?? 16,
    curveTo3: OPS.curveTo3 ?? 17,
    closePath: OPS.closePath ?? 18,
    rectangle: OPS.rectangle ?? 19,
  };

  fnArray.forEach((fn, index) => {
    const args = argsArray[index] || [];
    if (fn === op.save) {
      matrixStack.push([...matrixStack[matrixStack.length - 1]]);
      return;
    }
    if (fn === op.restore) {
      if (matrixStack.length > 1) matrixStack.pop();
      return;
    }
    if (fn === op.transform && args.length >= 6) {
      matrixStack[matrixStack.length - 1] = multiplyPdfMatrix(matrixStack[matrixStack.length - 1], args);
      return;
    }
    if (fn !== op.constructPath) return;

    const pathOps = args[0] || [];
    const pathArgs = args[1] || [];
    let cursor = 0;
    let current = null;
    const matrix = matrixStack[matrixStack.length - 1];

    pathOps.forEach((pathOp) => {
      if (pathOp === op.moveTo) {
        current = transformPdfPoint(pathArgs[cursor], pathArgs[cursor + 1], matrix);
        cursor += 2;
        return;
      }
      if (pathOp === op.lineTo) {
        const next = transformPdfPoint(pathArgs[cursor], pathArgs[cursor + 1], matrix);
        cursor += 2;
        if (current) segments.push(normalizePdfLineSegment(current, next));
        current = next;
        return;
      }
      if (pathOp === op.rectangle) {
        const x = pathArgs[cursor];
        const y = pathArgs[cursor + 1];
        const w = pathArgs[cursor + 2];
        const h = pathArgs[cursor + 3];
        cursor += 4;
        const p1 = transformPdfPoint(x, y, matrix);
        const p2 = transformPdfPoint(x + w, y, matrix);
        const p3 = transformPdfPoint(x + w, y + h, matrix);
        const p4 = transformPdfPoint(x, y + h, matrix);
        segments.push(normalizePdfLineSegment(p1, p2), normalizePdfLineSegment(p2, p3), normalizePdfLineSegment(p3, p4), normalizePdfLineSegment(p4, p1));
        current = p1;
        return;
      }
      if (pathOp === op.curveTo) {
        cursor += 6;
        return;
      }
      if (pathOp === op.curveTo2 || pathOp === op.curveTo3) {
        cursor += 4;
      }
    });
  });

  return segments.filter((segment) => {
    const horizontal = Math.abs(segment.y1 - segment.y2) <= 1.2 && Math.abs(segment.x2 - segment.x1) >= 10;
    const vertical = Math.abs(segment.x1 - segment.x2) <= 1.2 && Math.abs(segment.y2 - segment.y1) >= 10;
    const insidePage = segment.x2 >= -2 && segment.y2 >= -2 && segment.x1 <= viewport.width + 2 && segment.y1 <= viewport.height + 2;
    return insidePage && (horizontal || vertical);
  });
}

function multiplyPdfMatrix(a, b) {
  return [
    a[0] * b[0] + a[2] * b[1],
    a[1] * b[0] + a[3] * b[1],
    a[0] * b[2] + a[2] * b[3],
    a[1] * b[2] + a[3] * b[3],
    a[0] * b[4] + a[2] * b[5] + a[4],
    a[1] * b[4] + a[3] * b[5] + a[5],
  ];
}

function transformPdfPoint(x, y, matrix) {
  return {
    x: matrix[0] * x + matrix[2] * y + matrix[4],
    y: matrix[1] * x + matrix[3] * y + matrix[5],
  };
}

function normalizePdfLineSegment(a, b) {
  return {
    x1: Math.min(a.x, b.x),
    y1: Math.min(a.y, b.y),
    x2: Math.max(a.x, b.x),
    y2: Math.max(a.y, b.y),
  };
}

function buildPdfTableCellsFromLines(segments, pageWidth, pageHeight) {
  const horizontal = segments.filter((segment) => Math.abs(segment.y1 - segment.y2) <= 1.2);
  const vertical = segments.filter((segment) => Math.abs(segment.x1 - segment.x2) <= 1.2);
  if (horizontal.length < 2 || vertical.length < 2) return [];

  const xs = clusterPdfCoords(vertical.flatMap((segment) => [segment.x1]));
  const ys = clusterPdfCoords(horizontal.flatMap((segment) => [segment.y1]));
  const cells = [];

  for (let xi = 0; xi < xs.length - 1; xi += 1) {
    for (let yi = 0; yi < ys.length - 1; yi += 1) {
      const x1 = xs[xi];
      const x2 = xs[xi + 1];
      const y1 = ys[yi];
      const y2 = ys[yi + 1];
      const width = x2 - x1;
      const height = y2 - y1;
      if (width < 14 || height < 8 || width > pageWidth * 0.95 || height > pageHeight * 0.6) continue;

      const hasTop = hasPdfHorizontalCover(horizontal, y2, x1, x2);
      const hasBottom = hasPdfHorizontalCover(horizontal, y1, x1, x2);
      const hasLeft = hasPdfVerticalCover(vertical, x1, y1, y2);
      const hasRight = hasPdfVerticalCover(vertical, x2, y1, y2);
      if (hasTop && hasBottom && hasLeft && hasRight) {
        cells.push({ x: x1, y: y1, width, height, x2, y2 });
      }
    }
  }

  return mergePdfTableCells(cells);
}

function clusterPdfCoords(values, tolerance = 2) {
  const sorted = values.filter((value) => Number.isFinite(value)).sort((a, b) => a - b);
  const clusters = [];
  sorted.forEach((value) => {
    const cluster = clusters[clusters.length - 1];
    if (cluster && Math.abs(cluster[cluster.length - 1] - value) <= tolerance) {
      cluster.push(value);
    } else {
      clusters.push([value]);
    }
  });
  return clusters.map((cluster) => cluster.reduce((sum, value) => sum + value, 0) / cluster.length);
}

function hasPdfHorizontalCover(lines, y, x1, x2) {
  return lines.some((line) => Math.abs(line.y1 - y) <= 2.2 && line.x1 <= x1 + 3 && line.x2 >= x2 - 3);
}

function hasPdfVerticalCover(lines, x, y1, y2) {
  return lines.some((line) => Math.abs(line.x1 - x) <= 2.2 && line.y1 <= y1 + 3 && line.y2 >= y2 - 3);
}

function mergePdfTableCells(cells) {
  const map = new Map();
  cells.forEach((cell) => {
    const key = `${Math.round(cell.x)}:${Math.round(cell.y)}:${Math.round(cell.width)}:${Math.round(cell.height)}`;
    map.set(key, cell);
  });
  return [...map.values()];
}

function findPdfTableCell(bounds, cells) {
  if (!cells.length) return null;
  const cx = bounds.x + bounds.width / 2;
  const cy = bounds.y + bounds.height / 2;
  const candidates = cells.filter((cell) => cx >= cell.x - 1 && cx <= cell.x2 + 1 && cy >= cell.y - 1 && cy <= cell.y2 + 1);
  if (!candidates.length) return null;
  return candidates.sort((a, b) => a.width * a.height - b.width * b.height)[0];
}

function inferPdfCellAlign(bounds, cell, text) {
  const centerDelta = Math.abs(bounds.x + bounds.width / 2 - (cell.x + cell.width / 2));
  const shortText = [...String(text || "")].length <= 16;
  return shortText || centerDelta < cell.width * 0.18 ? "center" : "left";
}

function extractPdfLineSegments(items, pageWidth = 595.28) {
  const rows = new Map();
  items.forEach((item) => {
    const text = String(item.str || "").trim();
    if (!text) return;

    const transform = item.transform || [];
    const x = Number(transform[4] || 0);
    const y = Number(transform[5] || 0);
    const key = String(Math.round(y / 3) * 3);
    const row = rows.get(key) || [];
    const height = Math.max(6, Math.abs(Number(transform[3] || item.height || 10)));
    row.push({
      x,
      y,
      text,
      width: Number(item.width || 0),
      height,
      hasEol: Boolean(item.hasEOL),
    });
    rows.set(key, row);
  });

  const sortedRows = [...rows.entries()].sort((a, b) => Number(b[0]) - Number(a[0]));

  return sortedRows
    .sort((a, b) => Number(b[0]) - Number(a[0]))
    .flatMap(([key, row], index) => {
      const nextKey = sortedRows[index + 1]?.[0];
      const rowGap = nextKey === undefined ? 0 : Math.max(0, Number(key) - Number(nextKey));
      return splitPdfRowIntoSegments(row, pageWidth, rowGap);
    })
    .filter((line) => line.text);
}

function splitPdfRowIntoSegments(row, pageWidth, rowGap = 0) {
  const sorted = row.sort((a, b) => a.x - b.x);
  const maxHeight = Math.max(...sorted.map((item) => item.height));
  const charWidths = sorted
    .map((item) => item.width / Math.max(1, [...item.text].length))
    .filter((width) => Number.isFinite(width) && width > 0);
  const averageCharWidth = charWidths.length ? charWidths.reduce((sum, width) => sum + width, 0) / charWidths.length : maxHeight * 0.45;
  const gapThreshold = Math.max(10, maxHeight * 1.1, averageCharWidth * 3.2);
  const groups = [];
  let current = [];

  sorted.forEach((item) => {
    const previous = current[current.length - 1];
    const gap = previous ? item.x - (previous.x + previous.width) : 0;
    if (previous && (previous.hasEol || gap > gapThreshold)) {
      groups.push(current);
      current = [];
    }
    current.push(item);
  });

  if (current.length) groups.push(current);

  const groupBounds = groups.map((group) => ({
    group,
    minX: Math.min(...group.map((item) => item.x)),
    minY: Math.min(...group.map((item) => item.y)),
    maxX: Math.max(...group.map((item) => item.x + item.width)),
    maxHeight: Math.max(...group.map((item) => item.height)),
  }));

  return groupBounds.map((item, index) => {
    const group = item.group;
    const next = groupBounds[index + 1];
    const minX = Math.min(...group.map((item) => item.x));
    const minY = Math.min(...group.map((item) => item.y));
    const maxX = Math.max(...group.map((item) => item.x + item.width));
    const localMaxHeight = Math.max(...group.map((item) => item.height));
    const text = group.map((item) => item.text).join(" ").replace(/\s+/g, " ").trim();
    const textWidth = Math.max(8, maxX - minX);
    const availableRight = next ? next.minX - 3 : pageWidth - 30;
    const availableWidth = Math.max(textWidth, availableRight - minX);
    const isTableLike = groupBounds.length > 1;
    const baseHeight = Math.max(7, localMaxHeight * 1.32);
    const availableHeight = isTableLike && rowGap > localMaxHeight * 1.6 ? Math.max(baseHeight, rowGap * 0.72) : baseHeight;

    return {
      text,
      fontSize: localMaxHeight,
      availableWidth,
      availableHeight,
      rowSegmentCount: groups.length,
      bounds: {
        x: minX,
        y: minY - localMaxHeight * 0.25,
        width: textWidth,
        height: baseHeight,
      },
    };
  });
}

async function loadWordDocument() {
  const wordFiles = Object.keys(state.zip.files)
    .filter((path) => wordPathPattern.test(path))
    .sort((a, b) => a.localeCompare(b));

  if (!wordFiles.includes("word/document.xml")) {
    throw new Error("未找到 Word 正文 document.xml，请确认这是有效的 DOCX 文件。");
  }

  state.slideCount = wordFiles.length;

  for (const path of wordFiles) {
    const xmlText = await state.zip.file(path).async("text");
    const doc = parser.parseFromString(xmlText, "application/xml");
    const parseError = doc.querySelector("parsererror");
    if (parseError) {
      throw new Error(`${path} XML 解析失败。`);
    }

    const partLabel = getWordPartLabel(path);
    const paragraphs = [...doc.getElementsByTagNameNS(WORD_NS, "p")];
    paragraphs.forEach((paragraph, paragraphIndex) => {
      if (hasWordField(paragraph)) return;
      const textNodes = [...paragraph.getElementsByTagNameNS(WORD_NS, "t")];
      const rawOriginal = textNodes.map((node) => node.textContent || "").join("");
      const original = rawOriginal.trim();
      if (!original) return;
      const leadingWhitespace = rawOriginal.match(/^\s*/)?.[0] || "";
      const trailingWhitespace = rawOriginal.match(/\s*$/)?.[0] || "";
      const syntheticAlignment = detectWordSyntheticAlignment(path, paragraph, original, leadingWhitespace, trailingWhitespace);

      state.segments.push({
        id: `${path}-${paragraphIndex}`,
        type: "docx",
        slideNumber: partLabel,
        locationLabel: partLabel,
        path,
        paragraphIndex,
        textNodeCount: textNodes.length,
        wordLeadingWhitespace: leadingWhitespace,
        wordTrailingWhitespace: trailingWhitespace,
        wordSyntheticAlignment: syntheticAlignment,
        overrides: createSegmentOverrides(),
        original,
        translation: "",
      });
    });
  }
}

function renderSegments() {
  if (!state.segments.length) {
    els.segmentTable.innerHTML = `
      <tr class="empty-row">
        <td colspan="3">
          <div class="empty-state">
            <strong>未发现可翻译文字</strong>
            <span>请确认文件中包含普通文本框文字。</span>
          </div>
        </td>
      </tr>
    `;
    return;
  }

  const fragment = document.createDocumentFragment();
  state.segments.forEach((segment, index) => {
    const row = document.createElement("tr");

    const slideCell = document.createElement("td");
    slideCell.textContent = String(segment.locationLabel || segment.slideNumber);

    const sourceCell = document.createElement("td");
    const source = document.createElement("div");
    source.className = "source-text";
    source.textContent = segment.original;
    sourceCell.append(source);

    const targetCell = document.createElement("td");
    const textarea = document.createElement("textarea");
    textarea.value = segment.translation;
    textarea.placeholder = "输入译文";
    textarea.dataset.index = String(index);
    textarea.addEventListener("input", () => {
      state.segments[Number(textarea.dataset.index)].translation = textarea.value;
      updateStats();
      rerenderPreviewIfOpen();
      scheduleCurrentDraftSave();
    });
    targetCell.append(textarea, createSegmentControls(segment, index));

    row.append(slideCell, sourceCell, targetCell);
    fragment.append(row);
  });

  els.segmentTable.replaceChildren(fragment);
}

function handleTranslateButtonClick() {
  if (state.translationRunning) {
    cancelTranslation();
    return;
  }
  translateAll();
}

function cancelTranslation() {
  if (!state.translationRunning) return;
  state.translationAbortController?.abort();
  setStatus("正在停止翻译，请稍候...");
}

async function translateAll() {
  const apiKey = els.apiKey.value.trim();
  const apiBase = DEEPSEEK_API_BASE;
  const apiProxy = TRANSLATE_PROXY;
  const model = els.modelName.value.trim();
  const direction = getDirectionConfig(els.translationDirection.value);

  if (!apiKey || !model) {
    showToast("请先填写模型和 API Key。", true);
    return;
  }

  if (state.translationRunning) return false;
  const abortController = new AbortController();
  state.translationRunning = true;
  state.translationAbortController = abortController;

  try {
    setBusy(true, "正在翻译...");
    const untranslated = state.segments.filter((segment) => !segment.translation.trim());

    for (let index = 0; index < untranslated.length; index += 1) {
      if (abortController.signal.aborted) throw new DOMException("Translation stopped", "AbortError");
      const segment = untranslated[index];
      setProgress(index / untranslated.length);
      setStatus(`正在翻译 ${segment.locationLabel || segment.slideNumber}：${index + 1}/${untranslated.length}`);
      segment.translation = await translateText({
        apiBase,
        apiProxy,
        apiKey,
        model,
        direction,
        text: segment.original,
        signal: abortController.signal,
      });
      updateTextarea(segment);
      updateStats();
      scheduleCurrentDraftSave();
    }

    setProgress(1);
    setStatus("翻译完成，请检查译文后导出。");
    showToast("自动翻译完成。");
    scheduleCurrentDraftSave({ immediate: true });
    return true;
  } catch (error) {
    if (error?.name === "AbortError") {
      setStatus("已停止翻译。已完成的译文会保留，可继续编辑或再次点击自动翻译续翻。");
      showToast("已停止翻译。");
      scheduleCurrentDraftSave({ immediate: true });
      return false;
    }
    showToast(error.message || "翻译失败。", true);
    setStatus("翻译中断，请检查接口配置或网络连接。");
    return false;
  } finally {
    if (state.translationAbortController === abortController) {
      state.translationAbortController = null;
    }
    state.translationRunning = false;
    setBusy(false);
  }
}

async function translateText({ apiBase, apiProxy, apiKey, model, direction, text, signal }) {
  const response = await fetch(apiProxy, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    signal,
    body: JSON.stringify({
      apiBase,
      apiKey,
      model,
      instruction: direction.instruction,
      text,
    }),
  });

  if (!response.ok) {
    const detail = await response.json().catch(() => null);
    throw new Error(detail?.error || `接口返回 ${response.status}`);
  }

  const data = await response.json();
  return normalizeTranslation(data.translation || "", text);
}

function queueBatchFiles(files) {
  const supported = files.filter((file) => /\.(pptx|docx|pdf)$/i.test(file.name));
  const skipped = files.length - supported.length;
  state.batchFiles = supported;
  renderBatchQueue();
  updateBatchButton();

  if (supported.length) {
    setStatus(`已加入 ${supported.length} 个批量文件。点击“批量翻译并保存”开始处理。`);
  }
  if (skipped) {
    showToast(`${skipped} 个文件格式暂不支持，已跳过。`, true);
  }
}

function renderBatchQueue() {
  if (!els.batchQueue) return;
  els.batchQueue.hidden = !state.batchFiles.length;
  if (!state.batchFiles.length) {
    els.batchQueue.replaceChildren();
    return;
  }

  const fragment = document.createDocumentFragment();
  const title = document.createElement("strong");
  title.textContent = `待批量处理：${state.batchFiles.length} 个`;
  fragment.append(title);

  state.batchFiles.forEach((file) => {
    const item = document.createElement("div");
    item.className = "batch-file";
    item.textContent = file.name;
    fragment.append(item);
  });

  els.batchQueue.replaceChildren(fragment);
}

function updateBatchButton() {
  if (!els.batchTranslateButton) return;
  els.batchTranslateButton.disabled = !state.batchFiles.length || state.batchRunning || document.body.classList.contains("is-busy");
}

async function processBatchQueue() {
  if (!state.batchFiles.length) return;

  const apiKey = els.apiKey.value.trim();
  const model = els.modelName.value.trim();
  if (!apiKey || !model) {
    showToast("请先填写模型和 API Key。", true);
    return;
  }

  const files = [...state.batchFiles];
  const saved = [];
  const failed = [];
  state.batchRunning = true;

  try {
    setBusy(true, `正在批量翻译 0/${files.length}...`);

    for (let index = 0; index < files.length; index += 1) {
      const file = files[index];
      setProgress(index / files.length);
      setStatus(`批量处理中：${index + 1}/${files.length} · ${file.name}`);
      await waitForUiFrame();

      try {
        await loadOfficeFile(file);
        const translated = await translateAll();
        if (!translated) {
          failed.push(file.name);
          break;
        }
        const { blob, filename } = await generateTranslatedFile();
        await saveGeneratedFile(blob, filename);
        saved.push(filename);
      } catch (error) {
        console.error("Batch file failed", file.name, error);
        failed.push(file.name);
      }
    }

    state.batchFiles = [];
    renderBatchQueue();
    updateBatchButton();
    await loadSavedFiles();
    setProgress(1);
    setStatus(`批量完成：成功 ${saved.length} 个，失败 ${failed.length} 个。`);
    showToast(`批量翻译完成，已保存 ${saved.length} 个文件。${failed.length ? `失败 ${failed.length} 个。` : ""}`, Boolean(failed.length));
  } finally {
    state.batchRunning = false;
    setBusy(false);
    updateBatchButton();
  }
}

async function downloadPresentation() {
  if (!state.file) return;

  try {
    setBusy(true, `正在生成 ${getFileTypeName()}...`);
    const { blob, filename } = await generateTranslatedFile();
    await saveGeneratedFile(blob, filename);
    saveBlobAsFile(blob, filename);
    showToast(`已生成翻译版 ${getFileTypeName()}。`);
  } catch (error) {
    showToast(error.message || "导出失败。", true);
  } finally {
    setBusy(false);
  }
}

async function sharePresentation() {
  if (!state.file) return;

  let fallbackBlob = null;
  let fallbackFilename = "";
  try {
    setBusy(true, `正在生成可分享的 ${getFileTypeName()}...`);
    const { blob, filename } = await generateTranslatedFile();
    await saveGeneratedFile(blob, filename);
    fallbackBlob = blob;
    fallbackFilename = filename;
    setStatus("File ready. Opening system share panel...");
    await waitForUiFrame();
    const file = new File([blob], filename, { type: blob.type || getOutputMimeType() });

    if (navigator.canShare?.({ files: [file] })) {
      await navigator.share({
        files: [file],
        title: filename,
        text: "CuraWay 文档翻译工具生成的翻译文件",
      });
      showToast("已打开系统分享面板。");
      return;
    }

    saveBlobAsFile(blob, filename);
    showToast("当前浏览器不支持直接分享文件，已先下载翻译文件，可从微信或文件管理器转发。", true);
  } catch (error) {
    if (error?.name !== "AbortError" && fallbackBlob) {
      saveBlobAsFile(fallbackBlob, fallbackFilename);
      showToast("Share panel could not open. The translated file has been downloaded instead.", true);
      return;
    }
    if (error?.name === "AbortError") {
      showToast("已取消分享。");
    } else {
      showToast(error.message || "分享失败。", true);
    }
  } finally {
    setBusy(false);
  }
}

async function generateTranslatedFile() {
  if (!state.file) throw new Error("请先选择文件。");

  normalizeAllTranslations();

  if (state.fileType === "pdf") {
    const blob = await downloadPdfTranslation();
    return { blob, filename: buildOutputName(state.file.name) };
  }

  if (!state.zip) throw new Error("缺少文档数据，请重新上传文件。");

  const grouped = groupByPath(state.segments);

  for (const [path, segments] of grouped.entries()) {
    const xmlText = await state.zip.file(path).async("text");
    const doc = parser.parseFromString(xmlText, "application/xml");

    if (state.fileType === "docx") {
      writeWordSegments(doc, segments);
    } else {
      writePresentationSegments(doc, segments);
    }

    state.zip.file(path, serializer.serializeToString(doc));
  }

  const blob = await state.zip.generateAsync({
    type: "blob",
    mimeType: getOutputMimeType(),
  });

  return { blob, filename: buildOutputName(state.file.name) };
}

function normalizeAllTranslations() {
  state.segments.forEach((segment) => {
    if (!segment.translation) return;
    const normalized = applyTerminologyRules(segment.translation, segment.original);
    if (normalized !== segment.translation) {
      segment.translation = normalized;
      updateTextarea(segment);
    }
  });
}

async function saveGeneratedFile(blob, filename) {
  const record = {
    id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
    filename,
    type: blob.type || getOutputMimeType(),
    size: blob.size,
    createdAt: new Date().toISOString(),
    blob,
  };

  state.savedFiles = [record, ...state.savedFiles].slice(0, 30);
  renderSavedFiles();

  try {
    const db = await openSavedFilesDb();
    await idbPut(db, record);
  } catch (error) {
    console.warn("Saved file persistence failed", error);
  }

  return record;
}

async function loadSavedFiles() {
  try {
    const db = await openSavedFilesDb();
    state.savedFiles = (await idbGetAll(db)).sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt)));
  } catch (error) {
    console.warn("Saved files unavailable", error);
    state.savedFiles = [];
  }
  renderSavedFiles();
}

function renderSavedFiles() {
  if (!els.savedFileList) return;
  els.savedFileCount && (els.savedFileCount.textContent = String(state.savedFiles.length));
  els.savedFileList.replaceChildren();

  if (!state.savedFiles.length) {
    const empty = document.createElement("p");
    empty.className = "saved-empty";
    empty.textContent = "生成后的文件会保存在这里，可再次下载或分享。";
    els.savedFileList.append(empty);
    return;
  }

  const fragment = document.createDocumentFragment();
  state.savedFiles.forEach((record) => {
    const item = document.createElement("article");
    item.className = "saved-file";

    const meta = document.createElement("div");
    meta.className = "saved-file-meta";
    const name = document.createElement("strong");
    name.textContent = record.filename;
    const detail = document.createElement("span");
    detail.textContent = `${formatBytes(record.size)} · ${formatSavedTime(record.createdAt)}`;
    meta.append(name, detail);

    const actions = document.createElement("div");
    actions.className = "saved-file-actions";
    const download = document.createElement("button");
    download.type = "button";
    download.textContent = "下载";
    download.addEventListener("click", () => downloadSavedFile(record.id));
    const share = document.createElement("button");
    share.type = "button";
    share.textContent = "分享";
    share.addEventListener("click", () => shareSavedFile(record.id));
    const remove = document.createElement("button");
    remove.type = "button";
    remove.textContent = "删除";
    remove.addEventListener("click", () => deleteSavedFile(record.id));
    actions.append(download, share, remove);

    item.append(meta, actions);
    fragment.append(item);
  });

  els.savedFileList.append(fragment);
}

function getSavedRecord(id) {
  return state.savedFiles.find((record) => record.id === id);
}

async function downloadSavedFile(id) {
  const record = getSavedRecord(id);
  if (!record) return;
  saveBlobAsFile(record.blob, record.filename);
}

async function shareSavedFile(id) {
  const record = getSavedRecord(id);
  if (!record) return;
  await shareBlobFile(record.blob, record.filename, record.type);
}

async function deleteSavedFile(id) {
  state.savedFiles = state.savedFiles.filter((record) => record.id !== id);
  renderSavedFiles();
  try {
    const db = await openSavedFilesDb();
    await idbDelete(db, id);
  } catch (error) {
    console.warn("Saved file delete failed", error);
  }
}

async function shareBlobFile(blob, filename, type = "") {
  const file = new File([blob], filename, { type: type || blob.type || getOutputMimeType() });
  if (navigator.canShare?.({ files: [file] })) {
    await navigator.share({
      files: [file],
      title: filename,
      text: "CuraWay 文档翻译工具生成的翻译文件",
    });
    showToast("已打开系统分享面板。");
    return;
  }

  saveBlobAsFile(blob, filename);
  showToast("当前浏览器不支持直接分享文件，已先下载文件。", true);
}

function openSavedFilesDb() {
  return new Promise((resolve, reject) => {
    if (!("indexedDB" in window)) {
      reject(new Error("IndexedDB unavailable"));
      return;
    }
    const request = indexedDB.open(SAVED_FILES_DB, 1);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(SAVED_FILES_STORE)) {
        db.createObjectStore(SAVED_FILES_STORE, { keyPath: "id" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

function idbPut(db, record) {
  return idbRequest(db, SAVED_FILES_STORE, "readwrite", (store) => store.put(record));
}

function idbGetAll(db) {
  return idbRequest(db, SAVED_FILES_STORE, "readonly", (store) => store.getAll());
}

function idbDelete(db, id) {
  return idbRequest(db, SAVED_FILES_STORE, "readwrite", (store) => store.delete(id));
}

function idbRequest(db, storeName, mode, action) {
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(storeName, mode);
    const request = action(transaction.objectStore(storeName));
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
    transaction.onerror = () => reject(transaction.error);
  });
}

function formatBytes(bytes = 0) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

function formatSavedTime(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  return date.toLocaleString("zh-CN", { month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit" });
}

function saveBlobAsFile(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.target = "_blank";
  link.rel = "noopener";
  link.style.display = "none";
  document.body.append(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 30000);
}

function getOutputMimeType() {
  if (state.fileType === "docx") return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  if (state.fileType === "pdf") return "application/pdf";
  return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
}

function groupByPath(segments) {
  return segments.reduce((map, segment) => {
    if (!map.has(segment.path)) {
      map.set(segment.path, []);
    }
    map.get(segment.path).push(segment);
    return map;
  }, new Map());
}

function writePresentationSegments(doc, segments) {
  const paragraphs = [...doc.getElementsByTagNameNS(DRAWING_NS, "p")];

  segments.forEach((segment) => {
    const paragraph = paragraphs[segment.paragraphIndex];
    if (!paragraph || !segment.translation.trim()) return;

    const textNodes = [...paragraph.getElementsByTagNameNS(DRAWING_NS, "t")];
    if (!textNodes.length) return;

    textNodes[0].textContent = segment.translation.trim();
    normalizePresentationRunStyle(textNodes[0], segment);
    applyPresentationBounds(paragraph, segment);
    applyPresentationLayout(paragraph, segment);
    clearRemainingTextNodes(textNodes);
  });
}

function writeWordSegments(doc, segments) {
  const paragraphs = [...doc.getElementsByTagNameNS(WORD_NS, "p")];

  segments.forEach((segment) => {
    const paragraph = paragraphs[segment.paragraphIndex];
    if (!paragraph || !segment.translation.trim()) return;
    if (hasWordField(paragraph)) return;

    const textNodes = [...paragraph.getElementsByTagNameNS(WORD_NS, "t")];
    if (!textNodes.length) return;

    if (segment.wordSyntheticAlignment) {
      applyWordParagraphJustification(paragraph, segment.wordSyntheticAlignment);
    }
    writeWordTextNodes(textNodes, segment);
  });
}

function writeWordTextNodes(textNodes, segment) {
  const translation = segment.translation.trim();
  const originalTexts = textNodes.map((node) => node.textContent || "");
  const firstContentIndex = originalTexts.findIndex((text) => text.trim());
  const shouldUseContentRuns = Boolean(segment.wordSyntheticAlignment) && firstContentIndex >= 0;
  const activeTextNodes = shouldUseContentRuns
    ? textNodes.filter((node, index) => originalTexts[index].trim())
    : textNodes;
  const text = shouldUseContentRuns
    ? translation
    : `${segment.wordLeadingWhitespace || ""}${translation}${segment.wordTrailingWhitespace || ""}`;
  const characters = [...text];
  const originalLengths = activeTextNodes.map((node) => [...(node.textContent || "")].length);
  const hasOriginalText = originalLengths.some((length) => length > 0);

  if (shouldUseContentRuns) {
    for (let index = 0; index < textNodes.length; index += 1) {
      if (originalTexts[index].trim()) continue;
      textNodes[index].textContent = "";
      updateWordTextSpacePreserve(textNodes[index]);
    }
  }

  writeWordTextIntoNodes(activeTextNodes, characters, originalLengths, hasOriginalText);
}

function writeWordTextIntoNodes(textNodes, characters, originalLengths, hasOriginalText) {
  if (!hasOriginalText) {
    textNodes[0].textContent = characters.join("");
    normalizeWordTextRunForTranslation(textNodes[0]);
    updateWordTextSpacePreserve(textNodes[0]);
    clearRemainingTextNodes(textNodes);
    return;
  }

  let offset = 0;
  textNodes.forEach((node, index) => {
    const isLast = index === textNodes.length - 1;
    const originalLength = originalLengths[index];
    const take = isLast ? characters.length - offset : Math.min(originalLength, characters.length - offset);
    node.textContent = take > 0 ? characters.slice(offset, offset + take).join("") : "";
    offset += Math.max(0, take);
    normalizeWordTextRunForTranslation(node);
    updateWordTextSpacePreserve(node);
  });
}

function detectWordSyntheticAlignment(path, paragraph, original, leadingWhitespace, trailingWhitespace) {
  if (!/^word\/header\d+\.xml$/.test(path)) return "";
  if (getWordParagraphJustification(paragraph)) return "";

  if (leadingWhitespace.length >= 20 && !trailingWhitespace) {
    return "right";
  }
  if (leadingWhitespace && trailingWhitespace && original.length >= 6) {
    return "center";
  }
  return "";
}

function getWordParagraphJustification(paragraph) {
  const pPr = getWordDirectChild(paragraph, "pPr");
  if (!pPr) return "";
  const jc = getWordDirectChild(pPr, "jc");
  return jc?.getAttributeNS(WORD_NS, "val") || jc?.getAttribute("w:val") || jc?.getAttribute("val") || "";
}

function applyWordParagraphJustification(paragraph, alignment) {
  const pPr = getOrCreateWordDirectChild(paragraph, "pPr", paragraph.firstChild);
  const jc = getOrCreateWordDirectChild(pPr, "jc");
  jc.setAttributeNS(WORD_NS, "w:val", alignment);
  if (pPr.parentNode !== paragraph) {
    paragraph.insertBefore(pPr, paragraph.firstChild);
  }
}

function normalizeWordTextRunForTranslation(textNode) {
  const text = textNode.textContent || "";
  if (!/[A-Za-z]/.test(text)) return;
  const run = getWordAncestor(textNode, "r");
  if (!run) return;
  const rPr = getOrCreateWordDirectChild(run, "rPr", run.firstChild);
  const sz = getWordDirectChild(rPr, "sz");
  const szCs = getWordDirectChild(rPr, "szCs");
  const szValue = sz?.getAttributeNS(WORD_NS, "val") || sz?.getAttribute("w:val") || sz?.getAttribute("val");
  const szCsValue = szCs?.getAttributeNS(WORD_NS, "val") || szCs?.getAttribute("w:val") || szCs?.getAttribute("val");

  if (!szValue && szCsValue) {
    const nextSz = textNode.ownerDocument.createElementNS(WORD_NS, "w:sz");
    nextSz.setAttributeNS(WORD_NS, "w:val", szCsValue);
    rPr.insertBefore(nextSz, szCs || rPr.firstChild);
  }
}

function getWordAncestor(node, localName) {
  let current = node?.parentNode || null;
  while (current) {
    if (current.namespaceURI === WORD_NS && current.localName === localName) return current;
    current = current.parentNode;
  }
  return null;
}

function getWordDirectChild(node, localName) {
  return [...node.childNodes].find((child) => child.namespaceURI === WORD_NS && child.localName === localName) || null;
}

function getOrCreateWordDirectChild(node, localName, beforeNode = null) {
  const existing = getWordDirectChild(node, localName);
  if (existing) return existing;
  const child = node.ownerDocument.createElementNS(WORD_NS, `w:${localName}`);
  node.insertBefore(child, beforeNode || null);
  return child;
}

function updateWordTextSpacePreserve(textNode) {
  const text = textNode.textContent || "";
  if (/^\s|\s$|\s{2,}/.test(text)) {
    textNode.setAttribute("xml:space", "preserve");
  }
}

function hasWordField(paragraph) {
  return [...paragraph.getElementsByTagName("*")].some((node) =>
    ["fldChar", "instrText", "fldSimple"].includes(node.localName)
  );
}

function clearRemainingTextNodes(textNodes) {
  for (let index = 1; index < textNodes.length; index += 1) {
    textNodes[index].textContent = "";
  }
}

function updateTextarea(segment) {
  const index = state.segments.indexOf(segment);
  const textarea = els.segmentTable.querySelector(`textarea[data-index="${index}"]`);
  if (textarea) textarea.value = segment.translation;
}

function updateStats() {
  const translated = state.segments.filter((segment) => segment.translation.trim()).length;
  els.slideCount.textContent = String(state.slideCount);
  els.segmentCount.textContent = String(state.segments.length);
  els.translatedCount.textContent = String(translated);
  updateTranslateButtonState(document.body.classList.contains("is-busy") || state.batchRunning);
  els.previewButton.disabled = !state.segments.length;
  els.downloadButton.disabled = !state.segments.length;
  els.shareButton.disabled = !state.segments.length;
  updateBatchButton();
  if (els.previewDownloadButton) {
    els.previewDownloadButton.disabled = !state.segments.length;
  }
  if (els.previewShareButton) {
    els.previewShareButton.disabled = !state.segments.length;
  }
  setProgress(state.segments.length ? translated / state.segments.length : 0);
}

function setMobileMenuOpen(isOpen) {
  if (!els.mobileViewButton || !els.mobileViewMenu) return;
  els.mobileViewButton.setAttribute("aria-expanded", String(isOpen));
  els.mobileViewMenu.hidden = !isOpen;
}

function setMobileView(view) {
  const nextView = ["translate", "settings", "output", "library"].includes(view) ? view : "translate";
  state.mobileView = nextView;
  if (els.workspace) els.workspace.dataset.mobileView = nextView;
  els.mobileViewTargets.forEach((button) => {
    const isActive = button.dataset.mobileViewTarget === nextView;
    button.classList.toggle("active", isActive);
    button.setAttribute("aria-current", isActive ? "page" : "false");
  });
}

function setBusy(isBusy, message = "") {
  const effectiveBusy = isBusy || state.batchRunning;
  updateTranslateButtonState(effectiveBusy);
  els.previewButton.disabled = effectiveBusy || !state.segments.length;
  els.downloadButton.disabled = effectiveBusy || !state.segments.length;
  els.shareButton.disabled = effectiveBusy || !state.segments.length;
  if (els.batchTranslateButton) {
    els.batchTranslateButton.disabled = effectiveBusy || !state.batchFiles.length;
  }
  if (els.previewDownloadButton) {
    els.previewDownloadButton.disabled = effectiveBusy || !state.segments.length;
  }
  if (els.previewShareButton) {
    els.previewShareButton.disabled = effectiveBusy || !state.segments.length;
  }
  els.fileInput.disabled = effectiveBusy;
  document.body.classList.toggle("is-busy", effectiveBusy);
  if (els.statusVisual) {
    els.statusVisual.classList.toggle("active", effectiveBusy);
  }
  if (effectiveBusy) {
    requestScreenWakeLock();
  } else {
    releaseScreenWakeLock();
  }
  if (message) setStatus(message);
}

function updateTranslateButtonState(effectiveBusy = false) {
  const canStop = state.translationRunning;
  els.translateButton.textContent = canStop ? "停止翻译" : "自动翻译";
  els.translateButton.classList.toggle("danger-action", canStop);
  els.translateButton.disabled = canStop ? false : effectiveBusy || !state.segments.length;
}

async function requestScreenWakeLock() {
  if (!("wakeLock" in navigator)) {
    if (!state.wakeLockWarningShown) {
      state.wakeLockWarningShown = true;
      showToast("当前浏览器不支持自动保持屏幕常亮，请尽量停留在当前页面。", true);
    }
    return false;
  }

  try {
    if (state.wakeLock) return true;
    const lock = await navigator.wakeLock.request("screen");
    state.wakeLock = lock;
    lock.addEventListener("release", () => {
      if (state.wakeLock === lock) state.wakeLock = null;
    });

    if (!state.wakeLockNoticeShown) {
      state.wakeLockNoticeShown = true;
      showToast("处理期间已尝试保持屏幕常亮，请尽量不要切到后台。");
    }
    return true;
  } catch (error) {
    console.warn("Screen wake lock unavailable", error);
    if (!state.wakeLockWarningShown) {
      state.wakeLockWarningShown = true;
      showToast("当前系统未允许屏幕常亮，请避免锁屏或切到后台。", true);
    }
    return false;
  }
}

async function releaseScreenWakeLock() {
  if (!state.wakeLock) return;
  const lock = state.wakeLock;
  state.wakeLock = null;
  try {
    await lock.release();
  } catch (error) {
    console.warn("Screen wake lock release failed", error);
  }
}

function openPreview() {
  renderPreview();
  if (typeof els.previewDialog.showModal === "function") {
    els.previewDialog.showModal();
  } else {
    els.previewDialog.setAttribute("open", "");
  }
}

function closePreview() {
  els.previewDialog.close();
}

function openHelp() {
  if (!els.helpDialog) return;
  if (typeof els.helpDialog.showModal === "function") {
    els.helpDialog.showModal();
  } else {
    els.helpDialog.setAttribute("open", "");
  }
}

function closeHelp() {
  if (!els.helpDialog) return;
  if (typeof els.helpDialog.close === "function") {
    els.helpDialog.close();
  } else {
    els.helpDialog.removeAttribute("open");
  }
}

function renderPreview() {
  const previousScrollTop = els.previewBody.scrollTop;
  const selectedIndex = els.previewBody.querySelector(".slide-text-box.selected")?.dataset.index || "";
  const translated = state.segments.filter((segment) => segment.translation.trim()).length;
  els.previewMeta.textContent = `${getFileTypeName()} · ${state.segments.length} 段文字 · ${translated} 段已有译文`;
  els.previewBody.replaceChildren();

  if (!state.segments.length) {
    const empty = document.createElement("div");
    empty.className = "empty-state";
    empty.textContent = "还没有可预览的内容。";
    els.previewBody.append(empty);
    return;
  }

  const groups = groupSegmentsForPreview();
  groups.forEach((group) => {
    const section = document.createElement("section");
    section.className = "preview-card";

    const title = document.createElement("h3");
    title.textContent = group.label;
    section.append(title);

    if (state.fileType === "pptx") {
      try {
        section.append(createSlidePreview(group.segments));
      } catch (error) {
        console.warn("PPT preview fallback", error);
        group.segments.forEach((segment) => {
          section.append(createPreviewItem(segment));
        });
      }
    } else {
      group.segments.forEach((segment) => {
        section.append(createPreviewItem(segment));
      });
    }

    els.previewBody.append(section);
  });

  if (selectedIndex) {
    const selectedBox = els.previewBody.querySelector(`.slide-text-box[data-index="${selectedIndex}"]`);
    selectedBox?.classList.add("selected");
  }
  els.previewBody.scrollTop = previousScrollTop;
  requestAnimationFrame(() => {
    els.previewBody.scrollTop = previousScrollTop;
  });
}

async function downloadPdfTranslation() {
  setProgress(0.01);
  setStatus("PDF 导出准备中，请稍等...");
  await waitForUiFrame();

  setProgress(0.06);
  setStatus("正在加载 PDF 导出组件和中文字体，首次使用可能需要更久...");
  await waitForUiFrame();
  const { PDFDocument, rgb, fontkit, fontBytes } = await loadPdfExportTools();
  if (!state.pdfBytes) throw new Error("缺少原 PDF 数据，请重新上传 PDF。");

  setProgress(0.18);
  setStatus("正在打开原始 PDF，并保留图片、表格和页面结构...");
  await waitForUiFrame();
  const pdfjs = await loadPdfJs();
  const sourcePdf = await pdfjs.getDocument({ data: state.pdfBytes.slice() }).promise;
  const pdfDoc = await PDFDocument.create();
  pdfDoc.registerFontkit(fontkit);
  pdfDoc.setTitle(`${state.file.name} translated`);
  const pages = [];

  for (let pageNumber = 1; pageNumber <= sourcePdf.numPages; pageNumber += 1) {
    const sourcePage = await sourcePdf.getPage(pageNumber);
    const pageSize = state.pdfPageSizes.get(`pdf/page-${pageNumber}`) || sourcePage.getViewport({ scale: 1 });
    const page = pdfDoc.addPage([pageSize.width || 595.28, pageSize.height || 841.89]);
    const background = await renderPdfPageBackgroundForExport(sourcePage, pdfDoc, pageSize);
    page.drawImage(background.image, {
      x: 0,
      y: 0,
      width: pageSize.width || background.width,
      height: pageSize.height || background.height,
    });
    pages.push(page);

    const ratio = sourcePdf.numPages ? pageNumber / sourcePdf.numPages : 1;
    setProgress(0.18 + ratio * 0.22);
    setStatus(`正在生成高清页面背景：${pageNumber}/${sourcePdf.numPages}`);
    await waitForUiFrame();
  }

  setProgress(0.43);
  setStatus("正在嵌入译文字体...");
  await waitForUiFrame();
  const font = await pdfDoc.embedFont(fontBytes, { subset: true });
  const translatedSegments = state.segments.filter((segment) => segment.type === "pdf" && segment.layout?.bounds);

  for (let index = 0; index < translatedSegments.length; index += 1) {
    const segment = translatedSegments[index];
    const page = pages[Number(segment.slideNumber) - 1];
    if (page) drawPdfOverlayTranslation(page, segment, font, rgb);

    if (index % 12 === 0 || index === translatedSegments.length - 1) {
      const ratio = translatedSegments.length ? (index + 1) / translatedSegments.length : 1;
      setProgress(0.45 + ratio * 0.38);
      setStatus(`正在回写 PDF 译文：${index + 1}/${translatedSegments.length} 段，请勿关闭页面。`);
      await waitForUiFrame();
    }
  }

  setProgress(0.9);
  setStatus("正在保存翻译版 PDF，文件较大时可能需要几十秒...");
  await waitForUiFrame();
  setProgress(0.91);
  setStatus("正在重新绘制 PDF 表格线，修复文字覆盖造成的断线...");
  await waitForUiFrame();
  pages.forEach((page, index) => {
    const cells = state.pdfTableCells.get(`pdf/page-${index + 1}`) || [];
    drawPdfTableCellBorders(page, cells, rgb);
  });

  setProgress(0.94);
  setStatus("Saving translated PDF. On mobile this may take 30-90 seconds; keep this page open...");
  await waitForUiFrame();
  await delay(80);
  const pdfBytes = await pdfDoc.save();
  const blob = new Blob([pdfBytes], { type: "application/pdf" });
  setProgress(1);
  setStatus("PDF 导出完成。");
  return blob;
}

function drawPdfOverlayTranslation(page, segment, font, rgb) {
  const sourceBounds = segment.layout.bounds;
  const cell = segment.layout.tableCell;
  const direction = getDirectionConfig(els.translationDirection?.value || "");
  if (shouldSkipSegmentForDirection(segment, direction)) return;

  const fitBounds = cell
    ? { x: cell.x + 2, y: cell.y + 2, width: Math.max(4, cell.width - 4), height: Math.max(4, cell.height - 4) }
    : {
        ...sourceBounds,
        width: Math.max(sourceBounds.width, Number(segment.layout.availableWidth || 0)),
        height: Math.max(sourceBounds.height, Number(segment.layout.availableHeight || 0)),
      };
  const isTableLike = Boolean(cell) || Number(segment.layout.rowSegmentCount || 1) > 1;
  const sourceFontSize = Math.max(4, Math.min(28, Number(segment.layout.fontSize || 10)));
  const paddingX = isTableLike ? Math.max(1.2, sourceFontSize * 0.12) : Math.max(2, sourceFontSize * 0.16);
  const paddingY = isTableLike ? Math.max(1, sourceFontSize * 0.1) : Math.max(1.5, sourceFontSize * 0.12);
  const eraseWidth = Math.max(2, sourceBounds.width + paddingX * 2);
  const eraseHeight = Math.max(
    sourceBounds.height + paddingY * 2,
    sourceFontSize * (isTableLike ? 1.35 : 1.55)
  );
  const eraseX = Math.max(0, sourceBounds.x - paddingX);
  const eraseY = Math.max(
    0,
    sourceBounds.y - paddingY - Math.max(0, (eraseHeight - sourceBounds.height) / 2)
  );
  const coverColor = segment.layout.backgroundColor || { r: 1, g: 1, b: 1 };
  const textColor = segment.layout.textColor || getReadablePdfTextColor(coverColor);

  page.drawRectangle({
    x: eraseX,
    y: eraseY,
    width: eraseWidth,
    height: eraseHeight,
    color: rgb(coverColor.r, coverColor.g, coverColor.b),
    opacity: 1,
  });

  const text = getPdfExportText(segment);
  if (!text) return;

  const fit = fitPdfTextSize(text, font, fitBounds, segment.layout.fontSize || 10, segment.original, segment.layout);
  const fontSize = fit.fontSize;
  const lines = fit.lines;
  const lineHeight = fit.lineHeight;
  const textHeight = lines.length * lineHeight;
  const fitY = isTableLike ? Math.max(0, sourceBounds.y - Math.max(0, (fitBounds.height - sourceBounds.height) / 2)) : sourceBounds.y;
  let y = fitY + Math.max(0, (fitBounds.height - textHeight) / 2) + (lines.length - 1) * lineHeight;
  lines.forEach((line) => {
    const lineWidth = font.widthOfTextAtSize(line, fontSize);
    const x = segment.layout.cellAlign === "center" ? fitBounds.x + Math.max(0, (fitBounds.width - lineWidth) / 2) : fitBounds.x;
    page.drawText(line, {
      x,
      y,
      size: fontSize,
      font,
      color: rgb(textColor.r, textColor.g, textColor.b),
    });
    y -= lineHeight;
  });
}

function getPdfExportText(segment) {
  const text = applyTerminologyRules(String(segment.translation || "").trim(), segment.original);
  if (!text) return "";

  const direction = els.translationDirection?.value || "";
  if (getDirectionConfig(direction).targetCode === "en" && containsHanCharacters(text)) {
    return "";
  }

  return text;
}

function shouldSkipSegmentForDirection(segment, direction) {
  return direction.sourceLanguage === "Chinese" && !containsHanCharacters(segment.original || "");
}

function containsHanCharacters(text) {
  return /[\u3400-\u9fff]/.test(text);
}

function drawPdfTableCellBorders(page, cells, rgb) {
  if (!cells.length) return;

  const lines = new Map();
  cells.forEach((cell) => {
    addPdfBorderLine(lines, cell.x, cell.y, cell.x2, cell.y);
    addPdfBorderLine(lines, cell.x, cell.y2, cell.x2, cell.y2);
    addPdfBorderLine(lines, cell.x, cell.y, cell.x, cell.y2);
    addPdfBorderLine(lines, cell.x2, cell.y, cell.x2, cell.y2);
  });

  lines.forEach((line) => {
    page.drawLine({
      start: { x: line.x1, y: line.y1 },
      end: { x: line.x2, y: line.y2 },
      thickness: 0.45,
      color: rgb(0, 0, 0),
    });
  });
}

function addPdfBorderLine(lines, x1, y1, x2, y2) {
  const rounded = [x1, y1, x2, y2].map((value) => Math.round(value * 2) / 2);
  lines.set(rounded.join(":"), {
    x1: rounded[0],
    y1: rounded[1],
    x2: rounded[2],
    y2: rounded[3],
  });
}

function fitPdfTextSize(text, font, bounds, sourceSize, sourceText = "", layout = null) {
  const sourceFontSize = Math.max(4, Math.min(24, Number(sourceSize || 10) * 0.96));
  const isTableLike = Number(layout?.rowSegmentCount || 1) > 1;
  const preferred = sourceFontSize;
  const targetWidth = Math.max(4, bounds.width);
  const maxHeight = Math.max(5, bounds.height * (isTableLike ? 1.12 : 1.55));

  if (!isTableLike) {
    const lineHeight = getPdfLineHeight(preferred, false);
    const lines = wrapPdfText(text, font, preferred, targetWidth);
    return { fontSize: preferred, lines, lineHeight };
  }

  const minSize = Math.max(4.8, preferred * 0.72);

  const textWidthAtPreferred = Math.max(0.1, font.widthOfTextAtSize(text, preferred));
  const estimatedSize = Math.min(preferred, (preferred * targetWidth) / textWidthAtPreferred);

  for (let size = preferred; size >= minSize; size -= 0.25) {
    const lineHeight = getPdfLineHeight(size, true);
    const lines = wrapPdfText(text, font, size, targetWidth);
    if (lines.length * lineHeight <= maxHeight) {
      return { fontSize: size, lines, lineHeight };
    }
  }

  for (let size = estimatedSize; size >= minSize; size -= 0.25) {
    const lineHeight = getPdfLineHeight(size, true);
    const lines = wrapPdfText(text, font, size, targetWidth);
    if (lines.length * lineHeight <= maxHeight) {
      return { fontSize: size, lines, lineHeight };
    }
  }

  for (let size = minSize; size >= 3.6; size -= 0.2) {
    const lineHeight = getPdfLineHeight(size, true);
    const lines = wrapPdfText(text, font, size, targetWidth);
    if (lines.length * lineHeight <= maxHeight) {
      return { fontSize: size, lines, lineHeight };
    }
  }

  return { fontSize: minSize, lines: [text], lineHeight: getPdfLineHeight(minSize, true) };
}

function getPdfLineHeight(size, isTableLike) {
  return size * (isTableLike ? 1.18 : 1.28);
}

async function loadPdfExportTools() {
  if (!window.pdfLibTools) {
    const [pdfLib, fontkitModule, fontResponse] = await Promise.all([
      import(PDFLIB_URL),
      import(FONTKIT_URL),
      fetch(PDF_FONT_URL),
    ]);

    if (!fontResponse.ok) {
      throw new Error(`PDF 字体加载失败：${fontResponse.status}`);
    }

    window.pdfLibTools = {
      PDFDocument: pdfLib.PDFDocument,
      rgb: pdfLib.rgb,
      fontkit: fontkitModule.default || fontkitModule,
      fontBytes: await fontResponse.arrayBuffer(),
    };
  }

  return window.pdfLibTools;
}

function wrapPdfText(text, font, fontSize, maxWidth) {
  const paragraphs = String(text || "").split(/\r\n|\r|\n/);
  const lines = [];

  paragraphs.forEach((paragraph) => {
    const tokens = /[\u3400-\u9fff]/.test(paragraph) ? [...paragraph] : paragraph.split(/(\s+)/).filter(Boolean);
    let current = "";

    tokens.forEach((token) => {
      const candidate = current ? `${current}${token}` : token;
      if (font.widthOfTextAtSize(candidate, fontSize) <= maxWidth) {
        current = candidate;
        return;
      }

      if (current.trim()) lines.push(current.trim());
      current = token.trim();

      while (font.widthOfTextAtSize(current, fontSize) > maxWidth && [...current].length > 1) {
        let part = "";
        for (const char of current) {
          if (font.widthOfTextAtSize(part + char, fontSize) > maxWidth) break;
          part += char;
        }
        lines.push(part);
        current = current.slice(part.length);
      }
    });

    if (current.trim()) lines.push(current.trim());
    if (!paragraph.trim()) lines.push("");
  });

  return lines.length ? lines : [""];
}

function createPreviewItem(segment) {
  const item = document.createElement("article");
  item.className = `preview-item${segment.translation.trim() ? "" : " pending"}`;

  const text = document.createElement("p");
  text.textContent = segment.translation.trim() || segment.original;

  const meta = document.createElement("span");
  meta.textContent = segment.translation.trim() ? "译文" : "未翻译，暂用原文";

  item.append(text, meta);
  return item;
}

function createSlidePreview(segments) {
  const slide = document.createElement("div");
  slide.className = "slide-preview";
  slide.style.aspectRatio = `${state.slideSize.cx} / ${state.slideSize.cy}`;
  const slidePath = segments[0]?.path || "";

  (state.slideVisuals.get(slidePath) || []).forEach((visual, index) => {
    const box = document.createElement("div");
    box.className = "slide-image-box";
    box.style.left = `${toSlidePercentX(visual.bounds.x)}%`;
    box.style.top = `${toSlidePercentY(visual.bounds.y)}%`;
    box.style.width = `${toSlidePercentX(visual.bounds.cx)}%`;
    box.style.height = `${toSlidePercentY(visual.bounds.cy)}%`;

    const label = document.createElement("span");
    label.textContent = visual.name || `图片 ${index + 1}`;
    box.append(label);
    slide.append(box);
  });

  segments.forEach((segment) => {
    if (!segment.layout?.bounds) return;

    const index = state.segments.indexOf(segment);
    const box = document.createElement("div");
    box.className = `slide-text-box${segment.translation.trim() ? "" : " pending"}`;
    box.dataset.index = String(index);
    box.tabIndex = 0;

    const { x, y, cx, cy } = getSegmentBounds(segment);
    if (y / state.slideSize.cy < 0.22) {
      box.classList.add("tools-below");
    }
    box.style.left = `${(x / state.slideSize.cx) * 100}%`;
    box.style.top = `${(y / state.slideSize.cy) * 100}%`;
    box.style.width = `${(cx / state.slideSize.cx) * 100}%`;
    box.style.height = `${(cy / state.slideSize.cy) * 100}%`;
    box.style.fontSize = `${getPreviewFontCqw(segment)}cqw`;
    box.style.whiteSpace = shouldUseSingleLine(segment) ? "nowrap" : "normal";
    applyPreviewTextStyle(box, segment);

    const text = document.createElement("span");
    text.className = "slide-box-text";
    text.textContent = segment.translation.trim() || segment.original;

    const tools = createPreviewBoxTools(segment, index);
    const resizeHandle = document.createElement("button");
    resizeHandle.type = "button";
    resizeHandle.className = "slide-resize-handle";
    resizeHandle.title = "拖动调整文本框大小";
    resizeHandle.setAttribute("aria-label", "拖动调整文本框大小");
    attachPreviewResize(resizeHandle, box, segment);

    box.addEventListener("pointerdown", () => selectPreviewBox(box));
    box.addEventListener("focus", () => selectPreviewBox(box));
    attachPreviewMove(box, segment);
    box.append(text, tools, resizeHandle);
    slide.append(box);
  });

  if (!slide.children.length) {
    const fallback = document.createElement("div");
    fallback.className = "preview-fallback";
    fallback.textContent = "这一页没有可定位的文本框，显示文字清单。";
    slide.append(fallback);
    segments.forEach((segment) => slide.append(createPreviewItem(segment)));
  }

  return slide;
}

function createPreviewBoxTools(segment, index) {
  const tools = document.createElement("div");
  tools.className = "slide-box-tools";

  const scaleInput = document.createElement("input");
  scaleInput.type = "range";
  scaleInput.min = "60";
  scaleInput.max = "115";
  scaleInput.step = "5";
  scaleInput.value = segment.overrides.fontScale || els.fontScale.value || "100";
  scaleInput.title = "本段字号";

  const scaleValue = document.createElement("output");
  scaleValue.textContent = getPreviewScaleLabel(segment);

  scaleInput.addEventListener("input", () => {
    const current = state.segments[index];
    current.overrides.fontScale = scaleInput.value;
    scaleValue.textContent = getPreviewScaleLabel(current);
    rerenderPreviewIfOpen();
    scheduleCurrentDraftSave();
  });

  const wrapButton = document.createElement("button");
  wrapButton.type = "button";
  wrapButton.textContent = shouldUseSingleLine(segment) ? "换行" : "单行";
  wrapButton.title = shouldUseSingleLine(segment) ? "允许自动换行" : "强制本段单行";
  wrapButton.addEventListener("click", () => {
    const current = state.segments[index];
    const nextSingleLine = !shouldUseSingleLine(current);
    current.overrides.singleLine = nextSingleLine;
    current.overrides.wrapMode = nextSingleLine ? "single" : "wrap";
    rerenderPreviewIfOpen();
    scheduleCurrentDraftSave();
  });

  const resetButton = document.createElement("button");
  resetButton.type = "button";
  resetButton.textContent = "重置";
  resetButton.title = "重置本段预览调整";
  resetButton.addEventListener("click", () => {
    state.segments[index].overrides = createSegmentOverrides();
    renderSegments();
    rerenderPreviewIfOpen();
    scheduleCurrentDraftSave();
  });

  const exportButton = document.createElement("button");
  exportButton.type = "button";
  exportButton.textContent = "导出";
  exportButton.title = "导出翻译文件";
  exportButton.addEventListener("click", downloadPresentation);

  const shareButton = document.createElement("button");
  shareButton.type = "button";
  shareButton.textContent = "分享";
  shareButton.title = "分享翻译文件";
  shareButton.addEventListener("click", sharePresentation);

  tools.append(scaleInput, scaleValue, wrapButton, resetButton, exportButton, shareButton);
  return tools;
}

function selectPreviewBox(box) {
  box.closest(".slide-preview")
    ?.querySelectorAll(".slide-text-box.selected")
    .forEach((item) => {
      if (item !== box) item.classList.remove("selected");
    });
  box.classList.add("selected");
}

function attachPreviewMove(box, segment) {
  box.addEventListener("pointerdown", (event) => {
    if (event.target.closest(".slide-box-tools") || event.target.closest(".slide-resize-handle")) return;

    event.preventDefault();
    selectPreviewBox(box);

    const slide = box.closest(".slide-preview");
    const slideRect = slide.getBoundingClientRect();
    const startX = event.clientX;
    const startY = event.clientY;
    const startLeft = box.offsetLeft;
    const startTop = box.offsetTop;
    const maxLeft = Math.max(0, slide.clientWidth - box.offsetWidth);
    const maxTop = Math.max(0, slide.clientHeight - box.offsetHeight);
    let didMove = false;

    box.setPointerCapture(event.pointerId);
    box.classList.add("moving");

    const onMove = (moveEvent) => {
      const left = Math.max(0, Math.min(maxLeft, startLeft + moveEvent.clientX - startX));
      const top = Math.max(0, Math.min(maxTop, startTop + moveEvent.clientY - startY));
      didMove = didMove || Math.abs(left - startLeft) > 1 || Math.abs(top - startTop) > 1;
      box.style.left = `${(left / slideRect.width) * 100}%`;
      box.style.top = `${(top / slideRect.height) * 100}%`;
      updateSegmentBoundsFromPreview(segment, box, slideRect);
    };

    const onEnd = () => {
      box.removeEventListener("pointermove", onMove);
      box.removeEventListener("pointerup", onEnd);
      box.removeEventListener("pointercancel", onEnd);
      box.classList.remove("moving");
      if (didMove) {
        renderSegments();
        scheduleCurrentDraftSave();
      }
    };

    box.addEventListener("pointermove", onMove);
    box.addEventListener("pointerup", onEnd);
    box.addEventListener("pointercancel", onEnd);
  });
}

function attachPreviewResize(handle, box, segment) {
  handle.addEventListener("pointerdown", (event) => {
    event.preventDefault();
    event.stopPropagation();
    selectPreviewBox(box);

    const slide = box.closest(".slide-preview");
    const slideRect = slide.getBoundingClientRect();
    const boxRect = box.getBoundingClientRect();
    const startX = event.clientX;
    const startY = event.clientY;
    const startWidth = boxRect.width;
    const startHeight = boxRect.height;
    const minWidth = Math.max(36, slideRect.width * 0.035);
    const minHeight = Math.max(20, slideRect.height * 0.035);
    const maxWidth = slideRect.right - boxRect.left;
    const maxHeight = slideRect.bottom - boxRect.top;

    handle.setPointerCapture(event.pointerId);

    const onMove = (moveEvent) => {
      const width = Math.max(minWidth, Math.min(maxWidth, startWidth + moveEvent.clientX - startX));
      const height = Math.max(minHeight, Math.min(maxHeight, startHeight + moveEvent.clientY - startY));
      box.style.width = `${(width / slideRect.width) * 100}%`;
      box.style.height = `${(height / slideRect.height) * 100}%`;
      updateSegmentBoundsFromPreview(segment, box, slideRect);
    };

    const onEnd = () => {
      handle.removeEventListener("pointermove", onMove);
      handle.removeEventListener("pointerup", onEnd);
      handle.removeEventListener("pointercancel", onEnd);
      renderSegments();
      scheduleCurrentDraftSave();
    };

    handle.addEventListener("pointermove", onMove);
    handle.addEventListener("pointerup", onEnd);
    handle.addEventListener("pointercancel", onEnd);
  });
}

function groupSegmentsForPreview() {
  const map = new Map();
  state.segments.forEach((segment) => {
    const key = `${segment.path}:${segment.locationLabel || segment.slideNumber}`;
    if (!map.has(key)) {
      map.set(key, {
        label: state.fileType === "pptx" ? `第 ${segment.locationLabel} 页` : segment.locationLabel,
        segments: [],
      });
    }
    map.get(key).segments.push(segment);
  });
  return [...map.values()];
}

function setStatus(message) {
  els.statusText.textContent = message;
}

function setProgress(value) {
  const percent = Math.max(0, Math.min(100, Math.round(value * 100)));
  els.progressFill.style.width = `${percent}%`;
  els.progressFill.style.setProperty("--progress-percent", `${percent}%`);
  els.statusVisual?.style.setProperty("--progress-deg", `${Math.round(percent * 3.6)}deg`);
}

function waitForUiFrame() {
  return new Promise((resolve) => {
    requestAnimationFrame(() => setTimeout(resolve, 0));
  });
}

function delay(ms) {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}

function scheduleCurrentDraftSave(options = {}) {
  if (isRestoringDraft || state.batchRunning || !state.file || !state.segments.length) return;
  window.clearTimeout(draftSaveTimer);

  if (options.immediate) {
    draftSaveTimer = 0;
    saveCurrentDraft().catch((error) => console.warn("Draft save failed", error));
    return;
  }

  draftSaveTimer = window.setTimeout(() => {
    draftSaveTimer = 0;
    saveCurrentDraft().catch((error) => console.warn("Draft save failed", error));
  }, DRAFT_SAVE_DELAY);
}

function flushCurrentDraftSave() {
  if (!draftSaveTimer && (!state.file || !state.segments.length)) return;
  window.clearTimeout(draftSaveTimer);
  draftSaveTimer = 0;
  saveCurrentDraft().catch((error) => console.warn("Draft flush failed", error));
}

async function saveCurrentDraft() {
  if (isRestoringDraft || state.batchRunning || !state.file || !state.segments.length) return;

  const db = await openCurrentDraftDb();
  await idbRequest(db, CURRENT_DRAFT_STORE, "readwrite", (store) =>
    store.put({
      id: CURRENT_DRAFT_ID,
      version: 1,
      savedAt: new Date().toISOString(),
      fileName: state.file.name,
      fileType: state.fileType,
      fileSize: state.file.size,
      fileLastModified: state.file.lastModified || 0,
      fileBlob: state.file,
      slideSize: state.slideSize,
      mobileView: state.mobileView,
      segments: state.segments.map((segment) => ({
        id: segment.id,
        type: segment.type,
        path: segment.path,
        paragraphIndex: segment.paragraphIndex,
        original: segment.original,
        wordLeadingWhitespace: segment.wordLeadingWhitespace || "",
        wordTrailingWhitespace: segment.wordTrailingWhitespace || "",
        translation: segment.translation || "",
        overrides: sanitizeSegmentOverrides(segment.overrides),
      })),
    })
  );
}

async function restoreCurrentDraft() {
  if (state.file || state.segments.length) return;

  const db = await openCurrentDraftDb();
  const draft = await idbRequest(db, CURRENT_DRAFT_STORE, "readonly", (store) => store.get(CURRENT_DRAFT_ID));
  if (!draft?.fileBlob || !draft.fileName) return;

  isRestoringDraft = true;
  try {
    setBusy(true, "正在恢复上次编辑内容...");
    const file = new File([draft.fileBlob], draft.fileName, {
      type: draft.fileBlob.type || getMimeTypeForFileType(draft.fileType),
      lastModified: draft.fileLastModified || Date.now(),
    });

    await loadOfficeFile(file, { silent: true, skipDraftSave: true });
    if (!state.file || !state.segments.length) {
      throw new Error("Draft source file could not be restored");
    }
    applyDraftSegments(draft);
    renderSegments();
    updateStats();
    if (draft.mobileView) setMobileView(draft.mobileView);
    setStatus(`已恢复 ${draft.fileName} 的上次编辑内容。`);
    showToast("已恢复上次未导出的翻译内容。");
  } catch (error) {
    console.warn("Draft restore failed", error);
    setStatus("上次编辑内容恢复失败，请重新选择文件。");
  } finally {
    isRestoringDraft = false;
    setBusy(false);
  }
}

function applyDraftSegments(draft) {
  const savedSegments = new Map((draft.segments || []).map((segment) => [segment.id, segment]));

  state.segments.forEach((segment) => {
    const saved = savedSegments.get(segment.id);
    if (!saved || saved.original !== segment.original) return;
    segment.translation = saved.translation || "";
    segment.overrides = {
      ...createSegmentOverrides(),
      ...sanitizeSegmentOverrides(saved.overrides),
    };
  });
}

function sanitizeSegmentOverrides(overrides = {}) {
  const clean = createSegmentOverrides();
  if (!overrides || typeof overrides !== "object") return clean;

  clean.fontScale = overrides.fontScale ? String(overrides.fontScale) : "";
  clean.singleLine = Boolean(overrides.singleLine);
  clean.wrapMode = ["single", "wrap"].includes(overrides.wrapMode) ? overrides.wrapMode : "";

  if (overrides.bounds && ["x", "y", "cx", "cy"].every((key) => Number.isFinite(Number(overrides.bounds[key])))) {
    clean.bounds = {
      x: Number(overrides.bounds.x),
      y: Number(overrides.bounds.y),
      cx: Number(overrides.bounds.cx),
      cy: Number(overrides.bounds.cy),
    };
  }

  return clean;
}

async function clearCurrentDraft() {
  window.clearTimeout(draftSaveTimer);
  draftSaveTimer = 0;

  try {
    const db = await openCurrentDraftDb();
    await idbRequest(db, CURRENT_DRAFT_STORE, "readwrite", (store) => store.delete(CURRENT_DRAFT_ID));
  } catch (error) {
    console.warn("Draft clear failed", error);
  }
}

function getMimeTypeForFileType(fileType) {
  if (fileType === "docx") return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  if (fileType === "pdf") return "application/pdf";
  return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
}

function openCurrentDraftDb() {
  return new Promise((resolve, reject) => {
    if (!("indexedDB" in window)) {
      reject(new Error("IndexedDB unavailable"));
      return;
    }
    const request = indexedDB.open(CURRENT_DRAFT_DB, 1);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(CURRENT_DRAFT_STORE)) {
        db.createObjectStore(CURRENT_DRAFT_STORE, { keyPath: "id" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

function resetApp(clearInput = true, options = {}) {
  state.file = null;
  state.fileType = "";
  state.zip = null;
  state.segments = [];
  state.slideCount = 0;
  state.slideVisuals = new Map();
  state.pdfPageSizes = new Map();
  state.pdfTableCells = new Map();
  state.pdfBytes = null;
  if (clearInput) els.fileInput.value = "";
  els.fileMeta.textContent = "支持 PPTX / DOCX / PDF；旧版 PPT/DOC 请先另存";
  renderSegments();
  updateStats();
  setStatus("请先选择一个 PPTX、DOCX 或 PDF 文件。");
  if (options.clearDraft !== false) clearCurrentDraft();
}

function buildOutputName(name) {
  if (state.fileType === "pdf") return name.replace(/\.pdf$/i, "") + "-translated.pdf";
  return name.replace(/\.(pptx|docx)$/i, "") + `-translated.${state.fileType || "pptx"}`;
}

function normalizeTranslation(translation, source) {
  const clean = translation.trim();
  if (looksLikeNonTranslation(clean, source)) return "";
  const normalized = /[\r\n]/.test(source)
    ? clean
    : clean.replace(/\s*[\r\n]+\s*/g, " ").replace(/[ \t]{2,}/g, " ");
  return applyTerminologyRules(normalized, source);
}

function applyTerminologyRules(text, source = "") {
  if (!text) return text;

  let normalized = text
    .replace(/\bCura\s*way\b/gi, "CuraWay")
    .replace(/\bCuraway\b/g, "CuraWay")
    .replace(/\bCURAWAY\b/g, "CuraWay")
    .replace(/\bGanavi\b/gi, "CuraWay")
    .replace(/\bGanawei\b/gi, "CuraWay")
    .replace(/\bGanaiwei\b/gi, "CuraWay")
    .replace(/\bJianaiwei\b/gi, "CuraWay")
    .replace(/\bJia\s*Nai\s*Wei\b/gi, "CuraWay");

  if (/伽奈维/.test(source)) {
    normalized = normalized
      .replace(/\bZhejiang\s+(?:CuraWay|Ganavi|Ganawei|Ganaiwei|Jianaiwei|Jia\s*Nai\s*Wei)\s+Medical\s+Technology\s+Co\.?,?\s*Ltd\.?/gi, "Zhejiang CuraWay Medical Technology Co., Ltd.")
      .replace(/\b(?:Ganavi|Ganawei|Ganaiwei|Jianaiwei|Jia\s*Nai\s*Wei)\s+Medical\s+Technology\s+Co\.?,?\s*Ltd\.?/gi, "CuraWay Medical Technology Co., Ltd.");
  }

  return normalized;
}

function looksLikeNonTranslation(text, source) {
  if (!text) return false;
  const normalized = text.toLowerCase().replace(/\s+/g, " ");
  const refusalPatterns = [
    /^i['’]?m sorry\b/,
    /\bi cannot\b/,
    /\bunable to translate\b/,
    /\bno source text\b/,
    /\bplease provide\b.*\b(text|source|chinese)\b/,
    /\bas an ai\b/,
    /抱歉/,
    /无法.*翻译/,
    /请提供.*文本/,
  ];

  if (refusalPatterns.some((pattern) => pattern.test(normalized))) return true;
  return [...source].length <= 12 && [...text].length > 80 && /\b(sorry|provide|source|translate)\b/.test(normalized);
}

function normalizePresentationRunStyle(textNode, segment) {
  const run = textNode.parentElement;
  if (!run) return;

  const runProperties = ensureChild(run, "rPr");
  runProperties.setAttribute("sz", String(getPresentationExportFontSize(segment)));

  setTypeface(runProperties, "latin", "Arial");
  setTypeface(runProperties, "ea", "Microsoft YaHei");
  setTypeface(runProperties, "cs", "Arial");

  [...runProperties.getElementsByTagNameNS(DRAWING_NS, "sym")].forEach((node) => node.remove());
}

function applyPresentationBounds(paragraph, segment) {
  const bounds = segment.overrides?.bounds;
  if (!bounds) return;

  const transform = getParagraphTransform(paragraph);
  if (!transform?.offset || !transform?.extent) return;

  transform.offset.setAttribute("x", String(Math.max(0, Math.round(bounds.x))));
  transform.offset.setAttribute("y", String(Math.max(0, Math.round(bounds.y))));
  transform.extent.setAttribute("cx", String(Math.max(1, Math.round(bounds.cx))));
  transform.extent.setAttribute("cy", String(Math.max(1, Math.round(bounds.cy))));
}

function applyPresentationLayout(paragraph, segment) {
  const textBody = paragraph.parentElement;
  if (!textBody) return;

  const bodyProperties = [...textBody.children].find((node) => node.localName === "bodyPr");
  if (!bodyProperties) return;

  const layout = segment.layout || {};
  const shouldForceSingleLine = shouldUseSingleLine(segment);

  [...bodyProperties.children]
    .filter((node) => ["noAutofit", "spAutoFit", "normAutofit"].includes(node.localName))
    .forEach((node) => node.remove());

  if (shouldForceSingleLine) {
    bodyProperties.setAttribute("wrap", "none");
    const autoFit = document.createElementNS(DRAWING_NS, "a:normAutofit");
    autoFit.setAttribute("fontScale", "85000");
    autoFit.setAttribute("lnSpcReduction", "12000");
    bodyProperties.append(autoFit);
    return;
  } else if (layout.wrap && layout.wrap !== "none") {
    bodyProperties.setAttribute("wrap", layout.wrap);
  } else {
    bodyProperties.removeAttribute("wrap");
  }

  if (getPptLayoutMode() === "compact-fit") {
    const autoFit = document.createElementNS(DRAWING_NS, "a:normAutofit");
    autoFit.setAttribute("fontScale", "88000");
    autoFit.setAttribute("lnSpcReduction", "12000");
    bodyProperties.append(autoFit);
    return;
  }

  if (layout.autofit === "normAutofit") {
    const autoFit = document.createElementNS(DRAWING_NS, "a:normAutofit");
    Object.entries(layout.autofitAttrs || {}).forEach(([key, value]) => autoFit.setAttribute(key, value));
    bodyProperties.append(autoFit);
  } else if (layout.autofit === "spAutoFit") {
    bodyProperties.append(document.createElementNS(DRAWING_NS, "a:spAutoFit"));
  } else {
    const autoFit = document.createElementNS(DRAWING_NS, "a:normAutofit");
    autoFit.setAttribute("fontScale", "85000");
    autoFit.setAttribute("lnSpcReduction", "12000");
    bodyProperties.append(autoFit);
  }
}

function ensureChild(parent, localName) {
  const existing = [...parent.children].find((node) => node.localName === localName);
  if (existing) return existing;

  const child = document.createElementNS(DRAWING_NS, `a:${localName}`);
  if (localName === "rPr") {
    parent.prepend(child);
  } else {
    parent.append(child);
  }
  return child;
}

function ensureChildNS(parent, namespace, qualifiedName, localName) {
  const existing = [...parent.children].find((node) => node.localName === localName);
  if (existing) return existing;

  const child = document.createElementNS(namespace, qualifiedName);
  if (localName === "rPr") {
    parent.prepend(child);
  } else {
    parent.append(child);
  }
  return child;
}

function setTypeface(parent, localName, typeface) {
  const node = ensureChild(parent, localName);
  node.setAttribute("typeface", typeface);
}

function getLengthScale(source, translation, segment) {
  const sourceLength = Math.max(1, [...source].length);
  const translationLength = Math.max(1, [...translation].length);
  const ratio = translationLength / sourceLength;
  const userScale = getActiveFontScale(segment);

  if (ratio <= 1.25) return userScale;
  if (ratio <= 1.75) return Math.min(userScale, 0.94);
  if (ratio <= 2.4) return Math.min(userScale, 0.9);
  return Math.min(userScale, 0.88);
}

function getPresentationLengthScale(segment) {
  if (getPptLayoutMode() === "compact-fit" || shouldUseSingleLine(segment)) {
    return getLengthScale(segment.original, segment.translation, segment);
  }

  return getActiveFontScale(segment);
}

function getPptLayoutMode() {
  return els.pptLayoutMode?.value || "smart";
}

function getActiveFontScale(segment) {
  const override = segment?.overrides?.fontScale;
  if (override) {
    return Math.max(0.6, Math.min(1.15, Number(override) / 100));
  }

  return getUserFontScale();
}

async function readSlideSize() {
  const file = state.zip.file("ppt/presentation.xml");
  if (!file) return { cx: 12192000, cy: 6858000 };

  const xmlText = await file.async("text");
  const doc = parser.parseFromString(xmlText, "application/xml");
  const size = doc.getElementsByTagName("p:sldSz")[0] || [...doc.getElementsByTagName("*")].find((node) => node.localName === "sldSz");

  return {
    cx: Number(size?.getAttribute("cx")) || 12192000,
    cy: Number(size?.getAttribute("cy")) || 6858000,
  };
}

function inspectPresentationLayout(paragraph, original) {
  const textBody = paragraph.parentElement;
  const bodyProperties = [...(textBody?.children || [])].find((node) => node.localName === "bodyPr");
  const autofit = [...(bodyProperties?.children || [])].find((node) =>
    ["noAutofit", "spAutoFit", "normAutofit"].includes(node.localName)
  );
  const lines = original.split(/\r\n|\r|\n/).filter((line) => line.trim()).length || 1;

  return {
    wrap: bodyProperties?.getAttribute("wrap") || "",
    autofit: autofit?.localName || "",
    autofitAttrs: autofit ? Object.fromEntries([...autofit.attributes].map((attr) => [attr.name, attr.value])) : {},
    hasManualBreaks: lines > 1,
    textLength: [...original].length,
    textNodeCount: [...paragraph.getElementsByTagNameNS(DRAWING_NS, "t")].length,
    textBodyParagraphCount: countTextBodyParagraphs(textBody),
    bounds: getTextBounds(paragraph),
    fontSize: getParagraphFontSize(paragraph),
    insets: getTextBodyInsets(bodyProperties),
    runStyle: getParagraphRunStyle(paragraph),
    paragraphStyle: getParagraphStyle(paragraph),
  };
}

function inspectSlideVisuals(doc) {
  return [...doc.getElementsByTagName("*")]
    .filter((node) => node.localName === "pic")
    .map((picture, index) => {
      const shapeProperties = [...picture.children].find((node) => node.localName === "spPr");
      const transform = [...(shapeProperties?.children || [])].find((node) => node.localName === "xfrm");
      const offset = [...(transform?.children || [])].find((node) => node.localName === "off");
      const extent = [...(transform?.children || [])].find((node) => node.localName === "ext");
      if (!offset || !extent) return null;

      const nameNode = [...picture.getElementsByTagName("*")].find((node) => node.localName === "cNvPr");
      const bounds = {
        x: Number(offset.getAttribute("x")) || 0,
        y: Number(offset.getAttribute("y")) || 0,
        cx: Number(extent.getAttribute("cx")) || 0,
        cy: Number(extent.getAttribute("cy")) || 0,
      };

      if (bounds.cx <= 0 || bounds.cy <= 0) return null;

      return {
        type: "image",
        name: nameNode?.getAttribute("name") || `图片 ${index + 1}`,
        bounds,
      };
    })
    .filter(Boolean);
}

function shouldUseSingleLine(segment) {
  if (segment?.overrides?.wrapMode === "single") return true;
  if (segment?.overrides?.wrapMode === "wrap") return false;
  if (segment?.overrides?.singleLine) return true;
  const mode = getPptLayoutMode();
  if (mode === "compact-fit") return false;
  if (mode === "keep-size") return true;
  return false;
}

function createSegmentControls(segment, index) {
  const controls = document.createElement("div");
  controls.className = "segment-controls";

  const scaleLabel = document.createElement("label");
  scaleLabel.className = "segment-range";

  const scaleText = document.createElement("span");
  scaleText.textContent = "本段字号";

  const scaleInput = document.createElement("input");
  scaleInput.type = "range";
  scaleInput.min = "60";
  scaleInput.max = "115";
  scaleInput.step = "5";
  scaleInput.value = segment.overrides.fontScale || els.fontScale.value || "100";
  scaleInput.dataset.index = String(index);

  const scaleValue = document.createElement("output");
  scaleValue.textContent = segment.overrides.fontScale ? `${segment.overrides.fontScale}%` : "跟随全局";

  scaleInput.addEventListener("input", () => {
    const current = state.segments[Number(scaleInput.dataset.index)];
    current.overrides.fontScale = scaleInput.value;
    scaleValue.textContent = `${scaleInput.value}%`;
    rerenderPreviewIfOpen();
    scheduleCurrentDraftSave();
  });

  scaleLabel.append(scaleText, scaleInput, scaleValue);

  const singleLineLabel = document.createElement("label");
  singleLineLabel.className = "segment-check";

  const singleLineInput = document.createElement("input");
  singleLineInput.type = "checkbox";
  singleLineInput.checked = segment.overrides.wrapMode === "single" || Boolean(segment.overrides.singleLine);
  singleLineInput.dataset.index = String(index);
  singleLineInput.addEventListener("change", () => {
    const current = state.segments[Number(singleLineInput.dataset.index)];
    current.overrides.singleLine = singleLineInput.checked;
    current.overrides.wrapMode = singleLineInput.checked ? "single" : "";
    rerenderPreviewIfOpen();
    scheduleCurrentDraftSave();
  });

  const singleLineText = document.createElement("span");
  singleLineText.textContent = "本段单行";
  singleLineLabel.append(singleLineInput, singleLineText);

  const resetButton = document.createElement("button");
  resetButton.type = "button";
  resetButton.className = "segment-reset";
  resetButton.textContent = "重置";
  resetButton.dataset.index = String(index);
  resetButton.addEventListener("click", () => {
    const current = state.segments[Number(resetButton.dataset.index)];
    current.overrides = createSegmentOverrides();
    renderSegments();
    updateStats();
    rerenderPreviewIfOpen();
    scheduleCurrentDraftSave();
  });

  const pageButton = document.createElement("button");
  pageButton.type = "button";
  pageButton.className = "segment-reset";
  pageButton.textContent = "本页";
  pageButton.title = "把本段设置应用到同一页";
  pageButton.dataset.index = String(index);
  pageButton.addEventListener("click", () => {
    const current = state.segments[Number(pageButton.dataset.index)];
    state.segments.forEach((item) => {
      if (item.path !== current.path || item.locationLabel !== current.locationLabel) return;
      item.overrides = {
        fontScale: current.overrides.fontScale,
        singleLine: current.type === "pptx" ? current.overrides.singleLine : false,
        wrapMode: current.type === "pptx" ? current.overrides.wrapMode : "",
        bounds: null,
      };
    });
    renderSegments();
    updateStats();
    rerenderPreviewIfOpen();
    scheduleCurrentDraftSave();
  });

  controls.append(scaleLabel);
  if (segment.type === "pptx") {
    controls.append(singleLineLabel);
  }
  controls.append(pageButton, resetButton);
  return controls;
}

function createSegmentOverrides() {
  return {
    fontScale: "",
    singleLine: false,
    wrapMode: "",
    bounds: null,
  };
}

function countTextBodyParagraphs(textBody) {
  if (!textBody) return 1;
  return [...textBody.getElementsByTagNameNS(DRAWING_NS, "p")].filter((paragraph) => {
    const text = [...paragraph.getElementsByTagNameNS(DRAWING_NS, "t")]
      .map((node) => node.textContent || "")
      .join("")
      .trim();
    return Boolean(text);
  }).length;
}

function getFileTypeName() {
  if (state.fileType === "docx") return "DOCX";
  if (state.fileType === "pdf") return "PDF";
  return "PPTX";
}

function getFileTypeFromName(name) {
  if (/\.docx$/i.test(name)) return "docx";
  if (/\.pdf$/i.test(name)) return "pdf";
  return "pptx";
}

function getWordPartLabel(path) {
  if (path === "word/document.xml") return "正文";
  if (/header\d+\.xml$/.test(path)) return "页眉";
  if (/footer\d+\.xml$/.test(path)) return "页脚";
  if (/footnotes\.xml$/.test(path)) return "脚注";
  if (/endnotes\.xml$/.test(path)) return "尾注";
  if (/comments\.xml$/.test(path)) return "批注";
  return "文档";
}

function loadSettings() {
  try {
    const settings = JSON.parse(localStorage.getItem(SETTINGS_KEY) || "{}");
    if (settings.settingsVersion !== SETTINGS_VERSION && settings.pptLayoutMode === "keep-size") {
      settings.pptLayoutMode = "smart";
    }
    setElementValue(els.translationDirection, settings.translationDirection);
    setElementValue(els.pptLayoutMode, settings.pptLayoutMode);
    setElementValue(els.fontScale, settings.fontScale);
    setElementValue(els.modelName, settings.modelName);
    setElementValue(els.apiKey, settings.apiKey);
    updateFontScaleLabel();
  } catch {
    localStorage.removeItem(SETTINGS_KEY);
  }
}

function saveSettings() {
  const settings = {
    settingsVersion: SETTINGS_VERSION,
    translationDirection: els.translationDirection?.value || "",
    pptLayoutMode: els.pptLayoutMode?.value || "",
    fontScale: els.fontScale?.value || "100",
    modelName: els.modelName?.value || "",
    apiKey: els.apiKey?.value || "",
  };

  localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
}

function getTextBounds(paragraph) {
  const { offset, extent } = getParagraphTransform(paragraph) || {};

  if (!offset || !extent) return null;

  return {
    x: Number(offset.getAttribute("x")) || 0,
    y: Number(offset.getAttribute("y")) || 0,
    cx: Number(extent.getAttribute("cx")) || 0,
    cy: Number(extent.getAttribute("cy")) || 0,
  };
}

function getParagraphTransform(paragraph) {
  const shape = paragraph.parentElement?.parentElement;
  const transform = [...(shape?.getElementsByTagNameNS(DRAWING_NS, "xfrm") || [])][0];
  const offset = [...(transform?.getElementsByTagNameNS(DRAWING_NS, "off") || [])][0];
  const extent = [...(transform?.getElementsByTagNameNS(DRAWING_NS, "ext") || [])][0];

  return { transform, offset, extent };
}

function getSegmentBounds(segment) {
  return segment.overrides?.bounds || segment.layout?.bounds;
}

function updateSegmentBoundsFromPreview(segment, box, slideRect) {
  if (!segment.layout?.bounds) return;

  segment.overrides.bounds = {
    x: Math.max(0, (box.offsetLeft / slideRect.width) * state.slideSize.cx),
    y: Math.max(0, (box.offsetTop / slideRect.height) * state.slideSize.cy),
    cx: Math.max(1, (box.offsetWidth / slideRect.width) * state.slideSize.cx),
    cy: Math.max(1, (box.offsetHeight / slideRect.height) * state.slideSize.cy),
  };
}

function getParagraphFontSize(paragraph) {
  const runProperties = [...paragraph.getElementsByTagNameNS(DRAWING_NS, "rPr")].find((node) => node.getAttribute("sz"));
  return Number(runProperties?.getAttribute("sz")) || 1800;
}

function getTextBodyInsets(bodyProperties) {
  const defaultInset = 91440;
  return {
    left: getNumericAttribute(bodyProperties, "lIns", defaultInset),
    right: getNumericAttribute(bodyProperties, "rIns", defaultInset),
    top: getNumericAttribute(bodyProperties, "tIns", 45720),
    bottom: getNumericAttribute(bodyProperties, "bIns", 45720),
  };
}

function getParagraphRunStyle(paragraph) {
  const runProperties = [...paragraph.getElementsByTagNameNS(DRAWING_NS, "rPr")].find((node) => node.getAttribute("sz")) ||
    [...paragraph.getElementsByTagNameNS(DRAWING_NS, "rPr")][0];
  const latin = [...(runProperties?.children || [])].find((node) => node.localName === "latin")?.getAttribute("typeface");
  const eastAsian = [...(runProperties?.children || [])].find((node) => node.localName === "ea")?.getAttribute("typeface");

  return {
    latin: latin || "Arial",
    eastAsian: eastAsian || "Microsoft YaHei",
    bold: runProperties?.getAttribute("b") === "1",
    italic: runProperties?.getAttribute("i") === "1",
  };
}

function getParagraphStyle(paragraph) {
  const paragraphProperties = [...paragraph.children].find((node) => node.localName === "pPr");
  return {
    align: paragraphProperties?.getAttribute("algn") || "",
    lineHeight: getParagraphLineHeight(paragraphProperties),
  };
}

function getParagraphLineHeight(paragraphProperties) {
  const lineSpacing = [...(paragraphProperties?.children || [])].find((node) => node.localName === "lnSpc");
  const percent = [...(lineSpacing?.children || [])].find((node) => node.localName === "spcPct");
  const points = [...(lineSpacing?.children || [])].find((node) => node.localName === "spcPts");
  const percentValue = Number(percent?.getAttribute("val"));
  const pointValue = Number(points?.getAttribute("val"));

  if (percentValue > 0) return Math.max(0.8, Math.min(2.2, percentValue / 100000));
  if (pointValue > 0) return Math.max(0.8, Math.min(2.2, pointValue / 100));
  return 1;
}

function applyPreviewTextStyle(box, segment) {
  const layout = segment.layout || {};
  const insets = layout.insets || {};
  const runStyle = layout.runStyle || {};
  const paragraphStyle = layout.paragraphStyle || {};

  box.style.padding = [
    `${emuToPreviewCqw(insets.top || 0)}cqw`,
    `${emuToPreviewCqw(insets.right || 0)}cqw`,
    `${emuToPreviewCqw(insets.bottom || 0)}cqw`,
    `${emuToPreviewCqw(insets.left || 0)}cqw`,
  ].join(" ");
  box.style.fontFamily = `"${runStyle.latin || "Arial"}", "${runStyle.eastAsian || "Microsoft YaHei"}", sans-serif`;
  box.style.fontWeight = runStyle.bold ? "700" : "400";
  box.style.fontStyle = runStyle.italic ? "italic" : "normal";
  box.style.lineHeight = String(paragraphStyle.lineHeight || 1);
  box.style.textAlign = mapPptAlignment(paragraphStyle.align);
}

function emuToPreviewCqw(value) {
  return (Number(value || 0) / state.slideSize.cx) * 100;
}

function getNumericAttribute(node, name, fallback) {
  const value = Number(node?.getAttribute(name));
  return Number.isFinite(value) ? value : fallback;
}

function mapPptAlignment(value) {
  const map = {
    ctr: "center",
    r: "right",
    just: "justify",
    dist: "justify",
  };
  return map[value] || "left";
}

function getPreviewFontCqw(segment) {
  const scaledPt = (getPresentationExportFontSize(segment) / 100) * getPreviewAutofitScale(segment);
  const slideWidthPt = state.slideSize.cx / 12700;
  return (scaledPt / slideWidthPt) * 100;
}

function getPreviewAutofitScale(segment) {
  if (shouldUseSingleLine(segment)) return 0.85;
  if (getPptLayoutMode() === "compact-fit") return 0.88;

  const layout = segment.layout || {};
  if (layout.autofit === "normAutofit") {
    const fontScale = Number(layout.autofitAttrs?.fontScale);
    return fontScale > 0 ? Math.max(0.6, Math.min(1, fontScale / 100000)) : 1;
  }

  if (layout.autofit === "spAutoFit") return 1;
  return 0.85;
}

function getPresentationExportFontSize(segment) {
  const baseSize = Number(segment.layout?.fontSize || 1800);
  return Math.max(700, Math.round(baseSize * getPresentationLengthScale(segment)));
}

function getPreviewScaleLabel(segment) {
  const scale = Math.round(getActiveFontScale(segment) * 100);
  const points = Math.round(getPresentationExportFontSize(segment) / 100);
  return `${scale}% · ${points}pt`;
}

function handleSettingsChange() {
  updateFontScaleLabel();
  saveSettings();
  rerenderPreviewIfOpen();
}

function rerenderPreviewIfOpen() {
  if (els.previewDialog?.open) renderPreview();
}

function updateFontScaleLabel() {
  if (!els.fontScale || !els.fontScaleValue) return;
  els.fontScaleValue.textContent = `${els.fontScale.value}%`;
}

function getUserFontScale() {
  const value = Number(els.fontScale?.value || 100);
  return Math.max(0.6, Math.min(1.15, value / 100));
}

function setElementValue(element, value) {
  if (!element || !value) return;
  element.value = value;
}

function showToast(message, isError = false) {
  const toast = document.createElement("div");
  toast.className = `toast${isError ? " error" : ""}`;
  toast.textContent = message;
  document.body.append(toast);
  window.setTimeout(() => toast.remove(), 4200);
}

function showCompletionDialog(title, message) {
  const dialog = document.createElement("dialog");
  dialog.className = "completion-dialog";
  dialog.innerHTML = `
    <form method="dialog" class="completion-card">
      <h2></h2>
      <p></p>
      <button class="primary" value="ok">???</button>
    </form>
  `;
  dialog.querySelector("h2").textContent = title;
  dialog.querySelector("p").textContent = message;
  document.body.append(dialog);
  dialog.addEventListener("close", () => dialog.remove());
  if (typeof dialog.showModal === "function") {
    dialog.showModal();
  } else {
    window.alert(`${title}\n${message}`);
    dialog.remove();
  }
}

function getDirectionConfig(value) {
  const targets = {
    en: { name: "English", style: "fluent, concise English" },
    zh: { name: "Simplified Chinese", style: "natural Simplified Chinese" },
    ja: { name: "Japanese", style: "natural Japanese" },
    ko: { name: "Korean", style: "natural Korean" },
    fr: { name: "French", style: "natural French" },
    de: { name: "German", style: "natural German" },
    es: { name: "Spanish", style: "natural Spanish" },
    pt: { name: "Portuguese", style: "natural Portuguese" },
    it: { name: "Italian", style: "natural Italian" },
    ru: { name: "Russian", style: "natural Russian" },
    ar: { name: "Arabic", style: "natural Arabic" },
    th: { name: "Thai", style: "natural Thai" },
    vi: { name: "Vietnamese", style: "natural Vietnamese" },
    id: { name: "Indonesian", style: "natural Indonesian" },
  };

  const directConfigs = {
    "zh-en": { sourceLanguage: "Chinese", targetCode: "en" },
    "en-zh": { sourceLanguage: "English", targetCode: "zh" },
  };
  const direct = directConfigs[value];
  const targetCode = direct?.targetCode || String(value || "zh-en").replace(/^auto-/, "");
  const target = targets[targetCode] || targets.en;
  const sourcePhrase = direct?.sourceLanguage
    ? `Translate ${direct.sourceLanguage} presentation text`
    : "Detect the source language and translate the text";
  const compactGuidance = targetCode === "en"
    ? "Use short noun phrases for titles, labels, and table-of-contents entries."
    : "Keep titles, labels, and table-of-contents entries short and slide-friendly.";

  return {
    sourceLanguage: direct?.sourceLanguage || "auto",
    targetLanguage: target.name,
    targetCode,
    instruction: [
      `${sourcePhrase} into ${target.style} for business slides.`,
      compactGuidance,
      "Preserve numbers, units, product names, acronyms, and placeholders.",
      "Always translate the company name 伽奈维 as CuraWay exactly; never use Ganavi, Ganawei, Jianaiwei, or Curaway.",
    ].join(" "),
  };
}
