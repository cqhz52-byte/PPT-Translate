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
  pdfEraseFragments: new Map(),
  pdfPreservedRegions: new Map(),
  pdfManualPreservedRegions: new Map(),
  pdfManualRegionSelectionEnabled: false,
  pdfLayoutShowSourceBackground: false,
  pdfLayoutShowParsedMap: false,
  previewFullscreen: false,
  pdfBytes: null,
  pdfParseSource: "",
  pdfParsedMarkdown: "",
  pdfParsedPages: [],
  batchFiles: [],
  batchRunning: false,
  batchCancelRequested: false,
  translationRunning: false,
  translationAbortController: null,
  summaryRunning: false,
  savedFiles: [],
  wakeLock: null,
  wakeLockNoticeShown: false,
  wakeLockWarningShown: false,
  mobileView: "translate",
  serviceWorkerRegistration: null,
  updatePromptShown: false,
  updateCheckRunning: false,
  userRole: "",
  sessionChecked: false,
  pullStartY: 0,
  pullDistance: 0,
  pullCheckReady: false,
  pullIndicator: null,
  previewMode: "translation",
  activeSummary: null,
  originalPreviewUrl: "",
  loadToken: 0,
  importRunning: false,
  jsZipLoadPromise: null,
};

const els = {
  fileInput: document.querySelector("#fileInput"),
  uploadZone: document.querySelector(".upload-zone"),
  fileMeta: document.querySelector("#fileMeta"),
  segmentTable: document.querySelector("#segmentTable"),
  translateButton: document.querySelector("#translateButton"),
  batchTranslateButton: document.querySelector("#batchTranslateButton"),
  sourcePreviewButton: document.querySelector("#sourcePreviewButton"),
  previewButton: document.querySelector("#previewButton"),
  layoutPreviewButton: document.querySelector("#layoutPreviewButton"),
  downloadButton: document.querySelector("#downloadButton"),
  shareButton: document.querySelector("#shareButton"),
  summaryButton: document.querySelector("#summaryButton"),
  summaryDetail: document.querySelector("#summaryDetail"),
  summaryDialog: document.querySelector("#summaryDialog"),
  summaryCloseButton: document.querySelector("#summaryCloseButton"),
  summaryCopyButton: document.querySelector("#summaryCopyButton"),
  summaryShareButton: document.querySelector("#summaryShareButton"),
  summaryMeta: document.querySelector("#summaryMeta"),
  summaryOutput: document.querySelector("#summaryOutput"),
  resetButton: document.querySelector("#resetButton"),
  helpButton: document.querySelector("#helpButton"),
  helpDialog: document.querySelector("#helpDialog"),
  helpCloseButton: document.querySelector("#helpCloseButton"),
  appShareButton: document.querySelector("#appShareButton"),
  adminButton: document.querySelector("#adminButton"),
  libraryButton: document.querySelector("#libraryButton"),
  workspace: document.querySelector("#workspace"),
  mobileViewButton: document.querySelector("#mobileViewButton"),
  mobileViewMenu: document.querySelector("#mobileViewMenu"),
  mobileViewTargets: [...document.querySelectorAll("[data-mobile-view-target]")],
  returnHomeButtons: [...document.querySelectorAll("[data-return-home]")],
  previewDialog: document.querySelector("#previewDialog"),
  previewCloseButton: document.querySelector("#previewCloseButton"),
  previewFullscreenButton: document.querySelector("#previewFullscreenButton"),
  previewDownloadButton: document.querySelector("#previewDownloadButton"),
  previewShareButton: document.querySelector("#previewShareButton"),
  previewBody: document.querySelector("#previewBody"),
  previewTitle: document.querySelector("#previewTitle"),
  previewMeta: document.querySelector("#previewMeta"),
  translationDirection: document.querySelector("#translationDirection"),
  pptLayoutMode: document.querySelector("#pptLayoutMode"),
  fontScale: document.querySelector("#fontScale"),
  fontScaleValue: document.querySelector("#fontScaleValue"),
  modelName: document.querySelector("#modelName"),
  apiKey: document.querySelector("#apiKey"),
  pdfOutputMode: document.querySelector("#pdfOutputMode"),
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
const wordPathPattern = /^word\/(?:document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/;
const DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const PPT_EMU_PER_PT = 12700;
const SETTINGS_KEY = "deepseek-document-translator-settings-v1";
const SETTINGS_VERSION = 5;
const SAVED_FILES_DB = "curaway-translated-files-v1";
const SAVED_FILES_STORE = "files";
const CURRENT_DRAFT_DB = "curaway-current-draft-v1";
const CURRENT_DRAFT_STORE = "drafts";
const CURRENT_DRAFT_ID = "current";
const SUMMARY_CACHE_DB = "curaway-summary-cache-v1";
const SUMMARY_CACHE_STORE = "summaries";
const DRAFT_SAVE_DELAY = 600;
const APP_VERSION = "v136";
const VERSION_URL = "./version.json";
const JSZIP_URL = "./vendor/jszip.min.js";
const UPDATE_CHECK_INTERVAL = 5 * 60 * 1000;
const PULL_UPDATE_THRESHOLD = 76;
const DEEPSEEK_API_BASE = "https://api.deepseek.com";
const TRANSLATE_PROXY = "./api/translate";
const TRANSLATION_REQUEST_TIMEOUT_MS = 45000;
const TRANSLATION_RETRY_DELAY_MS = 900;
const PDF_PARSE_PROXY = "./api/pdf-parse";
const PDFJS_URL = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.min.mjs";
const PDFJS_WORKER_URL = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.worker.min.mjs";
const PDFLIB_URLS = [
  "./vendor/pdf-lib.esm.min.js",
  "https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/+esm",
];
const FONTKIT_URLS = [
  "./vendor/fontkit.bundle.js",
  "https://cdn.jsdelivr.net/npm/@pdf-lib/fontkit@1.1.1/+esm",
];
const PDF_FONT_URLS = [
  "./vendor/fonts/NotoSansCJKsc-Regular.otf",
  "https://cdn.jsdelivr.net/gh/notofonts/noto-cjk@main/Sans/OTF/SimplifiedChinese/NotoSansCJKsc-Regular.otf",
];
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

  navigator.serviceWorker.register("sw.js?v=81").then((registration) => {
    state.serviceWorkerRegistration = registration;
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

initUpdateChecks();
decorateActionButtons();
updateAdminButtonVisibility();
refreshCurrentUserSession();

document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "visible" && document.body.classList.contains("is-busy")) {
    requestScreenWakeLock();
  } else if (document.visibilityState === "hidden") {
    flushCurrentDraftSave();
  }
  if (document.visibilityState === "visible") {
    checkForAppUpdate();
  }
});

window.addEventListener("pagehide", flushCurrentDraftSave);

function initUpdateChecks() {
  createPullUpdateIndicator();
  window.setTimeout(() => checkForAppUpdate(), 2500);
  window.setInterval(() => {
    if (document.visibilityState === "visible") checkForAppUpdate();
  }, UPDATE_CHECK_INTERVAL);

  window.addEventListener("touchstart", (event) => {
    if (window.scrollY > 0 || document.body.classList.contains("is-busy")) return;
    state.pullStartY = event.touches[0]?.clientY || 0;
    state.pullDistance = 0;
    state.pullCheckReady = false;
  }, { passive: true });

  window.addEventListener("touchmove", (event) => {
    if (!state.pullStartY || window.scrollY > 0 || document.body.classList.contains("is-busy")) return;
    const delta = (event.touches[0]?.clientY || 0) - state.pullStartY;
    state.pullDistance = Math.max(0, delta);
    updatePullUpdateIndicator(state.pullDistance);
    if (delta > PULL_UPDATE_THRESHOLD && !state.pullCheckReady) {
      state.pullCheckReady = true;
      setStatus("松手检查是否有新版本...");
    }
  }, { passive: true });

  window.addEventListener("touchend", () => {
    const shouldCheck = state.pullCheckReady;
    state.pullStartY = 0;
    state.pullDistance = 0;
    state.pullCheckReady = false;
    if (shouldCheck) {
      updatePullUpdateIndicator(PULL_UPDATE_THRESHOLD, "checking");
      checkForAppUpdate({ manual: true });
    } else {
      hidePullUpdateIndicator();
    }
  }, { passive: true });

  window.addEventListener("touchcancel", () => {
    state.pullStartY = 0;
    state.pullDistance = 0;
    state.pullCheckReady = false;
    hidePullUpdateIndicator();
  }, { passive: true });
}

function createPullUpdateIndicator() {
  if (state.pullIndicator) return state.pullIndicator;
  const indicator = document.createElement("div");
  indicator.className = "pull-update-indicator";
  indicator.setAttribute("aria-live", "polite");
  indicator.innerHTML = `
    <span class="pull-update-spinner" aria-hidden="true"></span>
    <span class="pull-update-text">下拉检查更新</span>
  `;
  document.body.append(indicator);
  state.pullIndicator = indicator;
  return indicator;
}

function updatePullUpdateIndicator(distance, mode = "") {
  const indicator = state.pullIndicator || createPullUpdateIndicator();
  const progress = Math.max(0, Math.min(1, distance / PULL_UPDATE_THRESHOLD));
  const ready = mode === "checking" || progress >= 1;
  const translate = Math.round(-58 + progress * 76);
  indicator.classList.add("visible");
  indicator.classList.toggle("ready", ready && mode !== "checking");
  indicator.classList.toggle("checking", mode === "checking");
  indicator.style.setProperty("--pull-progress", progress.toFixed(2));
  indicator.style.transform = `translate(-50%, ${translate}px)`;
  const text = indicator.querySelector(".pull-update-text");
  if (text) {
    text.textContent = mode === "checking"
      ? "正在检查更新..."
      : ready
        ? "松手检查更新"
        : "继续下拉检查更新";
  }
}

function hidePullUpdateIndicator() {
  const indicator = state.pullIndicator;
  if (!indicator) return;
  indicator.classList.remove("visible", "ready", "checking");
  indicator.style.transform = "";
}

async function checkForAppUpdate(options = {}) {
  if (state.updateCheckRunning) {
    if (options.manual) window.setTimeout(hidePullUpdateIndicator, 500);
    return false;
  }
  state.updateCheckRunning = true;

  try {
    const response = await fetch(`${VERSION_URL}?t=${Date.now()}`, {
      cache: "no-store",
      headers: { "Cache-Control": "no-cache" },
    });
    if (!response.ok) throw new Error("Version check failed");
    const info = await response.json();
    const latestVersion = String(info.version || "").trim();
    if (!latestVersion) return false;

    if (isNewerAppVersion(latestVersion, APP_VERSION)) {
      showAppUpdatePrompt(info);
      return true;
    }

    if (options.manual) {
      showToast(`当前已经是最新版本 ${APP_VERSION}。`);
      setStatus(`当前已经是最新版本 ${APP_VERSION}。`);
      state.serviceWorkerRegistration?.update?.().catch(() => {});
    }
    return false;
  } catch (error) {
    if (options.manual) {
      showToast("暂时无法检查更新，请稍后再试。", true);
    }
    return false;
  } finally {
    state.updateCheckRunning = false;
    if (options.manual) {
      window.setTimeout(hidePullUpdateIndicator, 500);
    }
  }
}

function isNewerAppVersion(latestVersion, currentVersion) {
  const latest = Number(String(latestVersion).replace(/[^\d.]/g, "").split(".")[0] || 0);
  const current = Number(String(currentVersion).replace(/[^\d.]/g, "").split(".")[0] || 0);
  if (latest && current) return latest > current;
  return latestVersion !== currentVersion;
}

function showAppUpdatePrompt(info) {
  if (state.updatePromptShown) return;
  state.updatePromptShown = true;

  const latestVersion = String(info.version || "").trim();
  const message = info.message || `发现新版本 ${latestVersion}，建议立即更新。`;
  const dialog = document.createElement("dialog");
  dialog.className = "completion-dialog";
  dialog.innerHTML = `
    <form method="dialog" class="completion-card">
      <h2></h2>
      <p></p>
      <div class="dialog-actions">
        <button value="later" type="submit">稍后</button>
        <button class="primary" value="update" type="submit">立即更新</button>
      </div>
    </form>
  `;
  dialog.querySelector("h2").textContent = "发现新版本";
  dialog.querySelector("p").textContent = message;
  document.body.append(dialog);
  dialog.addEventListener("close", () => {
    const shouldUpdate = dialog.returnValue === "update";
    dialog.remove();
    state.updatePromptShown = false;
    if (shouldUpdate) {
      applyAppUpdate();
    }
  });
  if (typeof dialog.showModal === "function") {
    dialog.showModal();
  } else if (window.confirm(message)) {
    applyAppUpdate();
  }
}

async function applyAppUpdate() {
  setStatus("正在更新到最新版本...");
  try {
    const registration = state.serviceWorkerRegistration || await navigator.serviceWorker?.getRegistration?.();
    await registration?.update?.();
    if (registration?.waiting) {
      registration.waiting.postMessage({ type: "SKIP_WAITING" });
      return;
    }
  } catch (error) {
    console.warn("App update failed", error);
  }
  window.location.reload();
}

els.libraryButton?.addEventListener("click", openFileLibrary);

els.uploadZone?.addEventListener("click", (event) => {
  if (event.target === els.fileInput) return;
  event.preventDefault();
  openFilePicker();
});

els.fileInput.addEventListener("click", (event) => {
  event.stopPropagation();
  els.fileInput.value = "";
});

els.fileInput.addEventListener("change", async (event) => {
  const files = [...(event.target.files || [])];
  event.target.value = "";
  await handleSelectedFiles(files);
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

els.returnHomeButtons.forEach((button) => {
  button.addEventListener("click", () => {
    setMobileView("translate");
    setMobileMenuOpen(false);
    els.uploadZone?.scrollIntoView({ block: "nearest", behavior: "smooth" });
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
  els.uploadZone.addEventListener(eventName, async (event) => {
    event.preventDefault();
    if (eventName === "drop") {
      const files = [...(event.dataTransfer?.files || [])];
      await handleSelectedFiles(files);
    }
    els.uploadZone.classList.remove("drag-over");
  });
});

els.translateButton.addEventListener("click", handleTranslateButtonClick);
els.batchTranslateButton?.addEventListener("click", processBatchQueue);
els.sourcePreviewButton?.addEventListener("click", openSourcePreview);
els.previewButton.addEventListener("click", openPreview);
els.layoutPreviewButton?.addEventListener("click", openPreview);
els.summaryButton?.addEventListener("click", summarizeDocument);
els.summaryCloseButton?.addEventListener("click", closeSummary);
els.summaryCopyButton?.addEventListener("click", copySummary);
els.summaryShareButton?.addEventListener("click", shareSummary);
els.summaryDialog?.addEventListener("click", (event) => {
  if (event.target === els.summaryDialog) closeSummary();
});
els.downloadButton.addEventListener("click", downloadPresentation);
els.shareButton.addEventListener("click", sharePresentation);
els.resetButton.addEventListener("click", handleResetButtonClick);
els.helpButton?.addEventListener("click", openHelp);
els.helpCloseButton?.addEventListener("click", closeHelp);
els.appShareButton?.addEventListener("click", shareApp);
els.adminButton?.addEventListener("click", openAdminManagement);
els.previewCloseButton.addEventListener("click", closePreview);
els.previewFullscreenButton?.addEventListener("click", togglePreviewFullscreen);
els.previewDownloadButton.addEventListener("click", handlePreviewDownload);
els.previewShareButton.addEventListener("click", handlePreviewShare);
els.previewDialog.addEventListener("click", (event) => {
  if (event.target === els.previewDialog) closePreview();
});
els.helpDialog?.addEventListener("click", (event) => {
  if (event.target === els.helpDialog) closeHelp();
});
document.addEventListener("fullscreenchange", syncPreviewFullscreenState);

[
  els.translationDirection,
  els.pptLayoutMode,
  els.fontScale,
  els.summaryDetail,
  els.modelName,
  els.apiKey,
  els.pdfOutputMode,
].forEach((element) => {
  element?.addEventListener("change", handleSettingsChange);
  element?.addEventListener("input", handleSettingsChange);
});

function openFilePicker() {
  if (document.body.classList.contains("is-busy") || state.importRunning) {
    showToast("正在处理当前文档，请稍候再选择新文件。", true);
    return;
  }
  els.fileInput.disabled = false;
  els.fileInput.value = "";
  els.fileInput.click();
}

async function handleSelectedFiles(files) {
  const selectedFiles = [...(files || [])].filter(Boolean);
  els.fileInput.value = "";
  if (!selectedFiles.length) return;
  if (document.body.classList.contains("is-busy") || state.importRunning) {
    showToast("当前任务还没有完成，请稍候再选择新文件。", true);
    return;
  }
  if (selectedFiles.length > 1) {
    queueBatchFiles(selectedFiles);
    return;
  }
  await loadOfficeFile(selectedFiles[0]);
}

updateFontScaleLabel();
initializeAppStartup();

async function refreshCurrentUserSession() {
  try {
    const response = await fetch("/api/auth/session", {
      cache: "no-store",
      headers: { Accept: "application/json" },
    });
    if (!response.ok) throw new Error(`Session check failed: ${response.status}`);
    const session = await response.json();
    state.userRole = session?.role || "";
  } catch (error) {
    console.warn("User session check failed", error);
    state.userRole = "";
  } finally {
    state.sessionChecked = true;
    updateAdminButtonVisibility();
  }
}

function isSuperAdmin() {
  return state.userRole === "admin";
}

function updateAdminButtonVisibility() {
  if (!els.adminButton) return;
  const shouldShow = isSuperAdmin();
  els.adminButton.hidden = !shouldShow;
  els.adminButton.disabled = !shouldShow;
}

async function initializeAppStartup() {
  updateStartupStatus("正在加载应用界面...", "正在准备按钮、文件库和离线能力。", 28);

  try {
    updateStartupStatus("正在读取文件库...", "读取本机已保存的翻译文件。", 46);
    await loadSavedFiles();
  } catch (error) {
    console.warn("Saved file list unavailable", error);
  }

  try {
    updateStartupStatus("正在检查上次编辑内容...", "如果上次有未导出的文件，会自动恢复。", 64);
    await restoreCurrentDraft();
  } catch (error) {
    console.warn("Draft restore unavailable", error);
  }

  updateStartupStatus("准备完成", "可以选择文档开始翻译。", 100);
  window.setTimeout(hideStartupOverlay, 260);
}

function updateStartupStatus(message, detail = "", progress = 0) {
  const overlay = document.querySelector("#startupOverlay");
  if (!overlay) return;
  const messageEl = overlay.querySelector("#startupMessage");
  const detailEl = overlay.querySelector("#startupDetail");
  const progressEl = overlay.querySelector("#startupProgress");
  if (messageEl) messageEl.textContent = message;
  if (detailEl) detailEl.textContent = detail;
  if (progressEl) progressEl.style.setProperty("--startup-progress", `${Math.max(8, Math.min(100, progress))}%`);
}

function hideStartupOverlay() {
  const overlay = document.querySelector("#startupOverlay");
  if (!overlay) return;
  overlay.classList.add("hidden");
  window.setTimeout(() => overlay.remove(), 320);
}

async function loadOfficeFile(file, options = {}) {
  const silent = Boolean(options.silent);
  const loadToken = Date.now() + Math.random();
  state.loadToken = loadToken;
  state.importRunning = true;
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

    setProgress(0.02);
    setBusy(true, `正在导入 ${file.name}...`);
    els.fileMeta.textContent = `${file.name} · 正在导入...`;
    revokeOriginalPreviewUrl();
    state.file = file;
    state.fileType = getFileTypeFromName(file.name);
    state.zip = null;
    state.segments = [];
    state.slideCount = 0;
    state.slideSize = { cx: 12192000, cy: 6858000 };
    state.slideVisuals = new Map();
    state.pdfPageSizes = new Map();
    state.pdfTableCells = new Map();
    state.pdfEraseFragments = new Map();
    state.pdfPreservedRegions = new Map();
    state.pdfManualPreservedRegions = new Map();
    state.pdfManualRegionSelectionEnabled = false;
    state.pdfBytes = null;
    state.pdfParseSource = "";
    state.pdfParsedMarkdown = "";
    state.pdfParsedPages = [];
    document.body.classList.toggle("has-pdf-file", state.fileType === "pdf");

    if (state.fileType === "pdf") {
      await loadPdfDocument(file, loadToken);
    } else {
      setProgress(0.08);
      setStatus(`正在展开 ${getFileTypeName()} 文件结构，请稍候...`);
      await waitForUiFrame();
      const JSZipLib = await loadJsZip();
      state.zip = await JSZipLib.loadAsync(file);
    }

    if (state.loadToken !== loadToken) return;

    if (state.fileType === "docx") {
      await loadWordDocument(loadToken);
    } else if (state.fileType === "pptx") {
      await loadPresentation(loadToken);
    }

    if (state.loadToken !== loadToken) return;

    renderSegments();
    updateStats();
    els.fileMeta.textContent = `${file.name} · ${(file.size / 1024 / 1024).toFixed(2)} MB`;
    const pdfSourceLabel = state.fileType === "pdf"
      ? (state.pdfParseSource === "llamaparse" ? "（Cloudflare + LlamaParse 结构化解析）" : "（浏览器兼容解析）")
      : "";
    setStatus(`已解析 ${getFileTypeName()}${pdfSourceLabel}，共 ${state.segments.length} 段文字。默认使用 DeepSeek 翻译。`);
    if (!options.skipDraftSave) scheduleCurrentDraftSave({ immediate: true });
    if (!silent) showToast(`${getFileTypeName()} 已载入，可以开始翻译或手动编辑。`);
  } catch (error) {
    if (state.loadToken === loadToken) {
      showToast(error.message || "读取文件失败。", true);
      resetApp(false, { clearDraft: false });
    }
  } finally {
    if (state.loadToken === loadToken) {
      state.importRunning = false;
      setBusy(false);
    }
  }
}

async function loadPresentation(loadToken = state.loadToken) {
    setProgress(0.14);
    setStatus("正在读取 PPTX 幻灯片列表...");
    await waitForUiFrame();
    if (state.loadToken !== loadToken) return;
    state.slideSize = await readSlideSize();
    const slideFiles = Object.keys(state.zip.files)
      .map((path) => ({ path, match: path.match(slidePathPattern) }))
      .filter((item) => item.match)
      .sort((a, b) => Number(a.match[1]) - Number(b.match[1]));

    state.slideCount = slideFiles.length;

    for (let index = 0; index < slideFiles.length; index += 1) {
      if (state.loadToken !== loadToken) return;
      const item = slideFiles[index];
      const slideNumber = Number(item.match[1]);
      setProgress(0.14 + ((index + 1) / Math.max(1, slideFiles.length)) * 0.76);
      setStatus(`正在解析 PPTX 第 ${slideNumber}/${slideFiles.length} 页...`);
      await waitForUiFrame();
      if (state.loadToken !== loadToken) return;
      const xmlText = await state.zip.file(item.path).async("text");
      if (state.loadToken !== loadToken) return;
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

async function loadPdfDocument(file, loadToken = state.loadToken) {
  const pdfjs = await loadPdfJs();
  if (state.loadToken !== loadToken) return;
  const pdfBytes = new Uint8Array(await file.arrayBuffer());
  if (state.loadToken !== loadToken) return;
  state.pdfBytes = pdfBytes.slice();
  const pdf = await pdfjs.getDocument({ data: pdfBytes }).promise;
  if (state.loadToken !== loadToken) return;
  state.slideCount = pdf.numPages;
  state.pdfParseSource = "pdfjs";
  state.pdfParsedMarkdown = "";
  state.pdfParsedPages = [];

  let pdfReferenceSectionStarted = false;

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    if (state.loadToken !== loadToken) return;
    setProgress(0.08 + (pdf.numPages ? pageNumber / pdf.numPages : 0) * 0.7);
    setStatus(`正在解析 PDF 第 ${pageNumber}/${pdf.numPages} 页，提取文字和页面结构...`);
    await waitForUiFrame();
    if (state.loadToken !== loadToken) return;
    const page = await pdf.getPage(pageNumber);
    if (state.loadToken !== loadToken) return;
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
    const pageWidth = viewport.width || 595.28;
    const pageHeight = viewport.height || 841.89;
    const content = await page.getTextContent();
    if (state.loadToken !== loadToken) return;
    state.pdfEraseFragments.set(`pdf/page-${pageNumber}`, extractPdfTextEraseFragments(content.items));
    const lines = extractPdfLineSegments(content.items, pageWidth);
    const blocks = groupPdfLineSegmentsIntoBlocks(lines, pageWidth, pageHeight);
    const referenceContext = getPdfReferenceContext(lines, pageHeight, pdfReferenceSectionStarted);
    const preservedRegions = [
      ...detectPdfPreservedFigureRegions(page, blocks, pageWidth, pageHeight),
    ];
    state.pdfPreservedRegions.set(`pdf/page-${pageNumber}`, preservedRegions);
    pdfReferenceSectionStarted = pdfReferenceSectionStarted || referenceContext.startedOnPage;
    blocks.forEach((block, index) => {
      const tableCell = findPdfTableCell(block.bounds, tableCells);
      const sampledBackground = samplePdfBackgroundColor(pageSample, block.bounds, pageWidth, pageHeight);
      const backgroundColor = normalizePdfBackgroundColor(sampledBackground, Boolean(tableCell));
      state.segments.push({
        id: `pdf-${pageNumber}-${index}`,
        type: "pdf",
        slideNumber: pageNumber,
        locationLabel: `第 ${pageNumber} 页`,
        path: `pdf/page-${pageNumber}`,
        paragraphIndex: index,
        textNodeCount: 1,
        layout: {
          bounds: block.bounds,
          fontSize: block.fontSize,
          availableWidth: tableCell ? Math.max(block.availableWidth, tableCell.width - 4) : block.availableWidth,
          availableHeight: tableCell ? Math.max(block.availableHeight, tableCell.height - 4) : block.availableHeight,
          rowSegmentCount: tableCell ? Math.max(2, block.rowSegmentCount) : block.rowSegmentCount,
          lineCount: block.lineCount || 1,
          columnKey: block.columnKey || "",
          tableCell,
          cellAlign: tableCell ? inferPdfCellAlign(block.bounds, tableCell, block.text) : "",
          backgroundColor,
          textColor: getReadablePdfTextColor(backgroundColor),
          isReferenceText: isPdfReferenceLine(block, referenceContext, pageWidth, pageHeight),
          isFigureCaption: isPdfFigureCaptionText(block.text),
        },
        overrides: createSegmentOverrides(),
        original: block.text,
        translation: "",
      });
    });
  }

  if (state.segments.length) {
    setPdfOutputMode("overlay", { persist: true });
  } else {
    await loadPdfWithBackendFallback(file, loadToken);
  }

  if (!state.segments.length) {
    throw new Error("这个 PDF 没有可提取的文本。若是扫描版 PDF，请确认 Cloudflare 已配置 LLAMAPARSE_API_KEY 后再试。");
  }
}

async function loadPdfWithBackendFallback(file, loadToken) {
  let parsed = null;
  try {
    setProgress(0.82);
    setStatus("浏览器未提取到可选文字，正在调用 Cloudflare 后端 OCR/结构化解析...");
    await waitForUiFrame();
    parsed = await parsePdfWithBackend(file);
    if (state.loadToken !== loadToken) return;
  } catch (error) {
    console.warn("Backend PDF parsing unavailable", error);
    if (state.loadToken !== loadToken) return;
    setStatus(`后端 PDF 解析不可用：${error.message || "未知错误"}`);
    return;
  }

  if (parsed?.markdown || parsed?.pages?.length) {
    state.pdfParseSource = parsed.source || "llamaparse";
    state.pdfParsedMarkdown = parsed.markdown || "";
    state.pdfParsedPages = parsed.pages || [];
  }

  const structuredSegments = createPdfStructuredSegmentsFromLlamaParse();
  if (!structuredSegments.length) return;

  state.segments = structuredSegments;
  const hasOverlayCoordinates = structuredSegments.some((segment) => segment.layout?.bounds);
  setPdfOutputMode(hasOverlayCoordinates ? "overlay" : "reflow", { persist: true });
  setStatus(`已使用 Cloudflare 后端解析 PDF，识别到 ${structuredSegments.length} 个结构化段落。`);
}

async function parsePdfWithBackend(file) {
  const form = new FormData();
  form.append("file", file, file.name || "document.pdf");
  const response = await fetch(PDF_PARSE_PROXY, {
    method: "POST",
    body: form,
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(data.error || `PDF 后端解析失败：${response.status}`);
  }
  return data;
}

function createPdfStructuredSegmentsFromLlamaParse() {
  const fromPages = createPdfSegmentsFromParsedPages(state.pdfParsedPages);
  if (fromPages.length) return fromPages;
  return createPdfSegmentsFromMarkdown(state.pdfParsedMarkdown);
}

function createPdfSegmentsFromParsedPages(pages) {
  const segments = [];
  (pages || []).forEach((page, pageIndex) => {
    const pageNumber = Number(page.page || page.page_number || page.pageNumber || pageIndex + 1);
    const pageSize = state.pdfPageSizes.get(`pdf/page-${pageNumber}`) || { width: 595.28, height: 841.89 };
    const items = Array.isArray(page.items) ? page.items : (Array.isArray(page.elements) ? page.elements : []);
    items.forEach((item, itemIndex) => {
      const text = normalizePdfParsedText(item.text || item.value || item.md || item.markdown || "");
      if (!isUsefulPdfParsedText(text)) return;
      const bounds = normalizeParsedPdfBounds(item, pageSize);
      segments.push(createPdfParsedSegment({
        id: `pdf-llama-${pageNumber}-${itemIndex}`,
        pageNumber,
        index: itemIndex,
        text,
        bounds,
        itemType: String(item.type || item.category || ""),
      }));
    });

    if (!items.length) {
      splitParsedMarkdownIntoBlocks(page.md || page.markdown || page.text || "").forEach((text, blockIndex) => {
        segments.push(createPdfParsedSegment({
          id: `pdf-llama-${pageNumber}-${blockIndex}`,
          pageNumber,
          index: blockIndex,
          text,
          bounds: null,
          itemType: "",
        }));
      });
    }
  });
  return segments.slice(0, 1200);
}

function createPdfSegmentsFromMarkdown(markdown) {
  return splitParsedMarkdownIntoBlocks(markdown).slice(0, 1200).map((text, index) => {
    const pageNumber = inferParsedMarkdownPage(index);
    return createPdfParsedSegment({
      id: `pdf-llama-md-${index}`,
      pageNumber,
      index,
      text,
      bounds: null,
      itemType: "",
    });
  });
}

function createPdfParsedSegment({ id, pageNumber, index, text, bounds, itemType }) {
  const pagePath = `pdf/page-${pageNumber}`;
  const pageSize = state.pdfPageSizes.get(pagePath) || { width: 595.28, height: 841.89 };
  const tableCell = bounds ? findPdfTableCell(bounds, state.pdfTableCells.get(pagePath) || []) : null;
  const backgroundColor = { r: 1, g: 1, b: 1 };
  return {
    id,
    type: "pdf",
    slideNumber: pageNumber,
    locationLabel: `第 ${pageNumber} 页`,
    path: pagePath,
    paragraphIndex: index,
    textNodeCount: 1,
    layout: {
      bounds,
      fontSize: inferParsedPdfFontSize(text, itemType),
      availableWidth: bounds?.width || pageSize.width - 72,
      availableHeight: bounds?.height || 80,
      rowSegmentCount: tableCell ? 2 : 1,
      tableCell,
      cellAlign: tableCell && bounds ? inferPdfCellAlign(bounds, tableCell, text) : "",
      backgroundColor,
      textColor: getReadablePdfTextColor(backgroundColor),
      isReferenceText: isPdfReferencesHeading(text) || isLikelyPdfReferenceCitationLine(text),
      source: "llamaparse",
      itemType,
    },
    overrides: createSegmentOverrides(),
    original: text,
    translation: "",
  };
}

function splitParsedMarkdownIntoBlocks(markdown) {
  return String(markdown || "")
    .replace(/\r\n/g, "\n")
    .split(/\n{2,}/)
    .map(normalizePdfParsedText)
    .filter(isUsefulPdfParsedText);
}

function normalizePdfParsedText(text) {
  return String(text || "")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function isUsefulPdfParsedText(text) {
  const clean = String(text || "").replace(/[#*_`|:-]/g, "").replace(/\s+/g, "");
  return clean.length >= 2;
}

function inferParsedMarkdownPage(index) {
  const pageCount = Math.max(1, state.slideCount || 1);
  return Math.min(pageCount, Math.max(1, Math.floor(index / 12) + 1));
}

function inferParsedPdfFontSize(text, itemType) {
  if (/title|heading/i.test(itemType)) return 14;
  if (/table/i.test(itemType)) return 9;
  return String(text || "").length > 180 ? 10 : 11;
}

function normalizeParsedPdfBounds(item, pageSize) {
  const raw = item.bbox || item.bounding_box || item.boundingBox || item.bBox || item.bounds;
  if (!raw) return null;
  const box = Array.isArray(raw)
    ? { x: raw[0], y: raw[1], width: raw[2] - raw[0], height: raw[3] - raw[1] }
    : {
        x: raw.x ?? raw.left ?? raw.l,
        y: raw.y ?? raw.top ?? raw.t,
        width: raw.width ?? raw.w ?? ((raw.right ?? raw.r) - (raw.x ?? raw.left ?? raw.l)),
        height: raw.height ?? raw.h ?? ((raw.bottom ?? raw.b) - (raw.y ?? raw.top ?? raw.t)),
      };
  const x = Number(box.x);
  const y = Number(box.y);
  const width = Number(box.width);
  const height = Number(box.height);
  if (![x, y, width, height].every(Number.isFinite) || width <= 0 || height <= 0) return null;

  const looksNormalized = Math.max(x + width, y + height) <= 1.5;
  const scaled = looksNormalized
    ? { x: x * pageSize.width, y: y * pageSize.height, width: width * pageSize.width, height: height * pageSize.height }
    : { x, y, width, height };

  const topOrigin = Boolean(item.bbox || item.bounding_box || item.boundingBox || item.bBox);
  return {
    x: clampNumber(scaled.x, 0, pageSize.width),
    y: topOrigin
      ? clampNumber(pageSize.height - scaled.y - scaled.height, 0, pageSize.height)
      : clampNumber(scaled.y, 0, pageSize.height),
    width: clampNumber(scaled.width, 1, pageSize.width),
    height: clampNumber(scaled.height, 1, pageSize.height),
  };
}

async function loadPdfJs() {
  if (!window.pdfjsLib) {
    const module = await import(PDFJS_URL);
    window.pdfjsLib = module;
  }

  window.pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER_URL;
  return window.pdfjsLib;
}

async function loadJsZip() {
  if (window.JSZip) return window.JSZip;
  if (!state.jsZipLoadPromise) {
    setStatus("正在加载 Office 文档解析组件...");
    state.jsZipLoadPromise = loadScriptOnce(JSZIP_URL, "JSZip 文档解析组件")
      .then(() => {
        if (!window.JSZip) throw new Error("JSZip 文档解析组件加载失败。");
        return window.JSZip;
      })
      .finally(() => {
        state.jsZipLoadPromise = null;
      });
  }
  return state.jsZipLoadPromise;
}

function loadScriptOnce(src, label = "组件") {
  return new Promise((resolve, reject) => {
    const existing = document.querySelector(`script[data-dynamic-src="${src}"]`);
    if (existing?.dataset.loaded === "true") {
      resolve();
      return;
    }
    if (existing) {
      existing.addEventListener("load", resolve, { once: true });
      existing.addEventListener("error", () => reject(new Error(`${label}加载失败。`)), { once: true });
      return;
    }

    const script = document.createElement("script");
    script.src = src;
    script.async = true;
    script.dataset.dynamicSrc = src;
    script.addEventListener("load", () => {
      script.dataset.loaded = "true";
      resolve();
    }, { once: true });
    script.addEventListener("error", () => reject(new Error(`${label}加载失败。`)), { once: true });
    document.head.append(script);
  });
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

async function rebuildPdfPageNonTextElements(sourcePage, targetPage, pdfDoc, pdfjs, pageSize, rgb, options = {}) {
  const operatorList = await sourcePage.getOperatorList();
  const OPS = pdfjs.OPS || {};
  const fnArray = operatorList.fnArray || [];
  const argsArray = operatorList.argsArray || [];
  const pageWidth = Number(pageSize.width || 595.28);
  const pageHeight = Number(pageSize.height || 841.89);
  const stateStack = [createPdfGraphicsState()];
  let currentPath = null;
  const deadline = Number(options.deadline || 0);

  const op = {
    save: OPS.save ?? 10,
    restore: OPS.restore ?? 11,
    transform: OPS.transform ?? 12,
    setLineWidth: OPS.setLineWidth ?? 2,
    dependency: OPS.dependency ?? 1,
    setStrokeRGBColor: OPS.setStrokeRGBColor ?? 58,
    setFillRGBColor: OPS.setFillRGBColor ?? 59,
    setStrokeCMYKColor: OPS.setStrokeCMYKColor ?? 56,
    setFillCMYKColor: OPS.setFillCMYKColor ?? 57,
    setStrokeColorN: OPS.setStrokeColorN ?? 60,
    setFillColorN: OPS.setFillColorN ?? 61,
    setStrokeGray: OPS.setStrokeGray ?? 54,
    setFillGray: OPS.setFillGray ?? 53,
    constructPath: OPS.constructPath ?? 91,
    stroke: OPS.stroke ?? 20,
    closeStroke: OPS.closeStroke ?? 22,
    fill: OPS.fill ?? 21,
    eoFill: OPS.eoFill ?? 24,
    fillStroke: OPS.fillStroke ?? 25,
    eoFillStroke: OPS.eoFillStroke ?? 26,
    closeFillStroke: OPS.closeFillStroke ?? 27,
    endPath: OPS.endPath ?? 23,
    moveTo: OPS.moveTo ?? 13,
    lineTo: OPS.lineTo ?? 14,
    curveTo: OPS.curveTo ?? 15,
    curveTo2: OPS.curveTo2 ?? 16,
    curveTo3: OPS.curveTo3 ?? 17,
    closePath: OPS.closePath ?? 18,
    rectangle: OPS.rectangle ?? 19,
    paintImageXObject: OPS.paintImageXObject ?? 85,
    paintImageMaskXObject: OPS.paintImageMaskXObject ?? 83,
    paintImageMaskXObjectGroup: OPS.paintImageMaskXObjectGroup ?? 84,
    paintSolidColorImageMask: OPS.paintSolidColorImageMask ?? 90,
    paintImageXObjectRepeat: OPS.paintImageXObjectRepeat ?? 88,
    paintJpegXObject: OPS.paintJpegXObject ?? 82,
    paintInlineImageXObject: OPS.paintInlineImageXObject ?? 86,
    paintInlineImageXObjectGroup: OPS.paintInlineImageXObjectGroup ?? 87,
    paintFormXObjectBegin: OPS.paintFormXObjectBegin ?? 74,
    paintFormXObjectEnd: OPS.paintFormXObjectEnd ?? 75,
    beginGroup: OPS.beginGroup ?? 77,
    endGroup: OPS.endGroup ?? 78,
  };

  for (let index = 0; index < fnArray.length; index += 1) {
    if (deadline && Date.now() > deadline) {
      console.warn("PDF page rebuild skipped remaining operators after time budget.");
      return;
    }
    if (index && index % 80 === 0) await waitForUiFrame();
    const fn = fnArray[index];
    const args = argsArray[index] || [];
    const graphics = stateStack[stateStack.length - 1];

    if (fn === op.save) {
      stateStack.push(clonePdfGraphicsState(graphics));
      continue;
    }
    if (fn === op.restore) {
      if (stateStack.length > 1) stateStack.pop();
      continue;
    }
    if (fn === op.transform && args.length >= 6) {
      graphics.matrix = multiplyPdfMatrix(graphics.matrix, args);
      continue;
    }
    if (fn === op.dependency) {
      await waitForPdfJsDependencies(sourcePage, args);
      continue;
    }
    if (fn === op.paintFormXObjectBegin || fn === op.beginGroup) {
      const next = clonePdfGraphicsState(graphics);
      const matrix = getPdfFormTransform(args);
      if (matrix) next.matrix = multiplyPdfMatrix(next.matrix, matrix);
      stateStack.push(next);
      continue;
    }
    if (fn === op.paintFormXObjectEnd || fn === op.endGroup) {
      if (stateStack.length > 1) stateStack.pop();
      continue;
    }
    if (fn === op.setLineWidth) {
      graphics.lineWidth = Math.max(0.2, Math.min(8, Number(args[0] || 1)));
      continue;
    }
    if (fn === op.setStrokeRGBColor) {
      graphics.strokeColor = normalizePdfOperatorColor(args, graphics.strokeColor);
      continue;
    }
    if (fn === op.setFillRGBColor) {
      graphics.fillColor = normalizePdfOperatorColor(args, graphics.fillColor);
      continue;
    }
    if (fn === op.setStrokeCMYKColor) {
      graphics.strokeColor = normalizePdfOperatorCmyk(args, graphics.strokeColor);
      continue;
    }
    if (fn === op.setFillCMYKColor) {
      graphics.fillColor = normalizePdfOperatorCmyk(args, graphics.fillColor);
      continue;
    }
    if (fn === op.setStrokeColorN) {
      graphics.strokeColor = normalizePdfOperatorColorN(args, graphics.strokeColor);
      continue;
    }
    if (fn === op.setFillColorN) {
      graphics.fillColor = normalizePdfOperatorColorN(args, graphics.fillColor);
      continue;
    }
    if (fn === op.setStrokeGray) {
      graphics.strokeColor = normalizePdfOperatorGray(args, graphics.strokeColor);
      continue;
    }
    if (fn === op.setFillGray) {
      graphics.fillColor = normalizePdfOperatorGray(args, graphics.fillColor);
      continue;
    }
    if (fn === op.constructPath) {
      currentPath = extractPdfDrawablePath(args, graphics.matrix, op, pageWidth, pageHeight);
      continue;
    }
    if (fn === op.stroke || fn === op.closeStroke || fn === op.fillStroke || fn === op.eoFillStroke || fn === op.closeFillStroke) {
      if (fn === op.fillStroke || fn === op.eoFillStroke || fn === op.closeFillStroke) {
        drawPdfRebuiltPathFills(targetPage, currentPath, graphics, rgb);
      }
      drawPdfRebuiltPathStrokes(targetPage, currentPath, graphics, rgb, options);
      currentPath = null;
      continue;
    }
    if (fn === op.fill || fn === op.eoFill) {
      drawPdfRebuiltPathFills(targetPage, currentPath, graphics, rgb);
      currentPath = null;
      continue;
    }
    if (fn === op.endPath) {
      currentPath = null;
      continue;
    }
    if (fn === op.paintImageXObject || fn === op.paintJpegXObject) {
      await drawPdfRebuiltImage(targetPage, pdfDoc, sourcePage, args[0], graphics.matrix, pageWidth, pageHeight, options);
      continue;
    }
    if (fn === op.paintInlineImageXObject || fn === op.paintImageMaskXObject) {
      await drawPdfRebuiltInlineImage(targetPage, pdfDoc, args[0], graphics.matrix, pageWidth, pageHeight);
      continue;
    }
    if (fn === op.paintImageXObjectRepeat || fn === op.paintInlineImageXObjectGroup || fn === op.paintImageMaskXObjectGroup) {
      await drawPdfRebuiltImageGroup(targetPage, pdfDoc, sourcePage, args, graphics.matrix, pageWidth, pageHeight, options);
      continue;
    }
    if (fn === op.paintSolidColorImageMask) {
      drawPdfRebuiltImageMaskRect(targetPage, graphics.matrix, graphics.fillColor, pageWidth, pageHeight, rgb);
    }
  }
}

function getPdfFormTransform(args) {
  const first = Array.isArray(args) ? args[0] : null;
  if (Array.isArray(first) && first.length >= 6) return first;
  const matrix = Array.isArray(args) && args.length >= 6 ? args.slice(0, 6) : null;
  return matrix && matrix.every((value) => Number.isFinite(Number(value))) ? matrix : null;
}

async function waitForPdfJsDependencies(sourcePage, args) {
  const ids = Array.isArray(args?.[0]) ? args[0] : (Array.isArray(args) ? args : []);
  const waits = ids
    .filter((id) => typeof id === "string")
    .map((id) => getPdfJsObjectData(sourcePage.objs, id, 500).then((data) => data || getPdfJsObjectData(sourcePage.commonObjs, id, 500)));
  if (waits.length) await promiseWithTimeout(Promise.allSettled(waits), 1200, []);
}

function createPdfGraphicsState() {
  return {
    matrix: [1, 0, 0, 1, 0, 0],
    lineWidth: 1,
    strokeColor: { r: 0, g: 0, b: 0 },
    fillColor: { r: 1, g: 1, b: 1 },
  };
}

function clonePdfGraphicsState(state) {
  return {
    matrix: [...state.matrix],
    lineWidth: state.lineWidth,
    strokeColor: { ...state.strokeColor },
    fillColor: { ...state.fillColor },
  };
}

function normalizePdfOperatorColor(args, fallback) {
  const values = Array.isArray(args) ? args : [];
  if (values.length < 3) return fallback || { r: 0, g: 0, b: 0 };
  const max = Math.max(Number(values[0] || 0), Number(values[1] || 0), Number(values[2] || 0));
  const divisor = max > 1 ? 255 : 1;
  return {
    r: clampNumber(Number(values[0] || 0) / divisor, 0, 1),
    g: clampNumber(Number(values[1] || 0) / divisor, 0, 1),
    b: clampNumber(Number(values[2] || 0) / divisor, 0, 1),
  };
}

function normalizePdfOperatorGray(args, fallback) {
  if (!Array.isArray(args) || !args.length) return fallback || { r: 0, g: 0, b: 0 };
  const value = clampNumber(Number(args[0] || 0), 0, 1);
  return { r: value, g: value, b: value };
}

function normalizePdfOperatorCmyk(args, fallback) {
  const values = normalizePdfColorArgs(args);
  if (values.length < 4) return fallback || { r: 0, g: 0, b: 0 };
  const max = Math.max(...values.map((value) => Math.abs(Number(value || 0))));
  const divisor = max > 1 ? (max <= 100 ? 100 : 255) : 1;
  const c = clampNumber(Number(values[0] || 0) / divisor, 0, 1);
  const m = clampNumber(Number(values[1] || 0) / divisor, 0, 1);
  const y = clampNumber(Number(values[2] || 0) / divisor, 0, 1);
  const k = clampNumber(Number(values[3] || 0) / divisor, 0, 1);
  return {
    r: clampNumber(1 - Math.min(1, c + k), 0, 1),
    g: clampNumber(1 - Math.min(1, m + k), 0, 1),
    b: clampNumber(1 - Math.min(1, y + k), 0, 1),
  };
}

function normalizePdfOperatorColorN(args, fallback) {
  const values = normalizePdfColorArgs(args);
  if (values.length >= 4) return normalizePdfOperatorCmyk(values, fallback);
  if (values.length >= 3) return normalizePdfOperatorColor(values, fallback);
  if (values.length === 1) return normalizePdfOperatorGray(values, fallback);
  return fallback || { r: 0, g: 0, b: 0 };
}

function normalizePdfColorArgs(args) {
  if (!Array.isArray(args)) return [];
  if (Array.isArray(args[0])) return args[0].map(Number).filter(Number.isFinite);
  return args.map(Number).filter(Number.isFinite);
}

function extractPdfDrawablePath(args, matrix, op, pageWidth, pageHeight) {
  const pathOps = args?.[0] || [];
  const pathArgs = args?.[1] || [];
  const lines = [];
  const rectangles = [];
  let cursor = 0;
  let current = null;
  let firstPoint = null;

  pathOps.forEach((pathOp) => {
    if (pathOp === op.moveTo) {
      current = transformPdfPoint(pathArgs[cursor], pathArgs[cursor + 1], matrix);
      firstPoint = current;
      cursor += 2;
      return;
    }
    if (pathOp === op.lineTo) {
      const next = transformPdfPoint(pathArgs[cursor], pathArgs[cursor + 1], matrix);
      cursor += 2;
      if (current) lines.push(normalizePdfLineSegment(current, next));
      current = next;
      return;
    }
    if (pathOp === op.rectangle) {
      const x = Number(pathArgs[cursor] || 0);
      const y = Number(pathArgs[cursor + 1] || 0);
      const w = Number(pathArgs[cursor + 2] || 0);
      const h = Number(pathArgs[cursor + 3] || 0);
      cursor += 4;
      const p1 = transformPdfPoint(x, y, matrix);
      const p2 = transformPdfPoint(x + w, y, matrix);
      const p3 = transformPdfPoint(x + w, y + h, matrix);
      const p4 = transformPdfPoint(x, y + h, matrix);
      const bounds = pdfBoundsFromPoints([p1, p2, p3, p4]);
      if (isPdfBoundsOnPage(bounds, pageWidth, pageHeight)) {
        rectangles.push(bounds);
        lines.push(normalizePdfLineSegment(p1, p2), normalizePdfLineSegment(p2, p3), normalizePdfLineSegment(p3, p4), normalizePdfLineSegment(p4, p1));
      }
      current = p1;
      firstPoint = p1;
      return;
    }
    if (pathOp === op.closePath) {
      if (current && firstPoint) lines.push(normalizePdfLineSegment(current, firstPoint));
      current = firstPoint;
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

  return {
    lines: lines.filter((line) => isPdfDrawableLine(line, pageWidth, pageHeight)),
    rectangles: rectangles.filter((rect) => isPdfDrawableRect(rect, pageWidth, pageHeight)),
  };
}

function drawPdfRebuiltPathStrokes(page, path, graphics, rgb, options = {}) {
  if (!path?.lines?.length) return;
  const color = graphics.strokeColor || { r: 0, g: 0, b: 0 };
  const thickness = Math.max(0.2, Math.min(6, Number(graphics.lineWidth || 1)));
  path.lines.forEach((line) => {
    page.drawLine({
      start: { x: line.x1, y: line.y1 },
      end: { x: line.x2, y: line.y2 },
      thickness,
      color: rgb(color.r, color.g, color.b),
      opacity: 1,
    });
    if (Array.isArray(options.rebuiltLines)) {
      options.rebuiltLines.push({ ...line, thickness, color: { ...color } });
    }
  });
}

function drawPdfRebuiltPathFills(page, path, graphics, rgb) {
  if (!path?.rectangles?.length) return;
  const color = graphics.fillColor || { r: 1, g: 1, b: 1 };
  path.rectangles.forEach((rect) => {
    page.drawRectangle({
      x: rect.x,
      y: rect.y,
      width: rect.width,
      height: rect.height,
      color: rgb(color.r, color.g, color.b),
      opacity: 1,
    });
  });
}

function isPdfDrawableLine(line, pageWidth, pageHeight) {
  if (!line) return false;
  const length = Math.hypot(Number(line.x2 || 0) - Number(line.x1 || 0), Number(line.y2 || 0) - Number(line.y1 || 0));
  const horizontal = Math.abs(Number(line.y1 || 0) - Number(line.y2 || 0)) <= 1.5;
  const vertical = Math.abs(Number(line.x1 || 0) - Number(line.x2 || 0)) <= 1.5;
  const insidePage = line.x2 >= -4 && line.y2 >= -4 && line.x1 <= pageWidth + 4 && line.y1 <= pageHeight + 4;
  return insidePage && length >= 3 && (horizontal || vertical);
}

function isPdfDrawableRect(rect, pageWidth, pageHeight) {
  if (!rect) return false;
  const area = Number(rect.width || 0) * Number(rect.height || 0);
  if (area < 16) return false;
  if (rect.width < 1.5 || rect.height < 1.5) return false;
  return isPdfBoundsOnPage(rect, pageWidth, pageHeight);
}

function isPdfBoundsOnPage(bounds, pageWidth, pageHeight) {
  return bounds &&
    bounds.x + bounds.width >= -4 &&
    bounds.y + bounds.height >= -4 &&
    bounds.x <= pageWidth + 4 &&
    bounds.y <= pageHeight + 4;
}

function pdfBoundsFromPoints(points) {
  const xs = points.map((point) => Number(point.x || 0));
  const ys = points.map((point) => Number(point.y || 0));
  const x = Math.min(...xs);
  const y = Math.min(...ys);
  return {
    x,
    y,
    width: Math.max(0, Math.max(...xs) - x),
    height: Math.max(0, Math.max(...ys) - y),
  };
}

async function drawPdfRebuiltImage(targetPage, pdfDoc, sourcePage, objectId, matrix, pageWidth, pageHeight, options = {}) {
  if (!objectId) return;
  const imageData = await getPdfJsObjectData(sourcePage.objs, objectId, 3600) ||
    await getPdfJsObjectData(sourcePage.commonObjs, objectId, 3600);
  const drawn = await drawPdfRebuiltInlineImage(targetPage, pdfDoc, imageData, matrix, pageWidth, pageHeight);
  if (!drawn) {
    await drawPdfRenderedImageFallback(targetPage, pdfDoc, sourcePage, matrix, pageWidth, pageHeight, options);
  }
}

async function drawPdfRebuiltImageGroup(targetPage, pdfDoc, sourcePage, args, matrix, pageWidth, pageHeight, options = {}) {
  const deadline = Number(options.deadline || 0);
  const imageData = args?.[0]?.data ? args[0] : null;
  const objectId = typeof args?.[0] === "string" ? args[0] : null;
  const repeatPositions = Array.isArray(args?.[3]) ? args[3] : null;
  const positions = repeatPositions || (Array.isArray(args?.[1]) ? args[1] : []);
  const scaleX = repeatPositions ? Number(args?.[1] || 1) : 1;
  const scaleY = repeatPositions ? Number(args?.[2] || 1) : 1;
  if (!positions.length) {
    if (objectId) await drawPdfRebuiltImage(targetPage, pdfDoc, sourcePage, objectId, matrix, pageWidth, pageHeight, options);
    else if (imageData) await drawPdfRebuiltInlineImage(targetPage, pdfDoc, imageData, matrix, pageWidth, pageHeight);
    return;
  }

  const maxPositionValues = isLikelyMobileDevice() ? 120 : 240;
  const cappedPositions = positions.slice(0, maxPositionValues);
  for (let index = 0; index < cappedPositions.length; index += 2) {
    if (deadline && Date.now() > deadline) return;
    if (index && index % 40 === 0) await waitForUiFrame();
    const positionedMatrix = multiplyPdfMatrix(matrix, [scaleX, 0, 0, scaleY, Number(cappedPositions[index] || 0), Number(cappedPositions[index + 1] || 0)]);
    if (objectId) await drawPdfRebuiltImage(targetPage, pdfDoc, sourcePage, objectId, positionedMatrix, pageWidth, pageHeight, options);
    else if (imageData) await drawPdfRebuiltInlineImage(targetPage, pdfDoc, imageData, positionedMatrix, pageWidth, pageHeight);
  }
}

function getPdfJsObjectData(store, objectId, timeoutMs = 900) {
  if (!store || !objectId) return Promise.resolve(null);
  return promiseWithTimeout(new Promise((resolve) => {
    try {
      const direct = store.get(objectId);
      if (direct) {
        resolve(direct);
        return;
      }
    } catch (error) {
      // Some pdf.js object stores throw until the callback form is used.
    }

    try {
      store.get(objectId, (data) => resolve(data || null));
    } catch (error) {
      resolve(null);
    }
  }), timeoutMs, null);
}

function promiseWithTimeout(promise, timeoutMs, fallback = null) {
  let timer = null;
  return Promise.race([
    promise,
    new Promise((resolve) => {
      timer = window.setTimeout(() => resolve(fallback), timeoutMs);
    }),
  ]).finally(() => {
    if (timer) window.clearTimeout(timer);
  });
}

async function drawPdfRebuiltInlineImage(targetPage, pdfDoc, imageData, matrix, pageWidth, pageHeight) {
  if (!imageData) return false;
  const bounds = getPdfImageBounds(matrix);
  if (!isPdfBoundsOnPage(bounds, pageWidth, pageHeight) || bounds.width < 2 || bounds.height < 2) return false;
  const pngBytes = await pdfImageDataToPngBytes(imageData);
  if (!pngBytes) return false;
  const image = await pdfDoc.embedPng(pngBytes);
  targetPage.drawImage(image, {
    x: Math.max(0, bounds.x),
    y: Math.max(0, bounds.y),
    width: Math.max(1, Math.min(pageWidth, bounds.width)),
    height: Math.max(1, Math.min(pageHeight, bounds.height)),
  });
  return true;
}

async function drawPdfRenderedImageFallback(targetPage, pdfDoc, sourcePage, matrix, pageWidth, pageHeight, options = {}) {
  const bounds = expandPdfImageFallbackBounds(getPdfImageBounds(matrix), pageWidth, pageHeight);
  if (!isPdfBoundsOnPage(bounds, pageWidth, pageHeight) || bounds.width < 4 || bounds.height < 4) return false;
  const cropBytes = await renderPdfPageCropToPng(sourcePage, bounds, pageWidth, pageHeight, options);
  if (!cropBytes) return false;
  const image = await pdfDoc.embedPng(cropBytes);
  targetPage.drawImage(image, {
    x: Math.max(0, bounds.x),
    y: Math.max(0, bounds.y),
    width: Math.max(1, Math.min(pageWidth, bounds.width)),
    height: Math.max(1, Math.min(pageHeight, bounds.height)),
  });
  return true;
}

function expandPdfImageFallbackBounds(bounds, pageWidth, pageHeight) {
  if (!bounds) return bounds;
  const padX = Math.max(4, Math.min(18, Number(bounds.width || 0) * 0.035));
  const padY = Math.max(3, Math.min(14, Number(bounds.height || 0) * 0.028));
  const x = Math.max(0, Number(bounds.x || 0) - padX);
  const y = Math.max(0, Number(bounds.y || 0) - padY);
  const right = Math.min(pageWidth, Number(bounds.x || 0) + Number(bounds.width || 0) + padX);
  const top = Math.min(pageHeight, Number(bounds.y || 0) + Number(bounds.height || 0) + padY);
  return {
    x,
    y,
    width: Math.max(1, right - x),
    height: Math.max(1, top - y),
  };
}

async function drawPdfPreservedPageFurniture(sourcePage, targetPage, pdfDoc, pagePath, pageSize, options = {}) {
  const pageWidth = Number(pageSize.width || 595.28);
  const pageHeight = Number(pageSize.height || 841.89);
  const regions = getPdfPageFurnitureRegions(pagePath, pageWidth, pageHeight);

  for (const region of regions) {
    const pngBytes = await renderPdfPageCropToPng(sourcePage, region, pageWidth, pageHeight, options);
    if (!pngBytes) continue;
    const image = await pdfDoc.embedPng(pngBytes);
    targetPage.drawImage(image, {
      x: region.x,
      y: region.y,
      width: region.width,
      height: region.height,
    });
  }
}

function getPdfPageFurnitureMetrics(pageHeight) {
  const height = Number(pageHeight || 841.89);
  return {
    topHeight: Math.min(72, Math.max(52, height * 0.085)),
    bottomHeight: Math.min(58, Math.max(42, height * 0.064)),
  };
}

function getDefaultPdfPageFurnitureRegions(pageWidth, pageHeight) {
  const { topHeight, bottomHeight } = getPdfPageFurnitureMetrics(pageHeight);
  return [
    { id: "auto-header", kind: "header", x: 0, y: pageHeight - topHeight, width: pageWidth, height: topHeight },
    { id: "auto-footer", kind: "footer", x: 0, y: 0, width: pageWidth, height: bottomHeight },
  ];
}

function getPdfPageFurnitureRegions(pagePath, pageWidth, pageHeight) {
  const defaults = getDefaultPdfPageFurnitureRegions(pageWidth, pageHeight);
  const manual = getPdfManualFurnitureRegions(pagePath);
  return defaults.map((region) => {
    const replacement = manual.find((item) => item.kind === region.kind);
    return replacement || region;
  });
}

async function drawPdfPreservedRegions(sourcePage, targetPage, pdfDoc, pagePath, pageSize, options = {}) {
  const regions = getPdfPreservedRegionsForPage(pagePath);
  if (!regions.length) return;
  const pageWidth = Number(pageSize.width || 595.28);
  const pageHeight = Number(pageSize.height || 841.89);
  for (const region of regions) {
    const sourceBounds = getPdfRegionSourceBounds(region);
    const targetBounds = getPdfRegionTargetBounds(region);
    const pngBytes = await renderPdfPageCropToPng(sourcePage, sourceBounds, pageWidth, pageHeight, options);
    if (!pngBytes) continue;
    const image = await pdfDoc.embedPng(pngBytes);
    targetPage.drawImage(image, {
      x: targetBounds.x,
      y: targetBounds.y,
      width: targetBounds.width,
      height: targetBounds.height,
    });
  }
}

function getPdfPreservedRegionsForPage(pagePath) {
  const manualRegions = getPdfManualImageRegions(pagePath);
  const automaticRegions = state.pdfPreservedRegions.get(pagePath) || [];
  const effectiveAutomaticRegions = automaticRegions.filter((region) => {
    if (region.kind === "reference") return false;
    if (region.kind !== "figure") return true;
    return !manualRegions.some((manual) => rectsIntersect(getPdfRegionSourceBounds(manual), getPdfRegionSourceBounds(region)));
  });
  return [
    ...effectiveAutomaticRegions,
    ...manualRegions,
  ];
}

function getPdfManualRegions(pagePath) {
  return state.pdfManualPreservedRegions.get(pagePath) || [];
}

function getPdfManualImageRegions(pagePath) {
  return getPdfManualRegions(pagePath).filter((region) => !isPdfPageFurnitureRegion(region));
}

function getPdfManualFurnitureRegions(pagePath) {
  return getPdfManualRegions(pagePath).filter(isPdfPageFurnitureRegion);
}

function isPdfPageFurnitureRegion(region) {
  return region?.kind === "header" || region?.kind === "footer";
}

function getPdfRegionSourceBounds(region) {
  return {
    x: Number(region?.x || 0),
    y: Number(region?.y || 0),
    width: Number(region?.width || 0),
    height: Number(region?.height || 0),
  };
}

function getPdfRegionTargetBounds(region) {
  const target = region?.targetBounds;
  if (target && ["x", "y", "width", "height"].every((key) => Number.isFinite(Number(target[key])))) {
    return {
      x: Number(target.x),
      y: Number(target.y),
      width: Number(target.width),
      height: Number(target.height),
    };
  }
  return getPdfRegionSourceBounds(region);
}

function getPdfSegmentTargetBounds(segment) {
  const bounds = segment?.overrides?.pdfBounds;
  if (bounds && ["x", "y", "width", "height"].every((key) => Number.isFinite(Number(bounds[key])))) {
    return {
      x: Number(bounds.x),
      y: Number(bounds.y),
      width: Math.max(1, Number(bounds.width)),
      height: Math.max(1, Number(bounds.height)),
    };
  }
  return segment?.layout?.bounds || { x: 0, y: 0, width: 1, height: 1 };
}

async function renderPdfPageCropToPng(sourcePage, bounds, pageWidth, pageHeight, options = {}) {
  const scale = isLikelyMobileDevice() ? 1.35 : 1.8;
  const cache = options.renderCache || (options.renderCache = new Map());
  const cacheKey = `${sourcePage.pageNumber || "page"}:${scale}`;
  let rendered = cache.get(cacheKey);
  if (!rendered) {
    const viewport = sourcePage.getViewport({ scale });
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.ceil(viewport.width));
    canvas.height = Math.max(1, Math.ceil(viewport.height));
    const context = canvas.getContext("2d", { willReadFrequently: true });
    if (!context) return null;
    await sourcePage.render({ canvasContext: context, viewport }).promise;
    rendered = { canvas, scale };
    cache.set(cacheKey, rendered);
  }

  const sourceX = clampNumber(Math.floor(bounds.x * rendered.scale), 0, rendered.canvas.width - 1);
  const sourceY = clampNumber(Math.floor((pageHeight - bounds.y - bounds.height) * rendered.scale), 0, rendered.canvas.height - 1);
  const sourceWidth = clampNumber(Math.ceil(bounds.width * rendered.scale), 1, rendered.canvas.width - sourceX);
  const sourceHeight = clampNumber(Math.ceil(bounds.height * rendered.scale), 1, rendered.canvas.height - sourceY);
  const crop = document.createElement("canvas");
  crop.width = sourceWidth;
  crop.height = sourceHeight;
  const cropContext = crop.getContext("2d");
  if (!cropContext) return null;
  cropContext.drawImage(rendered.canvas, sourceX, sourceY, sourceWidth, sourceHeight, 0, 0, sourceWidth, sourceHeight);
  return canvasToArrayBuffer(crop, "image/png");
}

function getPdfImageBounds(matrix) {
  return pdfBoundsFromPoints([
    transformPdfPoint(0, 0, matrix),
    transformPdfPoint(1, 0, matrix),
    transformPdfPoint(1, 1, matrix),
    transformPdfPoint(0, 1, matrix),
  ]);
}

function drawPdfRebuiltImageMaskRect(targetPage, matrix, color, pageWidth, pageHeight, rgb) {
  const bounds = getPdfImageBounds(matrix);
  if (!isPdfBoundsOnPage(bounds, pageWidth, pageHeight) || bounds.width < 0.5 || bounds.height < 0.5) return;
  const fill = color || { r: 0, g: 0, b: 0 };
  targetPage.drawRectangle({
    x: Math.max(0, bounds.x),
    y: Math.max(0, bounds.y),
    width: Math.max(0.5, Math.min(pageWidth, bounds.width)),
    height: Math.max(0.5, Math.min(pageHeight, bounds.height)),
    color: rgb(fill.r, fill.g, fill.b),
    opacity: 1,
  });
}

async function pdfImageDataToPngBytes(imageData) {
  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d");
  if (!context) return null;

  if (imageData instanceof HTMLCanvasElement) {
    canvas.width = Math.max(1, imageData.width);
    canvas.height = Math.max(1, imageData.height);
    context.drawImage(imageData, 0, 0);
    return canvasToArrayBuffer(canvas, "image/png");
  }

  if (typeof ImageBitmap !== "undefined" && imageData instanceof ImageBitmap) {
    canvas.width = Math.max(1, imageData.width);
    canvas.height = Math.max(1, imageData.height);
    context.drawImage(imageData, 0, 0);
    return canvasToArrayBuffer(canvas, "image/png");
  }

  if (typeof HTMLImageElement !== "undefined" && imageData instanceof HTMLImageElement) {
    canvas.width = Math.max(1, imageData.naturalWidth || imageData.width);
    canvas.height = Math.max(1, imageData.naturalHeight || imageData.height);
    context.drawImage(imageData, 0, 0);
    return canvasToArrayBuffer(canvas, "image/png");
  }

  const width = Number(imageData.width || 0);
  const height = Number(imageData.height || 0);
  if (!width || !height) return null;
  canvas.width = Math.max(1, width);
  canvas.height = Math.max(1, height);

  const data = imageData.data || imageData;
  if (data?.length === width * height * 4) {
    context.putImageData(new ImageData(new Uint8ClampedArray(data), width, height), 0, 0);
  } else if (data?.length === width * height * 3) {
    const rgba = new Uint8ClampedArray(width * height * 4);
    for (let source = 0, target = 0; source < data.length; source += 3, target += 4) {
      rgba[target] = data[source];
      rgba[target + 1] = data[source + 1];
      rgba[target + 2] = data[source + 2];
      rgba[target + 3] = 255;
    }
    context.putImageData(new ImageData(rgba, width, height), 0, 0);
  } else if (data?.length === width * height) {
    const rgba = new Uint8ClampedArray(width * height * 4);
    for (let source = 0, target = 0; source < data.length; source += 1, target += 4) {
      rgba[target] = data[source];
      rgba[target + 1] = data[source];
      rgba[target + 2] = data[source];
      rgba[target + 3] = 255;
    }
    context.putImageData(new ImageData(rgba, width, height), 0, 0);
  } else if (data?.length === Math.ceil(width * height / 8)) {
    const rgba = new Uint8ClampedArray(width * height * 4);
    for (let pixel = 0; pixel < width * height; pixel += 1) {
      const byte = data[pixel >> 3] || 0;
      const bit = (byte >> (7 - (pixel & 7))) & 1;
      const value = bit ? 0 : 255;
      const target = pixel * 4;
      rgba[target] = value;
      rgba[target + 1] = value;
      rgba[target + 2] = value;
      rgba[target + 3] = bit ? 255 : 0;
    }
    context.putImageData(new ImageData(rgba, width, height), 0, 0);
  } else {
    return null;
  }

  return canvasToArrayBuffer(canvas, "image/png");
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

function normalizePdfBackgroundColor(color, preserveTint = false) {
  const background = color || { r: 1, g: 1, b: 1 };
  if (preserveTint) return background;

  const luminance = 0.2126 * background.r + 0.7152 * background.g + 0.0722 * background.b;
  const maxChannel = Math.max(background.r, background.g, background.b);
  const minChannel = Math.min(background.r, background.g, background.b);
  const isNearWhite = luminance > 0.965 && minChannel > 0.94 && maxChannel - minChannel < 0.035;
  return isNearWhite ? { r: 1, g: 1, b: 1 } : background;
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

function extractPdfTextEraseFragments(items) {
  return (items || [])
    .map((item) => {
      const text = String(item?.str || "").trim();
      if (!text) return null;

      const transform = item.transform || [];
      const x = Number(transform[4] || 0);
      const y = Number(transform[5] || 0);
      const height = Math.max(6, Math.abs(Number(transform[3] || item.height || 10)));
      const width = Math.max(1, Number(item.width || 0));
      return {
        text,
        bounds: {
          x,
          y: y - height * 0.3,
          width,
          height: Math.max(7, height * 1.42),
        },
      };
    })
    .filter(Boolean);
}

function getPdfReferenceContext(lines, pageHeight, alreadyStarted = false) {
  const heading = lines.find((line) => isPdfReferencesHeading(line.text));

  return {
    active: alreadyStarted || Boolean(heading),
    startedOnPage: Boolean(heading),
    headingY: heading?.bounds?.y ?? (alreadyStarted ? pageHeight + 1 : -1),
  };
}

function isPdfReferencesHeading(text) {
  const normalized = String(text || "").replace(/\s+/g, " ").trim().toLowerCase();
  return /^(references|bibliography|参考文献|參考文獻)$/.test(normalized);
}

function isPdfReferenceLine(line, context, pageWidth, pageHeight) {
  const text = String(line?.text || "").replace(/\s+/g, " ").trim();
  if (!text || isPdfReferencesHeading(text)) return false;
  if (isLikelyPdfReferenceCitationLine(text)) return true;

  const bounds = line?.bounds || {};
  const belowSamePageHeading = context?.startedOnPage && Number(bounds.y || 0) < Number(context.headingY || 0) - 2;
  const rightColumnAfterHeading = context?.startedOnPage && Number(bounds.x || 0) > pageWidth * 0.46;
  const laterReferencePage = context?.active && !context?.startedOnPage && Number(bounds.y || 0) < pageHeight * 0.94;

  return (belowSamePageHeading || rightColumnAfterHeading || laterReferencePage) && isLikelyPdfReferenceContinuation(text);
}

function isLikelyPdfReferenceCitationLine(text) {
  return /^\s*(?:\[\s*\d+\s*\]|\d+\.)\s+/.test(text);
}

function isLikelyPdfReferenceContinuation(text) {
  const normalized = String(text || "").replace(/\s+/g, " ").trim();
  if (!normalized) return false;

  const initialCount = (normalized.match(/\b[A-Z]\./g) || []).length;
  const hasCitationSyntax = /(?:https?:\/\/|doi\.org|doi:|et al\.|\(\d{4}\)|\b\d+\s*\(\d+\)|\b\d{4}\s*[,;]\s*\d+|\b(?:J|Nat|Sci|Cancer|Oncol|Radiol|Med|Clin|Rep|BMJ|PLoS)\.)/i.test(normalized);
  const hasJournalShape = initialCount >= 2 && /[,;]/.test(normalized) && /\b\d{4}\b/.test(normalized);

  return hasCitationSyntax || hasJournalShape;
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
    if (previous && (previous.hasEol || shouldSplitPdfRowGroup(previous, item, gap, gapThreshold, pageWidth))) {
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

function shouldSplitPdfRowGroup(previous, item, gap, gapThreshold, pageWidth) {
  if (gap <= gapThreshold) return false;

  const previousRight = previous.x + previous.width;
  const crossesColumnGutter = previousRight < pageWidth * 0.52 && item.x > pageWidth * 0.46;
  const veryLargeGap = gap > pageWidth * 0.2;
  return crossesColumnGutter || veryLargeGap;
}

function groupPdfLineSegmentsIntoBlocks(lines, pageWidth = 595.28, pageHeight = 841.89) {
  if (!Array.isArray(lines) || !lines.length) return [];

  const ordered = lines
    .filter((line) => line?.bounds && String(line.text || "").trim())
    .map((line, index) => ({
      ...line,
      sourceIndex: index,
      columnKey: getPdfLineColumnKey(line, pageWidth),
    }));
  const columns = new Map();

  ordered.forEach((line) => {
    const bucket = columns.get(line.columnKey) || [];
    bucket.push(line);
    columns.set(line.columnKey, bucket);
  });

  const blocks = [];
  columns.forEach((columnLines) => {
    const sorted = columnLines.sort(comparePdfLinesTopDown);
    let current = null;

    sorted.forEach((line) => {
      if (!current || shouldStartNewPdfTextBlock(current, line, pageWidth, pageHeight)) {
        if (current) blocks.push(finalizePdfTextBlock(current, pageWidth));
        current = createPdfTextBlock(line);
        return;
      }

      mergePdfLineIntoBlock(current, line);
    });

    if (current) blocks.push(finalizePdfTextBlock(current, pageWidth));
  });

  return blocks.sort(comparePdfLinesTopDown);
}

function comparePdfLinesTopDown(a, b) {
  const ay = Number(a.bounds?.y || 0);
  const by = Number(b.bounds?.y || 0);
  if (Math.abs(by - ay) > 2) return by - ay;
  return Number(a.bounds?.x || 0) - Number(b.bounds?.x || 0);
}

function getPdfLineColumnKey(line, pageWidth) {
  const bounds = line?.bounds || {};
  const x = Number(bounds.x || 0);
  const width = Number(bounds.width || 0);
  const center = x + width / 2;

  if (x < pageWidth * 0.23 && width < pageWidth * 0.26) return "side";
  if (width >= pageWidth * 0.58) return "full";
  if (x <= pageWidth * 0.16 && width >= pageWidth * 0.46) return "full";
  if (center < pageWidth * 0.52) return "left";
  return "right";
}

function createPdfTextBlock(line) {
  return {
    columnKey: line.columnKey,
    lines: [line],
    text: String(line.text || "").trim(),
    fontSizes: [Number(line.fontSize || 10)],
    bounds: { ...line.bounds },
    availableWidth: Number(line.availableWidth || line.bounds?.width || 0),
    availableHeight: Number(line.availableHeight || line.bounds?.height || 0),
    rowSegmentCount: 1,
    sourceIndex: line.sourceIndex,
  };
}

function shouldStartNewPdfTextBlock(block, line, pageWidth, pageHeight) {
  if (!block?.lines?.length) return true;
  if (line.columnKey !== block.columnKey) return true;
  if (shouldMergePdfTitleLine(block, line, pageWidth, pageHeight)) return false;
  if (isPdfStandaloneTextLine(line, pageWidth, pageHeight)) return true;

  const previous = block.lines[block.lines.length - 1];
  if (isPdfStandaloneTextLine(previous, pageWidth, pageHeight)) return true;
  const previousBounds = previous.bounds || {};
  const lineBounds = line.bounds || {};
  const previousFont = Math.max(4, Number(previous.fontSize || previousBounds.height || 10));
  const lineFont = Math.max(4, Number(line.fontSize || lineBounds.height || 10));
  const fontDelta = Math.abs(previousFont - lineFont);
  if (fontDelta > Math.max(1.4, Math.min(previousFont, lineFont) * 0.18)) return true;

  const baselineGap = Number(previousBounds.y || 0) - Number(lineBounds.y || 0);
  const expectedLineGap = Math.max(previousFont, lineFont) * 1.65;
  if (baselineGap <= 0 || baselineGap > expectedLineGap) return true;

  const indentDelta = Math.abs(Number(lineBounds.x || 0) - Number(block.lines[0].bounds?.x || 0));
  const currentLineCount = block.lines.length;
  if (indentDelta > Math.max(18, lineFont * 2.4) && currentLineCount > 1) return true;
  if (currentLineCount >= 10) return true;

  const columnWidth = estimatePdfColumnWidth([...block.lines, line], pageWidth);
  const previousWidth = Number(previousBounds.width || 0);
  const previousText = String(previous.text || "").trim();
  const lineText = String(line.text || "").trim();
  const previousLooksFinal = previousWidth < columnWidth * 0.74 && /[.!?;:。！？；：)]$/.test(previousText);
  if (previousLooksFinal && !startsWithPdfContinuationText(lineText)) return true;

  if (looksLikePdfSectionHeading(lineText, line, pageWidth)) return true;
  return false;
}

function shouldMergePdfTitleLine(block, line, pageWidth, pageHeight) {
  const previous = block.lines[block.lines.length - 1];
  if (!isPdfMainTitleLine(previous, pageWidth, pageHeight) || !isPdfMainTitleLine(line, pageWidth, pageHeight)) return false;
  const previousBounds = previous.bounds || {};
  const lineBounds = line.bounds || {};
  const previousFont = Number(previous.fontSize || previousBounds.height || 0);
  const lineFont = Number(line.fontSize || lineBounds.height || 0);
  const baselineGap = Number(previousBounds.y || 0) - Number(lineBounds.y || 0);
  const xDelta = Math.abs(Number(previousBounds.x || 0) - Number(lineBounds.x || 0));
  const fontDelta = Math.abs(previousFont - lineFont);
  return baselineGap > 0 &&
    baselineGap < Math.max(previousFont, lineFont) * 1.75 &&
    xDelta < Math.max(18, lineFont * 1.2) &&
    fontDelta < Math.max(2.2, Math.min(previousFont, lineFont) * 0.18) &&
    block.lines.length < 9;
}

function isPdfMainTitleLine(line, pageWidth, pageHeight) {
  const bounds = line?.bounds || {};
  const text = String(line?.text || "").replace(/\s+/g, " ").trim();
  if (!text || text.length < 8) return false;
  const y = Number(bounds.y || 0);
  const x = Number(bounds.x || 0);
  const width = Number(bounds.width || 0);
  const fontSize = Number(line?.fontSize || bounds.height || 0);
  return fontSize >= 14 &&
    y > pageHeight * 0.48 &&
    y < pageHeight * 0.9 &&
    x > pageWidth * 0.16 &&
    width > pageWidth * 0.22;
}

function isPdfStandaloneTextLine(line, pageWidth, pageHeight) {
  const text = String(line?.text || "").replace(/\s+/g, " ").trim();
  if (!text) return true;
  if (line?.tableCell) return true;

  const bounds = line?.bounds || {};
  const y = Number(bounds.y || 0);
  const fontSize = Number(line?.fontSize || bounds.height || 10);
  if (fontSize >= 13.5) return true;
  if (y > pageHeight * 0.94 || y < pageHeight * 0.04) return true;
  if (/^\d+$/.test(text) && text.length <= 3) return true;

  return looksLikePdfSectionHeading(text, line, pageWidth);
}

function looksLikePdfSectionHeading(text, line, pageWidth) {
  const normalized = String(text || "").replace(/\s+/g, " ").trim();
  if (!normalized) return false;
  if (normalized.length > 46) return false;

  const bounds = line?.bounds || {};
  const width = Number(bounds.width || 0);
  const fontSize = Number(line?.fontSize || bounds.height || 10);
  const mostlyCaps = /^[A-Z0-9\s,/&().:-]+$/.test(normalized) && /[A-Z]/.test(normalized);
  const shortLabel = /^[A-Z][A-Z\s,/&-]{2,28}:?$/.test(normalized);

  if ((mostlyCaps || shortLabel) && width < pageWidth * 0.36) return true;
  if (fontSize >= 11.5 && normalized.length <= 34 && !/[.!?。！？]$/.test(normalized)) return true;
  return false;
}

function startsWithPdfContinuationText(text) {
  return /^(?:and|or|of|in|to|for|with|from|by|on|at|vs\.?|versus|than|that|which|where|when|while|because|as|the|a|an)\b/i.test(text) ||
    /^[,;:)\]]/.test(text);
}

function mergePdfLineIntoBlock(block, line) {
  block.lines.push(line);
  block.text = joinPdfBlockLineText(block.text, line.text);
  block.fontSizes.push(Number(line.fontSize || 10));
  block.bounds = unionPdfBounds(block.bounds, line.bounds);
  block.availableWidth = Math.max(block.availableWidth, Number(line.availableWidth || line.bounds?.width || 0));
  block.availableHeight = Math.max(
    block.availableHeight,
    Math.max(0, block.bounds.height),
    block.lines.length * Math.max(...block.fontSizes) * 1.34
  );
  block.rowSegmentCount = 1;
}

function joinPdfBlockLineText(existing, next) {
  const left = String(existing || "").trim();
  const right = String(next || "").trim();
  if (!left) return right;
  if (!right) return left;
  if (/-$/.test(left) && /^[A-Za-z]/.test(right)) return `${left.slice(0, -1)}${right}`;
  return `${left} ${right}`.replace(/\s+/g, " ").trim();
}

function unionPdfBounds(a, b) {
  const ax = Number(a?.x || 0);
  const ay = Number(a?.y || 0);
  const bx = Number(b?.x || 0);
  const by = Number(b?.y || 0);
  const minX = Math.min(ax, bx);
  const minY = Math.min(ay, by);
  const maxX = Math.max(ax + Number(a?.width || 0), bx + Number(b?.width || 0));
  const maxY = Math.max(ay + Number(a?.height || 0), by + Number(b?.height || 0));
  return {
    x: minX,
    y: minY,
    width: Math.max(1, maxX - minX),
    height: Math.max(1, maxY - minY),
  };
}

function finalizePdfTextBlock(block, pageWidth) {
  const lineWidths = block.lines
    .map((line) => Number(line.bounds?.width || 0))
    .filter((width) => width > 0)
    .sort((a, b) => a - b);
  const fontSizes = block.fontSizes.filter((size) => Number.isFinite(size) && size > 0).sort((a, b) => a - b);
  const medianFontSize = fontSizes.length ? fontSizes[Math.floor(fontSizes.length / 2)] : 10;
  const columnWidth = estimatePdfColumnWidth(block.lines, pageWidth);
  const widthP85 = lineWidths.length ? lineWidths[Math.min(lineWidths.length - 1, Math.floor(lineWidths.length * 0.85))] : block.bounds.width;
  const targetWidth = Math.max(block.bounds.width, widthP85, columnWidth * 0.82);
  const rightLimit = block.columnKey === "side" ? pageWidth * 0.25 : (block.columnKey === "left" ? pageWidth * 0.51 : pageWidth - 24);
  const boundedWidth = block.columnKey === "full"
    ? Math.min(targetWidth, pageWidth - block.bounds.x - 24)
    : Math.min(targetWidth, rightLimit - block.bounds.x);

  return {
    text: block.text,
    fontSize: medianFontSize,
    availableWidth: Math.max(8, boundedWidth),
    availableHeight: Math.max(block.availableHeight, block.bounds.height),
    rowSegmentCount: block.rowSegmentCount,
    lineCount: block.lines.length,
    columnKey: block.columnKey,
    bounds: {
      ...block.bounds,
      width: Math.max(block.bounds.width, Math.min(Math.max(8, boundedWidth), pageWidth - block.bounds.x - 18)),
    },
  };
}

function detectPdfPreservedFigureRegions(page, blocks, pageWidth, pageHeight) {
  const captions = (blocks || []).filter((block) => isPdfFigureCaptionText(block.text));
  if (!captions.length) return [];

  return captions.map((caption) => {
    const captionBounds = caption.bounds || {};
    const captionY = Number(captionBounds.y || 0);
    const captionX = Number(captionBounds.x || 0);
    const captionWidth = Number(captionBounds.width || 0);
    const nearbyTextBounds = (blocks || [])
      .filter((block) => block !== caption && isLikelyPdfFigureInnerLabel(block.text, block.bounds, caption.bounds, pageWidth, pageHeight))
      .map((block) => block.bounds);
    const nearbyRegions = nearbyTextBounds.filter((bounds) => {
      const verticallyAboveCaption = Number(bounds.y || 0) >= captionY + Number(captionBounds.height || 0) * 0.25;
      const closeEnough = Number(bounds.y || 0) <= captionY + Math.max(260, pageHeight * 0.38);
      const horizontalOverlap = rectHorizontalOverlap(bounds, {
        x: Math.max(0, captionX - pageWidth * 0.12),
        width: Math.min(pageWidth, captionWidth + pageWidth * 0.24),
      }) > Math.min(Number(bounds.width || 0), Math.max(1, captionWidth)) * 0.16;
      return verticallyAboveCaption && closeEnough && horizontalOverlap;
    });

    const union = nearbyRegions.reduce((area, bounds) => area ? unionPdfBounds(area, bounds) : { ...bounds }, null);
    if (!union) return null;
    const region = unionPdfBounds(union, captionBounds);
    const padX = Math.max(10, pageWidth * 0.025);
    const padY = Math.max(8, pageHeight * 0.014);
    return {
      x: Math.max(0, region.x - padX),
      y: Math.max(0, region.y - padY),
      width: Math.min(pageWidth, region.x + region.width + padX) - Math.max(0, region.x - padX),
      height: Math.min(pageHeight, region.y + region.height + padY) - Math.max(0, region.y - padY),
      kind: "figure",
    };
  }).filter((region) => region && region.width > 24 && region.height > 24);
}

function detectPdfPreservedReferenceRegions(blocks, referenceContext, pageWidth, pageHeight) {
  if (!referenceContext?.active) return [];

  const referenceBlocks = (blocks || []).filter((block) =>
    isPdfReferencesHeading(block.text) ||
    isPdfReferenceLine(block, referenceContext, pageWidth, pageHeight)
  );
  if (!referenceBlocks.length) return [];

  const groups = new Map();
  referenceBlocks.forEach((block) => {
    const bounds = block.bounds || {};
    const cx = Number(bounds.x || 0) + Number(bounds.width || 0) / 2;
    const key = cx < pageWidth * 0.44 ? "left" : (cx > pageWidth * 0.56 ? "right" : "full");
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(block);
  });

  return [...groups.entries()].map(([key, group]) => {
    const union = group.reduce((area, block) => area ? unionPdfBounds(area, block.bounds) : { ...block.bounds }, null);
    if (!union) return null;
    const column = getPdfReferenceColumnBounds(key, union, pageWidth);
    const padX = Math.max(5, pageWidth * 0.008);
    const padTop = Math.max(10, pageHeight * 0.014);
    const padBottom = Math.max(10, pageHeight * 0.014);
    const y = Math.max(0, union.y - padBottom);
    const top = Math.min(pageHeight, union.y + union.height + padTop);
    return {
      x: Math.max(0, column.x - padX),
      y,
      width: Math.min(pageWidth, column.x + column.width + padX) - Math.max(0, column.x - padX),
      height: Math.max(1, top - y),
      kind: "reference",
    };
  }).filter((region) => region && region.width > 24 && region.height > 24);
}

function getPdfReferenceColumnBounds(key, union, pageWidth) {
  if (key === "left") {
    return { x: Math.max(0, pageWidth * 0.055), width: Math.max(1, pageWidth * 0.44) };
  }
  if (key === "right") {
    const x = Math.max(0, pageWidth * 0.505);
    return { x, width: Math.max(1, pageWidth * 0.45) };
  }
  return {
    x: Math.max(0, Number(union.x || 0) - pageWidth * 0.015),
    width: Math.min(pageWidth, Number(union.width || 0) + pageWidth * 0.03),
  };
}

function isLikelyPdfFigureInnerLabel(text, bounds, captionBounds, pageWidth, pageHeight) {
  const normalized = String(text || "").replace(/\s+/g, " ").trim();
  if (!normalized || !bounds || !captionBounds) return false;
  if (isPdfFigureCaptionText(normalized)) return false;
  const y = Number(bounds.y || 0);
  const captionY = Number(captionBounds.y || 0);
  if (y <= captionY) return false;
  if (Number(bounds.height || 0) > pageHeight * 0.14) return false;
  if (Number(bounds.width || 0) > pageWidth * 0.34) return false;
  return /^[\d(). -]+$/.test(normalized) ||
    /(?:ultrasound|generator|meter|thermometer|electrode|muscle|transducer|spacer|probe|distance|surface|center|NO[- ]?\d+|DT[- ]?\d+)/i.test(normalized);
}

function rectHorizontalOverlap(a, b) {
  const left = Math.max(Number(a.x || 0), Number(b.x || 0));
  const right = Math.min(Number(a.x || 0) + Number(a.width || 0), Number(b.x || 0) + Number(b.width || 0));
  return Math.max(0, right - left);
}

function isPdfFigureCaptionText(text) {
  return /^\s*(?:fig\.?|figure)\s*\d+[.:]/i.test(String(text || "").trim());
}

function estimatePdfColumnWidth(lines, pageWidth) {
  const widths = lines
    .map((line) => Number(line.bounds?.width || 0))
    .filter((width) => width > 0)
    .sort((a, b) => a - b);
  const widest = widths.length ? widths[widths.length - 1] : pageWidth * 0.42;
  const first = lines[0];
  const columnKey = first?.columnKey || getPdfLineColumnKey(first, pageWidth);

  if (columnKey === "full") return Math.min(pageWidth - 48, Math.max(widest, Number(first?.availableWidth || 0)));
  if (columnKey === "side") return Math.min(pageWidth * 0.22, Math.max(widest, pageWidth * 0.16));
  return Math.min(pageWidth * 0.46, Math.max(widest, pageWidth * 0.36));
}

async function loadWordDocument(loadToken = state.loadToken) {
  const wordFiles = Object.keys(state.zip.files)
    .filter((path) => wordPathPattern.test(path))
    .sort((a, b) => a.localeCompare(b));

  if (!wordFiles.includes("word/document.xml")) {
    throw new Error("未找到 Word 正文 document.xml，请确认这是有效的 DOCX 文件。");
  }

  state.slideCount = wordFiles.length;

  for (const path of wordFiles) {
    if (state.loadToken !== loadToken) return;
    const xmlText = await state.zip.file(path).async("text");
    if (state.loadToken !== loadToken) return;
    const doc = parser.parseFromString(xmlText, "application/xml");
    const parseError = doc.querySelector("parsererror");
    if (parseError) {
      throw new Error(`${path} XML 解析失败。`);
    }

    const partLabel = getWordPartLabel(path);
    const paragraphs = [...doc.getElementsByTagNameNS(WORD_NS, "p")];
    paragraphs.forEach((paragraph, paragraphIndex) => {
      const textNodes = [...paragraph.getElementsByTagNameNS(WORD_NS, "t")];
      const rawOriginal = textNodes.map((node) => node.textContent || "").join("");
      const original = rawOriginal.trim();
      if (shouldSkipWordParagraph(original)) return;
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

async function handleResetButtonClick() {
  if (els.resetButton.disabled) return;

  state.batchCancelRequested = true;
  cancelTranslation();
  els.resetButton.disabled = true;

  try {
    await resetApp(true, { message: "已清空当前内容。", toast: true });
  } finally {
    els.resetButton.disabled = false;
  }
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

  if (!model) {
    showToast("请先填写模型。API Key 可留空使用 Cloudflare 后端配置。", true);
    return;
  }

  return translateAllConcurrent({ apiKey, apiBase, apiProxy, model, direction });

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
        segment,
        text: segment.original,
        signal: abortController.signal,
      });
      let shortenAttempts = 0;
      while (shouldShortenPptSingleLineTranslation(segment, direction) && shortenAttempts < 2) {
        setStatus(`正在压缩标题译文 ${segment.locationLabel || segment.slideNumber}：${index + 1}/${untranslated.length}`);
        const shortened = await translateText({
          apiBase,
          apiProxy,
          apiKey,
          model,
          direction,
          segment,
          text: buildPptShorteningRequest(segment, shortenAttempts),
          mode: shortenAttempts === 0 ? "shorten" : "ultraShort",
          signal: abortController.signal,
        });
        if (shortened && estimatePptTextWidthPt(shortened, getPptSegmentFontPt(segment)) <= estimatePptTextWidthPt(segment.translation, getPptSegmentFontPt(segment))) {
          segment.translation = shortened;
        } else {
          break;
        }
        shortenAttempts += 1;
      }
      let boxFitAttempts = 0;
      while (shouldShortenPptBoxTranslation(segment, direction) && boxFitAttempts < 2) {
        setStatus(`正在压缩文本框译文 ${segment.locationLabel || segment.slideNumber}：${index + 1}/${untranslated.length}`);
        const shortened = await translateText({
          apiBase,
          apiProxy,
          apiKey,
          model,
          direction,
          segment,
          text: buildPptBoxFitRequest(segment, boxFitAttempts),
          mode: boxFitAttempts === 0 ? "boxFit" : "boxFitStrict",
          signal: abortController.signal,
        });
        if (shortened && isPptBoxFitTranslationBetter(segment, shortened)) {
          segment.translation = normalizeTranslation(shortened, segment.original, segment);
        } else {
          break;
        }
        boxFitAttempts += 1;
      }
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

async function translateAllConcurrent({ apiKey, apiBase, apiProxy, model, direction }) {
  if (state.translationRunning) return false;
  const abortController = new AbortController();
  state.translationRunning = true;
  state.translationAbortController = abortController;

  try {
    setBusy(true, "正在翻译...");
    const untranslated = state.segments.filter((segment) => !segment.translation.trim());
    const total = untranslated.length;
    const concurrency = Math.min(total || 1, getTranslationConcurrency());
    let cursor = 0;
    let completed = 0;
    let failed = 0;

    const workers = Array.from({ length: concurrency }, async () => {
      while (cursor < total) {
        if (abortController.signal.aborted) throw new DOMException("Translation stopped", "AbortError");
        const index = cursor;
        cursor += 1;
        const segment = untranslated[index];
        setStatus(`正在并行翻译：${completed + failed + 1}/${total}，并发 ${concurrency} 路...`);
        try {
          await translateSegmentWithLayoutFit({
            apiBase,
            apiProxy,
            apiKey,
            model,
            direction,
            segment,
            index,
            total,
            signal: abortController.signal,
          });
          segment.translationError = "";
          completed += 1;
        } catch (error) {
          if (error?.name === "AbortError") throw error;
          failed += 1;
          segment.translationError = error?.message || "翻译失败";
          console.warn("Segment translation failed", {
            index,
            label: segment.locationLabel || segment.slideNumber,
            error,
          });
        }
        setProgress(total ? (completed + failed) / total : 1);
        updateTextarea(segment);
        updateStats();
        scheduleCurrentDraftSave();
      }
    });

    await Promise.all(workers);
    setProgress(1);
    const qualityWarning = getPdfTargetLanguageQualityWarning(direction);
    if (failed) {
      const message = `翻译完成 ${completed}/${total} 段，${failed} 段请求超时或失败；可再次点击自动翻译续翻失败段。`;
      setStatus(message);
      showToast(message, true);
    } else if (qualityWarning) {
      setStatus(qualityWarning);
      showToast(qualityWarning, true);
    } else {
      setStatus("翻译完成，请检查译文后导出。");
      showToast("自动翻译完成。");
    }
    scheduleCurrentDraftSave({ immediate: true });
    return !qualityWarning && failed === 0;
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

async function translateSegmentWithLayoutFit({ apiBase, apiProxy, apiKey, model, direction, segment, index, total, signal }) {
  segment.translation = await translateText({
    apiBase,
    apiProxy,
    apiKey,
    model,
    direction,
    segment,
    text: segment.original,
    signal,
  });

  let shortenAttempts = 0;
  while (shouldShortenPptSingleLineTranslation(segment, direction) && shortenAttempts < 2) {
    setStatus(`正在压缩标题译文 ${segment.locationLabel || segment.slideNumber}：${index + 1}/${total}`);
    const shortened = await translateText({
      apiBase,
      apiProxy,
      apiKey,
      model,
      direction,
      segment,
      text: buildPptShorteningRequest(segment, shortenAttempts),
      mode: shortenAttempts === 0 ? "shorten" : "ultraShort",
      signal,
    });
    if (shortened && estimatePptTextWidthPt(shortened, getPptSegmentFontPt(segment)) <= estimatePptTextWidthPt(segment.translation, getPptSegmentFontPt(segment))) {
      segment.translation = shortened;
    } else {
      break;
    }
    shortenAttempts += 1;
  }

  let boxFitAttempts = 0;
  while (shouldShortenPptBoxTranslation(segment, direction) && boxFitAttempts < 2) {
    setStatus(`正在压缩文本框译文 ${segment.locationLabel || segment.slideNumber}：${index + 1}/${total}`);
    const shortened = await translateText({
      apiBase,
      apiProxy,
      apiKey,
      model,
      direction,
      segment,
      text: buildPptBoxFitRequest(segment, boxFitAttempts),
      mode: boxFitAttempts === 0 ? "boxFit" : "boxFitStrict",
      signal,
    });
    if (shortened && isPptBoxFitTranslationBetter(segment, shortened)) {
      segment.translation = normalizeTranslation(shortened, segment.original, segment);
    } else {
      break;
    }
    boxFitAttempts += 1;
  }
}

function getTranslationConcurrency() {
  if (state.fileType === "pdf") return 5;
  if (state.fileType === "docx") return 5;
  if (state.fileType === "pptx") return 3;
  return 4;
}

async function translateText({ apiBase, apiProxy, apiKey, model, direction, segment = null, text, mode = "translate", signal }) {
  const result = await requestAiTextWithRetry({
    apiBase,
    apiProxy,
    apiKey,
    model,
    instruction: buildSegmentTranslationInstruction(direction, segment, mode),
    text,
    task: "translate",
    signal,
  });

  let normalized = normalizeTranslation(result, text, segment);
  if (shouldRetryTranslationForTarget(normalized, text, direction, segment)) {
    const retry = await requestAiTextWithRetry({
      apiBase,
      apiProxy,
      apiKey,
      model,
      instruction: buildTargetEnforcementInstruction(direction, segment),
      text,
      task: "translate",
      signal,
    });
    normalized = normalizeTranslation(retry, text, segment);
  }

  return normalized;
}

async function requestAiTextWithRetry(options) {
  let lastError = null;
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      return await requestAiText(options);
    } catch (error) {
      if (error?.name === "AbortError") throw error;
      lastError = error;
      if (!isRetriableTranslationError(error) || attempt === 1) break;
      await delay(TRANSLATION_RETRY_DELAY_MS);
    }
  }
  throw lastError || new Error("翻译请求失败。");
}

function isRetriableTranslationError(error) {
  const status = Number(error?.status || 0);
  if (status && ![408, 429, 500, 502, 503, 504].includes(status)) return false;
  return true;
}

async function requestAiText({ apiBase, apiProxy, apiKey, model, instruction, text, task = "translate", signal }) {
  const timeoutController = new AbortController();
  const timeoutId = window.setTimeout(() => {
    timeoutController.abort(new DOMException("Translation request timed out", "TimeoutError"));
  }, TRANSLATION_REQUEST_TIMEOUT_MS);
  const requestSignal = mergeAbortSignals(signal, timeoutController.signal);

  const response = await fetch(apiProxy, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    signal: requestSignal,
    body: JSON.stringify({
      apiBase,
      apiKey,
      model,
      task,
      instruction,
      text,
    }),
  }).catch((error) => {
    if (timeoutController.signal.aborted && !signal?.aborted) {
      const timeoutError = new Error("单段翻译请求超过 45 秒，已跳过并继续后续段落。");
      timeoutError.name = "TimeoutError";
      timeoutError.status = 408;
      throw timeoutError;
    }
    throw error;
  }).finally(() => {
    window.clearTimeout(timeoutId);
  });

  if (!response.ok) {
    const detail = await response.json().catch(() => null);
    const error = new Error(detail?.error || `接口返回 ${response.status}`);
    error.status = response.status;
    throw error;
  }

  const data = await response.json();
  return data.translation || "";
}

function mergeAbortSignals(...signals) {
  const activeSignals = signals.filter(Boolean);
  if (!activeSignals.length) return undefined;
  if (activeSignals.length === 1) return activeSignals[0];
  const controller = new AbortController();
  const abort = (event) => {
    if (controller.signal.aborted) return;
    controller.abort(event?.target?.reason);
  };
  activeSignals.forEach((signal) => {
    if (signal.aborted) {
      abort({ target: signal });
    } else {
      signal.addEventListener("abort", abort, { once: true });
    }
  });
  return controller.signal;
}

async function summarizeDocument() {
  if (state.summaryRunning || !state.segments.length) return;

  const apiKey = els.apiKey.value.trim();
  const apiBase = DEEPSEEK_API_BASE;
  const apiProxy = TRANSLATE_PROXY;
  const model = els.modelName.value.trim();
  const detail = els.summaryDetail?.value || "standard";

  if (!model) {
    showToast("请先填写模型。API Key 可留空使用 Cloudflare 后端配置。", true);
    return;
  }

  const sourceText = buildDocumentSummarySource();
  if (!sourceText.trim()) {
    showToast("当前文档没有可总结的文本。", true);
    return;
  }

  const cacheKey = await buildSummaryCacheKey(sourceText, detail);
  const cached = await getCachedSummary(cacheKey).catch((error) => {
    console.warn("Summary cache read failed", error);
    return null;
  });
  if (cached?.summary) {
    openSummary(cached.summary, detail, { cached: true, savedAt: cached.savedAt });
    setStatus("已打开上次缓存的文献总结。");
    showToast("已打开缓存总结，无需重新 AI 总结。");
    return;
  }

  state.summaryRunning = true;
  setBusy(true, "正在总结文献主要内容...");
  setProgress(0.08);

  try {
    setProgress(0.24);
    const summary = await requestAiText({
      apiBase,
      apiProxy,
      apiKey,
      model,
      task: "summary",
      instruction: buildDocumentSummaryInstruction(detail),
      text: sourceText,
    });

    setProgress(1);
    const cleanedSummary = cleanSummaryText(summary.trim() || "未生成总结。");
    await putCachedSummary({
      id: cacheKey,
      summary: cleanedSummary,
      detail,
      fileName: state.file?.name || "",
      fileSize: state.file?.size || 0,
      segmentCount: state.segments.length,
      savedAt: new Date().toISOString(),
    }).catch((error) => console.warn("Summary cache write failed", error));
    openSummary(cleanedSummary, detail, { cached: false });
    setStatus("文献总结完成。");
    showToast("文献主要内容总结完成。");
  } catch (error) {
    showToast(error.message || "文献总结失败。", true);
    setStatus("文献总结失败，请检查接口配置或网络连接。");
  } finally {
    state.summaryRunning = false;
    setBusy(false);
  }
}

function buildDocumentSummarySource() {
  const chunks = state.segments
    .map((segment) => {
      const label = segment.locationLabel || `第 ${segment.slideNumber || ""} 页`;
      const text = String(segment.original || "").replace(/\s+/g, " ").trim();
      return text ? `[${label}] ${text}` : "";
    })
    .filter(Boolean);

  return trimDocumentForSummary(chunks.join("\n"));
}

function trimDocumentForSummary(text) {
  const normalized = String(text || "").replace(/\n{3,}/g, "\n\n").trim();
  const maxCharacters = 42000;
  if ([...normalized].length <= maxCharacters) return normalized;

  const chars = [...normalized];
  const headLength = Math.floor(maxCharacters * 0.42);
  const middleLength = Math.floor(maxCharacters * 0.22);
  const tailLength = maxCharacters - headLength - middleLength;
  const middleStart = Math.max(headLength, Math.floor((chars.length - middleLength) / 2));
  return [
    chars.slice(0, headLength).join(""),
    "\n\n[中间内容节选]\n",
    chars.slice(middleStart, middleStart + middleLength).join(""),
    "\n\n[末尾内容节选]\n",
    chars.slice(-tailLength).join(""),
  ].join("");
}

function buildDocumentSummaryInstruction(detail) {
  const detailMap = {
    brief: "输出简要总结，约 5-8 条要点，重点说明主题、目的、核心结论和适用场景。",
    standard: "输出标准结构化总结，包含：一句话概述、背景/目的、主要方法或结构、关键发现/主张、重要数据/术语、应用价值、注意事项。",
    detailed: "输出详细总结，包含：一句话概述、文献结构、逐部分要点、关键术语解释、主要数据/专利权利要求或技术方案、创新点、局限/风险、可执行建议。必要时使用小标题和项目符号。",
  };

  return [
    "你是严谨的文献阅读助手，请用简体中文总结用户提供的 PDF、Word 或 PPT 文档主要内容。",
    detailMap[detail] || detailMap.standard,
    "保留关键数字、年份、专利号、产品名、机构名、术语缩写和单位。",
    "如果来源是专利，请重点提炼技术问题、解决方案、系统/方法构成、核心权利要求、应用场景和潜在价值。",
    "如果来源是论文或报告，请重点提炼研究目的、方法、结果、结论和局限。",
    "不要编造未出现的信息；不确定时标注“原文未明确”。",
    "直接输出总结正文，不要说你无法访问文件。",
  ].join(" ");
}

function openSummary(summary, detail, options = {}) {
  if (!els.summaryDialog || !els.summaryOutput) return;
  const cleaned = cleanSummaryText(summary);
  state.activeSummary = {
    text: cleaned,
    detail,
    cached: Boolean(options.cached),
    savedAt: options.savedAt || "",
  };
  els.summaryOutput.textContent = cleaned;
  if (els.summaryMeta) {
    const labels = { brief: "简要", standard: "标准", detailed: "详细" };
    const cacheLabel = options.cached ? " · 已缓存" : "";
    els.summaryMeta.textContent = `${getFileTypeName()} · ${state.segments.length} 段文本 · ${labels[detail] || "标准"}总结${cacheLabel}`;
  }
  if (typeof els.summaryDialog.showModal === "function") {
    els.summaryDialog.showModal();
  } else {
    els.summaryDialog.setAttribute("open", "");
  }
}

function cleanSummaryText(text) {
  return String(text || "")
    .replace(/\*+/g, "")
    .replace(/#+/g, "")
    .replace(/[ \t]+\n/g, "\n")
    .trim();
}

function closeSummary() {
  if (!els.summaryDialog) return;
  if (typeof els.summaryDialog.close === "function") {
    els.summaryDialog.close();
  } else {
    els.summaryDialog.removeAttribute("open");
  }
}

async function copySummary() {
  const text = els.summaryOutput?.textContent || "";
  if (!text.trim()) return;
  try {
    await navigator.clipboard.writeText(text);
    showToast("总结已复制。");
  } catch {
    showToast("当前浏览器不允许自动复制，请手动选择文本复制。", true);
  }
}

async function shareSummary() {
  const text = els.summaryOutput?.textContent || state.activeSummary?.text || "";
  if (!text.trim()) return;
  const title = `${state.file?.name || "文献"} - 主要内容总结`;

  try {
    if (navigator.share) {
      await navigator.share({ title, text });
      showToast("已打开系统分享面板。");
      return;
    }
  } catch (error) {
    if (error?.name === "AbortError") {
      showToast("已取消分享。");
      return;
    }
    console.warn("Summary share failed", error);
  }

  try {
    await navigator.clipboard.writeText(`${title}\n\n${text}`);
    showToast("当前浏览器不支持直接分享总结，已复制到剪贴板。");
  } catch {
    showToast("当前浏览器不支持直接分享，请手动复制总结文本。", true);
  }
}

async function buildSummaryCacheKey(sourceText, detail) {
  const fingerprintSource = [
    state.fileType || "",
    detail || "standard",
    sourceText,
  ].join("\n---summary-cache---\n");
  return `summary:${detail}:${await digestText(fingerprintSource)}`;
}

async function digestText(text) {
  const value = String(text || "");
  if (window.crypto?.subtle && window.TextEncoder) {
    const bytes = new TextEncoder().encode(value);
    const hash = await window.crypto.subtle.digest("SHA-256", bytes);
    return [...new Uint8Array(hash)].map((byte) => byte.toString(16).padStart(2, "0")).join("");
  }

  let hash = 2166136261;
  for (let index = 0; index < value.length; index += 1) {
    hash ^= value.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0).toString(16);
}

async function getCachedSummary(id) {
  const db = await openSummaryCacheDb();
  return idbRequest(db, SUMMARY_CACHE_STORE, "readonly", (store) => store.get(id));
}

async function putCachedSummary(record) {
  const db = await openSummaryCacheDb();
  return idbRequest(db, SUMMARY_CACHE_STORE, "readwrite", (store) => store.put(record));
}

function openSummaryCacheDb() {
  return new Promise((resolve, reject) => {
    if (!("indexedDB" in window)) {
      reject(new Error("IndexedDB unavailable"));
      return;
    }
    const request = indexedDB.open(SUMMARY_CACHE_DB, 1);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(SUMMARY_CACHE_STORE)) {
        db.createObjectStore(SUMMARY_CACHE_STORE, { keyPath: "id" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
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
  if (!model) {
    showToast("请先填写模型。API Key 可留空使用 Cloudflare 后端配置。", true);
    return;
  }

  const files = [...state.batchFiles];
  const saved = [];
  const failed = [];
  state.batchRunning = true;
  state.batchCancelRequested = false;

  try {
    setBusy(true, `正在批量翻译 0/${files.length}...`);

    for (let index = 0; index < files.length; index += 1) {
      if (state.batchCancelRequested) break;
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

    if (state.batchCancelRequested) {
      state.batchFiles = [];
      renderBatchQueue();
      return;
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

async function shareOriginalFile() {
  if (!state.file) return;

  try {
    setBusy(true, "正在准备原始文件分享...");
    await waitForUiFrame();
    const file = new File([state.file], state.file.name || "source-document", {
      type: state.file.type || getOriginalMimeType(),
    });

    if (navigator.canShare?.({ files: [file] })) {
      await navigator.share({
        files: [file],
        title: state.file.name || "原始文献",
        text: "CuraWay 文档翻译工具中的原始文献",
      });
      showToast("已打开原始文件分享面板。");
      return;
    }

    saveBlobAsFile(state.file, state.file.name || "source-document");
    showToast("当前浏览器不支持直接分享原始文件，已先下载原始文件。", true);
  } catch (error) {
    if (error?.name === "AbortError") {
      showToast("已取消分享。");
    } else {
      console.error("Original file share failed", error);
      showToast(error.message || "原始文件分享失败。", true);
    }
  } finally {
    setBusy(false);
  }
}

function getOriginalMimeType() {
  if (state.fileType === "pdf") return "application/pdf";
  if (state.fileType === "docx") return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  if (state.fileType === "pptx") return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
  return state.file?.type || "application/octet-stream";
}

async function generateTranslatedFile() {
  if (!state.file) throw new Error("请先选择文件。");

  normalizeAllTranslations();

  if (state.fileType === "pdf") {
    validatePdfTranslationsBeforeExport();
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
    const titleRow = document.createElement("div");
    titleRow.className = "saved-file-title";
    const name = document.createElement("strong");
    name.textContent = record.filename;
    const typeBadge = document.createElement("span");
    const typeInfo = getSavedFileTypeInfo(record);
    typeBadge.className = `file-type-badge ${typeInfo.kind}`;
    typeBadge.textContent = typeInfo.label;
    titleRow.append(name, typeBadge);
    const detail = document.createElement("span");
    detail.textContent = `${typeInfo.description} · ${formatBytes(record.size)} · ${formatSavedTime(record.createdAt)}`;
    meta.append(titleRow, detail);

    const actions = document.createElement("div");
    actions.className = "saved-file-actions";
    const download = document.createElement("button");
    download.type = "button";
    download.className = "saved-download-action";
    download.dataset.icon = "download";
    download.textContent = "下载";
    decorateButtonIcon(download);
    download.addEventListener("click", () => downloadSavedFile(record.id));
    const share = document.createElement("button");
    share.type = "button";
    share.className = "saved-share-action";
    share.dataset.icon = "share";
    share.textContent = typeInfo.kind === "pdf" ? "分享" : "转PDF分享";
    decorateButtonIcon(share);
    share.addEventListener("click", () => shareSavedFile(record.id));
    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "saved-remove-action";
    remove.dataset.icon = "trash";
    remove.textContent = "删除";
    decorateButtonIcon(remove);
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
  if (getSavedFileKind(record) !== "pdf") {
    await shareSavedOfficeAsPdf(record);
    return;
  }
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

async function shareSavedOfficeAsPdf(record) {
  const typeInfo = getSavedFileTypeInfo(record);
  const confirmed = window.confirm(
    `${typeInfo.label} 文件在手机微信里通常不能直接作为文件分享。\n\n是否生成一个 PDF 分享版（文本版）并分享？`
  );
  if (!confirmed) return;

  try {
    setBusy(true, `正在生成 ${typeInfo.label} 的 PDF 分享版...`);
    setStatus("正在提取已保存文件中的译文文本，生成 PDF 分享版...");
    await waitForUiFrame();
    const { blob, filename } = await createSavedOfficePdfShareFile(record);
    await saveGeneratedFile(blob, filename);
    await shareBlobFile(blob, filename, "application/pdf");
  } catch (error) {
    console.error("Office PDF share conversion failed", error);
    showToast(error.message || "PDF 分享版生成失败，请先下载文件后用 WPS/Office 转成 PDF 再分享。", true);
  } finally {
    setBusy(false);
  }
}

async function createSavedOfficePdfShareFile(record) {
  const kind = getSavedFileKind(record);
  if (!["docx", "pptx"].includes(kind)) throw new Error("这个文件暂不支持生成 PDF 分享版。");

  const JSZipLib = await loadJsZip();
  const zip = await JSZipLib.loadAsync(record.blob);
  const sections = kind === "docx"
    ? await extractSavedDocxTextSections(zip)
    : await extractSavedPptxTextSections(zip);

  if (!sections.some((section) => section.lines.length)) {
    throw new Error("没有提取到可生成 PDF 的文本内容。");
  }

  const blob = await createTextPdfBlob({
    title: record.filename,
    subtitle: `${getSavedFileTypeInfo(record).label} PDF 分享版（文本版）`,
    sections,
  });

  return { blob, filename: buildPdfShareFilename(record.filename) };
}

async function extractSavedDocxTextSections(zip) {
  const paths = Object.keys(zip.files)
    .filter((path) => wordPathPattern.test(path))
    .sort((a, b) => (a === "word/document.xml" ? -1 : b === "word/document.xml" ? 1 : a.localeCompare(b)));

  const sections = [];
  for (const path of paths) {
    const xmlText = await zip.file(path)?.async("text");
    if (!xmlText) continue;
    const doc = parser.parseFromString(xmlText, "application/xml");
    const lines = [...doc.getElementsByTagNameNS(WORD_NS, "p")]
      .map((paragraph) => [...paragraph.getElementsByTagNameNS(WORD_NS, "t")].map((node) => node.textContent || "").join("").trim())
      .filter(Boolean);
    if (lines.length) sections.push({ heading: getWordPartLabel(path), lines });
  }
  return sections;
}

async function extractSavedPptxTextSections(zip) {
  const slideFiles = Object.keys(zip.files)
    .map((path) => ({ path, match: path.match(slidePathPattern) }))
    .filter((item) => item.match)
    .sort((a, b) => Number(a.match[1]) - Number(b.match[1]));

  const sections = [];
  for (const item of slideFiles) {
    const xmlText = await zip.file(item.path)?.async("text");
    if (!xmlText) continue;
    const doc = parser.parseFromString(xmlText, "application/xml");
    const lines = [...doc.getElementsByTagNameNS(DRAWING_NS, "p")]
      .map((paragraph) => [...paragraph.getElementsByTagNameNS(DRAWING_NS, "t")].map((node) => node.textContent || "").join("").trim())
      .filter(Boolean);
    if (lines.length) sections.push({ heading: `第 ${Number(item.match[1])} 页`, lines });
  }
  return sections;
}

async function createTextPdfBlob({ title, subtitle, sections, showHeader = true }) {
  const { PDFDocument, rgb, fontkit, fontBytes } = await loadPdfExportTools();
  const pdfDoc = await PDFDocument.create();
  pdfDoc.registerFontkit(fontkit);
  pdfDoc.setTitle(buildPdfShareFilename(title));
  const font = await pdfDoc.embedFont(fontBytes, { subset: false });
  const pageWidth = 595.28;
  const pageHeight = 841.89;
  const margin = 44;
  const maxWidth = pageWidth - margin * 2;
  let page = pdfDoc.addPage([pageWidth, pageHeight]);
  let y = pageHeight - margin;

  const ensurePage = (needed = 16) => {
    if (y - needed >= margin) return;
    page = pdfDoc.addPage([pageWidth, pageHeight]);
    y = pageHeight - margin;
  };

  const drawLines = (text, size = 11, lineHeight = 17, color = rgb(0.08, 0.11, 0.12)) => {
    wrapPdfText(text, font, size, maxWidth).forEach((line) => {
      ensurePage(lineHeight);
      if (line) page.drawText(line, { x: margin, y, size, font, color });
      y -= lineHeight;
    });
  };

  if (showHeader) {
    drawLines(title, 15, 22, rgb(0.05, 0.22, 0.28));
    drawLines(subtitle, 10, 18, rgb(0.36, 0.42, 0.4));
    y -= 8;
  }

  sections.forEach((section) => {
    ensurePage(34);
    drawLines(section.heading, 12, 20, rgb(0.08, 0.34, 0.43));
    section.lines.forEach((line) => drawLines(line, 10.5, 16));
    y -= 8;
  });

  const bytes = await pdfDoc.save();
  return new Blob([bytes], { type: "application/pdf" });
}

function buildPdfShareFilename(filename) {
  return String(filename || "translated-file")
    .replace(/\.(pptx|docx|pdf)$/i, "")
    .replace(/[\\/:*?"<>|]+/g, "_") + "-PDF分享版.pdf";
}

function getSavedFileKind(record) {
  const filename = String(record?.filename || "").toLowerCase();
  const type = String(record?.type || "").toLowerCase();
  if (filename.endsWith(".pdf") || type.includes("pdf")) return "pdf";
  if (filename.endsWith(".docx") || type.includes("wordprocessingml")) return "docx";
  if (filename.endsWith(".pptx") || type.includes("presentationml")) return "pptx";
  return "file";
}

function getSavedFileTypeInfo(record) {
  const kind = getSavedFileKind(record);
  if (kind === "pdf") return { kind, label: "PDF", description: "可直接分享" };
  if (kind === "docx") return { kind, label: "Word", description: "可转 PDF 分享" };
  if (kind === "pptx") return { kind, label: "PPT", description: "可转 PDF 分享" };
  return { kind, label: "文件", description: "保存文件" };
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
  const compactedSegments = compactPresentationTextBoxGroups(paragraphs, segments);

  segments.forEach((segment) => {
    if (compactedSegments.has(segment)) return;
    const paragraph = paragraphs[segment.paragraphIndex];
    if (!paragraph || !segment.translation.trim()) return;

    const textNodes = [...paragraph.getElementsByTagNameNS(DRAWING_NS, "t")];
    if (!textNodes.length) return;

    if (shouldSuppressZeroBoundsPresentationText(segment, paragraph)) {
      textNodes[0].textContent = "";
      clearRemainingTextNodes(textNodes);
      return;
    }

    textNodes[0].textContent = cleanPptTranslationText(segment.translation, segment.original);
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

    const textNodes = [...paragraph.getElementsByTagNameNS(WORD_NS, "t")];
    if (!textNodes.length) return;

    if (segment.wordSyntheticAlignment) {
      applyWordParagraphJustification(paragraph, segment.wordSyntheticAlignment);
    }
    writeWordTextNodes(textNodes, segment);
  });
}

function compactPresentationTextBoxGroups(paragraphs, segments) {
  const compacted = new Set();
  const groups = new Map();

  segments.forEach((segment) => {
    if (segment.type !== "pptx" || !segment.translation.trim()) return;
    const key = getPptShapeGroupKey(segment);
    if (!key) return;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(segment);
  });

  groups.forEach((group) => {
    const ordered = group.sort((a, b) => a.paragraphIndex - b.paragraphIndex);
    if (!shouldCompactPresentationTextBoxGroup(ordered)) return;

    const titleSegment = ordered[0];
    const bodySegments = ordered.slice(1);
    const titleParagraph = paragraphs[titleSegment.paragraphIndex];
    const bodyParagraph = paragraphs[bodySegments[0]?.paragraphIndex];
    if (!titleParagraph || !bodyParagraph) return;

    writePresentationParagraphText(titleParagraph, titleSegment, cleanPptTranslationText(titleSegment.translation, titleSegment.original));
    writePresentationParagraphText(
      bodyParagraph,
      bodySegments[0],
      bodySegments.map((segment) => cleanPptTranslationText(segment.translation, segment.original)).filter(Boolean).join("\n"),
    );

    bodySegments.slice(1).forEach((segment) => {
      const paragraph = paragraphs[segment.paragraphIndex];
      paragraph?.parentNode?.removeChild(paragraph);
    });

    ordered.forEach((segment) => compacted.add(segment));
  });

  return compacted;
}

function shouldCompactPresentationTextBoxGroup(group) {
  if (group.length < 3) return false;
  const textLength = group.reduce((sum, segment) => sum + [...String(segment.original || "").trim()].length, 0);
  const box = getPptContentBoxPt(group[0].layout || {});
  return textLength <= 140 && Boolean(box);
}

function writePresentationParagraphText(paragraph, segment, text) {
  const textNodes = [...paragraph.getElementsByTagNameNS(DRAWING_NS, "t")];
  if (!textNodes.length) return;
  if (shouldSuppressZeroBoundsPresentationText(segment, paragraph)) {
    textNodes[0].textContent = "";
    clearRemainingTextNodes(textNodes);
    return;
  }
  textNodes[0].textContent = text;
  normalizePresentationRunStyle(textNodes[0], segment);
  applyPresentationBounds(paragraph, segment);
  applyPresentationLayout(paragraph, segment);
  clearRemainingTextNodes(textNodes);
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

function shouldSkipWordParagraph(text) {
  const value = String(text || "").trim();
  if (!value) return true;
  return !/[A-Za-z\u3400-\u9fff]/.test(value);
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
  const hasTranslations = translated > 0;
  els.slideCount.textContent = String(state.slideCount);
  els.segmentCount.textContent = String(state.segments.length);
  els.translatedCount.textContent = String(translated);
  updateTranslateButtonState(document.body.classList.contains("is-busy") || state.batchRunning);
  if (els.sourcePreviewButton) {
    els.sourcePreviewButton.disabled = !state.segments.length;
  }
  els.previewButton.disabled = !state.segments.length;
  if (els.layoutPreviewButton) {
    els.layoutPreviewButton.disabled = !state.segments.length;
  }
  els.downloadButton.disabled = !state.segments.length || !hasTranslations;
  els.shareButton.disabled = !state.segments.length || !hasTranslations;
  if (els.summaryButton) {
    els.summaryButton.disabled = !state.segments.length;
  }
  updateBatchButton();
  if (els.previewDownloadButton) {
    els.previewDownloadButton.disabled = !state.segments.length || !hasTranslations;
  }
  if (els.previewShareButton) {
    els.previewShareButton.disabled = !state.segments.length || !hasTranslations;
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

function openFileLibrary() {
  if (state.mobileView === "library") {
    setMobileView("translate");
    setMobileMenuOpen(false);
    els.uploadZone?.scrollIntoView({ block: "nearest", behavior: "smooth" });
    return;
  }
  setMobileView("library");
  setMobileMenuOpen(false);
  document.querySelector(".saved-files")?.scrollIntoView({ block: "nearest", behavior: "smooth" });
}

function setBusy(isBusy, message = "") {
  const effectiveBusy = isBusy || state.batchRunning;
  const hasTranslations = state.segments.some((segment) => segment.translation.trim());
  updateTranslateButtonState(effectiveBusy);
  if (els.sourcePreviewButton) {
    els.sourcePreviewButton.disabled = effectiveBusy || !state.segments.length;
  }
  els.previewButton.disabled = effectiveBusy || !state.segments.length;
  if (els.layoutPreviewButton) {
    els.layoutPreviewButton.disabled = effectiveBusy || !state.segments.length;
  }
  els.downloadButton.disabled = effectiveBusy || !state.segments.length || !hasTranslations;
  els.shareButton.disabled = effectiveBusy || !state.segments.length || !hasTranslations;
  if (els.summaryButton) {
    els.summaryButton.disabled = effectiveBusy || !state.segments.length;
  }
  if (els.batchTranslateButton) {
    els.batchTranslateButton.disabled = effectiveBusy || !state.batchFiles.length;
  }
  if (els.previewDownloadButton) {
    els.previewDownloadButton.disabled = effectiveBusy || !state.segments.length || !hasTranslations;
  }
  if (els.previewShareButton) {
    els.previewShareButton.disabled = effectiveBusy || !state.segments.length || !hasTranslations;
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
  setButtonLabel(els.translateButton, canStop ? "停止翻译" : "自动翻译");
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
  state.pdfManualRegionSelectionEnabled = false;
  updatePdfRegionSelectionControls();
  state.previewMode = "translation";
  renderPreview();
  openPreviewDialog();
}

function openSourcePreview() {
  state.previewMode = "original";
  renderPreview();
  openPreviewDialog();
}

function openPreviewDialog() {
  if (typeof els.previewDialog.showModal === "function") {
    els.previewDialog.showModal();
  } else {
    els.previewDialog.setAttribute("open", "");
  }
}

function handlePreviewDownload() {
  if (state.previewMode === "original") {
    if (state.file) saveBlobAsFile(state.file, state.file.name || "source-document");
    return;
  }
  downloadPresentation();
}

function handlePreviewShare() {
  if (state.previewMode === "original") {
    return shareOriginalFile();
  }
  return sharePresentation();
}

function closePreview() {
  state.pdfManualRegionSelectionEnabled = false;
  updatePdfRegionSelectionControls();
  exitPreviewFullscreen();
  if (typeof els.previewDialog.close === "function" && els.previewDialog.open) {
    els.previewDialog.close();
  } else {
    els.previewDialog.removeAttribute("open");
  }
  revokeOriginalPreviewUrl();
}

async function togglePreviewFullscreen() {
  if (!els.previewDialog) return;
  if (state.previewFullscreen || document.fullscreenElement === els.previewDialog) {
    await exitPreviewFullscreen();
    return;
  }
  state.previewFullscreen = true;
  syncPreviewFullscreenClass();
  try {
    if (els.previewDialog.requestFullscreen) {
      await els.previewDialog.requestFullscreen();
    }
  } catch (error) {
    console.warn("Preview fullscreen request failed; using CSS fullscreen.", error);
  }
  syncPreviewFullscreenClass();
}

async function exitPreviewFullscreen() {
  state.previewFullscreen = false;
  if (document.fullscreenElement === els.previewDialog && document.exitFullscreen) {
    try {
      await document.exitFullscreen();
    } catch (error) {
      console.warn("Preview fullscreen exit failed", error);
    }
  }
  syncPreviewFullscreenClass();
}

function syncPreviewFullscreenState() {
  state.previewFullscreen = document.fullscreenElement === els.previewDialog;
  syncPreviewFullscreenClass();
}

function syncPreviewFullscreenClass() {
  els.previewDialog?.classList.toggle("fullscreen", Boolean(state.previewFullscreen));
  if (els.previewFullscreenButton) {
    els.previewFullscreenButton.title = state.previewFullscreen ? "退出全屏" : "全屏预览";
    els.previewFullscreenButton.setAttribute("aria-label", els.previewFullscreenButton.title);
  }
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

async function shareApp() {
  const url = new URL(window.location.href);
  url.hash = "";
  const shareUrl = url.toString();
  const title = "CuraWay 文档翻译工具";
  const text = "打开 CuraWay 文档翻译工具";

  try {
    if (navigator.share) {
      await navigator.share({ title, text, url: shareUrl });
      setStatus("已打开 App 分享面板。");
      return;
    }

    await navigator.clipboard.writeText(shareUrl);
    showCompletionDialog("App 链接已复制", "已复制当前 App 链接，可以粘贴到微信或浏览器发送。");
    setStatus("App 链接已复制。");
  } catch (error) {
    console.warn("App share failed", error);
    showCompletionDialog("无法自动分享", `请手动复制当前链接：${shareUrl}`);
    setStatus("App 分享未完成。");
  }
}

function decorateActionButtons() {
  [
    [els.batchTranslateButton, "layers"],
    [els.sourcePreviewButton, "file-search"],
    [els.translateButton, "spark"],
    [els.summaryButton, "list"],
    [els.previewButton, "layout"],
    [els.layoutPreviewButton, "layout"],
    [els.shareButton, "share"],
    [els.previewShareButton, "share"],
    [els.previewDownloadButton, "download"],
  ].forEach(([button, icon]) => {
    if (!button) return;
    button.dataset.icon = button.dataset.icon || icon;
    decorateButtonIcon(button);
  });
}

function decorateButtonIcon(button) {
  if (!button || button.querySelector(".button-label")) return;
  const text = button.textContent.trim();
  button.textContent = "";
  const icon = document.createElement("span");
  icon.className = "button-icon";
  icon.setAttribute("aria-hidden", "true");
  icon.innerHTML = getButtonIconSvg(button.dataset.icon || "spark");
  const label = document.createElement("span");
  label.className = "button-label";
  label.textContent = text;
  button.append(icon, label);
}

function setButtonLabel(button, text) {
  if (!button) return;
  const label = button.querySelector(".button-label");
  if (label) {
    label.textContent = text;
  } else {
    button.textContent = text;
  }
}

function getButtonIconSvg(name) {
  const icons = {
    share: '<svg viewBox="0 0 24 24"><path d="M4 12v7a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-7"/><path d="M12 16V4"/><path d="m7 9 5-5 5 5"/></svg>',
    download: '<svg viewBox="0 0 24 24"><path d="M12 3v12"/><path d="m7 10 5 5 5-5"/><path d="M5 21h14"/></svg>',
    spark: '<svg viewBox="0 0 24 24"><path d="M12 3l1.7 5.3L19 10l-5.3 1.7L12 17l-1.7-5.3L5 10l5.3-1.7L12 3Z"/><path d="M19 15l.8 2.2L22 18l-2.2.8L19 21l-.8-2.2L16 18l2.2-.8L19 15Z"/></svg>',
    "file-search": '<svg viewBox="0 0 24 24"><path d="M14 2H7a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h7"/><path d="M14 2v6h6"/><path d="M20 11V8l-6-6"/><circle cx="16.5" cy="16.5" r="3.5"/><path d="m19 19 2 2"/></svg>',
    layout: '<svg viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 10h18"/><path d="M9 10v10"/></svg>',
    list: '<svg viewBox="0 0 24 24"><path d="M8 6h13"/><path d="M8 12h13"/><path d="M8 18h13"/><path d="M3 6h.01"/><path d="M3 12h.01"/><path d="M3 18h.01"/></svg>',
    layers: '<svg viewBox="0 0 24 24"><path d="m12 3 9 5-9 5-9-5 9-5Z"/><path d="m3 12 9 5 9-5"/><path d="m3 16 9 5 9-5"/></svg>',
    trash: '<svg viewBox="0 0 24 24"><path d="M3 6h18"/><path d="M8 6V4h8v2"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v5"/><path d="M14 11v5"/></svg>',
  };
  return icons[name] || icons.spark;
}

async function openAdminManagement() {
  if (!isSuperAdmin()) {
    await refreshCurrentUserSession();
  }
  if (!isSuperAdmin()) {
    showToast("普通用户没有管理权限。", true);
    setStatus("普通用户没有管理权限。");
    return;
  }

  setStatus("正在检查用户管理后台...");
  try {
    const response = await fetch("/api/admin/session", {
      cache: "no-store",
      headers: { Accept: "application/json" },
    });

    if (response.status === 404 || response.status === 405) {
      showAdminUnavailableDialog();
      return;
    }

    window.location.href = "/admin";
  } catch (error) {
    console.warn("Admin backend check failed", error);
    showAdminUnavailableDialog();
  }
}

function showAdminUnavailableDialog() {
  showCompletionDialog(
    "请打开 Cloudflare 正式地址",
    "用户管理依赖 Cloudflare Pages Functions 和 PHONE_AUTH_KV。当前地址没有检测到后台接口，请在新手机上打开已部署到 Cloudflare 的正式链接。"
  );
  setStatus("当前地址不是可用的 Cloudflare 用户管理后台。");
}

function renderPreview() {
  const previousScrollTop = els.previewBody.scrollTop;
  const selectedIndex = els.previewBody.querySelector(".slide-text-box.selected")?.dataset.index || "";
  const translated = state.segments.filter((segment) => segment.translation.trim()).length;
  const isOriginalMode = state.previewMode === "original";
  if (isOriginalMode) {
    renderOriginalDocumentPreview();
    return;
  }
  const isSourceMode = state.previewMode === "source";
  if (els.previewTitle) {
    els.previewTitle.textContent = isSourceMode ? "待翻译文献预览" : "译文预览";
  }
  if (els.previewDownloadButton) els.previewDownloadButton.hidden = true;
  if (els.previewShareButton) {
    els.previewShareButton.hidden = isSourceMode;
    setButtonLabel(els.previewShareButton, "分享翻译文件");
    els.previewShareButton.dataset.icon = "share";
    decorateButtonIcon(els.previewShareButton);
  }
  els.previewMeta.textContent = isSourceMode
    ? `${getFileTypeName()} · ${state.segments.length} 段原文 · 翻译前预览`
    : `${getFileTypeName()} · ${state.segments.length} 段文字 · ${translated} 段已有译文`;
  els.previewBody.replaceChildren();

  if (!state.segments.length) {
    const empty = document.createElement("div");
    empty.className = "empty-state";
    empty.textContent = "还没有可预览的内容。";
    els.previewBody.append(empty);
    return;
  }

  if (state.fileType === "pdf" && !isSourceMode) {
    renderPdfLayoutPreview();
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
  const mode = els.pdfOutputMode?.value || "overlay";
  if (mode === "overlay") return downloadPdfOverlayTranslation();
  return downloadPdfReflowTranslation();
}

async function downloadPdfReflowTranslation() {
  setProgress(0.01);
  setStatus("正在生成重排版译文 PDF...");
  await waitForUiFrame();

  const { PDFDocument, rgb, fontkit, fontBytes } = await loadPdfExportTools();
  const pdfDoc = await PDFDocument.create();
  pdfDoc.registerFontkit(fontkit);
  pdfDoc.setTitle(`${state.file.name} translated reflow`);
  const font = await pdfDoc.embedFont(fontBytes, { subset: false });

  const pageWidth = 595.28;
  const pageHeight = 841.89;
  const margin = 42;
  const contentWidth = pageWidth - margin * 2;
  let page = pdfDoc.addPage([pageWidth, pageHeight]);
  let y = pageHeight - margin;

  page.drawText(stripPdfExtension(state.file?.name || "Translated PDF"), {
    x: margin,
    y,
    size: 15,
    font,
    color: rgb(0.06, 0.08, 0.09),
  });
  y -= 24;

  const meta = state.pdfParseSource === "llamaparse"
    ? "Parsed by Cloudflare + LlamaParse, translated by DeepSeek"
    : "Parsed by browser compatible mode, translated by DeepSeek";
  page.drawText(meta, {
    x: margin,
    y,
    size: 8.5,
    font,
    color: rgb(0.38, 0.44, 0.42),
  });
  y -= 26;

  const segments = state.segments.filter((segment) => segment.type === "pdf" && String(segment.translation || "").trim());
  if (!segments.length) throw new Error("没有可导出的 PDF 译文。");

  for (let index = 0; index < segments.length; index += 1) {
    const segment = segments[index];
    const source = normalizePdfParsedText(segment.original);
    const translation = getPdfExportText(segment) || normalizePdfParsedText(segment.translation);
    if (!translation) continue;

    const block = createPdfReflowBlock(segment, source, translation, font, contentWidth);
    if (y - block.height < margin) {
      page = pdfDoc.addPage([pageWidth, pageHeight]);
      y = pageHeight - margin;
    }

    y = drawPdfReflowBlock(page, block, margin, y, font, rgb);

    if (index % 10 === 0 || index === segments.length - 1) {
      setProgress(0.12 + ((index + 1) / segments.length) * 0.72);
      setStatus(`正在排版重排版 PDF：${index + 1}/${segments.length}`);
      await waitForUiFrame();
    }
  }

  setProgress(0.92);
  setStatus("正在保存重排版译文 PDF...");
  await waitForUiFrame();
  const pdfBytes = await pdfDoc.save();
  const blob = new Blob([pdfBytes], { type: "application/pdf" });
  setProgress(1);
  setStatus("PDF 重排版导出完成。");
  return blob;
}

function createPdfReflowBlock(segment, source, translation, font, width) {
  const isHeading = /heading|title/i.test(segment.layout?.itemType || "") || /^#{1,4}\s+/.test(source);
  const isTable = looksLikeMarkdownTable(source) || looksLikeMarkdownTable(translation);
  const fontSize = isHeading ? 13 : 10.5;
  const sourceSize = 7.4;
  const labelHeight = 12;
  const translationLines = isTable
    ? wrapMarkdownTableForPdf(translation, font, fontSize, width)
    : wrapPdfText(stripMarkdownDecorations(translation), font, fontSize, width);
  const sourceLines = wrapPdfText(stripMarkdownDecorations(source), font, sourceSize, width);
  const lineHeight = fontSize * 1.38;
  const sourceLineHeight = sourceSize * 1.35;
  const height = labelHeight + translationLines.length * lineHeight + Math.min(4, sourceLines.length) * sourceLineHeight + 18;
  return { segment, isHeading, translationLines, sourceLines: sourceLines.slice(0, 4), fontSize, sourceSize, lineHeight, sourceLineHeight, height };
}

function drawPdfReflowBlock(page, block, x, y, font, rgb) {
  page.drawText(block.segment.locationLabel || "", {
    x,
    y,
    size: 7,
    font,
    color: rgb(0.48, 0.53, 0.5),
  });
  y -= 12;

  block.translationLines.forEach((line) => {
    page.drawText(line, {
      x,
      y,
      size: block.fontSize,
      font,
      color: rgb(0.06, 0.08, 0.09),
    });
    y -= block.lineHeight;
  });

  if (block.sourceLines.length) {
    y -= 2;
    block.sourceLines.forEach((line) => {
      page.drawText(line, {
        x,
        y,
        size: block.sourceSize,
        font,
        color: rgb(0.46, 0.5, 0.49),
      });
      y -= block.sourceLineHeight;
    });
  }

  return y - 12;
}

function looksLikeMarkdownTable(text) {
  const lines = String(text || "").split(/\r\n|\r|\n/);
  return lines.length >= 2 && lines.some((line) => /\|/.test(line)) && lines.some((line) => /^\s*\|?\s*:?-{3,}/.test(line));
}

function wrapMarkdownTableForPdf(text, font, fontSize, maxWidth) {
  return String(text || "")
    .split(/\r\n|\r|\n/)
    .filter((line) => !/^\s*\|?\s*:?-{3,}/.test(line))
    .flatMap((line) => wrapPdfText(line.replace(/^\s*\|/, "").replace(/\|\s*$/, "").replace(/\s*\|\s*/g, "  |  "), font, fontSize, maxWidth));
}

function stripMarkdownDecorations(text) {
  return String(text || "")
    .replace(/^#{1,6}\s+/gm, "")
    .replace(/\*\*(.*?)\*\*/g, "$1")
    .replace(/__(.*?)__/g, "$1")
    .replace(/`([^`]+)`/g, "$1")
    .trim();
}

function stripPdfExtension(name) {
  return String(name || "").replace(/\.pdf$/i, "");
}

async function downloadPdfOverlayTranslation() {
  setProgress(0.01);
  setStatus("PDF 导出准备中，请稍等...");
  await waitForUiFrame();

  setProgress(0.06);
  setStatus("正在加载 PDF 导出组件和中文字体，首次使用可能需要更久...");
  await waitForUiFrame();
  const { PDFDocument, rgb, fontkit, fontBytes } = await loadPdfExportTools();
  if (!state.pdfBytes) throw new Error("缺少原 PDF 数据，请重新上传 PDF。");

  setProgress(0.18);
  setStatus("正在解析原始 PDF，并重建图片、表格、线条等非文字元素...");
  await waitForUiFrame();
  const pdfjs = await loadPdfJs();
  const sourcePdf = await pdfjs.getDocument({ data: state.pdfBytes.slice() }).promise;
  const pdfDoc = await PDFDocument.create();
  pdfDoc.registerFontkit(fontkit);
  pdfDoc.setTitle(`${state.file.name} translated`);
  const pages = [];
  const pageVectorLines = [];

  for (let pageNumber = 1; pageNumber <= sourcePdf.numPages; pageNumber += 1) {
    const sourcePage = await sourcePdf.getPage(pageNumber);
    const pageSize = state.pdfPageSizes.get(`pdf/page-${pageNumber}`) || sourcePage.getViewport({ scale: 1 });
    const page = pdfDoc.addPage([pageSize.width || 595.28, pageSize.height || 841.89]);
    const startRatio = sourcePdf.numPages ? (pageNumber - 1) / sourcePdf.numPages : 0;
    const rebuiltLines = [];
    const pageRenderOptions = { renderCache: new Map(), rebuiltLines };
    setProgress(0.18 + startRatio * 0.22);
    setStatus(`正在重建 PDF 非文字元素：第 ${pageNumber}/${sourcePdf.numPages} 页（复杂图片会自动跳过，避免卡住）`);
    await waitForUiFrame();
    try {
      const pageBudgetMs = isLikelyMobileDevice() ? 8500 : 13500;
      await rebuildPdfPageNonTextElements(sourcePage, page, pdfDoc, pdfjs, pageSize, rgb, {
        ...pageRenderOptions,
        deadline: Date.now() + pageBudgetMs,
      });
      await drawPdfPreservedRegions(sourcePage, page, pdfDoc, `pdf/page-${pageNumber}`, pageSize, pageRenderOptions);
      await drawPdfPreservedPageFurniture(sourcePage, page, pdfDoc, `pdf/page-${pageNumber}`, pageSize, pageRenderOptions);
    } catch (error) {
      console.warn(`PDF page ${pageNumber} non-text rebuild failed`, error);
      setStatus(`第 ${pageNumber} 页部分图片或复杂元素重建超时，正在继续写入译文...`);
      await waitForUiFrame();
    }
    pages.push(page);
    pageVectorLines[pageNumber - 1] = rebuiltLines;

    const ratio = sourcePdf.numPages ? pageNumber / sourcePdf.numPages : 1;
    setProgress(0.18 + ratio * 0.22);
    setStatus(`已完成 PDF 非文字元素重建：${pageNumber}/${sourcePdf.numPages}`);
    await waitForUiFrame();
  }

  setProgress(0.43);
  setStatus("正在嵌入译文字体...");
  await waitForUiFrame();
  const font = await pdfDoc.embedFont(fontBytes, { subset: false });
  const translatedSegments = state.segments.filter((segment) => segment.type === "pdf" && segment.layout?.bounds);
  if (!translatedSegments.length) {
    throw new Error("当前 PDF 段落没有可用于重建版译文 PDF 的坐标。请切换为“备用：重排版译文 PDF”。");
  }
  const overlayPlans = [];

  for (let index = 0; index < translatedSegments.length; index += 1) {
    const segment = translatedSegments[index];
    const page = pages[Number(segment.slideNumber) - 1];
    if (page) {
      const plan = createPdfOverlayPlan(segment, font);
      if (plan) {
        overlayPlans.push({ page, plan });
      }
    }

    if (index % 12 === 0 || index === translatedSegments.length - 1) {
      const ratio = translatedSegments.length ? (index + 1) / translatedSegments.length : 1;
      setProgress(0.45 + ratio * 0.38);
      setStatus(`正在回写 PDF 译文：${index + 1}/${translatedSegments.length} 段，请勿关闭页面。`);
      await waitForUiFrame();
    }
  }

  resolvePdfOverlayPlanCollisions(overlayPlans);
  overlayPlans.forEach(({ page, plan }) => drawPdfOverlayErase(page, plan, rgb));
  overlayPlans.forEach(({ page, plan }) => drawPdfOverlayText(page, plan, font, rgb));
  pages.forEach((page, index) => {
    const pagePlans = overlayPlans
      .filter((entry) => entry.page === page)
      .map((entry) => entry.plan);
    redrawPdfLinesCoveredByOverlay(page, pageVectorLines[index] || [], pagePlans, rgb);
  });

  setProgress(0.9);
  setStatus("正在保存翻译版 PDF，文件较大时可能需要几十秒...");
  await waitForUiFrame();
  setProgress(0.91);
  setStatus("正在补绘 PDF 表格线，增强重建版表格边框...");
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

function createPdfOverlayPlan(segment, font) {
  const sourceBounds = segment.layout.bounds;
  const targetBounds = getPdfSegmentTargetBounds(segment);
  const cell = segment.layout.tableCell;
  const direction = getDirectionConfig(els.translationDirection?.value || "");
  if (shouldSkipSegmentForDirection(segment, direction)) return null;

  const text = getPdfExportText(segment);
  if (!text) return null;

  let fitBounds;
  if (segment.overrides?.pdfBounds) {
    fitBounds = targetBounds;
  } else if (cell) {
    fitBounds = { x: cell.x + 2, y: cell.y + 2, width: Math.max(4, cell.width - 4), height: Math.max(4, cell.height - 4) };
  } else {
    fitBounds = getPdfOverlayFitBounds(segment, font, text, targetBounds);
  }
  const isTableLike = Boolean(cell) && !segment.overrides?.pdfBounds;
  const sourceFontSize = Math.max(4, Math.min(28, Number(segment.layout.fontSize || 10)));
  const paddingX = isTableLike ? Math.max(1.2, sourceFontSize * 0.12) : Math.max(2, sourceFontSize * 0.16);
  const paddingY = isTableLike ? Math.max(1, sourceFontSize * 0.1) : Math.max(1.5, sourceFontSize * 0.12);
  const eraseBaseX = Math.min(sourceBounds.x, fitBounds.x);
  const eraseBaseY = Math.min(sourceBounds.y, fitBounds.y);
  const eraseBaseRight = Math.max(sourceBounds.x + sourceBounds.width, fitBounds.x + fitBounds.width);
  const eraseBaseTop = Math.max(sourceBounds.y + sourceBounds.height, fitBounds.y + fitBounds.height);
  const eraseWidth = Math.max(2, eraseBaseRight - eraseBaseX + paddingX * 2);
  const eraseHeight = Math.max(
    eraseBaseTop - eraseBaseY + paddingY * 2,
    sourceFontSize * (isTableLike ? 1.35 : 1.55)
  );
  const eraseX = Math.max(0, eraseBaseX - paddingX);
  const eraseY = Math.max(
    0,
    eraseBaseY - paddingY - Math.max(0, (eraseHeight - (eraseBaseTop - eraseBaseY)) / 2)
  );
  const coverColor = segment.layout.backgroundColor || { r: 1, g: 1, b: 1 };
  const textColor = segment.layout.textColor || getReadablePdfTextColor(coverColor);

  const fit = fitPdfTextSize(text, font, fitBounds, segment.layout.fontSize || 10, segment.original, segment.layout);
  const fontSize = fit.fontSize;
  const lines = fit.lines;
  const lineHeight = fit.lineHeight;
  const textHeight = getPdfTextBlockHeight(lines, fontSize, lineHeight);
  const fitY = isTableLike ? Math.max(0, targetBounds.y - Math.max(0, (fitBounds.height - targetBounds.height) / 2)) : targetBounds.y;
  const baselineOffset = getPdfBaselineOffset(fontSize);
  const isTitleLike = isPdfTitleSegment(segment);
  let y = isTitleLike
    ? fitY + fitBounds.height - fontSize + baselineOffset
    : fitY + Math.max(0, (fitBounds.height - textHeight) / 2) + baselineOffset + (lines.length - 1) * lineHeight;

  return {
    segmentId: segment.id,
    path: segment.path,
    erase: {
      x: eraseX,
      y: eraseY,
      width: eraseWidth,
      height: eraseHeight,
      color: coverColor,
    },
    fitBounds,
    lines,
    fontSize,
    lineHeight,
    textColor,
    cellAlign: segment.layout.cellAlign,
    columnKey: segment.layout.columnKey || "",
    isFigureCaption: Boolean(segment.layout.isFigureCaption),
    isTitleLike,
    isTableLike,
    y,
  };
}

function isPdfTitleSegment(segment) {
  if (segment?.type !== "pdf") return false;
  const bounds = segment.layout?.bounds || {};
  const pageSize = state.pdfPageSizes.get(segment.path) || {};
  const pageWidth = Number(pageSize.width || 595.28);
  const pageHeight = Number(pageSize.height || 841.89);
  const fontSize = Number(segment.layout?.fontSize || 0);
  const x = Number(bounds.x || 0);
  const y = Number(bounds.y || 0);
  const width = Number(bounds.width || 0);
  return fontSize >= 13.5 &&
    x > pageWidth * 0.16 &&
    y > pageHeight * 0.45 &&
    y < pageHeight * 0.88 &&
    width > pageWidth * 0.25;
}

function createPdfSourceErasePlan(segment) {
  const sourceBounds = segment.layout?.bounds;
  if (!sourceBounds) return null;

  const cell = segment.layout?.tableCell;
  const sourceFontSize = Math.max(4, Math.min(28, Number(segment.layout?.fontSize || 10)));
  const lineCount = Number(segment.layout?.lineCount || 1);
  const pageSize = state.pdfPageSizes.get(segment.path) || {};
  const pageWidth = Number(pageSize.width || 595.28);
  const coverColor = segment.layout?.backgroundColor || { r: 1, g: 1, b: 1 };
  const paddingX = cell ? Math.max(1.2, sourceFontSize * 0.12) : Math.max(2, sourceFontSize * 0.2);
  const paddingY = cell ? Math.max(1, sourceFontSize * 0.1) : Math.max(1.5, sourceFontSize * 0.18);

  if (cell) {
    return {
      erase: {
        x: Math.max(0, cell.x - paddingX),
        y: Math.max(0, cell.y - paddingY),
        width: Math.max(2, cell.width + paddingX * 2),
        height: Math.max(2, cell.height + paddingY * 2),
        color: coverColor,
      },
    };
  }

  const availableWidth = Number(segment.layout?.availableWidth || 0);
  const availableHeight = Number(segment.layout?.availableHeight || 0);
  const eraseWidthBase = lineCount > 1
    ? Math.max(Number(sourceBounds.width || 0), availableWidth)
    : Number(sourceBounds.width || 0);
  const eraseHeightBase = lineCount > 1
    ? Math.max(Number(sourceBounds.height || 0), availableHeight)
    : Number(sourceBounds.height || 0);
  const columnKey = segment.layout?.columnKey || "";
  let eraseBaseX = Number(sourceBounds.x || 0);
  if (lineCount > 1 && columnKey === "left") eraseBaseX = Math.min(eraseBaseX, pageWidth * 0.055);
  if (lineCount > 1 && columnKey === "right") eraseBaseX = Math.min(eraseBaseX, pageWidth * 0.495);
  let eraseX = Math.max(0, eraseBaseX - paddingX);
  const columnRight = lineCount > 1 && columnKey === "left"
    ? pageWidth * 0.49
    : (lineCount > 1 && columnKey === "right" ? pageWidth * 0.955 : pageWidth - 4);
  const maxWidth = Math.max(2, columnRight - eraseX);
  let eraseY = Math.max(0, Number(sourceBounds.y || 0) - paddingY);
  let eraseWidth = lineCount > 1 && (columnKey === "left" || columnKey === "right")
    ? maxWidth
    : Math.max(2, Math.min(maxWidth, eraseWidthBase + paddingX * 2));
  let eraseHeight = Math.max(2, eraseHeightBase + paddingY * (lineCount > 1 ? 4.8 : 2));

  const fragmentUnion = getPdfEraseFragmentUnionForSegment(segment, {
    x: eraseX,
    y: eraseY,
    width: eraseWidth,
    height: eraseHeight,
  }, pageWidth, sourceFontSize);
  if (fragmentUnion) {
    const leftSafety = lineCount > 1 ? Math.max(paddingX * 2.6, sourceFontSize * 0.55) : paddingX;
    const bottomSafety = lineCount > 1 ? Math.max(paddingY * 4.8, sourceFontSize * 1.05) : paddingY;
    const topSafety = lineCount > 1 ? Math.max(paddingY * 2.2, sourceFontSize * 0.44) : paddingY;
    const rightSafety = lineCount > 1 ? Math.max(paddingX * 1.6, sourceFontSize * 0.32) : paddingX;
    const right = Math.max(eraseX + eraseWidth, fragmentUnion.x + fragmentUnion.width + rightSafety);
    const top = Math.max(eraseY + eraseHeight, fragmentUnion.y + fragmentUnion.height + topSafety);
    eraseX = Math.max(0, Math.min(eraseX, fragmentUnion.x - leftSafety));
    eraseY = Math.max(0, Math.min(eraseY, fragmentUnion.y - bottomSafety));
    eraseWidth = Math.max(2, Math.min(pageWidth - eraseX - 4, right - eraseX));
    eraseHeight = Math.max(2, top - eraseY);
  }

  return {
    erase: {
      x: eraseX,
      y: eraseY,
      width: eraseWidth,
      height: eraseHeight,
      color: coverColor,
    },
  };
}

function getPdfEraseFragmentUnionForSegment(segment, area, pageWidth, sourceFontSize) {
  const fragments = state.pdfEraseFragments.get(segment.path) || [];
  if (!fragments.length || !area) return null;

  const lineCount = Number(segment.layout?.lineCount || 1);
  const columnKey = segment.layout?.columnKey || "";
  const verticalSlack = Math.max(sourceFontSize * (lineCount > 1 ? 2.2 : 0.7), 6);
  const horizontalSlack = Math.max(sourceFontSize * (lineCount > 1 ? 3.2 : 1), 6);
  const searchArea = {
    x: Math.max(0, Number(area.x || 0) - horizontalSlack),
    y: Math.max(0, Number(area.y || 0) - verticalSlack),
    width: Number(area.width || 0) + horizontalSlack * 2,
    height: Number(area.height || 0) + verticalSlack * 2,
  };
  const matches = fragments.filter((fragment) => {
    const bounds = fragment.bounds;
    if (!bounds) return false;
    const cx = Number(bounds.x || 0) + Number(bounds.width || 0) / 2;
    const cy = Number(bounds.y || 0) + Number(bounds.height || 0) / 2;
    if (!pointInRect(cx, cy, searchArea) && !rectsIntersect(bounds, searchArea)) return false;
    if (columnKey === "left" && cx > pageWidth * 0.54) return false;
    if (columnKey === "right" && cx < pageWidth * 0.44) return false;
    return true;
  });
  if (!matches.length) return null;

  return matches.reduce((bounds, fragment) => bounds ? unionPdfBounds(bounds, fragment.bounds) : { ...fragment.bounds }, null);
}

function drawPdfRawTextErasesForSegment(page, segment, rgb) {
  const fragments = state.pdfEraseFragments.get(segment.path) || [];
  if (!fragments.length) return;

  const sourceErase = createPdfSourceErasePlan(segment);
  const area = sourceErase?.erase;
  if (!area) return;

  const coverColor = segment.layout?.backgroundColor || { r: 1, g: 1, b: 1 };
  const fontSize = Math.max(4, Math.min(28, Number(segment.layout?.fontSize || 10)));
  const padX = Math.max(1, fontSize * 0.16);
  const padY = Math.max(0.8, fontSize * 0.12);

  fragments.forEach((fragment) => {
    const bounds = fragment.bounds;
    if (!bounds || !rectsIntersect(area, bounds)) return;
    page.drawRectangle({
      x: Math.max(0, bounds.x - padX),
      y: Math.max(0, bounds.y - padY),
      width: Math.max(1, bounds.width + padX * 2),
      height: Math.max(1, bounds.height + padY * 2),
      color: rgb(coverColor.r, coverColor.g, coverColor.b),
      opacity: 1,
    });
  });
}

function rectsIntersect(a, b) {
  return Number(a.x || 0) < Number(b.x || 0) + Number(b.width || 0) &&
    Number(a.x || 0) + Number(a.width || 0) > Number(b.x || 0) &&
    Number(a.y || 0) < Number(b.y || 0) + Number(b.height || 0) &&
    Number(a.y || 0) + Number(a.height || 0) > Number(b.y || 0);
}

function pointInRect(x, y, rect) {
  return x >= Number(rect.x || 0) &&
    x <= Number(rect.x || 0) + Number(rect.width || 0) &&
    y >= Number(rect.y || 0) &&
    y <= Number(rect.y || 0) + Number(rect.height || 0);
}

function getPdfOverlayFitBounds(segment, font, text, sourceBounds) {
  const layout = segment.layout || {};
  const sourceWidth = Math.max(4, Number(sourceBounds.width || 0));
  const sourceHeight = Math.max(5, Number(sourceBounds.height || 0));
  const sourceFontSize = Math.max(4, Math.min(28, Number(layout.fontSize || 10)));
  const textWidth = Math.max(0, font.widthOfTextAtSize(String(text || ""), sourceFontSize));
  const rowSegments = Number(layout.rowSegmentCount || 1);
  const lineCount = Number(layout.lineCount || 1);
  const compactRow = rowSegments > 1 || sourceWidth < 110 || sourceFontSize <= 8.5;
  const widthMultiplier = compactRow ? 1.14 : 1.38;
  const widthPad = Math.max(8, sourceFontSize * (compactRow ? 0.8 : 1.4));
  const rawTargetWidth = Math.max(sourceWidth, Math.min(textWidth * 1.04, sourceWidth * widthMultiplier + widthPad));
  const availableWidth = Number(layout.availableWidth || 0);
  const pageSize = state.pdfPageSizes.get(segment.path) || {};
  const pageWidth = Number(pageSize.width || 595.28);
  const columnKey = layout.columnKey || "";
  const columnLeft = columnKey === "right" ? pageWidth * 0.505 : (columnKey === "left" ? pageWidth * 0.055 : 0);
  const columnRight = columnKey === "left" ? pageWidth * 0.49 : (columnKey === "right" ? pageWidth * 0.955 : pageWidth - 8);
  if (layout.isFigureCaption) {
    const captionX = Math.max(36, Math.min(Number(sourceBounds.x || 0), pageWidth * 0.08));
    const captionRight = Math.min(pageWidth - 32, Math.max(Number(sourceBounds.x || 0) + sourceWidth, pageWidth * 0.92));
    return {
      ...sourceBounds,
      x: captionX,
      width: Math.max(sourceWidth, captionRight - captionX),
      height: Math.max(sourceHeight, Math.min(54, sourceHeight * 2.2)),
    };
  }
  const boundedX = columnKey === "left" || columnKey === "right"
    ? clampNumber(Number(sourceBounds.x || 0), columnLeft, Math.max(columnLeft, columnRight - sourceWidth))
    : Number(sourceBounds.x || 0);
  const columnWidthCap = Math.max(sourceWidth, columnRight - boundedX);
  const widthCap = Math.min(
    columnWidthCap,
    availableWidth > sourceWidth ? availableWidth : columnWidthCap
  );
  const targetWidth = Math.min(rawTargetWidth, widthCap);
  const availableHeight = Number(layout.availableHeight || 0);
  const heightMultiplier = lineCount > 1 ? 1.08 : 1.35;
  const targetHeight = Math.max(sourceHeight, Math.min(Math.max(sourceHeight, availableHeight), sourceHeight * heightMultiplier));

  return {
    ...sourceBounds,
    x: boundedX,
    width: Math.max(4, targetWidth),
    height: Math.max(5, targetHeight),
  };
}

function renderOriginalDocumentPreview() {
  if (els.previewTitle) {
    els.previewTitle.textContent = "原始文献预览";
  }
  if (els.previewDownloadButton) els.previewDownloadButton.hidden = true;
  if (els.previewShareButton) {
    els.previewShareButton.hidden = false;
    els.previewShareButton.disabled = !state.file;
    setButtonLabel(els.previewShareButton, "分享原始文件");
    els.previewShareButton.dataset.icon = "share";
    decorateButtonIcon(els.previewShareButton);
  }
  if (els.previewMeta) {
    els.previewMeta.textContent = state.file
      ? `${state.file.name} · ${formatBytes(state.file.size)}`
      : "请先选择文献。";
  }
  els.previewBody.replaceChildren();

  if (!state.file) {
    const empty = document.createElement("div");
    empty.className = "empty-state";
    empty.textContent = "还没有可预览的原始文献。";
    els.previewBody.append(empty);
    return;
  }

  const url = getOriginalPreviewUrl();
  if (state.fileType === "pdf") {
    renderOriginalPdfPreview();
    return;
  }

  if (state.fileType === "pptx") {
    renderOriginalOfficePdfPreview("pptx");
    return;
  }

  if (state.fileType === "docx") {
    renderOriginalOfficePdfPreview("docx");
    return;
  }

  els.previewBody.append(
    createOriginalPreviewPanel(
      "无法预览此格式",
      "此格式无法在浏览器内直接预览，可使用下方入口打开或下载原始文件。",
      url
    )
  );
}

function renderOriginalPdfPreview() {
  els.previewBody.append(createPdfRegionSelectionPanel());

  const pages = document.createElement("div");
  pages.className = "original-pdf-pages";
  pages.append(createOriginalPreviewFallback("正在渲染 PDF 原始页面..."));
  els.previewBody.append(pages);

  renderOriginalPdfPages(pages).catch((error) => {
    console.warn("Original PDF preview failed", error);
    pages.replaceChildren(createOriginalPreviewFallback("PDF 页面渲染失败。"));
  });
}

async function renderOriginalPdfPages(container) {
  if (!state.pdfBytes) throw new Error("Missing PDF bytes.");
  const pdfjs = await loadPdfJs();
  const pdf = await pdfjs.getDocument({ data: state.pdfBytes.slice() }).promise;
  container.replaceChildren();
  await waitForUiFrame();

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const pdfPage = await pdf.getPage(pageNumber);
    const baseViewport = pdfPage.getViewport({ scale: 1 });
    const metrics = getOriginalPdfPreviewMetrics(container, baseViewport);
    const viewport = pdfPage.getViewport({ scale: metrics.renderScale });
    const pageShell = document.createElement("section");
    pageShell.className = "original-pdf-page";
    pageShell.style.width = `${metrics.cssWidth}px`;
    pageShell.dataset.pagePath = `pdf/page-${pageNumber}`;

    const label = document.createElement("span");
    label.textContent = `第 ${pageNumber} 页`;

    const frame = document.createElement("div");
    frame.className = "original-pdf-page-frame";
    frame.style.width = `${metrics.cssWidth}px`;
    frame.style.height = `${metrics.cssHeight}px`;

    const canvas = document.createElement("canvas");
    canvas.width = Math.ceil(viewport.width);
    canvas.height = Math.ceil(viewport.height);
    canvas.style.width = `${metrics.cssWidth}px`;
    canvas.style.height = `${metrics.cssHeight}px`;
    canvas.style.aspectRatio = `${baseViewport.width} / ${baseViewport.height}`;
    const context = canvas.getContext("2d", { alpha: false });
    if (!context) throw new Error("Canvas unavailable.");
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, canvas.width, canvas.height);

    const selectionLayer = createPdfRegionSelectionLayer({
      pagePath: `pdf/page-${pageNumber}`,
      pageWidth: baseViewport.width,
      pageHeight: baseViewport.height,
      cssWidth: metrics.cssWidth,
      cssHeight: metrics.cssHeight,
    });
    frame.append(canvas, selectionLayer);
    pageShell.append(label, frame);
    container.append(pageShell);
    await pdfPage.render({ canvasContext: context, viewport }).promise;
    renderPdfManualRegionBoxes(selectionLayer);
    await waitForUiFrame();
  }

  if (!container.children.length) {
    container.append(createOriginalPreviewFallback("PDF 没有可渲染页面。"));
  }
}

function createPdfRegionSelectionPanel() {
  const panel = document.createElement("section");
  panel.className = "pdf-region-tool";

  const copy = document.createElement("div");
  copy.className = "pdf-region-tool-copy";
  const title = document.createElement("h3");
  title.textContent = "图片原样复制区域";
  const detail = document.createElement("p");
  detail.textContent = "点击“开始框选”后，在页面上拖出需要原样保留的图片/图题区域；同一页有手动框选时，会优先使用手动框，不再自动整块保留图区。";
  const count = document.createElement("span");
  count.className = "pdf-region-count";
  copy.append(title, detail, count);

  const actions = document.createElement("div");
  actions.className = "pdf-region-tool-actions";
  const toggle = document.createElement("button");
  toggle.type = "button";
  toggle.dataset.pdfRegionToggle = "true";
  const clear = document.createElement("button");
  clear.type = "button";
  clear.dataset.pdfRegionClear = "true";
  clear.textContent = "清空框选";
  toggle.addEventListener("click", () => {
    state.pdfManualRegionSelectionEnabled = !state.pdfManualRegionSelectionEnabled;
    updatePdfRegionSelectionControls();
  });
  clear.addEventListener("click", () => {
    clearPdfManualImageRegions();
    document.querySelectorAll(".pdf-region-layer").forEach((layer) => renderPdfManualRegionBoxes(layer));
    updatePdfRegionSelectionControls();
    scheduleCurrentDraftSave();
  });
  actions.append(toggle, clear);
  panel.append(copy, actions);
  updatePdfRegionSelectionControls(panel);
  return panel;
}

function createPdfRegionSelectionLayer({ pagePath, pageWidth, pageHeight, cssWidth, cssHeight }) {
  const layer = document.createElement("div");
  layer.className = "pdf-region-layer";
  layer.dataset.pagePath = pagePath;
  layer.dataset.pageWidth = String(pageWidth);
  layer.dataset.pageHeight = String(pageHeight);
  layer.dataset.cssWidth = String(cssWidth);
  layer.dataset.cssHeight = String(cssHeight);

  let drag = null;

  layer.addEventListener("pointerdown", (event) => {
    if (!state.pdfManualRegionSelectionEnabled) return;
    if (event.button !== undefined && event.button !== 0) return;
    event.preventDefault();
    const start = getPdfRegionLayerPoint(layer, event);
    const draft = document.createElement("div");
    draft.className = "pdf-region-draft";
    layer.append(draft);
    drag = { start, current: start, draft };
    layer.setPointerCapture?.(event.pointerId);
    updatePdfRegionDraftBox(drag);
  });

  layer.addEventListener("pointermove", (event) => {
    if (!drag) return;
    event.preventDefault();
    drag.current = getPdfRegionLayerPoint(layer, event);
    updatePdfRegionDraftBox(drag);
  });

  const finishDrag = (event) => {
    if (!drag) return;
    event.preventDefault();
    const rect = getPdfRegionCssRect(drag.start, drag.current);
    drag.draft.remove();
    drag = null;
    layer.releasePointerCapture?.(event.pointerId);
    if (rect.width < 8 || rect.height < 8) return;
    addPdfManualPreservedRegionFromCss(layer, rect);
    renderAllPdfManualRegionBoxes();
    updatePdfRegionSelectionControls();
    scheduleCurrentDraftSave();
  };

  layer.addEventListener("pointerup", finishDrag);
  layer.addEventListener("pointercancel", (event) => {
    if (!drag) return;
    event.preventDefault();
    drag.draft.remove();
    drag = null;
  });

  return layer;
}

function getPdfRegionLayerPoint(layer, event) {
  const rect = layer.getBoundingClientRect();
  const x = clampNumber(event.clientX - rect.left, 0, rect.width);
  const y = clampNumber(event.clientY - rect.top, 0, rect.height);
  return { x, y };
}

function getPdfRegionCssRect(a, b) {
  const left = Math.min(a.x, b.x);
  const top = Math.min(a.y, b.y);
  return {
    left,
    top,
    width: Math.abs(a.x - b.x),
    height: Math.abs(a.y - b.y),
  };
}

function updatePdfRegionDraftBox(drag) {
  const rect = getPdfRegionCssRect(drag.start, drag.current);
  Object.assign(drag.draft.style, {
    left: `${rect.left}px`,
    top: `${rect.top}px`,
    width: `${rect.width}px`,
    height: `${rect.height}px`,
  });
}

function addPdfManualPreservedRegionFromCss(layer, rect) {
  const pagePath = layer.dataset.pagePath || "";
  const pageWidth = Number(layer.dataset.pageWidth || 0);
  const pageHeight = Number(layer.dataset.pageHeight || 0);
  const cssWidth = Number(layer.dataset.cssWidth || layer.getBoundingClientRect().width || 1);
  const cssHeight = Number(layer.dataset.cssHeight || layer.getBoundingClientRect().height || 1);
  if (!pagePath || !pageWidth || !pageHeight) return;

  const x = clampNumber((rect.left / cssWidth) * pageWidth, 0, pageWidth);
  const width = clampNumber((rect.width / cssWidth) * pageWidth, 1, pageWidth - x);
  const topPdf = pageHeight - (rect.top / cssHeight) * pageHeight;
  const bottomPdf = pageHeight - ((rect.top + rect.height) / cssHeight) * pageHeight;
  const y = clampNumber(bottomPdf, 0, pageHeight);
  const height = clampNumber(topPdf - bottomPdf, 1, pageHeight - y);
  const regions = state.pdfManualPreservedRegions.get(pagePath) || [];
  regions.push({
    id: `manual-${Date.now()}-${Math.random().toString(16).slice(2)}`,
    kind: "manual",
    x,
    y,
    width,
    height,
  });
  state.pdfManualPreservedRegions.set(pagePath, regions);
}

function renderPdfManualRegionBoxes(layer) {
  layer.querySelectorAll(".pdf-manual-region-box").forEach((box) => box.remove());
  const pagePath = layer.dataset.pagePath || "";
  const regions = getPdfManualImageRegions(pagePath);
  const pageWidth = Number(layer.dataset.pageWidth || 1);
  const pageHeight = Number(layer.dataset.pageHeight || 1);
  const cssWidth = Number(layer.dataset.cssWidth || layer.getBoundingClientRect().width || 1);
  const cssHeight = Number(layer.dataset.cssHeight || layer.getBoundingClientRect().height || 1);

  regions.forEach((region) => {
    const box = document.createElement("div");
    box.className = "pdf-manual-region-box";
    box.style.left = `${(Number(region.x || 0) / pageWidth) * cssWidth}px`;
    box.style.top = `${((pageHeight - Number(region.y || 0) - Number(region.height || 0)) / pageHeight) * cssHeight}px`;
    box.style.width = `${(Number(region.width || 0) / pageWidth) * cssWidth}px`;
    box.style.height = `${(Number(region.height || 0) / pageHeight) * cssHeight}px`;
    const label = document.createElement("span");
    label.textContent = `保留 ${getPdfManualRegionDisplayNumber(pagePath, region.id)}`;
    const remove = document.createElement("button");
    remove.type = "button";
    remove.title = "删除此区域";
    remove.setAttribute("aria-label", "删除此区域");
    remove.textContent = "×";
    remove.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      removePdfManualPreservedRegion(pagePath, region.id);
      renderAllPdfManualRegionBoxes();
      updatePdfRegionSelectionControls();
      scheduleCurrentDraftSave();
    });
    box.append(label, remove);
    layer.append(box);
  });
}

function getPdfManualRegionDisplayNumber(pagePath, regionId) {
  const ordered = getOrderedPdfManualRegions();
  const index = ordered.findIndex((item) => item.pagePath === pagePath && item.region.id === regionId);
  return index >= 0 ? index + 1 : ordered.length + 1;
}

function getOrderedPdfManualRegions() {
  return [...state.pdfManualPreservedRegions.entries()]
    .sort((a, b) => getPdfPageNumberFromPath(a[0]) - getPdfPageNumberFromPath(b[0]))
    .flatMap(([pagePath, regions]) => (regions || []).filter((region) => !isPdfPageFurnitureRegion(region)).map((region) => ({ pagePath, region })));
}

function getPdfPageNumberFromPath(pagePath) {
  const match = String(pagePath || "").match(/page-(\d+)/);
  return match ? Number(match[1]) : Number.MAX_SAFE_INTEGER;
}

function renderAllPdfManualRegionBoxes() {
  document.querySelectorAll(".pdf-region-layer").forEach((layer) => renderPdfManualRegionBoxes(layer));
}

function removePdfManualPreservedRegion(pagePath, regionId) {
  const regions = state.pdfManualPreservedRegions.get(pagePath) || [];
  const next = regions.filter((region) => region.id !== regionId);
  if (next.length) state.pdfManualPreservedRegions.set(pagePath, next);
  else state.pdfManualPreservedRegions.delete(pagePath);
}

function clearPdfManualImageRegions() {
  const next = new Map();
  state.pdfManualPreservedRegions.forEach((regions, pagePath) => {
    const furniture = (regions || []).filter(isPdfPageFurnitureRegion);
    if (furniture.length) next.set(pagePath, furniture);
  });
  state.pdfManualPreservedRegions = next;
}

function updatePdfRegionSelectionControls(root = document) {
  const count = getPdfManualPreservedRegionCount();
  const enabled = Boolean(state.pdfManualRegionSelectionEnabled);
  root.querySelectorAll?.("[data-pdf-region-toggle]").forEach((button) => {
    button.textContent = enabled ? "结束框选" : "开始框选";
    button.classList.toggle("active", enabled);
  });
  root.querySelectorAll?.("[data-pdf-region-clear]").forEach((button) => {
    button.disabled = count === 0;
  });
  root.querySelectorAll?.(".pdf-region-count").forEach((item) => {
    item.textContent = count ? `已框选 ${count} 个区域` : "还没有手动框选区域";
  });
  els.previewBody?.classList.toggle("pdf-region-selecting", enabled);
}

function getPdfManualPreservedRegionCount() {
  return [...state.pdfManualPreservedRegions.values()].reduce((sum, regions) => sum + (regions || []).filter((region) => !isPdfPageFurnitureRegion(region)).length, 0);
}

function renderOriginalOfficePdfPreview(kind) {
  const typeLabel = kind === "pptx" ? "PPTX" : "DOCX";
  const pages = document.createElement("div");
  pages.className = "original-pdf-pages";
  pages.append(createOriginalPreviewFallback("正在生成预览..."));
  els.previewBody.append(pages);

  createOriginalOfficePreviewPdfBlob(kind)
    .then((blob) => blob.arrayBuffer())
    .then((buffer) => renderPdfBytesIntoPreview(
      pages,
      new Uint8Array(buffer),
      "没有可渲染页面。"
    ))
    .catch((error) => {
      console.warn("Original Office PDF preview failed", error);
      pages.replaceChildren(
        createOriginalPreviewFallback(
          "预览生成失败，可回到文件列表后重新选择文件。"
        )
      );
    });
}

async function createOriginalOfficePreviewPdfBlob(kind) {
  if (!state.zip) throw new Error("Missing Office document data.");
  const sections = kind === "docx"
    ? await extractSavedDocxTextSections(state.zip)
    : await extractSavedPptxTextSections(state.zip);

  if (!sections.some((section) => section.lines.length)) {
    throw new Error("No extractable text for Office PDF preview.");
  }

  return createTextPdfBlob({
    title: state.file?.name || "original-file",
    subtitle: `${kind === "pptx" ? "PPTX" : "DOCX"} 原文 PDF 预览版（文本版）`,
    sections,
    showHeader: false,
  });
}

async function renderPdfBytesIntoPreview(container, pdfBytes, emptyMessage) {
  const pdfjs = await loadPdfJs();
  const pdf = await pdfjs.getDocument({ data: pdfBytes }).promise;
  container.replaceChildren();
  await waitForUiFrame();

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const pdfPage = await pdf.getPage(pageNumber);
    const baseViewport = pdfPage.getViewport({ scale: 1 });
    const metrics = getOriginalPdfPreviewMetrics(container, baseViewport);
    const viewport = pdfPage.getViewport({ scale: metrics.renderScale });
    const pageShell = document.createElement("section");
    pageShell.className = "original-pdf-page";
    pageShell.style.width = `${metrics.cssWidth}px`;

    const label = document.createElement("span");
    label.textContent = `第 ${pageNumber} 页`;

    const canvas = document.createElement("canvas");
    canvas.width = Math.ceil(viewport.width);
    canvas.height = Math.ceil(viewport.height);
    canvas.style.width = `${metrics.cssWidth}px`;
    canvas.style.height = `${metrics.cssHeight}px`;
    canvas.style.aspectRatio = `${baseViewport.width} / ${baseViewport.height}`;
    const context = canvas.getContext("2d", { alpha: false });
    if (!context) throw new Error("Canvas unavailable.");
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, canvas.width, canvas.height);

    pageShell.append(label, canvas);
    container.append(pageShell);
    await pdfPage.render({ canvasContext: context, viewport }).promise;
    await waitForUiFrame();
  }

  if (!container.children.length) {
    container.append(createOriginalPreviewFallback(emptyMessage));
  }
}

function getOriginalPdfPreviewMetrics(container, viewport) {
  const containerWidth = Math.floor(container.getBoundingClientRect().width || container.clientWidth || 0);
  const bodyWidth = Math.floor(els.previewBody?.getBoundingClientRect().width || 0);
  const viewportWidth = Math.floor(window.visualViewport?.width || window.innerWidth || 0);
  const fallbackWidth = Math.max(1, viewportWidth - (isLikelyMobileDevice() ? 20 : 56));
  const availableWidth = Math.max(1, containerWidth || bodyWidth || fallbackWidth);
  const cssWidth = Math.floor(Math.min(availableWidth, isLikelyMobileDevice() ? availableWidth : 880));
  const cssHeight = Math.max(1, Math.round(cssWidth * ((viewport.height || 1) / Math.max(1, viewport.width || 1))));
  const pixelRatio = Math.max(1, Math.min(window.devicePixelRatio || 1, isLikelyMobileDevice() ? 2.25 : 1.8));
  const renderScale = (cssWidth / Math.max(1, viewport.width || 1)) * pixelRatio;
  return { cssWidth, cssHeight, renderScale };
}

function renderOriginalPptxPreview(url) {
  els.previewBody.append(
    createOriginalPreviewPanel(
      "PPTX 原文版式预览",
      "浏览器无法直接执行 PowerPoint 渲染；下方使用已解析的文本框坐标和图片占位还原原文版式，用于快速核对源文件内容。",
      url
    )
  );

  const previousMode = state.previewMode;
  state.previewMode = "source";
  groupSegmentsForPreview().forEach((group) => {
    const section = document.createElement("section");
    section.className = "preview-card original-layout-card";
    const heading = document.createElement("h3");
    heading.textContent = group.label;
    section.append(heading);
    try {
      section.append(createSlidePreview(group.segments));
    } catch (error) {
      console.warn("Original PPTX preview fallback", error);
      group.segments.forEach((segment) => section.append(createOriginalPreviewItem(segment)));
    }
    els.previewBody.append(section);
  });
  state.previewMode = previousMode;
}

function renderOriginalDocxPreview(url) {
  els.previewBody.append(
    createOriginalPreviewPanel(
      "DOCX 原文内容预览",
      "浏览器无法完整复刻 Word/WPS 分页和样式；下方展示已解析出的完整原文段落，用于核对源文件内容。",
      url
    )
  );

  const doc = document.createElement("article");
  doc.className = "original-docx-preview";
  groupSegmentsForPreview().forEach((group) => {
    const section = document.createElement("section");
    const heading = document.createElement("h3");
    heading.textContent = group.label;
    section.append(heading);
    group.segments.forEach((segment) => {
      const paragraph = document.createElement("p");
      paragraph.textContent = segment.original;
      section.append(paragraph);
    });
    doc.append(section);
  });
  els.previewBody.append(doc);
}

function createOriginalPreviewPanel(titleText, messageText, url) {
  const panel = document.createElement("section");
  panel.className = "original-file-panel";
  const copy = document.createElement("div");
  copy.className = "original-file-copy";
  const title = document.createElement("h3");
  title.textContent = titleText;
  const message = document.createElement("p");
  message.textContent = messageText;
  copy.append(title, message);
  panel.append(copy, createOriginalFileActions(url));
  return panel;
}

function createOriginalPreviewItem(segment) {
  const item = document.createElement("article");
  item.className = "preview-item source";
  const text = document.createElement("p");
  text.textContent = segment.original;
  const meta = document.createElement("span");
  meta.textContent = "原文片段";
  item.append(text, meta);
  return item;
}

function createOriginalPreviewFallback(message, options = {}) {
  const fallback = document.createElement("div");
  fallback.className = "preview-fallback";
  if (!options.loading) {
    fallback.textContent = message;
    return fallback;
  }

  fallback.classList.add("loading");
  fallback.setAttribute("role", "status");
  fallback.setAttribute("aria-live", "polite");
  const spinner = document.createElement("span");
  spinner.className = "preview-loading-spinner";
  spinner.setAttribute("aria-hidden", "true");
  const copy = document.createElement("div");
  copy.className = "preview-loading-copy";
  const title = document.createElement("strong");
  title.className = "preview-loading-title";
  title.textContent = message;
  const detail = document.createElement("span");
  detail.className = "preview-loading-detail";
  detail.textContent = options.detail || "正在处理，请保持页面打开。";
  const bar = document.createElement("span");
  bar.className = "preview-loading-bar";
  bar.setAttribute("aria-hidden", "true");
  copy.append(title, detail, bar);
  fallback.append(spinner, copy);
  return fallback;
}

function updatePreviewLoadingFallback(fallback, message, detail = "") {
  if (!fallback) return;
  const title = fallback.querySelector(".preview-loading-title");
  const detailNode = fallback.querySelector(".preview-loading-detail");
  if (title && message) title.textContent = message;
  if (detailNode && detail) detailNode.textContent = detail;
}

function renderOriginalDocumentPreviewLegacy() {
  if (els.previewTitle) {
    els.previewTitle.textContent = "原始文献预览";
  }
  if (els.previewDownloadButton) els.previewDownloadButton.hidden = true;
  if (els.previewShareButton) els.previewShareButton.hidden = true;
  if (els.previewMeta) {
    els.previewMeta.textContent = state.file
      ? `${state.file.name} · ${formatBytes(state.file.size)} · 原始文件核对`
      : "请先选择文献。";
  }
  els.previewBody.replaceChildren();

  if (!state.file) {
    const empty = document.createElement("div");
    empty.className = "empty-state";
    empty.textContent = "还没有可预览的原始文献。";
    els.previewBody.append(empty);
    return;
  }

  const url = getOriginalPreviewUrl();
  if (state.fileType === "pdf") {
    const panel = document.createElement("section");
    panel.className = "original-file-panel";
    const title = document.createElement("h3");
    title.textContent = "原始 PDF 文件";
    const message = document.createElement("p");
    message.textContent = "下方显示的是上传的原始 PDF，用于核对是否为需要翻译的文献。若手机浏览器不显示内嵌 PDF，可直接打开原始文件。";
    panel.append(title, message, createOriginalFileActions(url));
    els.previewBody.append(panel);

    const viewer = document.createElement("iframe");
    viewer.className = "original-pdf-viewer";
    viewer.title = "原始 PDF 文献预览";
    viewer.src = url;
    els.previewBody.append(viewer);
    return;
  }

  const panel = document.createElement("section");
  panel.className = "original-file-panel";

  const title = document.createElement("h3");
  title.textContent = "浏览器无法完整内嵌预览此格式";
  const message = document.createElement("p");
  message.textContent = "DOCX/PPTX 的原始版式需要用 Word、PowerPoint 或 WPS 打开核对。下面提供原始文件入口，并列出已提取的原文片段用于快速确认。";

  panel.append(title, message, createOriginalFileActions(url));
  els.previewBody.append(panel);

  const groups = groupSegmentsForPreview();
  groups.slice(0, 8).forEach((group) => {
    const section = document.createElement("section");
    section.className = "preview-card";
    const heading = document.createElement("h3");
    heading.textContent = group.label;
    section.append(heading);
    group.segments.slice(0, 8).forEach((segment) => {
      const item = document.createElement("article");
      item.className = "preview-item source";
      const text = document.createElement("p");
      text.textContent = segment.original;
      const meta = document.createElement("span");
      meta.textContent = "原文片段";
      item.append(text, meta);
      section.append(item);
    });
    els.previewBody.append(section);
  });
}

function createOriginalFileActions(url) {
  const actions = document.createElement("div");
  actions.className = "original-file-actions";
  const openLink = document.createElement("a");
  openLink.href = url;
  openLink.target = "_blank";
  openLink.rel = "noopener";
  openLink.textContent = "打开原始文件";
  const downloadLink = document.createElement("a");
  downloadLink.href = url;
  downloadLink.download = state.file?.name || "source-document";
  downloadLink.textContent = "下载原始文件";
  actions.append(openLink, downloadLink);
  return actions;
}

function renderPdfLayoutPreview() {
  if (els.previewTitle) {
    els.previewTitle.textContent = "PDF 真实排版预览/调整";
  }
  const stats = getPdfLayoutPreviewStats();
  if (els.previewMeta) {
    els.previewMeta.textContent = `PDF · 已解析 ${stats.parsed} 段 · 可定位 ${stats.positioned} 段 · 将回写 ${stats.exportable} 段 · 跳过 ${stats.skipped} 段 · 无坐标 ${stats.unpositioned} 段 · ${getPdfManualPreservedRegionCount()} 个手动图片区域`;
  }

  const panel = document.createElement("section");
  panel.className = "pdf-layout-tool";
  const title = document.createElement("h3");
  title.textContent = "局部调整";
  const detail = document.createElement("p");
  detail.textContent = "默认只显示系统会回写的译文层和手动图片框。原文底图仅用于对齐参考，不参与导出。";
  const actions = document.createElement("div");
  actions.className = "pdf-layout-tool-actions";
  title.textContent = "真实排版校准";
  detail.textContent = "页面底图由当前导出 PDF 实时渲染；上层只显示可校准边框，避免文字重复堆叠。";
  const sourceToggle = document.createElement("button");
  sourceToggle.type = "button";
  sourceToggle.textContent = state.pdfLayoutShowSourceBackground ? "隐藏编辑文字" : "显示编辑文字";
  sourceToggle.classList.toggle("active", state.pdfLayoutShowSourceBackground);
  sourceToggle.addEventListener("click", () => {
    state.pdfLayoutShowSourceBackground = !state.pdfLayoutShowSourceBackground;
    rerenderPreviewIfOpen();
  });
  const parsedToggle = document.createElement("button");
  parsedToggle.type = "button";
  parsedToggle.textContent = state.pdfLayoutShowParsedMap ? "隐藏解析范围" : "显示解析范围";
  parsedToggle.classList.toggle("active", state.pdfLayoutShowParsedMap);
  parsedToggle.addEventListener("click", () => {
    state.pdfLayoutShowParsedMap = !state.pdfLayoutShowParsedMap;
    rerenderPreviewIfOpen();
  });
  const refreshButton = document.createElement("button");
  refreshButton.type = "button";
  refreshButton.textContent = "刷新真实预览";
  refreshButton.title = "重新生成真实 PDF 预览。移动或缩放单个框时不会自动刷新，避免长时间等待。";
  refreshButton.addEventListener("click", () => {
    rerenderPreviewIfOpen();
  });
  actions.append(sourceToggle, parsedToggle, refreshButton);
  panel.append(title, detail, actions);
  els.previewBody.append(panel);

  const pages = document.createElement("div");
  pages.className = "pdf-layout-pages";
  const loadingFallback = createOriginalPreviewFallback("正在生成真实 PDF 预览", {
    loading: true,
    detail: "会先生成一份和导出完全一致的 PDF，再渲染到预览区；大文件可能需要几十秒。",
  });
  pages.replaceChildren(loadingFallback);
  els.previewBody.append(pages);

  renderPdfLayoutPreviewPages(pages, (message, detail) => {
    updatePreviewLoadingFallback(loadingFallback, message, detail);
  }).catch((error) => {
    console.warn("PDF layout preview failed", error);
    pages.replaceChildren(createOriginalPreviewFallback("PDF 排版预览生成失败。"));
  });
}

async function renderPdfLayoutPreviewPages(container, onProgress = null) {
  if (!state.pdfBytes) throw new Error("Missing PDF bytes.");
  onProgress?.("正在加载真实预览组件", "正在准备 PDF 渲染库和中文字体，首次打开会稍慢。");
  const pdfjs = await loadPdfJs();
  const { PDFDocument, fontkit, fontBytes } = await loadPdfExportTools();
  const previewDoc = await PDFDocument.create();
  previewDoc.registerFontkit(fontkit);
  const previewFont = await previewDoc.embedFont(fontBytes, { subset: false });
  onProgress?.("正在生成真实 PDF 预览", "这一步会调用实际导出 PDF 的同一套逻辑，用来保证预览和导出一致。");
  const previewBlob = await downloadPdfOverlayTranslation();
  onProgress?.("正在读取真实 PDF 页面", "已生成临时 PDF，正在交给浏览器渲染预览页面。");
  const previewBytes = await previewBlob.arrayBuffer();
  const pdf = await pdfjs.getDocument({ data: previewBytes.slice(0) }).promise;
  container.replaceChildren();
  await waitForUiFrame();

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    onProgress?.(
      `正在渲染真实预览：第 ${pageNumber}/${pdf.numPages} 页`,
      "页面会陆续显示出来；上层只保留可校准边框，避免文字重复堆叠。"
    );
    const pdfPage = await pdf.getPage(pageNumber);
    const baseViewport = pdfPage.getViewport({ scale: 1 });
    const metrics = getOriginalPdfPreviewMetrics(container, baseViewport);
    const viewport = pdfPage.getViewport({ scale: metrics.renderScale });
    const pagePath = `pdf/page-${pageNumber}`;
    const pageShell = document.createElement("section");
    pageShell.className = "pdf-layout-page";
    pageShell.style.width = `${metrics.cssWidth}px`;

    const label = document.createElement("span");
    label.textContent = `第 ${pageNumber} 页`;

    const frame = document.createElement("div");
    frame.className = "pdf-layout-page-frame";
    frame.classList.add("rendered-output");
    frame.classList.toggle("source-visible", state.pdfLayoutShowSourceBackground);
    frame.style.width = `${metrics.cssWidth}px`;
    frame.style.height = `${metrics.cssHeight}px`;

    const canvas = document.createElement("canvas");
    canvas.width = Math.ceil(viewport.width);
    canvas.height = Math.ceil(viewport.height);
    canvas.style.width = `${metrics.cssWidth}px`;
    canvas.style.height = `${metrics.cssHeight}px`;
    canvas.style.aspectRatio = `${baseViewport.width} / ${baseViewport.height}`;
    const context = canvas.getContext("2d", { alpha: false });
    if (!context) throw new Error("Canvas unavailable.");
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, canvas.width, canvas.height);

    const overlay = document.createElement("div");
    overlay.className = "pdf-layout-overlay";
    overlay.dataset.pagePath = pagePath;
    overlay.dataset.pageWidth = String(baseViewport.width);
    overlay.dataset.pageHeight = String(baseViewport.height);
    overlay.dataset.cssWidth = String(metrics.cssWidth);
    overlay.dataset.cssHeight = String(metrics.cssHeight);

    frame.append(canvas, overlay);
    pageShell.append(label, frame);
    container.append(pageShell);
    await pdfPage.render({ canvasContext: context, viewport }).promise;
    renderPdfLayoutTextBoxes(overlay, pagePath, previewFont);
    renderPdfLayoutImageBoxes(overlay, pagePath);
    renderPdfLayoutFurnitureBoxes(overlay, pagePath);
    await waitForUiFrame();
  }

  if (!container.children.length) {
    container.append(createOriginalPreviewFallback("PDF 没有可渲染页面。"));
  }
}

function renderPdfLayoutTextBoxes(overlay, pagePath, font = null) {
  const pageWidth = Number(overlay.dataset.pageWidth || 1);
  const pageHeight = Number(overlay.dataset.pageHeight || 1);
  const cssWidth = Number(overlay.dataset.cssWidth || 1);
  const cssHeight = Number(overlay.dataset.cssHeight || 1);
  const segments = state.segments
    .filter((segment) => segment.type === "pdf" && segment.path === pagePath && segment.layout?.bounds);
  const planEntries = [];
  const plansBySegmentId = new Map();

  if (font) {
    segments.forEach((segment) => {
      const plan = createPdfOverlayPlan(segment, font);
      if (!plan) return;
      plansBySegmentId.set(segment.id, plan);
      planEntries.push({ page: pagePath, plan });
    });
    resolvePdfOverlayPlanCollisions(planEntries);
  }

  segments.forEach((segment) => {
      const info = getPdfLayoutPreviewTextInfo(segment);
      const plan = plansBySegmentId.get(segment.id);
      if (!plan && !info.text) return;
      if (!plan && !info.exportable && !state.pdfLayoutShowParsedMap) return;
      const bounds = segment.layout.bounds;
      const box = document.createElement("div");
      box.className = `pdf-layout-text-box ${plan ? "exportable" : "skipped"}`;
      box.title = info.reason;
      positionPdfLayoutBox(box, plan ? plan.fitBounds : bounds, pageWidth, pageHeight, cssWidth, cssHeight);
      const content = document.createElement("div");
      content.className = "pdf-layout-text-content";
      content.textContent = plan ? plan.lines.join("\n") : info.text;
      if (plan) {
        applyPdfLayoutTextPlanStyles(content, box, plan, pageWidth, pageHeight, cssWidth, cssHeight);
        setPdfLayoutTextContentHidden(content, !state.pdfLayoutShowSourceBackground);
        content.contentEditable = "true";
        content.spellcheck = false;
        content.title = "可直接编辑译文";
        content.addEventListener("blur", () => {
          segment.translation = content.textContent.trim();
          scheduleCurrentDraftSave();
          renderSegments();
        });
        const moveHandle = document.createElement("span");
        moveHandle.className = "pdf-layout-text-drag";
        moveHandle.textContent = "移动";
        moveHandle.title = "拖动移动此段译文";
        attachPdfLayoutTextMove(moveHandle, box, segment, overlay);
        const handle = document.createElement("i");
        handle.className = "pdf-layout-text-resize-handle";
        handle.setAttribute("aria-hidden", "true");
        attachPdfLayoutTextResize(handle, box, segment, overlay);
        const reset = document.createElement("button");
        reset.type = "button";
        reset.className = "pdf-layout-text-reset";
        reset.textContent = "重置";
        reset.title = "恢复此段译文位置和大小";
        reset.addEventListener("click", (event) => {
          event.preventDefault();
          event.stopPropagation();
          delete segment.overrides.pdfBounds;
          positionPdfLayoutBoxFromBounds(box, segment.layout.bounds, overlay);
          clearPdfLayoutBoxLocalUpdate(box);
          setPdfLayoutTextContentHidden(content, !state.pdfLayoutShowSourceBackground);
          setStatus("已恢复默认位置。导出/分享时会使用默认排版；如需重新生成画布，请点“刷新真实预览”。");
          scheduleCurrentDraftSave();
        });
        box.append(content, moveHandle, reset, handle);
      } else {
        box.append(content);
      }
      overlay.append(box);
    });
}

function applyPdfLayoutTextPlanStyles(content, box, plan, pageWidth, pageHeight, cssWidth, cssHeight) {
  const scale = cssWidth / Math.max(1, pageWidth);
  const boxTop = ((pageHeight - Number(plan.fitBounds.y || 0) - Number(plan.fitBounds.height || 0)) / pageHeight) * cssHeight;
  const textTopPdf = Number(plan.y || 0) + Math.max(1, Number(plan.fontSize || 0) * 0.28);
  const textTop = ((pageHeight - textTopPdf) / pageHeight) * cssHeight;
  const contentTop = clampNumber(textTop - boxTop, 0, Math.max(0, cssHeight));
  content.style.position = "absolute";
  content.style.left = "2px";
  content.style.right = "2px";
  content.style.top = `${contentTop}px`;
  content.style.height = "auto";
  content.style.fontSize = `${Math.max(1, Number(plan.fontSize || 8) * scale)}px`;
  content.style.lineHeight = `${Math.max(1, Number(plan.lineHeight || 10) * scale)}px`;
  content.style.fontWeight = "600";
  content.style.textAlign = plan.cellAlign === "center" ? "center" : "left";
  content.style.whiteSpace = "pre-wrap";
}

function getPdfLayoutPreviewTextInfo(segment) {
  const exportText = getPdfExportText(segment);
  if (exportText) {
    return {
      exportable: true,
      text: exportText,
      reason: "将回写到译文 PDF",
    };
  }

  const sourceText = sanitizePdfExportText(segment.original || "");
  if (!sourceText) return { exportable: false, text: "", reason: "空文本" };
  if (isInsidePdfPreservedRegion(segment)) {
    return { exportable: false, text: sourceText, reason: "已解析，但位于图片/图题/参考文献保留区域，导出时原样保留" };
  }
  if (shouldKeepPdfSourceText(segment)) {
    return { exportable: false, text: sourceText, reason: "已解析，但属于页眉页脚、参考文献或出版信息，导出时不重写" };
  }
  if (!String(segment.translation || "").trim()) {
    return { exportable: false, text: sourceText, reason: "已解析，但还没有译文" };
  }
  return { exportable: false, text: sourceText, reason: "已解析，但当前译文未通过目标语言校验，导出时跳过" };
}

function getPdfLayoutPreviewStats() {
  const pdfSegments = state.segments.filter((segment) => segment.type === "pdf");
  const positionedSegments = pdfSegments.filter((segment) => segment.layout?.bounds);
  const exportable = positionedSegments.filter((segment) => Boolean(getPdfExportText(segment))).length;
  return {
    parsed: pdfSegments.length,
    positioned: positionedSegments.length,
    exportable,
    skipped: Math.max(0, positionedSegments.length - exportable),
    unpositioned: Math.max(0, pdfSegments.length - positionedSegments.length),
  };
}

function renderPdfLayoutImageBoxes(overlay, pagePath) {
  const regions = getPdfManualImageRegions(pagePath);
  const automaticRegions = getPdfAutomaticFigureRegions(pagePath).filter((region) =>
    !regions.some((manual) => rectsIntersect(getPdfRegionSourceBounds(manual), getPdfRegionSourceBounds(region)))
  );
  const pageWidth = Number(overlay.dataset.pageWidth || 1);
  const pageHeight = Number(overlay.dataset.pageHeight || 1);
  const cssWidth = Number(overlay.dataset.cssWidth || 1);
  const cssHeight = Number(overlay.dataset.cssHeight || 1);

  regions.forEach((region) => {
    const box = document.createElement("div");
    box.className = "pdf-layout-image-box manual";
    box.dataset.pagePath = pagePath;
    box.dataset.regionId = region.id;
    positionPdfLayoutBox(box, getPdfRegionTargetBounds(region), pageWidth, pageHeight, cssWidth, cssHeight);

    const label = document.createElement("span");
    label.textContent = `图片 ${getPdfManualRegionDisplayNumber(pagePath, region.id)}`;

    const reset = document.createElement("button");
    reset.type = "button";
    reset.textContent = "重置";
    reset.title = "恢复到原始框选位置";
    reset.addEventListener("click", (event) => {
      event.stopPropagation();
      delete region.targetBounds;
      rerenderPreviewIfOpen();
      scheduleCurrentDraftSave();
    });

    const handle = document.createElement("i");
    handle.className = "pdf-layout-resize-handle";
    handle.setAttribute("aria-hidden", "true");
    attachPdfLayoutImageMove(box, region, overlay);
    attachPdfLayoutImageResize(handle, box, region, overlay);
    box.append(label, reset, handle);
    overlay.append(box);
  });

  automaticRegions.forEach((region, index) => {
    const box = document.createElement("div");
    box.className = "pdf-layout-image-box automatic";
    positionPdfLayoutBox(box, getPdfRegionSourceBounds(region), pageWidth, pageHeight, cssWidth, cssHeight);

    const label = document.createElement("span");
    label.textContent = `系统图 ${index + 1}`;

    const adopt = document.createElement("button");
    adopt.type = "button";
    adopt.textContent = "校准";
    adopt.title = "采用系统识别区域，并允许手工校准";
    adopt.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      promotePdfAutomaticFigureRegion(pagePath, region, box, overlay);
      rerenderPreviewIfOpen();
      scheduleCurrentDraftSave();
    });

    const handle = document.createElement("i");
    handle.className = "pdf-layout-resize-handle";
    handle.setAttribute("aria-hidden", "true");
    attachPdfLayoutAutomaticImageMove(box, pagePath, region, overlay);
    attachPdfLayoutAutomaticImageResize(handle, box, pagePath, region, overlay);
    box.append(label, adopt, handle);
    overlay.append(box);
  });
}

function getPdfAutomaticFigureRegions(pagePath) {
  return (state.pdfPreservedRegions.get(pagePath) || []).filter((region) => region.kind === "figure");
}

function promotePdfAutomaticFigureRegion(pagePath, region, box, overlay) {
  const targetBounds = getPdfBoundsFromLayoutBox(box, overlay);
  const sourceBounds = getPdfRegionSourceBounds(region);
  const manual = {
    id: `manual-${Date.now()}-${Math.random().toString(16).slice(2)}`,
    kind: "manual",
    x: sourceBounds.x,
    y: sourceBounds.y,
    width: sourceBounds.width,
    height: sourceBounds.height,
    targetBounds,
  };
  const regions = getPdfManualRegions(pagePath).filter((item) =>
    isPdfPageFurnitureRegion(item) || !rectsIntersect(getPdfRegionSourceBounds(item), sourceBounds)
  );
  regions.push(manual);
  state.pdfManualPreservedRegions.set(pagePath, regions);
  return manual;
}

function renderPdfLayoutFurnitureBoxes(overlay, pagePath) {
  const pageWidth = Number(overlay.dataset.pageWidth || 1);
  const pageHeight = Number(overlay.dataset.pageHeight || 1);
  const cssWidth = Number(overlay.dataset.cssWidth || 1);
  const cssHeight = Number(overlay.dataset.cssHeight || 1);
  const regions = getPdfPageFurnitureRegions(pagePath, pageWidth, pageHeight);

  regions.forEach((region) => {
    const box = document.createElement("div");
    box.className = `pdf-layout-furniture-box ${region.kind}`;
    positionPdfLayoutBox(box, getPdfRegionSourceBounds(region), pageWidth, pageHeight, cssWidth, cssHeight);

    const label = document.createElement("span");
    label.textContent = region.kind === "header" ? "页眉区" : "页脚区";

    const reset = document.createElement("button");
    reset.type = "button";
    reset.textContent = "重置";
    reset.title = "恢复系统默认页眉/页脚范围";
    reset.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      removePdfManualFurnitureRegion(pagePath, region.kind);
      rerenderPreviewIfOpen();
      scheduleCurrentDraftSave();
    });

    const applyAll = document.createElement("button");
    applyAll.type = "button";
    applyAll.textContent = "应用全部";
    applyAll.title = "将当前页眉/页脚范围应用到全部页面";
    applyAll.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      applyPdfManualFurnitureRegionToAllPages(pagePath, region.kind, box, overlay);
      rerenderPreviewIfOpen();
      scheduleCurrentDraftSave();
    });

    const handle = document.createElement("i");
    handle.className = "pdf-layout-resize-handle";
    handle.setAttribute("aria-hidden", "true");
    attachPdfLayoutFurnitureMove(box, pagePath, region.kind, overlay);
    attachPdfLayoutFurnitureResize(handle, box, pagePath, region.kind, overlay);
    box.append(label, applyAll, reset, handle);
    overlay.append(box);
  });
}

function upsertPdfManualFurnitureRegion(pagePath, kind, box, overlay) {
  const bounds = getPdfBoundsFromLayoutBox(box, overlay);
  upsertPdfManualFurnitureBounds(pagePath, kind, bounds);
}

function upsertPdfManualFurnitureBounds(pagePath, kind, bounds) {
  const regions = getPdfManualRegions(pagePath).filter((region) => region.kind !== kind);
  regions.push({
    id: `${kind}-${Date.now()}-${Math.random().toString(16).slice(2)}`,
    kind,
    ...bounds,
  });
  state.pdfManualPreservedRegions.set(pagePath, regions);
}

function applyPdfManualFurnitureRegionToAllPages(sourcePagePath, kind, box, overlay) {
  const sourceBounds = getPdfBoundsFromLayoutBox(box, overlay);
  const sourceWidth = Number(overlay.dataset.pageWidth || 1);
  const sourceHeight = Number(overlay.dataset.pageHeight || 1);
  getPdfAllPagePaths().forEach((pagePath) => {
    const pageSize = state.pdfPageSizes.get(pagePath) || state.pdfPageSizes.get(sourcePagePath) || {};
    const pageWidth = Number(pageSize.width || sourceWidth || 595.28);
    const pageHeight = Number(pageSize.height || sourceHeight || 841.89);
    upsertPdfManualFurnitureBounds(pagePath, kind, scalePdfBoundsToPage(sourceBounds, sourceWidth, sourceHeight, pageWidth, pageHeight));
  });
}

function getPdfAllPagePaths() {
  const count = Number(state.slideCount || 0);
  if (count > 0) {
    return Array.from({ length: count }, (_, index) => `pdf/page-${index + 1}`);
  }
  return [...state.pdfPageSizes.keys()]
    .filter((pagePath) => /^pdf\/page-\d+$/.test(pagePath))
    .sort((a, b) => getPdfPageNumberFromPath(a) - getPdfPageNumberFromPath(b));
}

function scalePdfBoundsToPage(bounds, sourceWidth, sourceHeight, pageWidth, pageHeight) {
  const widthRatio = pageWidth / Math.max(1, Number(sourceWidth || pageWidth || 1));
  const heightRatio = pageHeight / Math.max(1, Number(sourceHeight || pageHeight || 1));
  const x = clampNumber(Number(bounds.x || 0) * widthRatio, 0, pageWidth);
  const width = clampNumber(Number(bounds.width || 0) * widthRatio, 1, pageWidth - x);
  const y = clampNumber(Number(bounds.y || 0) * heightRatio, 0, pageHeight);
  const height = clampNumber(Number(bounds.height || 0) * heightRatio, 1, pageHeight - y);
  return { x, y, width, height };
}

function removePdfManualFurnitureRegion(pagePath, kind) {
  const regions = getPdfManualRegions(pagePath).filter((region) => region.kind !== kind);
  if (regions.length) state.pdfManualPreservedRegions.set(pagePath, regions);
  else state.pdfManualPreservedRegions.delete(pagePath);
}

function positionPdfLayoutBox(box, bounds, pageWidth, pageHeight, cssWidth, cssHeight) {
  box.style.left = `${(Number(bounds.x || 0) / pageWidth) * cssWidth}px`;
  box.style.top = `${((pageHeight - Number(bounds.y || 0) - Number(bounds.height || 0)) / pageHeight) * cssHeight}px`;
  box.style.width = `${(Number(bounds.width || 0) / pageWidth) * cssWidth}px`;
  box.style.height = `${(Number(bounds.height || 0) / pageHeight) * cssHeight}px`;
}

function updatePdfManualRegionTargetFromBox(region, box, overlay) {
  region.targetBounds = getPdfBoundsFromLayoutBox(box, overlay);
}

function updatePdfManualRegionSourceFromBox(region, box, overlay) {
  const bounds = getPdfBoundsFromLayoutBox(box, overlay);
  region.x = bounds.x;
  region.y = bounds.y;
  region.width = bounds.width;
  region.height = bounds.height;
  delete region.targetBounds;
}

function getPdfBoundsFromLayoutBox(box, overlay) {
  const pageWidth = Number(overlay.dataset.pageWidth || 1);
  const pageHeight = Number(overlay.dataset.pageHeight || 1);
  const cssWidth = Number(overlay.dataset.cssWidth || overlay.getBoundingClientRect().width || 1);
  const cssHeight = Number(overlay.dataset.cssHeight || overlay.getBoundingClientRect().height || 1);
  const left = Math.max(0, box.offsetLeft);
  const top = Math.max(0, box.offsetTop);
  const width = Math.max(8, box.offsetWidth);
  const height = Math.max(8, box.offsetHeight);
  const x = clampNumber((left / cssWidth) * pageWidth, 0, pageWidth);
  const targetWidth = clampNumber((width / cssWidth) * pageWidth, 1, pageWidth - x);
  const topPdf = pageHeight - (top / cssHeight) * pageHeight;
  const bottomPdf = pageHeight - ((top + height) / cssHeight) * pageHeight;
  const y = clampNumber(bottomPdf, 0, pageHeight);
  const targetHeight = clampNumber(topPdf - bottomPdf, 1, pageHeight - y);
  return { x, y, width: targetWidth, height: targetHeight };
}

function updatePdfSegmentBoundsFromBox(segment, box, overlay) {
  segment.overrides.pdfBounds = getPdfBoundsFromLayoutBox(box, overlay);
}

function positionPdfLayoutBoxFromBounds(box, bounds, overlay) {
  const pageWidth = Number(overlay.dataset.pageWidth || 1);
  const pageHeight = Number(overlay.dataset.pageHeight || 1);
  const cssWidth = Number(overlay.dataset.cssWidth || overlay.getBoundingClientRect().width || 1);
  const cssHeight = Number(overlay.dataset.cssHeight || overlay.getBoundingClientRect().height || 1);
  positionPdfLayoutBox(box, bounds, pageWidth, pageHeight, cssWidth, cssHeight);
}

function markPdfLayoutBoxLocalUpdate(box, message = "已记录局部调整。导出/分享时会按新位置生成；如需重新生成画布，请点“刷新真实预览”。") {
  box.classList.add("local-updated");
  box.dataset.localUpdate = "true";
  setPdfLayoutTextContentHidden(box.querySelector(".pdf-layout-text-content"), false);
  box.title = `${box.title || ""}\n${message}`.trim();
  setStatus(message);
}

function clearPdfLayoutBoxLocalUpdate(box) {
  box.classList.remove("local-updated");
  delete box.dataset.localUpdate;
}

function setPdfLayoutTextContentHidden(content, hidden) {
  if (!content) return;
  content.classList.toggle("preview-text-hidden", Boolean(hidden));
  content.setAttribute("aria-hidden", hidden ? "true" : "false");
}

function attachPdfLayoutTextMove(handle, box, segment, overlay) {
  handle.addEventListener("pointerdown", (event) => {
    event.preventDefault();
    event.stopPropagation();
    handle.setPointerCapture?.(event.pointerId);
    const start = {
      x: event.clientX,
      y: event.clientY,
      left: box.offsetLeft,
      top: box.offsetTop,
    };
    const move = (moveEvent) => {
      const maxLeft = overlay.clientWidth - box.offsetWidth;
      const maxTop = overlay.clientHeight - box.offsetHeight;
      box.style.left = `${clampNumber(start.left + moveEvent.clientX - start.x, 0, Math.max(0, maxLeft))}px`;
      box.style.top = `${clampNumber(start.top + moveEvent.clientY - start.y, 0, Math.max(0, maxTop))}px`;
    };
    const done = () => {
      handle.removeEventListener("pointermove", move);
      handle.removeEventListener("pointerup", done);
      handle.removeEventListener("pointercancel", done);
      updatePdfSegmentBoundsFromBox(segment, box, overlay);
      markPdfLayoutBoxLocalUpdate(box);
      scheduleCurrentDraftSave();
    };
    handle.addEventListener("pointermove", move);
    handle.addEventListener("pointerup", done);
    handle.addEventListener("pointercancel", done);
  });
}

function attachPdfLayoutTextResize(handle, box, segment, overlay) {
  handle.addEventListener("pointerdown", (event) => {
    event.preventDefault();
    event.stopPropagation();
    handle.setPointerCapture?.(event.pointerId);
    const start = {
      x: event.clientX,
      y: event.clientY,
      width: box.offsetWidth,
      height: box.offsetHeight,
    };
    const move = (moveEvent) => {
      const maxWidth = overlay.clientWidth - box.offsetLeft;
      const maxHeight = overlay.clientHeight - box.offsetTop;
      const width = clampNumber(start.width + moveEvent.clientX - start.x, 24, Math.max(24, maxWidth));
      const height = clampNumber(start.height + moveEvent.clientY - start.y, 18, Math.max(18, maxHeight));
      box.style.width = `${width}px`;
      box.style.height = `${height}px`;
    };
    const done = () => {
      handle.removeEventListener("pointermove", move);
      handle.removeEventListener("pointerup", done);
      handle.removeEventListener("pointercancel", done);
      updatePdfSegmentBoundsFromBox(segment, box, overlay);
      markPdfLayoutBoxLocalUpdate(box);
      scheduleCurrentDraftSave();
    };
    handle.addEventListener("pointermove", move);
    handle.addEventListener("pointerup", done);
    handle.addEventListener("pointercancel", done);
  });
}

function attachPdfLayoutImageMove(box, region, overlay) {
  box.addEventListener("pointerdown", (event) => {
    if (event.target.closest("button") || event.target.closest(".pdf-layout-resize-handle")) return;
    event.preventDefault();
    box.setPointerCapture?.(event.pointerId);
    const start = {
      x: event.clientX,
      y: event.clientY,
      left: box.offsetLeft,
      top: box.offsetTop,
    };
    const move = (moveEvent) => {
      const maxLeft = overlay.clientWidth - box.offsetWidth;
      const maxTop = overlay.clientHeight - box.offsetHeight;
      box.style.left = `${clampNumber(start.left + moveEvent.clientX - start.x, 0, Math.max(0, maxLeft))}px`;
      box.style.top = `${clampNumber(start.top + moveEvent.clientY - start.y, 0, Math.max(0, maxTop))}px`;
    };
    const done = () => {
      box.removeEventListener("pointermove", move);
      box.removeEventListener("pointerup", done);
      box.removeEventListener("pointercancel", done);
      updatePdfManualRegionTargetFromBox(region, box, overlay);
      scheduleCurrentDraftSave();
    };
    box.addEventListener("pointermove", move);
    box.addEventListener("pointerup", done);
    box.addEventListener("pointercancel", done);
  });
}

function attachPdfLayoutImageResize(handle, box, region, overlay) {
  handle.addEventListener("pointerdown", (event) => {
    event.preventDefault();
    event.stopPropagation();
    handle.setPointerCapture?.(event.pointerId);
    const start = {
      x: event.clientX,
      y: event.clientY,
      width: box.offsetWidth,
      height: box.offsetHeight,
    };
    const ratio = Math.max(0.1, start.width / Math.max(1, start.height));
    const move = (moveEvent) => {
      const delta = Math.max(moveEvent.clientX - start.x, (moveEvent.clientY - start.y) * ratio);
      const maxWidth = overlay.clientWidth - box.offsetLeft;
      const maxHeight = overlay.clientHeight - box.offsetTop;
      const width = clampNumber(start.width + delta, 24, Math.max(24, maxWidth));
      const height = clampNumber(width / ratio, 18, Math.max(18, maxHeight));
      box.style.width = `${width}px`;
      box.style.height = `${height}px`;
    };
    const done = () => {
      handle.removeEventListener("pointermove", move);
      handle.removeEventListener("pointerup", done);
      handle.removeEventListener("pointercancel", done);
      updatePdfManualRegionTargetFromBox(region, box, overlay);
      scheduleCurrentDraftSave();
    };
    handle.addEventListener("pointermove", move);
    handle.addEventListener("pointerup", done);
    handle.addEventListener("pointercancel", done);
  });
}

function attachPdfLayoutAutomaticImageMove(box, pagePath, region, overlay) {
  box.addEventListener("pointerdown", (event) => {
    if (event.target.closest("button") || event.target.closest(".pdf-layout-resize-handle")) return;
    event.preventDefault();
    box.setPointerCapture?.(event.pointerId);
    const start = {
      x: event.clientX,
      y: event.clientY,
      left: box.offsetLeft,
      top: box.offsetTop,
    };
    const move = (moveEvent) => {
      const maxLeft = overlay.clientWidth - box.offsetWidth;
      const maxTop = overlay.clientHeight - box.offsetHeight;
      box.style.left = `${clampNumber(start.left + moveEvent.clientX - start.x, 0, Math.max(0, maxLeft))}px`;
      box.style.top = `${clampNumber(start.top + moveEvent.clientY - start.y, 0, Math.max(0, maxTop))}px`;
    };
    const done = () => {
      box.removeEventListener("pointermove", move);
      box.removeEventListener("pointerup", done);
      box.removeEventListener("pointercancel", done);
      promotePdfAutomaticFigureRegion(pagePath, region, box, overlay);
      scheduleCurrentDraftSave();
      rerenderPreviewIfOpen();
    };
    box.addEventListener("pointermove", move);
    box.addEventListener("pointerup", done);
    box.addEventListener("pointercancel", done);
  });
}

function attachPdfLayoutAutomaticImageResize(handle, box, pagePath, region, overlay) {
  handle.addEventListener("pointerdown", (event) => {
    event.preventDefault();
    event.stopPropagation();
    handle.setPointerCapture?.(event.pointerId);
    const start = {
      x: event.clientX,
      y: event.clientY,
      width: box.offsetWidth,
      height: box.offsetHeight,
    };
    const ratio = Math.max(0.1, start.width / Math.max(1, start.height));
    const move = (moveEvent) => {
      const delta = Math.max(moveEvent.clientX - start.x, (moveEvent.clientY - start.y) * ratio);
      const maxWidth = overlay.clientWidth - box.offsetLeft;
      const maxHeight = overlay.clientHeight - box.offsetTop;
      const width = clampNumber(start.width + delta, 24, Math.max(24, maxWidth));
      const height = clampNumber(width / ratio, 18, Math.max(18, maxHeight));
      box.style.width = `${width}px`;
      box.style.height = `${height}px`;
    };
    const done = () => {
      handle.removeEventListener("pointermove", move);
      handle.removeEventListener("pointerup", done);
      handle.removeEventListener("pointercancel", done);
      promotePdfAutomaticFigureRegion(pagePath, region, box, overlay);
      scheduleCurrentDraftSave();
      rerenderPreviewIfOpen();
    };
    handle.addEventListener("pointermove", move);
    handle.addEventListener("pointerup", done);
    handle.addEventListener("pointercancel", done);
  });
}

function attachPdfLayoutFurnitureMove(box, pagePath, kind, overlay) {
  box.addEventListener("pointerdown", (event) => {
    if (event.target.closest("button") || event.target.closest(".pdf-layout-resize-handle")) return;
    event.preventDefault();
    box.setPointerCapture?.(event.pointerId);
    const start = {
      x: event.clientX,
      y: event.clientY,
      left: box.offsetLeft,
      top: box.offsetTop,
    };
    const move = (moveEvent) => {
      const maxLeft = overlay.clientWidth - box.offsetWidth;
      const maxTop = overlay.clientHeight - box.offsetHeight;
      box.style.left = `${clampNumber(start.left + moveEvent.clientX - start.x, 0, Math.max(0, maxLeft))}px`;
      box.style.top = `${clampNumber(start.top + moveEvent.clientY - start.y, 0, Math.max(0, maxTop))}px`;
    };
    const done = () => {
      box.removeEventListener("pointermove", move);
      box.removeEventListener("pointerup", done);
      box.removeEventListener("pointercancel", done);
      upsertPdfManualFurnitureRegion(pagePath, kind, box, overlay);
      scheduleCurrentDraftSave();
    };
    box.addEventListener("pointermove", move);
    box.addEventListener("pointerup", done);
    box.addEventListener("pointercancel", done);
  });
}

function attachPdfLayoutFurnitureResize(handle, box, pagePath, kind, overlay) {
  handle.addEventListener("pointerdown", (event) => {
    event.preventDefault();
    event.stopPropagation();
    handle.setPointerCapture?.(event.pointerId);
    const start = {
      x: event.clientX,
      y: event.clientY,
      width: box.offsetWidth,
      height: box.offsetHeight,
    };
    const move = (moveEvent) => {
      const maxWidth = overlay.clientWidth - box.offsetLeft;
      const maxHeight = overlay.clientHeight - box.offsetTop;
      const width = clampNumber(start.width + moveEvent.clientX - start.x, 40, Math.max(40, maxWidth));
      const height = clampNumber(start.height + moveEvent.clientY - start.y, 14, Math.max(14, maxHeight));
      box.style.width = `${width}px`;
      box.style.height = `${height}px`;
    };
    const done = () => {
      handle.removeEventListener("pointermove", move);
      handle.removeEventListener("pointerup", done);
      handle.removeEventListener("pointercancel", done);
      upsertPdfManualFurnitureRegion(pagePath, kind, box, overlay);
      scheduleCurrentDraftSave();
    };
    handle.addEventListener("pointermove", move);
    handle.addEventListener("pointerup", done);
    handle.addEventListener("pointercancel", done);
  });
}

function getOriginalPreviewUrl() {
  if (!state.originalPreviewUrl && state.file) {
    state.originalPreviewUrl = URL.createObjectURL(state.file);
  }
  return state.originalPreviewUrl;
}

function revokeOriginalPreviewUrl() {
  if (!state.originalPreviewUrl) return;
  URL.revokeObjectURL(state.originalPreviewUrl);
  state.originalPreviewUrl = "";
}

function drawPdfOverlayErase(page, plan, rgb) {
  page.drawRectangle({
    x: plan.erase.x,
    y: plan.erase.y,
    width: plan.erase.width,
    height: plan.erase.height,
    color: rgb(plan.erase.color.r, plan.erase.color.g, plan.erase.color.b),
    opacity: 1,
  });
}

function redrawPdfLinesCoveredByOverlay(page, lines, plans, rgb) {
  if (!Array.isArray(lines) || !lines.length || !Array.isArray(plans) || !plans.length) return;
  const eraseBounds = plans
    .map((plan) => plan?.erase)
    .filter(Boolean)
    .map((erase) => ({
      x: Number(erase.x || 0),
      y: Number(erase.y || 0),
      width: Number(erase.width || 0),
      height: Number(erase.height || 0),
    }));
  if (!eraseBounds.length) return;

  const drawn = new Set();
  lines
    .filter((line) => shouldRestorePdfLineAfterOverlay(line, eraseBounds))
    .forEach((line) => {
      const key = [
        Math.round(Number(line.x1 || 0) * 2) / 2,
        Math.round(Number(line.y1 || 0) * 2) / 2,
        Math.round(Number(line.x2 || 0) * 2) / 2,
        Math.round(Number(line.y2 || 0) * 2) / 2,
      ].join(":");
      if (drawn.has(key)) return;
      drawn.add(key);
      const color = line.color || { r: 0, g: 0, b: 0 };
      page.drawLine({
        start: { x: Number(line.x1 || 0), y: Number(line.y1 || 0) },
        end: { x: Number(line.x2 || 0), y: Number(line.y2 || 0) },
        thickness: Math.max(0.25, Math.min(4, Number(line.thickness || 1))),
        color: rgb(color.r, color.g, color.b),
        opacity: 1,
      });
    });
}

function shouldRestorePdfLineAfterOverlay(line, eraseBounds) {
  if (!line) return false;
  const horizontal = Math.abs(Number(line.y1 || 0) - Number(line.y2 || 0)) <= 1.5;
  const vertical = Math.abs(Number(line.x1 || 0) - Number(line.x2 || 0)) <= 1.5;
  if (!horizontal && !vertical) return false;
  const length = Math.hypot(Number(line.x2 || 0) - Number(line.x1 || 0), Number(line.y2 || 0) - Number(line.y1 || 0));
  if (length < 18) return false;
  const thickness = Math.max(0.2, Number(line.thickness || 1));
  const pad = Math.max(1.5, thickness * 2);
  const bounds = {
    x: Math.min(Number(line.x1 || 0), Number(line.x2 || 0)) - pad,
    y: Math.min(Number(line.y1 || 0), Number(line.y2 || 0)) - pad,
    width: Math.abs(Number(line.x2 || 0) - Number(line.x1 || 0)) + pad * 2 || pad * 2,
    height: Math.abs(Number(line.y2 || 0) - Number(line.y1 || 0)) + pad * 2 || pad * 2,
  };
  return eraseBounds.some((erase) => rectsIntersect(bounds, erase));
}

function resolvePdfOverlayPlanCollisions(overlayPlans) {
  const pageGroups = new Map();
  overlayPlans.forEach((entry) => {
    if (!entry?.page || !entry.plan) return;
    if (!pageGroups.has(entry.page)) pageGroups.set(entry.page, []);
    pageGroups.get(entry.page).push(entry.plan);
  });

  pageGroups.forEach((plans) => {
    const blockers = plans
      .filter((plan) => plan.isFigureCaption)
      .map((plan) => ({ plan, bounds: getPdfPlanTextBounds(plan, 3) }))
      .filter((item) => item.bounds);
    if (!blockers.length) return;

    const sorted = plans
      .filter((plan) => !plan.isFigureCaption)
      .sort((a, b) => Number(b.fitBounds?.y || 0) - Number(a.fitBounds?.y || 0));

    sorted.forEach((plan) => {
      let bounds = getPdfPlanTextBounds(plan, 2);
      if (!bounds) return;
      blockers.forEach((blocker) => {
        if (!bounds || !blocker.bounds) return;
        if (!rectHorizontalOverlap(bounds, blocker.bounds)) return;
        if (!rectsIntersect(bounds, blocker.bounds)) return;
        const targetTop = Number(blocker.bounds.y || 0) - 4;
        const currentTop = Number(bounds.y || 0) + Number(bounds.height || 0);
        const shiftDown = Math.max(0, currentTop - targetTop);
        if (shiftDown <= 0) return;
        shiftPdfOverlayPlanDown(plan, Math.min(shiftDown, 42));
        bounds = getPdfPlanTextBounds(plan, 2);
      });
    });
  });
}

function shiftPdfOverlayPlanDown(plan, amount) {
  const delta = Number(amount || 0);
  if (!Number.isFinite(delta) || delta <= 0) return;
  plan.y -= delta;
  if (plan.fitBounds) plan.fitBounds.y = Math.max(0, Number(plan.fitBounds.y || 0) - delta);
  if (plan.erase) plan.erase.y = Math.max(0, Number(plan.erase.y || 0) - delta);
}

function getPdfPlanTextBounds(plan, padding = 0) {
  if (!plan?.fitBounds || !Array.isArray(plan.lines) || !plan.lines.length) return null;
  const baseline = getPdfBaselineOffset(plan.fontSize);
  const top = Number(plan.y || 0) + Math.max(1, Number(plan.fontSize || 0) * 0.28) + padding;
  const bottom = Number(plan.y || 0) - (plan.lines.length - 1) * Number(plan.lineHeight || 0) - baseline - padding;
  return {
    x: Math.max(0, Number(plan.fitBounds.x || 0) - padding),
    y: Math.max(0, bottom),
    width: Math.max(1, Number(plan.fitBounds.width || 0) + padding * 2),
    height: Math.max(1, top - Math.max(0, bottom)),
  };
}

function drawPdfOverlayText(page, plan, font, rgb) {
  let y = plan.y;

  plan.lines.forEach((line) => {
    const fontSize = plan.fontSize;
    const lineWidth = font.widthOfTextAtSize(line, fontSize);
    const x = plan.cellAlign === "center" ? plan.fitBounds.x + Math.max(0, (plan.fitBounds.width - lineWidth) / 2) : plan.fitBounds.x;
    page.drawText(line, {
      x,
      y,
      size: fontSize,
      font,
      color: rgb(plan.textColor.r, plan.textColor.g, plan.textColor.b),
    });
    y -= plan.lineHeight;
  });
}

function getPdfExportText(segment) {
  if (isInsidePdfPreservedRegion(segment)) return "";

  if (shouldKeepPdfSourceText(segment)) {
    return shouldRedrawKeptPdfSourceText(segment) ? sanitizePdfExportText(segment.original) : "";
  }

  const text = sanitizePdfExportText(applyTerminologyRules(String(segment.translation || "").trim(), segment.original));
  if (!text) return "";

  const direction = els.translationDirection?.value || "";
  const directionConfig = getDirectionConfig(direction);
  if (directionConfig.targetCode === "en" && containsHanCharacters(text)) {
    return "";
  }

  if (directionConfig.targetCode === "zh" && shouldExpectHanTranslation(segment.original, directionConfig, segment) && !containsHanCharacters(text)) {
    return "";
  }

  return text;
}

function isInsidePdfPreservedRegion(segment) {
  if (segment?.type !== "pdf" || !segment.layout?.bounds) return false;
  const regions = getPdfPreservedRegionsForPage(segment.path);
  if (!regions.length) return false;
  const bounds = segment.layout.bounds;
  const cx = Number(bounds.x || 0) + Number(bounds.width || 0) / 2;
  const cy = Number(bounds.y || 0) + Number(bounds.height || 0) / 2;
  return regions.some((region) => {
    const targetBounds = getPdfRegionTargetBounds(region);
    return pointInRect(cx, cy, targetBounds) || rectsIntersect(bounds, targetBounds);
  });
}

function shouldRedrawKeptPdfSourceText(segment) {
  if (!segment || isLikelyPdfHeaderFooterText(segment) || isLikelyPdfPublicationFurniture(segment) || isLikelyPdfArtifactText(segment)) {
    return false;
  }
  return false;
}

function shouldKeepPdfSourceText(segment) {
  return isLikelyPdfArtifactText(segment) ||
    isLikelyPdfHeaderFooterText(segment) ||
    isLikelyPdfPublicationFurniture(segment);
}

function sanitizePdfExportText(text) {
  return String(text || "")
    .replace(/[\uF6B1-\uF6BA]/g, (char) => String(char.charCodeAt(0) - 0xF6B1))
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, "")
    .replace(/\s+([,.;:!?，。；：！？%）\]])/g, "$1")
    .replace(/([（\[])\s+/g, "$1")
    .replace(/\s{2,}/g, " ")
    .trim();
}

function isLikelyPdfArtifactText(segment) {
  if (segment?.type !== "pdf") return false;
  const source = sanitizePdfExportText(segment.original);
  const translation = sanitizePdfExportText(segment.translation);
  const text = translation || source;
  if (!text) return true;

  const compact = text.replace(/\s+/g, "");
  const sourceCompact = source.replace(/\s+/g, "");
  if (/^[†‡*•·|¦_+\-=–—.,;:()[\]{}<>/\\'"\d]+$/.test(compact) && compact.length <= 8) return true;
  if (/^\d{1,3}$/.test(compact)) return true;
  if (/^[A-Za-z]$/.test(compact)) return true;
  if (/^[\uF6B1-\uF6BA]+$/.test(String(segment.original || "").replace(/\s+/g, ""))) return true;
  if (sourceCompact.length <= 2 && /^[†‡*•·\d]+$/.test(sourceCompact)) return true;
  return false;
}

function isLikelyPdfHeaderFooterText(segment) {
  if (segment?.type !== "pdf") return false;
  const text = String(segment.original || "").replace(/\s+/g, " ").trim();
  if (!text) return false;

  const bounds = segment.layout?.bounds;
  const pageHeight = state.pdfPageSizes.get(segment.path)?.height || 841.89;
  if (!bounds || !pageHeight) return false;

  const y = Number(bounds.y || 0);
  const height = Number(bounds.height || 0);
  const centerY = y + height / 2;
  const fontSize = Number(segment.layout?.fontSize || 10);
  const pageSize = state.pdfPageSizes.get(segment.path) || {};
  const pageWidth = Number(pageSize.width || 595.28);
  const furnitureRegions = getPdfPageFurnitureRegions(segment.path, pageWidth, pageHeight);
  const inHeader = furnitureRegions.some((region) => region.kind === "header" && pointInRect(Number(bounds.x || 0) + Number(bounds.width || 0) / 2, centerY, region));
  const inFooter = furnitureRegions.some((region) => region.kind === "footer" && pointInRect(Number(bounds.x || 0) + Number(bounds.width || 0) / 2, centerY, region));
  if (!inHeader && !inFooter) return false;

  const looksLikePageFurniture = /^[\d\s\-–—]+$/.test(text) ||
    /(?:doi|www\.|https?:\/\/|copyright|©|page\s+\d+|\d+\s*\/\s*\d+)/i.test(text) ||
    fontSize <= 10.5;
  return looksLikePageFurniture;
}

function isLikelyPdfPublicationFurniture(segment) {
  if (segment?.type !== "pdf") return false;
  const text = String(segment.original || "").replace(/\s+/g, " ").trim();
  if (!text || containsHanCharacters(text)) return false;

  const bounds = segment.layout?.bounds;
  const pageHeight = state.pdfPageSizes.get(segment.path)?.height || 841.89;
  const inTopFurnitureZone = bounds ? bounds.y > pageHeight * 0.68 : false;
  if (!inTopFurnitureZone) return false;

  return /(?:ScienceDirect|Elsevier|BMJ|JITC|Journal for ImmunoTherapy|Open access|Hypothesis|contents lists available|journal homepage|www\.|ISSN|doi:|Lung Cancer\s+\d|^Lung Cancer$)/i.test(text);
}

function isLikelyPdfSidebarMetadata(segment) {
  if (segment?.type !== "pdf") return false;
  const text = String(segment.original || "").replace(/\s+/g, " ").trim();
  if (!text || containsHanCharacters(text)) return false;

  const bounds = segment.layout?.bounds;
  const pageSize = state.pdfPageSizes.get(segment.path) || {};
  const pageWidth = Number(pageSize.width || 595.28);
  const pageHeight = Number(pageSize.height || 841.89);
  if (!bounds || !pageWidth || !pageHeight) return false;

  const x = Number(bounds.x || 0);
  const y = Number(bounds.y || 0);
  const width = Number(bounds.width || 0);
  const fontSize = Number(segment.layout?.fontSize || 10);
  const inLeftRail = x < pageWidth * 0.26 && width < pageWidth * 0.24;
  const inBottomFurniture = y < pageHeight * 0.14;
  const smallText = fontSize <= 8.8;

  if (!inLeftRail && !inBottomFurniture) return false;
  if (/^(?:abstract|introduction|methods?|results?|discussion|conclusion)s?$/i.test(text)) return false;
  if (segment.layout?.columnKey === "side" && smallText) return true;

  const metadataPattern = /(?:to cite|cite this|accepted|received|correspondence|email|@|©|copyright|author\(s\)|group|hospital|university|department|institute|license|reuse|commercial|doi:|journal|immunother|cancer|updates|check for updates|contributor|funding|competing interests|patient consent|ethics approval|provenance|data availability|open access|bmj)/i;
  const mostlyNamesOrAddress = /(?:professor|school|college|wuhan|china|road|avenue|street|tel|fax|^[A-Z][A-Za-z-]+(?:\s+[A-Z][A-Za-z-]+){1,5}$)/i.test(text);

  return smallText && (metadataPattern.test(text) || mostlyNamesOrAddress);
}

function isLikelyPdfAuthorByline(segment) {
  if (segment?.type !== "pdf") return false;
  const text = String(segment.original || "").replace(/\s+/g, " ").trim();
  if (!text || containsHanCharacters(text)) return false;

  const bounds = segment.layout?.bounds;
  const pageHeight = state.pdfPageSizes.get(segment.path)?.height || 841.89;
  const fontSize = Number(segment.layout?.fontSize || 10);
  const inTitleArea = bounds ? bounds.y > pageHeight * 0.42 && bounds.y < pageHeight * 0.72 : false;
  if (!inTitleArea || fontSize < 6 || fontSize > 14) return false;

  const nameWords = text.match(/\b[A-Z][a-z]+(?:[-'][A-Z][a-z]+)?\b/g) || [];
  const lowerWords = text.match(/\b[a-z]{3,}\b/g) || [];
  const hasAffiliationMarkers = /(?:\*|\b[a-z]\s*,|\b[a-z]\s*$|,\s*[a-z]\b)/.test(text);
  const hasNameShape = nameWords.length >= 2 && nameWords.length <= 18;
  const mostlyNames = lowerWords.length <= Math.max(1, Math.floor(nameWords.length * 0.45));

  return hasNameShape && hasAffiliationMarkers && /[,;]/.test(text) && mostlyNames;
}

function shouldSkipSegmentForDirection(segment, direction) {
  return direction.sourceLanguage === "Chinese" && !containsHanCharacters(segment.original || "");
}

function containsHanCharacters(text) {
  return /[\u3400-\u9fff]/.test(text);
}

function validatePdfTranslationsBeforeExport() {
  if (state.fileType !== "pdf") return;
  const direction = getDirectionConfig(els.translationDirection?.value || "");
  if (direction.targetCode !== "zh") return;

  const qualityWarning = getPdfTargetLanguageQualityWarning(direction);
  if (qualityWarning) {
    throw new Error(qualityWarning);
  }
}

function getPdfTargetLanguageQualityWarning(direction = getDirectionConfig(els.translationDirection?.value || "")) {
  if (state.fileType !== "pdf" || direction.targetCode !== "zh") return "";
  const candidates = state.segments.filter((segment) =>
    segment.type === "pdf" &&
    segment.layout?.bounds &&
    !shouldKeepPdfSourceText(segment) &&
    shouldExpectHanTranslation(segment.original, direction, segment)
  );
  if (!candidates.length) return "";

  const translated = candidates.filter((segment) => String(segment.translation || "").trim());
  const withChinese = candidates.filter((segment) => containsHanCharacters(segment.translation || ""));
  const chineseRatio = withChinese.length / candidates.length;
  const translatedRatio = translated.length / candidates.length;

  if (withChinese.length === 0 || chineseRatio < 0.45 || translatedRatio < 0.75) {
    return `PDF 英文转中文尚未生成足够中文译文（${withChinese.length}/${candidates.length} 段含中文）。请确认翻译语种为“英文 → 中文”，重新点击“自动翻译”完成后再导出。`;
  }

  return "";
}

function shouldRetryTranslationForTarget(translation, source, direction, segment = null) {
  if (direction?.targetCode !== "zh") return false;
  if (!shouldExpectHanTranslation(source, direction, segment)) return false;
  if (containsHanCharacters(translation || "")) return false;
  return true;
}

function shouldExpectHanTranslation(source, direction, segment = null) {
  if (direction?.targetCode !== "zh") return false;
  if (shouldKeepPdfSourceText(segment)) return false;
  const text = String(source || "").trim();
  if (!text || containsHanCharacters(text)) return false;
  const letters = text.match(/[A-Za-z]+/g) || [];
  if (!letters.length) return false;
  const words = letters.filter((word) => word.length > 1);
  const letterCount = letters.join("").length;
  const hasSentenceShape = /[.!?;:,)]/.test(text) || words.length >= 4 || [...text].length >= 28;
  const mostlyReferenceNoise = /^\s*(?:\[\d+\]|\(?\d+\)?|[-–—\d\s.,;:()[\]/]+|[A-Z]{1,6})+\s*$/.test(text);
  const mostlyAcronyms = words.length <= 3 && words.every((word) => /^[A-Z]{2,8}$/.test(word));
  const tooShortPdfLabel = segment?.type === "pdf" && [...text].length < 18 && words.length <= 2;

  return letterCount >= 8 && hasSentenceShape && !mostlyReferenceNoise && !mostlyAcronyms && !tooShortPdfLabel;
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
  const isTableLike = Number(layout?.rowSegmentCount || 1) > 1;
  const sourceFontSize = Math.max(3.8, Math.min(24, Number(sourceSize || 10) * 0.96));
  const preferred = sourceFontSize;
  const targetWidth = Math.max(4, bounds.width);
  const maxHeight = Math.max(5, bounds.height * (isTableLike ? 0.9 : 0.96));
  const minSize = Math.max(isTableLike ? 3.4 : 2.8, preferred * (isTableLike ? 0.56 : 0.34));

  const textWidthAtPreferred = Math.max(0.1, font.widthOfTextAtSize(text, preferred));
  const estimatedSize = Math.min(preferred, (preferred * targetWidth) / textWidthAtPreferred);

  for (let size = preferred; size >= minSize; size -= 0.25) {
    const lineHeight = getPdfLineHeight(size, isTableLike);
    const lines = wrapPdfText(text, font, size, targetWidth);
    if (getPdfTextBlockHeight(lines, size, lineHeight) <= maxHeight) {
      return { fontSize: size, lines, lineHeight };
    }
  }

  for (let size = estimatedSize; size >= minSize; size -= 0.25) {
    const lineHeight = getPdfLineHeight(size, isTableLike);
    const lines = wrapPdfText(text, font, size, targetWidth);
    if (getPdfTextBlockHeight(lines, size, lineHeight) <= maxHeight) {
      return { fontSize: size, lines, lineHeight };
    }
  }

  for (let size = minSize; size >= 2.8; size -= 0.2) {
    const lineHeight = getPdfLineHeight(size, isTableLike);
    const lines = wrapPdfText(text, font, size, targetWidth);
    if (getPdfTextBlockHeight(lines, size, lineHeight) <= maxHeight) {
      return { fontSize: size, lines, lineHeight };
    }
  }

  const fallbackSize = Math.max(2.4, minSize * 0.92);
  const fallbackLineHeight = getPdfLineHeight(fallbackSize, isTableLike);
  const fallbackLines = wrapPdfText(text, font, fallbackSize, targetWidth);
  return {
    fontSize: fallbackSize,
    lines: clampPdfLinesToHeight(fallbackLines, font, fallbackSize, fallbackLineHeight, maxHeight, targetWidth),
    lineHeight: fallbackLineHeight,
  };
}

function getPdfLineHeight(size, isTableLike) {
  return size * (isTableLike ? 1.28 : 1.36);
}

function clampPdfLinesToHeight(lines, font, fontSize, lineHeight, maxHeight, maxWidth) {
  const baseline = getPdfBaselineOffset(fontSize);
  const maxLines = Math.max(1, Math.floor((Math.max(1, maxHeight) - fontSize - baseline) / Math.max(1, lineHeight)) + 1);
  if (lines.length <= maxLines) return lines;
  const clamped = lines.slice(0, maxLines);
  clamped[clamped.length - 1] = trimPdfLineToWidth(`${clamped[clamped.length - 1]}...`, font, fontSize, maxWidth);
  return clamped;
}

function trimPdfLineToWidth(line, font, fontSize, maxWidth) {
  let value = String(line || "");
  while (value.length > 3 && font.widthOfTextAtSize(value, fontSize) > maxWidth) {
    value = `${value.slice(0, -4)}...`;
  }
  return value || "...";
}

function getPdfBaselineOffset(fontSize) {
  return Math.max(0.8, fontSize * 0.22);
}

function getPdfTextBlockHeight(lines, fontSize, lineHeight) {
  if (!lines.length) return 0;
  return (lines.length - 1) * lineHeight + fontSize + getPdfBaselineOffset(fontSize);
}

async function loadPdfExportTools() {
  if (!window.pdfLibTools) {
    const [pdfLib, fontkitModule, fontResponse] = await Promise.all([
      importFirstAvailable(PDFLIB_URLS, "PDF 导出库"),
      importFirstAvailable(FONTKIT_URLS, "PDF 中文字体引擎"),
      fetchFirstAvailable(PDF_FONT_URLS, "PDF 中文字体"),
    ]);

    window.pdfLibTools = {
      PDFDocument: pdfLib.PDFDocument,
      rgb: pdfLib.rgb,
      fontkit: fontkitModule.default || fontkitModule,
      fontBytes: await fontResponse.arrayBuffer(),
    };
  }

  return window.pdfLibTools;
}

async function importFirstAvailable(urls, label) {
  let lastError = null;

  for (const url of urls) {
    try {
      return await import(url);
    } catch (error) {
      lastError = error;
      console.warn(`${label} 加载失败，尝试备用地址：${url}`, error);
    }
  }

  throw new Error(`${label} 加载失败：${lastError?.message || "请检查网络或刷新更新 app"}`);
}

async function fetchFirstAvailable(urls, label) {
  let lastError = null;

  for (const url of urls) {
    try {
      const response = await fetch(url);
      if (response.ok) return response;
      lastError = new Error(`HTTP ${response.status}`);
      console.warn(`${label} 加载失败，尝试备用地址：${url}`, lastError);
    } catch (error) {
      lastError = error;
      console.warn(`${label} 加载失败，尝试备用地址：${url}`, error);
    }
  }

  throw new Error(`${label} 加载失败：${lastError?.message || "请检查网络或刷新更新 app"}`);
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
  const isSourceMode = state.previewMode === "source";
  item.className = `preview-item${isSourceMode ? " source" : segment.translation.trim() ? "" : " pending"}`;

  const text = document.createElement("p");
  text.textContent = isSourceMode ? segment.original : segment.translation.trim() || segment.original;

  const meta = document.createElement("span");
  meta.textContent = isSourceMode ? "原文" : segment.translation.trim() ? "译文" : "未翻译，暂用原文";

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

    const isSourceMode = state.previewMode === "source";
    const index = state.segments.indexOf(segment);
    const box = document.createElement("div");
    box.className = `slide-text-box${isSourceMode ? " source" : segment.translation.trim() ? "" : " pending"}`;
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
    text.textContent = isSourceMode ? segment.original : segment.translation.trim() || segment.original;

    if (isSourceMode) {
      box.append(text);
    } else {
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
    }
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
    const box = tools.closest(".slide-text-box");
    if (box) {
      box.style.fontSize = `${getPreviewFontCqw(current)}cqw`;
    }
    scheduleCurrentDraftSave();
  });

  scaleInput.addEventListener("change", () => {
    renderSegments();
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
      pdfManualPreservedRegions: serializePdfManualPreservedRegions(),
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
    updateStartupStatus("正在恢复上次文件...", "正在重新解析上次未导出的文档，请稍候。", 72);
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
    restorePdfManualPreservedRegions(draft.pdfManualPreservedRegions);
    renderSegments();
    updateStats();
    if (draft.mobileView) setMobileView(draft.mobileView);
    setStatus(`已恢复 ${draft.fileName} 的上次编辑内容。`);
    showToast("已恢复上次未导出的翻译内容。");
  } catch (error) {
    console.warn("Draft restore failed", error);
    await resetApp(false, { clearDraft: true });
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

function serializePdfManualPreservedRegions() {
  return [...state.pdfManualPreservedRegions.entries()].map(([pagePath, regions]) => ({
    pagePath,
    regions: (regions || []).map((region) => ({
      id: region.id || `manual-${Date.now()}-${Math.random().toString(16).slice(2)}`,
      kind: sanitizePdfManualRegionKind(region.kind),
      x: Number(region.x || 0),
      y: Number(region.y || 0),
      width: Number(region.width || 0),
      height: Number(region.height || 0),
      targetBounds: sanitizePdfRegionTargetBounds(region.targetBounds),
    })).filter((region) => region.width > 0 && region.height > 0),
  })).filter((entry) => entry.pagePath && entry.regions.length);
}

function restorePdfManualPreservedRegions(entries) {
  state.pdfManualPreservedRegions = new Map();
  (entries || []).forEach((entry) => {
    const pagePath = String(entry?.pagePath || "");
    if (!pagePath) return;
    const regions = (entry.regions || [])
      .map((region) => ({
        id: region.id || `manual-${Date.now()}-${Math.random().toString(16).slice(2)}`,
        kind: sanitizePdfManualRegionKind(region.kind),
        x: Number(region.x || 0),
        y: Number(region.y || 0),
        width: Number(region.width || 0),
        height: Number(region.height || 0),
        targetBounds: sanitizePdfRegionTargetBounds(region.targetBounds),
      }))
      .filter((region) => Number.isFinite(region.x) && Number.isFinite(region.y) && region.width > 0 && region.height > 0);
    if (regions.length) state.pdfManualPreservedRegions.set(pagePath, regions);
  });
}

function sanitizePdfManualRegionKind(kind) {
  return ["manual", "header", "footer"].includes(kind) ? kind : "manual";
}

function sanitizePdfRegionTargetBounds(bounds) {
  if (!bounds || !["x", "y", "width", "height"].every((key) => Number.isFinite(Number(bounds[key])))) return null;
  return {
    x: Number(bounds.x),
    y: Number(bounds.y),
    width: Number(bounds.width),
    height: Number(bounds.height),
  };
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

  if (overrides.pdfBounds && ["x", "y", "width", "height"].every((key) => Number.isFinite(Number(overrides.pdfBounds[key])))) {
    clean.pdfBounds = {
      x: Number(overrides.pdfBounds.x),
      y: Number(overrides.pdfBounds.y),
      width: Number(overrides.pdfBounds.width),
      height: Number(overrides.pdfBounds.height),
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

async function resetApp(clearInput = true, options = {}) {
  window.clearTimeout(draftSaveTimer);
  draftSaveTimer = 0;
  state.translationAbortController?.abort();
  state.translationAbortController = null;
  state.translationRunning = false;
  state.batchRunning = false;
  state.batchCancelRequested = true;
  state.importRunning = false;
  state.file = null;
  state.fileType = "";
  state.zip = null;
  state.segments = [];
  state.slideCount = 0;
  state.slideVisuals = new Map();
  state.pdfPageSizes = new Map();
  state.pdfTableCells = new Map();
  state.pdfEraseFragments = new Map();
  state.pdfPreservedRegions = new Map();
  state.pdfManualPreservedRegions = new Map();
  state.pdfManualRegionSelectionEnabled = false;
  state.pdfBytes = null;
  state.pdfParseSource = "";
  state.pdfParsedMarkdown = "";
  state.pdfParsedPages = [];
  state.batchFiles = [];
  state.activeSummary = null;
  revokeOriginalPreviewUrl();
  if (clearInput) els.fileInput.value = "";
  els.fileMeta.textContent = "支持 PPTX / DOCX / PDF；旧版 PPT/DOC 请先另存";
  document.body.classList.remove("has-pdf-file");
  closePreview();
  renderSegments();
  renderBatchQueue();
  updateStats();
  setBusy(false);
  setProgress(0);
  setStatus("请先选择一个 PPTX、DOCX 或 PDF 文件。");
  if (options.message) setStatus(options.message);
  if (options.clearDraft !== false) await clearCurrentDraft();
  if (options.toast) showToast(options.message || "已清空当前内容。");
}

function buildOutputName(name) {
  if (state.fileType === "pdf") {
    const mode = els.pdfOutputMode?.value === "reflow" ? "reflow" : "overlay";
    return name.replace(/\.pdf$/i, "") + `-translated-${mode}.pdf`;
  }
  return name.replace(/\.(pptx|docx)$/i, "") + `-translated.${state.fileType || "pptx"}`;
}

function setPdfOutputMode(mode, options = {}) {
  if (!els.pdfOutputMode) return;
  const normalized = mode === "reflow" ? "reflow" : "overlay";
  if (els.pdfOutputMode.value !== normalized) {
    els.pdfOutputMode.value = normalized;
  }
  if (options.persist) saveSettings();
}

function normalizeTranslation(translation, source, segment = null) {
  const clean = translation.trim();
  if (looksLikeNonTranslation(clean, source)) return "";
  if (segment?.type === "pptx") {
    return applyTerminologyRules(cleanPptTranslationText(clean, source), source);
  }
  const normalized = /[\r\n]/.test(source)
    ? clean
    : clean.replace(/\s*[\r\n]+\s*/g, " ").replace(/[ \t]{2,}/g, " ");
  return applyTerminologyRules(normalized, source);
}

function cleanPptTranslationText(text, source = "") {
  const normalizedLines = String(text || "")
    .replace(/\r\n?/g, "\n")
    .split("\n")
    .map((line) => line.replace(/[ \t]+/g, " ").trim())
    .filter(Boolean);

  if (!normalizedLines.length) return "";
  if (!/[\r\n]/.test(source) && normalizedLines.length <= 2) {
    return normalizedLines.join(" ");
  }
  return normalizedLines.join("\n");
}

function applyTerminologyRules(text, source = "") {
  if (!text) return text;

  let normalized = text
    .replace(/\bPori\s*Nova\b/gi, "PoriNova")
    .replace(/\bPorinova\b/g, "PoriNova")
    .replace(/\bBao\s*Rui\s*(?:Dao|Knife|Blade)\b/gi, "PoriNova")
    .replace(/\bBaorui\s*(?:Dao|Knife|Blade)\b/gi, "PoriNova")
    .replace(/\bsteep\s*pulse\b/gi, "steep-pulse")
    .replace(/\bhigh frequency bipolar\b/gi, "high-frequency bipolar")
    .replace(/\bradio frequency\b/gi, "radiofrequency")
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

  if (/宝瑞刀|PoriNova/i.test(source)) {
    normalized = normalized.replace(/\bPoriNova(?:\s*(?:Knife|Blade|Device|System))?\b/gi, "PoriNova");
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

function shouldSuppressZeroBoundsPresentationText(segment, paragraph = null) {
  if (segment?.type !== "pptx") return false;
  const bounds = paragraph ? getOwnShapeBounds(paragraph) : segment.layout?.bounds;
  return Number(bounds?.cx || 0) <= 1 || Number(bounds?.cy || 0) <= 1;
}

function getOwnShapeBounds(paragraph) {
  const shape = paragraph?.parentElement?.parentElement;
  const shapeProperties = [...(shape?.children || [])].find((node) => node.localName === "spPr");
  const transform = [...(shapeProperties?.children || [])].find((node) => node.localName === "xfrm");
  const offset = [...(transform?.children || [])].find((node) => node.localName === "off");
  const extent = [...(transform?.children || [])].find((node) => node.localName === "ext");
  if (!offset || !extent) return null;

  return {
    x: Number(offset.getAttribute("x")) || 0,
    y: Number(offset.getAttribute("y")) || 0,
    cx: Number(extent.getAttribute("cx")) || 0,
    cy: Number(extent.getAttribute("cy")) || 0,
  };
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
    bodyProperties.append(document.createElementNS(DRAWING_NS, "a:noAutofit"));
    return;
  } else if (layout.wrap && layout.wrap !== "none") {
    bodyProperties.setAttribute("wrap", layout.wrap);
  } else {
    bodyProperties.setAttribute("wrap", "square");
  }
  bodyProperties.setAttribute("anchor", "t");
  bodyProperties.setAttribute("anchorCtr", "0");

  applyPresentationLineSpacing(paragraph, segment);

  if (getPptLayoutMode() === "compact-fit") {
    const autoFit = document.createElementNS(DRAWING_NS, "a:normAutofit");
    autoFit.setAttribute("fontScale", String(Math.round(getPresentationLengthScale(segment) * 100000)));
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
    autoFit.setAttribute("fontScale", String(Math.round(getPresentationLengthScale(segment) * 100000)));
    autoFit.setAttribute("lnSpcReduction", "12000");
    bodyProperties.append(autoFit);
  }
}

function applyPresentationLineSpacing(paragraph, segment) {
  if (shouldUseSingleLine(segment)) return;
  const paragraphProperties = ensureChild(paragraph, "pPr");
  [...paragraphProperties.children]
    .filter((node) => node.localName === "lnSpc" || node.localName === "spcBef" || node.localName === "spcAft")
    .forEach((node) => node.remove());

  const lineSpacing = document.createElementNS(DRAWING_NS, "a:lnSpc");
  const percent = document.createElementNS(DRAWING_NS, "a:spcPct");
  percent.setAttribute("val", getPptShapeGroupSegments(segment).length > 1 ? "72000" : "85000");
  lineSpacing.append(percent);
  paragraphProperties.append(lineSpacing, createZeroParagraphSpacing("spcBef"), createZeroParagraphSpacing("spcAft"));
}

function createZeroParagraphSpacing(localName) {
  const spacing = document.createElementNS(DRAWING_NS, `a:${localName}`);
  const points = document.createElementNS(DRAWING_NS, "a:spcPts");
  points.setAttribute("val", "0");
  spacing.append(points);
  return spacing;
}

function ensureChild(parent, localName) {
  const existing = [...parent.children].find((node) => node.localName === localName);
  if (existing) return existing;

  const child = document.createElementNS(DRAWING_NS, `a:${localName}`);
  if (localName === "rPr" || localName === "pPr") {
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
  const userScale = getActiveFontScale(segment);
  if (segment?.type !== "pptx") return userScale;

  const mode = getPptLayoutMode();
  const shapeGroup = getPptShapeGroupSegments(segment);
  const groupedTextBox = shapeGroup.length > 1;
  const singleLine = shouldUseSingleLine(segment);

  if (singleLine && !groupedTextBox && !isPptOversizedSingleLineSegment(segment)) {
    return userScale;
  }

  const fitScale = groupedTextBox
    ? getPptShapeGroupFitScale(segment, shapeGroup)
    : getPptTextBoxFitScale(segment, { singleLine });

  if (mode === "compact-fit") {
    return Math.min(getLengthScale(segment.original, segment.translation, segment), fitScale);
  }

  return Math.min(userScale, fitScale);
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

function getPptTextBoxFitScale(segment, options = {}) {
  const layout = segment?.layout || {};
  const box = getPptContentBoxPt(layout);
  if (!box) return getActiveFontScale(segment);

  const text = cleanPptTranslationText(segment?.translation || segment?.original || "", segment?.original || "");
  if (!text) return getActiveFontScale(segment);

  const baseFontPt = Math.max(1, Number(layout.fontSize || 1800) / 100);
  const userScale = getActiveFontScale(segment);
  const maxScale = Math.min(1.15, userScale);
  const minScale = options.singleLine
    ? Math.min(0.24, userScale)
    : Math.max(0.32, Math.min(0.72, userScale));

  if (options.singleLine) {
    const widthAtBase = Math.max(0.1, estimatePptTextWidthPt(text, baseFontPt));
    const lineHeightAtBase = getPptLineHeightPt(segment, baseFontPt);
    const fitScale = Math.min((box.width * 0.98) / widthAtBase, (box.height * 0.92) / Math.max(0.1, lineHeightAtBase));
    return Math.max(minScale, Math.min(maxScale, fitScale));
  }

  for (let scale = maxScale; scale >= minScale; scale -= 0.02) {
    const fontPt = baseFontPt * scale;
    const lines = estimatePptWrappedLineCount(text, fontPt, box.width);
    const lineHeight = getPptLineHeightPt(segment, fontPt);
    if (lines * lineHeight <= box.height * 0.98) {
      return Math.max(minScale, Math.min(maxScale, scale));
    }
  }

  return minScale;
}

function getPptShapeGroupFitScale(segment, group = getPptShapeGroupSegments(segment)) {
  const layout = segment?.layout || {};
  const box = getPptContentBoxPt(layout);
  if (!box || group.length <= 1) return getPptTextBoxFitScale(segment);

  const userScale = getActiveFontScale(segment);
  const maxScale = Math.min(1.15, userScale);
  const minScale = Math.max(0.3, Math.min(0.72, userScale));
  const effectiveWidth = Math.max(1, box.width * 0.78);
  const effectiveHeight = Math.max(1, box.height * 0.82);

  for (let scale = maxScale; scale >= minScale; scale -= 0.02) {
    const totalHeight = group.reduce((sum, item) => {
      const text = cleanPptTranslationText(item.translation || item.original || "", item.original || "");
      if (!text) return sum;
      const fontPt = getPptSegmentFontPt(item) * scale;
      const lines = estimatePptWrappedLineCount(text, fontPt, effectiveWidth);
      return sum + lines * getPptLineHeightPt(item, fontPt) * 1.22;
    }, 0);

    if (totalHeight <= effectiveHeight) {
      return Math.max(minScale, Math.min(maxScale, scale));
    }
  }

  return minScale;
}

function getPptShapeGroupSegments(segment) {
  if (segment?.type !== "pptx") return [];
  const key = getPptShapeGroupKey(segment);
  if (!key) return [segment];

  return state.segments.filter((item) => item.type === "pptx" && getPptShapeGroupKey(item) === key);
}

function getPptShapeGroupKey(segment) {
  const bounds = segment?.layout?.bounds;
  if (!segment?.path || !bounds?.cx || !bounds?.cy) return "";
  return [
    segment.path,
    Math.round(Number(bounds.x || 0)),
    Math.round(Number(bounds.y || 0)),
    Math.round(Number(bounds.cx || 0)),
    Math.round(Number(bounds.cy || 0)),
  ].join("|");
}

function shouldShortenPptSingleLineTranslation(segment, direction) {
  if (direction?.targetCode !== "en" || segment?.type !== "pptx" || !shouldUseSingleLine(segment)) return false;
  const translation = String(segment.translation || "").trim();
  if (!translation || !/[A-Za-z]/.test(translation)) return false;

  const box = getPptContentBoxPt(segment.layout || {});
  if (!box) return false;

  const fontPt = getPptSegmentFontPt(segment);
  const width = estimatePptTextWidthPt(translation, fontPt);
  const budget = getPptTranslationLengthBudget(segment);
  const words = countEnglishWords(translation);

  return width > box.width * 1.08 || words > budget.words + 2 || [...translation].length > budget.characters * 1.35;
}

function isPptOversizedSingleLineSegment(segment) {
  if (segment?.type !== "pptx") return false;
  if (!isSourceSingleLinePptSegment(segment)) return false;

  const box = getPptContentBoxPt(segment.layout || {});
  if (!box) return false;

  const fontPt = getPptSegmentFontPt(segment);
  const lineHeight = getPptLineHeightPt(segment, fontPt);
  return fontPt >= 32 && box.height < lineHeight * 1.05;
}

function buildPptShorteningRequest(segment, attempt = 0) {
  const budget = getPptTranslationLengthBudget(segment);
  const strictLimit = {
    words: Math.max(2, Math.min(budget.words, attempt > 0 ? 3 : budget.words)),
    characters: Math.max(8, Math.min(budget.characters, attempt > 0 ? Math.round(budget.characters * 0.72) : budget.characters)),
  };
  const lines = [
    `Source Chinese: ${segment.original}`,
    `Current English: ${segment.translation}`,
    `Rewrite as a shorter PowerPoint title/label, ideally within ${strictLimit.words} words and about ${strictLimit.characters} characters.`,
  ];

  if (attempt > 0) {
    lines.push("The previous rewrite is still too long for the text box. Be more aggressive: keep only the core meaning, drop the brand or generic words if needed, and use accepted abbreviations.");
  }

  lines.push(
    "Return only the shorter English text.",
  );

  return lines.join("\n");
}

function shouldShortenPptBoxTranslation(segment, direction) {
  if (direction?.targetCode !== "en" || segment?.type !== "pptx" || shouldUseSingleLine(segment)) return false;
  const translation = cleanPptTranslationText(segment.translation || "", segment.original);
  if (!translation || !/[A-Za-z]/.test(translation)) return false;

  const box = getPptContentBoxPt(segment.layout || {});
  if (!box) return false;

  const fitScale = getPptTextBoxFitScale({ ...segment, translation });
  const baseFontPt = getPptSegmentFontPt(segment);
  const sourceLines = getPptSourceVisualLineCount(segment);
  const translatedLines = estimatePptWrappedLineCount(translation, baseFontPt * Math.max(0.35, fitScale), box.width);
  const wordBudget = getPptBoxWordBudget(segment);
  const words = countEnglishWords(translation);
  const isCompactLabel = isPptCompactLabelSegment(segment);

  if (isCompactLabel) {
    return fitScale < 0.82 || translatedLines > sourceLines || words > wordBudget;
  }

  return fitScale < 0.68 || translatedLines > Math.max(sourceLines + 1, Math.ceil(sourceLines * 1.25)) || words > wordBudget * 1.15;
}

function isPptBoxFitTranslationBetter(segment, candidate) {
  const current = cleanPptTranslationText(segment.translation || "", segment.original);
  const next = cleanPptTranslationText(candidate || "", segment.original);
  if (!next) return false;
  if (!current) return true;

  const currentScore = getPptBoxFitScore({ ...segment, translation: current });
  const nextScore = getPptBoxFitScore({ ...segment, translation: next });
  return nextScore <= currentScore * 0.98 || next.length < current.length;
}

function getPptBoxFitScore(segment) {
  const box = getPptContentBoxPt(segment.layout || {});
  const text = cleanPptTranslationText(segment.translation || "", segment.original);
  if (!box || !text) return 0;

  const fitScale = getPptTextBoxFitScale({ ...segment, translation: text });
  const fontPt = getPptSegmentFontPt(segment) * Math.max(0.35, fitScale);
  const lines = estimatePptWrappedLineCount(text, fontPt, box.width);
  const sourceLines = getPptSourceVisualLineCount(segment);
  const overflow = Math.max(0, lines - sourceLines);
  return (1 - fitScale) * 100 + overflow * 12 + countEnglishWords(text) * 0.35 + text.length * 0.05;
}

function buildPptBoxFitRequest(segment, attempt = 0) {
  const sourceLines = getPptSourceVisualLineCount(segment);
  const wordBudget = getPptBoxWordBudget(segment, attempt);
  const compactLabel = isPptCompactLabelSegment(segment);
  const lines = [
    `Source Chinese text box:\n${segment.original}`,
    `Current English:\n${cleanPptTranslationText(segment.translation, segment.original)}`,
    `Rewrite to fit the same PowerPoint text box height. Use no more than ${sourceLines} visual lines and about ${wordBudget} English words.`,
    compactLabel
      ? "This is a short label inside a diagram; use a terse label phrase, not a sentence."
      : "Use compact slide wording, short noun phrases, and accepted abbreviations.",
    "Preserve the core meaning only; remove filler verbs, repeated subjects, and explanatory wording.",
    "Do not add blank lines.",
  ];

  if (attempt > 0) {
    lines.push("The previous rewrite is still too tall. Be more aggressive and keep only essential bullet phrases.");
  }

  lines.push("Return only the rewritten English text.");
  return lines.join("\n\n");
}

function getPptSourceVisualLineCount(segment) {
  const layout = segment?.layout || {};
  const box = getPptContentBoxPt(layout);
  const original = String(segment?.original || "").trim();
  const manualLines = original.split(/\r\n|\r|\n/).filter((line) => line.trim()).length || 1;
  if (!box) return manualLines;

  const fontPt = getPptSegmentFontPt(segment);
  const wrappedLines = estimatePptWrappedLineCount(original, fontPt, box.width);
  return Math.max(1, manualLines, wrappedLines);
}

function getPptBoxWordBudget(segment, attempt = 0) {
  const originalChars = [...String(segment?.original || "").replace(/\s+/g, "")].length;
  const sourceLines = getPptSourceVisualLineCount(segment);
  const labelLimit = isPptCompactLabelSegment(segment)
    ? Math.max(2, Math.min(10, sourceLines * 2 + 1))
    : Infinity;
  const base = Math.min(labelLimit, Math.max(sourceLines * 2, Math.ceil(originalChars * 0.72)));
  const strict = attempt > 0 ? Math.floor(base * 0.72) : base;
  return Math.max(4, Math.min(42, strict));
}

function isPptCompactLabelSegment(segment) {
  if (segment?.type !== "pptx") return false;
  const text = String(segment.original || "").replace(/\s+/g, "");
  if (!text) return false;

  const sourceLines = getPptSourceVisualLineCount(segment);
  const box = getPptContentBoxPt(segment.layout || {});
  const fontPt = getPptSegmentFontPt(segment);
  const shortText = [...text].length <= 24 && sourceLines <= 4;
  const shallowBox = box ? box.height <= getPptLineHeightPt(segment, fontPt) * Math.max(1.8, sourceLines + 0.65) : false;
  return shortText && (shallowBox || sourceLines <= 2);
}

function getPptSegmentFontPt(segment) {
  return Math.max(1, Number(segment?.layout?.fontSize || 1800) / 100);
}

function countEnglishWords(text) {
  return (String(text || "").match(/[A-Za-z0-9]+(?:[-/][A-Za-z0-9]+)*/g) || []).length;
}

function getPptContentBoxPt(layout) {
  const bounds = layout?.bounds;
  if (!bounds?.cx || !bounds?.cy) return null;

  const insets = layout.insets || {};
  const widthEmu = Math.max(1, Number(bounds.cx || 0) - Number(insets.left || 0) - Number(insets.right || 0));
  const heightEmu = Math.max(1, Number(bounds.cy || 0) - Number(insets.top || 0) - Number(insets.bottom || 0));

  return {
    width: widthEmu / PPT_EMU_PER_PT,
    height: heightEmu / PPT_EMU_PER_PT,
  };
}

function getPptLineHeightPt(segment, fontPt) {
  const factor = Number(segment?.layout?.paragraphStyle?.lineHeight || 1);
  return fontPt * Math.max(1.05, Math.min(2.2, factor || 1.12));
}

function estimatePptWrappedLineCount(text, fontPt, maxWidthPt) {
  if (!maxWidthPt || maxWidthPt <= 0) return 1;

  return String(text || "")
    .split(/\r\n|\r|\n/)
    .reduce((total, paragraph) => total + estimatePptParagraphLineCount(paragraph, fontPt, maxWidthPt), 0);
}

function estimatePptParagraphLineCount(paragraph, fontPt, maxWidthPt) {
  const clean = String(paragraph || "").trim();
  if (!clean) return 1;

  const tokens = /[\u3400-\u9fff]/.test(clean)
    ? [...clean]
    : clean.split(/(\s+)/).filter(Boolean);
  let lines = 1;
  let currentWidth = 0;

  tokens.forEach((token) => {
    const width = estimatePptTextWidthPt(token, fontPt);
    if (currentWidth > 0 && currentWidth + width > maxWidthPt) {
      lines += 1;
      currentWidth = 0;
    }

    if (width > maxWidthPt) {
      const charWidths = [...token].map((char) => estimatePptTextWidthPt(char, fontPt));
      charWidths.forEach((charWidth) => {
        if (currentWidth > 0 && currentWidth + charWidth > maxWidthPt) {
          lines += 1;
          currentWidth = 0;
        }
        currentWidth += charWidth;
      });
      return;
    }

    currentWidth += width;
  });

  return lines;
}

function estimatePptTextWidthPt(text, fontPt) {
  const em = [...String(text || "")].reduce((sum, char) => sum + getPptCharWidthEm(char), 0);
  return em * Math.max(1, fontPt);
}

function getPptCharWidthEm(char) {
  if (/[\u3400-\u9fff]/.test(char)) return 1;
  if (/\s/.test(char)) return 0.32;
  if (/[ilI.,;:|!]/.test(char)) return 0.28;
  if (/[fjrt()[\]{}'"]/i.test(char)) return 0.38;
  if (/[A-ZMW@#%&]/.test(char)) return 0.68;
  if (/[0-9]/.test(char)) return 0.54;
  if (/[-_/\\]/.test(char)) return 0.36;
  return 0.52;
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
  if (getPptShapeGroupSegments(segment).length > 1) return false;
  const mode = getPptLayoutMode();
  if (mode === "compact-fit") return false;
  if (mode === "keep-size") return true;
  return isSourceSingleLinePptSegment(segment);
}

function isSourceSingleLinePptSegment(segment) {
  if (segment?.type !== "pptx") return false;
  const layout = segment.layout || {};
  const original = String(segment.original || "").trim();
  if (!original || /[\r\n]/.test(original)) return false;
  if (layout.hasManualBreaks || Number(layout.textBodyParagraphCount || 1) > 1) return false;

  const box = getPptContentBoxPt(layout);
  const baseFontPt = Math.max(1, Number(layout.fontSize || 1800) / 100);
  if (!box) return [...original].length <= 18;

  const sourceWidth = estimatePptTextWidthPt(original, baseFontPt);
  if (sourceWidth <= box.width * 1.08) return true;

  const sourceLines = estimatePptWrappedLineCount(original, baseFontPt, box.width);
  if (sourceLines <= 1) return true;

  const sourceLineHeight = getPptLineHeightPt(segment, baseFontPt);
  return box.height <= sourceLineHeight * 1.45 && [...original].length <= 28;
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
  singleLineInput.checked = shouldUseSingleLine(segment);
  singleLineInput.dataset.index = String(index);
  singleLineInput.addEventListener("change", () => {
    const current = state.segments[Number(singleLineInput.dataset.index)];
    current.overrides.singleLine = singleLineInput.checked;
    current.overrides.wrapMode = singleLineInput.checked ? "single" : "wrap";
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
        pdfBounds: item.overrides?.pdfBounds || null,
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
    pdfBounds: null,
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
    setElementValue(els.summaryDetail, settings.summaryDetail);
    setElementValue(els.modelName, settings.modelName);
    setElementValue(els.apiKey, settings.apiKey);
    setElementValue(els.pdfOutputMode, settings.pdfOutputMode);
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
    summaryDetail: els.summaryDetail?.value || "standard",
    modelName: els.modelName?.value || "",
    apiKey: els.apiKey?.value || "",
    pdfOutputMode: els.pdfOutputMode?.value || "overlay",
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

function toSlidePercentX(value) {
  const width = Number(state.slideSize?.cx || 0) || 1;
  return Math.max(0, (Number(value || 0) / width) * 100);
}

function toSlidePercentY(value) {
  const height = Number(state.slideSize?.cy || 0) || 1;
  return Math.max(0, (Number(value || 0) / height) * 100);
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
  if (shouldUseSingleLine(segment)) return 1;
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
  const minimum = segment?.type === "pptx" && !shouldUseSingleLine(segment) ? 450 : 700;
  return Math.max(minimum, Math.round(baseSize * getPresentationLengthScale(segment)));
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
      <button class="primary" value="ok">我知道了</button>
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

function buildSegmentTranslationInstruction(direction, segment, mode = "translate") {
  const extras = [];

  if (segment?.type === "pdf" || segment?.type === "docx") {
    extras.push([
      segment.type === "pdf"
        ? "This source is extracted from a PDF article or document, sometimes as a paragraph, caption, table cell, reference entry, or figure label."
        : "This source is extracted from a Word document paragraph, header, footer, table cell, or list item.",
      direction.targetCode === "zh"
        ? "Translate meaningful English sentences and phrases into natural Simplified Chinese. The output for ordinary English prose must contain Simplified Chinese characters."
        : "Translate ordinary prose naturally into the target language.",
      "Keep technical acronyms, gene/protein names, chemical symbols, formulas, reference numbers, author names, journal names, and units unchanged when they are standard terms.",
      "If the source is only an acronym, citation number, equation, or isolated symbol, return it unchanged.",
      "Do not summarize, omit, explain, or add markdown.",
    ].join(" "));
  }

  if (segment?.type === "pptx") {
    const targetIsEnglish = direction.targetCode === "en";
    const isSingleLine = shouldUseSingleLine(segment);
    const budget = getPptTranslationLengthBudget(segment);
    const role = getPptSegmentRole(segment);

    extras.push([
      "Use this fixed terminology when applicable:",
      "宝瑞刀/PoriNova = PoriNova; 陡脉冲治疗系统 = Steep-Pulse Therapy System; 不可逆电穿孔 = irreversible electroporation (IRE); 射频 = radiofrequency (RF); 针道止血 = tract hemostasis; 种植转移 = tract seeding; 高频双极性 = high-frequency bipolar; 局麻 = local anesthesia; 肌松 = muscle relaxation.",
      "Keep FDA, MDR, NMPA, CE, ANVISA, MDA, HSA, IRE, RFA, FSA, and PPS unchanged.",
    ].join(" "));

    if (mode === "shorten" || mode === "ultraShort") {
      extras.push([
        "Rewrite the provided current English only.",
        mode === "ultraShort"
          ? "Make it extremely short for a narrow PowerPoint single-line title or label."
          : "Make it shorter for a PowerPoint single-line title or label.",
        `Target: no more than ${budget.words} words and about ${budget.characters} characters if possible.`,
        "Prefer accepted abbreviations, compact noun phrases, ampersands, and slash forms.",
        mode === "ultraShort" ? "Keep only the core meaning; remove brand words or generic descriptors if they are not essential." : "",
        "Return only the shortened English text.",
      ].filter(Boolean).join(" "));
      return [direction.instruction, ...extras].filter(Boolean).join(" ");
    }

    if (mode === "boxFit" || mode === "boxFitStrict") {
      const sourceLines = getPptSourceVisualLineCount(segment);
      const wordBudget = getPptBoxWordBudget(segment, mode === "boxFitStrict" ? 1 : 0);
      const compactLabel = isPptCompactLabelSegment(segment);
      extras.push([
        "Rewrite the provided current English only to fit the original PowerPoint text box height.",
        `Use no more than ${sourceLines} visual lines and about ${wordBudget} English words.`,
        compactLabel ? "This is a short diagram label; use a terse label phrase, not a sentence." : "Use short slide bullet phrases, not explanatory sentences.",
        "Remove filler, repeated subjects, and nonessential adjectives.",
        "Do not add blank lines.",
        mode === "boxFitStrict" ? "Be aggressive: preserve only the essential meaning." : "",
        "Return only the rewritten English text.",
      ].filter(Boolean).join(" "));
      return [direction.instruction, ...extras].filter(Boolean).join(" ");
    }

    if (isSingleLine) {
      extras.push(targetIsEnglish
        ? [
            `This source is a single-line PowerPoint ${role}.`,
            "Return one concise single-line English phrase.",
            `Aim to stay within ${budget.words} words and about ${budget.characters} characters if possible.`,
            "Prefer terse heading wording, noun phrases, standard abbreviations, ampersands (&), slash forms (/), and colon labels when natural.",
            "Omit unnecessary articles, filler words, and explanatory connectors.",
            "Do not expand it into a sentence.",
          ].join(" ")
        : [
            `This source is a single-line PowerPoint ${role}.`,
            "Return one concise single-line translation with similar visual length.",
            "Do not expand it into a sentence.",
          ].join(" "));
    } else {
      const compactLabel = isPptCompactLabelSegment(segment);
      extras.push(targetIsEnglish
        ? (compactLabel ? [
            "This source is a compact PowerPoint diagram label, even if it visually wraps across two lines.",
            `Keep it within ${getPptSourceVisualLineCount(segment)} visual lines and about ${getPptBoxWordBudget(segment)} words.`,
            "Use a terse label phrase, not a sentence.",
            "Do not add blank lines.",
          ].join(" ") : [
            "This source is a multi-line PowerPoint text box.",
            "Keep the translation compact enough to fit the original text box.",
            `Keep close to the original visual height, ideally within ${getPptSourceVisualLineCount(segment)} lines.`,
            "Use compact bullet-style wording or concise sentence fragments where natural.",
            "Avoid explanatory expansion, long relative clauses, and repeated subjects.",
            "Do not add blank lines.",
          ].join(" "))
        : "This source is a multi-line PowerPoint text box; keep the translation compact enough to fit the original text box and do not add blank lines.");
    }
  }

  return [direction.instruction, ...extras].filter(Boolean).join(" ");
}

function buildTargetEnforcementInstruction(direction, segment = null) {
  const base = buildSegmentTranslationInstruction(direction, segment, "translate");
  if (direction?.targetCode === "zh") {
    return [
      base,
      "The previous output did not satisfy the target language requirement.",
      "Translate the provided English source into Simplified Chinese now.",
      "For meaningful English prose, the answer must contain Simplified Chinese characters.",
      "Return only the corrected translation.",
    ].join(" ");
  }
  return [
    base,
    "The previous output did not satisfy the target language requirement. Return only the corrected translation.",
  ].join(" ");
}

function getPptTranslationLengthBudget(segment) {
  const originalLength = Math.max(1, [...String(segment?.original || "").trim()].length);
  const layout = segment?.layout || {};
  const box = getPptContentBoxPt(layout);
  const fontPt = Math.max(1, Number(layout.fontSize || 1800) / 100);
  const visualCharacters = box ? Math.floor((box.width / Math.max(1, fontPt)) / 0.52) : originalLength * 2;
  const characters = Math.max(12, Math.min(72, Math.round(Math.max(originalLength * 1.6, visualCharacters))));
  const words = Math.max(2, Math.min(12, Math.round(characters / 6)));

  return { characters, words };
}

function getPptSegmentRole(segment) {
  const layout = segment?.layout || {};
  const fontPt = getPptSegmentFontPt(segment);
  const y = Number(layout.bounds?.y || 0);
  const topRatio = state.slideSize?.cy ? y / state.slideSize.cy : 1;

  if (fontPt >= 28 || (topRatio < 0.22 && fontPt >= 18)) return "title";
  if (shouldUseSingleLine(segment)) return "label";
  return "body text";
}
