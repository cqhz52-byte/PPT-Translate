const state = {
  file: null,
  fileType: "",
  zip: null,
  segments: [],
  slideCount: 0,
  installPrompt: null,
};

const els = {
  fileInput: document.querySelector("#fileInput"),
  fileMeta: document.querySelector("#fileMeta"),
  segmentTable: document.querySelector("#segmentTable"),
  translateButton: document.querySelector("#translateButton"),
  previewButton: document.querySelector("#previewButton"),
  downloadButton: document.querySelector("#downloadButton"),
  resetButton: document.querySelector("#resetButton"),
  installButton: document.querySelector("#installButton"),
  previewDialog: document.querySelector("#previewDialog"),
  previewCloseButton: document.querySelector("#previewCloseButton"),
  previewBody: document.querySelector("#previewBody"),
  previewMeta: document.querySelector("#previewMeta"),
  translationDirection: document.querySelector("#translationDirection"),
  pptLayoutMode: document.querySelector("#pptLayoutMode"),
  apiBase: document.querySelector("#apiBase"),
  apiProxy: document.querySelector("#apiProxy"),
  modelName: document.querySelector("#modelName"),
  apiKey: document.querySelector("#apiKey"),
  slideCount: document.querySelector("#slideCount"),
  segmentCount: document.querySelector("#segmentCount"),
  translatedCount: document.querySelector("#translatedCount"),
  statusText: document.querySelector("#statusText"),
  progressFill: document.querySelector("#progressFill"),
};

const slidePathPattern = /^ppt\/slides\/slide(\d+)\.xml$/;
const wordPathPattern = /^word\/(?:document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/;
const DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const SETTINGS_KEY = "deepseek-document-translator-settings-v1";
const SETTINGS_VERSION = 3;
const parser = new DOMParser();
const serializer = new XMLSerializer();

loadSettings();

if ("serviceWorker" in navigator) {
  navigator.serviceWorker.register("sw.js?v=10").catch(() => {
    showToast("PWA 缓存注册失败，应用仍可在浏览器中使用。", true);
  });
}

window.addEventListener("beforeinstallprompt", (event) => {
  event.preventDefault();
  state.installPrompt = event;
  els.installButton.hidden = false;
});

els.installButton.addEventListener("click", async () => {
  if (!state.installPrompt) return;
  state.installPrompt.prompt();
  await state.installPrompt.userChoice;
  state.installPrompt = null;
  els.installButton.hidden = true;
});

els.fileInput.addEventListener("change", async (event) => {
  const [file] = event.target.files;
  if (!file) return;
  await loadOfficeFile(file);
});

els.translateButton.addEventListener("click", translateAll);
els.previewButton.addEventListener("click", openPreview);
els.downloadButton.addEventListener("click", downloadPresentation);
els.resetButton.addEventListener("click", resetApp);
els.previewCloseButton.addEventListener("click", closePreview);
els.previewDialog.addEventListener("click", (event) => {
  if (event.target === els.previewDialog) closePreview();
});

[
  els.translationDirection,
  els.pptLayoutMode,
  els.apiBase,
  els.apiProxy,
  els.modelName,
  els.apiKey,
].forEach((element) => {
  element?.addEventListener("change", saveSettings);
  element?.addEventListener("input", saveSettings);
});

async function loadOfficeFile(file) {
  try {
    if (/\.ppt$/i.test(file.name) && !/\.pptx$/i.test(file.name)) {
      throw new Error("当前网页版只能可靠解析和回写 PPTX。请先用 PowerPoint/WPS 将 .ppt 另存为 .pptx。");
    }

    if (/\.doc$/i.test(file.name) && !/\.docx$/i.test(file.name)) {
      throw new Error("当前网页版只能可靠解析和回写 DOCX。请先用 Word/WPS 将 .doc 另存为 .docx。");
    }

    if (!/\.(pptx|docx)$/i.test(file.name)) {
      throw new Error("请选择 .pptx 或 .docx 文件。");
    }

    setBusy(true, "正在读取文件...");
    state.file = file;
    state.fileType = /\.docx$/i.test(file.name) ? "docx" : "pptx";
    state.zip = await JSZip.loadAsync(file);
    state.segments = [];
    state.slideCount = 0;

    if (state.fileType === "docx") {
      await loadWordDocument();
    } else {
      await loadPresentation();
    }

    renderSegments();
    updateStats();
    els.fileMeta.textContent = `${file.name} · ${(file.size / 1024 / 1024).toFixed(2)} MB`;
    setStatus(`已解析 ${getFileTypeName()}，共 ${state.segments.length} 段文字。默认使用 DeepSeek 翻译。`);
    showToast(`${getFileTypeName()} 已载入，可以开始翻译或手动编辑。`);
  } catch (error) {
    showToast(error.message || "读取文件失败。", true);
    resetApp(false);
  } finally {
    setBusy(false);
  }
}

async function loadPresentation() {
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
          original,
          translation: "",
        });
      });
    }
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
      const textNodes = [...paragraph.getElementsByTagNameNS(WORD_NS, "t")];
      const original = textNodes.map((node) => node.textContent || "").join("").trim();
      if (!original) return;

      state.segments.push({
        id: `${path}-${paragraphIndex}`,
        type: "docx",
        slideNumber: partLabel,
        locationLabel: partLabel,
        path,
        paragraphIndex,
        textNodeCount: textNodes.length,
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
    });
    targetCell.append(textarea);

    row.append(slideCell, sourceCell, targetCell);
    fragment.append(row);
  });

  els.segmentTable.replaceChildren(fragment);
}

async function translateAll() {
  const apiKey = els.apiKey.value.trim();
  const apiBase = els.apiBase.value.trim().replace(/\/$/, "");
  const apiProxy = els.apiProxy.value.trim() || "./api/translate";
  const model = els.modelName.value.trim();
  const direction = getDirectionConfig(els.translationDirection.value);

  if (!apiKey || !apiBase || !model) {
    showToast("请先填写接口地址、模型和 API Key。", true);
    return;
  }

  try {
    setBusy(true, "正在翻译...");
    const untranslated = state.segments.filter((segment) => !segment.translation.trim());

    for (let index = 0; index < untranslated.length; index += 1) {
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
      });
      updateTextarea(segment);
      updateStats();
    }

    setProgress(1);
    setStatus("翻译完成，请检查译文后导出。");
    showToast("自动翻译完成。");
  } catch (error) {
    showToast(error.message || "翻译失败。", true);
    setStatus("翻译中断，请检查接口配置或网络连接。");
  } finally {
    setBusy(false);
  }
}

async function translateText({ apiBase, apiProxy, apiKey, model, direction, text }) {
  const response = await fetch(apiProxy, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
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

async function downloadPresentation() {
  if (!state.zip || !state.file) return;

  try {
    setBusy(true, `正在生成 ${getFileTypeName()}...`);
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
      mimeType:
        state.fileType === "docx"
          ? "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          : "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    });

    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = buildOutputName(state.file.name);
    document.body.append(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    showToast(`已生成翻译版 ${getFileTypeName()}。`);
  } catch (error) {
    showToast(error.message || "导出失败。", true);
  } finally {
    setBusy(false);
  }
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

    textNodes[0].textContent = segment.translation.trim();
    normalizeWordRunStyle(textNodes[0], segment.original, segment.translation);
    clearRemainingTextNodes(textNodes);
  });
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
  els.translateButton.disabled = !state.segments.length;
  els.previewButton.disabled = !state.segments.length;
  els.downloadButton.disabled = !state.segments.length;
  setProgress(state.segments.length ? translated / state.segments.length : 0);
}

function setBusy(isBusy, message = "") {
  els.translateButton.disabled = isBusy || !state.segments.length;
  els.previewButton.disabled = isBusy || !state.segments.length;
  els.downloadButton.disabled = isBusy || !state.segments.length;
  els.fileInput.disabled = isBusy;
  if (message) setStatus(message);
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

function renderPreview() {
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

    group.segments.forEach((segment) => {
      const item = document.createElement("article");
      item.className = `preview-item${segment.translation.trim() ? "" : " pending"}`;

      const text = document.createElement("p");
      text.textContent = segment.translation.trim() || segment.original;

      const meta = document.createElement("span");
      meta.textContent = segment.translation.trim() ? "译文" : "未翻译，暂用原文";

      item.append(text, meta);
      section.append(item);
    });

    els.previewBody.append(section);
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
}

function resetApp(clearInput = true) {
  state.file = null;
  state.fileType = "";
  state.zip = null;
  state.segments = [];
  state.slideCount = 0;
  if (clearInput) els.fileInput.value = "";
  els.fileMeta.textContent = "支持 PPTX / DOCX；旧版 PPT/DOC 请先另存";
  renderSegments();
  updateStats();
  setStatus("请先选择一个 PPTX 或 DOCX 文件。");
}

function buildOutputName(name) {
  return name.replace(/\.(pptx|docx)$/i, "") + `-translated.${state.fileType || "pptx"}`;
}

function normalizeTranslation(translation, source) {
  const clean = translation.trim();
  if (/[\r\n]/.test(source)) return clean;
  return clean.replace(/\s*[\r\n]+\s*/g, " ").replace(/[ \t]{2,}/g, " ");
}

function normalizePresentationRunStyle(textNode, segment) {
  const run = textNode.parentElement;
  if (!run) return;

  const runProperties = ensureChild(run, "rPr");
  const scale = getPresentationLengthScale(segment);
  const currentSize = Number(runProperties.getAttribute("sz") || 0);

  if (currentSize > 0 && scale < 1) {
    const nextSize = Math.max(900, Math.round(currentSize * scale));
    runProperties.setAttribute("sz", String(nextSize));
  }

  setTypeface(runProperties, "latin", "Arial");
  setTypeface(runProperties, "ea", "Microsoft YaHei");
  setTypeface(runProperties, "cs", "Arial");

  [...runProperties.getElementsByTagNameNS(DRAWING_NS, "sym")].forEach((node) => node.remove());
}

function normalizeWordRunStyle(textNode, source, translation) {
  const run = textNode.parentElement;
  if (!run) return;

  const runProperties = ensureChildNS(run, WORD_NS, "w:rPr", "rPr");
  const fonts = ensureChildNS(runProperties, WORD_NS, "w:rFonts", "rFonts");
  fonts.setAttributeNS(WORD_NS, "w:ascii", "Arial");
  fonts.setAttributeNS(WORD_NS, "w:hAnsi", "Arial");
  fonts.setAttributeNS(WORD_NS, "w:eastAsia", "Microsoft YaHei");
  fonts.setAttributeNS(WORD_NS, "w:cs", "Arial");

  const scale = getLengthScale(source, translation);
  const sizeNode = [...runProperties.children].find((node) => node.localName === "sz");
  const currentSize = Number(sizeNode?.getAttributeNS(WORD_NS, "val") || sizeNode?.getAttribute("w:val") || 0);

  if (sizeNode && currentSize > 0 && scale < 1) {
    sizeNode.setAttributeNS(WORD_NS, "w:val", String(Math.max(16, Math.round(currentSize * scale))));
  }
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
    const noAutofit = document.createElementNS(DRAWING_NS, "a:noAutofit");
    bodyProperties.append(noAutofit);
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

function getLengthScale(source, translation) {
  const sourceLength = Math.max(1, [...source].length);
  const translationLength = Math.max(1, [...translation].length);
  const ratio = translationLength / sourceLength;

  if (ratio <= 1.25) return 1;
  if (ratio <= 1.75) return 0.94;
  if (ratio <= 2.4) return 0.9;
  return 0.88;
}

function getPresentationLengthScale(segment) {
  if (getPptLayoutMode() === "compact-fit" || shouldUseSingleLine(segment)) {
    return getLengthScale(segment.original, segment.translation);
  }

  return 1;
}

function getPptLayoutMode() {
  return els.pptLayoutMode?.value || "smart";
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
  };
}

function shouldUseSingleLine(segment) {
  const mode = getPptLayoutMode();
  if (mode === "compact-fit") return false;
  if (mode === "keep-size") return true;

  const layout = segment.layout || {};
  if (layout.hasManualBreaks) return false;
  if (layout.textBodyParagraphCount > 1) return false;

  return true;
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
  return state.fileType === "docx" ? "DOCX" : "PPTX";
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
    setElementValue(els.apiBase, settings.apiBase);
    setElementValue(els.apiProxy, settings.apiProxy);
    setElementValue(els.modelName, settings.modelName);
    setElementValue(els.apiKey, settings.apiKey);
  } catch {
    localStorage.removeItem(SETTINGS_KEY);
  }
}

function saveSettings() {
  const settings = {
    settingsVersion: SETTINGS_VERSION,
    translationDirection: els.translationDirection?.value || "",
    pptLayoutMode: els.pptLayoutMode?.value || "",
    apiBase: els.apiBase?.value || "",
    apiProxy: els.apiProxy?.value || "",
    modelName: els.modelName?.value || "",
    apiKey: els.apiKey?.value || "",
  };

  localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
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

function getDirectionConfig(value) {
  const configs = {
    "zh-en": {
      instruction:
        "Translate Simplified or Traditional Chinese presentation text into fluent, concise English for business slides. Use short noun phrases for titles, labels, and table-of-contents entries.",
    },
    "en-zh": {
      instruction:
        "Translate English presentation text into natural Simplified Chinese for business slides.",
    },
    "auto-en": {
      instruction:
        "Detect the source language and translate the text into fluent, concise English for business slides. Use short noun phrases for titles, labels, and table-of-contents entries.",
    },
    "auto-zh": {
      instruction:
        "Detect the source language and translate the text into natural Simplified Chinese for business slides.",
    },
  };

  return configs[value] || configs["zh-en"];
}
