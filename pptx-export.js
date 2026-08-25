(function () {
  "use strict";

  const MIME_PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
  const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const state = { ready: false, busy: false, objectUrl: null };
  const elements = {};

  document.addEventListener("DOMContentLoaded", () => {
    elements.exportButton = document.getElementById("export-button");
    elements.status = document.getElementById("status");
    elements.progressWrap = document.getElementById("progress-wrap");
    elements.progressBar = document.getElementById("progress-bar");
    elements.downloadCard = document.getElementById("download-card");
    elements.downloadLink = document.getElementById("download-link");
    elements.downloadSummary = document.getElementById("download-summary");
    elements.modeInputs = Array.from(document.querySelectorAll('input[name="export-mode"]'));

    elements.exportButton.addEventListener("click", exportSelectedSlides);

    Office.onReady(onOfficeReady).catch((error) => {
      showStatus("PowerPointへの接続に失敗しました。タスクペインを開き直してください。", "error");
      console.error(error);
    });
  });

  function onOfficeReady(info) {
    if (info.host !== Office.HostType.PowerPoint) {
      showStatus("このアドインはPowerPoint専用です。", "error");
      return;
    }
    if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.8")) {
      showStatus("PowerPointApi 1.8以上が必要です。PowerPointを更新してください。", "error");
      return;
    }
    if (typeof JSZip === "undefined") {
      showStatus("PPTX処理ライブラリを読み込めませんでした。GitHub上のlibフォルダを確認してください。", "error");
      return;
    }
    if (typeof PptxGenJS === "undefined") {
      showStatus("高速書き出しライブラリを読み込めませんでした。GitHub上のlibフォルダを確認してください。", "error");
      return;
    }

    state.ready = true;
    elements.exportButton.disabled = false;
    showStatus("準備完了。必要なスライドを選択してください。", "success");
  }

  async function exportSelectedSlides() {
    if (!state.ready || state.busy) return;

    setBusy(true);
    setProgress(4);
    hideDownload();

    try {
      showStatus("選択スライドを確認しています…", "neutral");
      const selection = await getSelectionInfo();
      if (selection.indexes.length === 0) {
        throw new UserFacingError("スライドが選択されていません。左側の一覧から1枚以上選択してください。");
      }

      const mode = getExportMode();
      let blob;
      if (mode === "fast") {
        showStatus(`${selection.indexes.length}枚を高速変換しています…`, "neutral");
        blob = await exportFastImagePresentation(selection.indexes.length);
        setProgress(94);
      } else if (selection.indexes.length === 1) {
        showStatus("選択スライドを書き出しています…", "neutral");
        setProgress(35);
        blob = await exportSingleSelectedSlide();
        setProgress(92);
      } else {
        showStatus(`${selection.indexes.length}枚を読み込んでいます…`, "neutral");
        const sourceBytes = await getWholePresentationBytes((ratio) => {
          setProgress(10 + Math.round(ratio * 40));
        });

        showStatus(`${selection.indexes.length}枚を1つのPPTXにまとめています…`, "neutral");
        blob = await extractSelectedSlides(sourceBytes, selection.indexes, (ratio) => {
          setProgress(52 + Math.round(ratio * 44));
        });
      }

      const fileName = makeFileName(selection.title, selection.indexes.length);
      setDownload(blob, fileName, selection.indexes.length, mode);
      setProgress(100);
      showStatus(`${selection.indexes.length}枚を「${fileName}」として保存しました。`, "success");

      // ユーザー操作から時間が経って自動ダウンロードが拒否されても、表示リンクから再実行できます。
      elements.downloadLink.click();
    } catch (error) {
      const message = error instanceof UserFacingError
        ? error.message
        : "PPTXの作成に失敗しました。資料を一度保存してから、もう一度お試しください。";
      showStatus(message, "error");
      console.error(error);
    } finally {
      setBusy(false);
    }
  }

  async function getSelectionInfo() {
    return PowerPoint.run(async (context) => {
      const presentation = context.presentation;
      const selected = presentation.getSelectedSlides();
      presentation.load("title");
      selected.load("items/index");
      await context.sync();

      const indexes = selected.items.map((slide) => slide.index).sort((a, b) => a - b);
      return { indexes, title: presentation.title || "presentation" };
    });
  }

  async function exportSingleSelectedSlide() {
    const base64 = await PowerPoint.run(async (context) => {
      const selected = context.presentation.getSelectedSlides();
      selected.load("items");
      await context.sync();
      if (selected.items.length !== 1) {
        throw new UserFacingError("選択状態が変わりました。スライドを選び直してください。");
      }
      const result = selected.items[0].exportAsBase64();
      await context.sync();
      return result.value;
    });
    return new Blob([base64ToBytes(base64)], { type: MIME_PPTX });
  }

  async function exportFastImagePresentation(expectedCount) {
    setProgress(12);
    const images = await PowerPoint.run(async (context) => {
      const selected = context.presentation.getSelectedSlides();
      selected.load("items/index");
      await context.sync();
      if (selected.items.length !== expectedCount) {
        throw new UserFacingError("選択状態が変わりました。スライドを選び直してください。");
      }

      const orderedSlides = selected.items.slice().sort((a, b) => a.index - b.index);
      const imageResults = orderedSlides.map((slide) => slide.getImageAsBase64({ width: 1920 }));
      await context.sync();
      return imageResults.map((result) => result.value);
    });

    if (images.length === 0) throw new UserFacingError("スライドが選択されていません。");
    setProgress(68);

    const dimensions = getPngDimensions(images[0]);
    const slideWidth = 10;
    const slideHeight = slideWidth * dimensions.height / dimensions.width;
    const output = new PptxGenJS();
    output.defineLayout({ name: "PPTX_EXPORT_CUSTOM", width: slideWidth, height: slideHeight });
    output.layout = "PPTX_EXPORT_CUSTOM";
    output.author = "PPTX EXPORT";
    output.subject = "Selected slides exported from PowerPoint";
    output.title = "Selected slides";

    images.forEach((base64) => {
      const slide = output.addSlide();
      slide.addImage({
        data: `image/png;base64,${stripDataUrl(base64)}`,
        x: 0,
        y: 0,
        w: slideWidth,
        h: slideHeight
      });
    });

    setProgress(78);
    const result = await output.write({ outputType: "blob", compression: false });
    return result instanceof Blob ? result : new Blob([result], { type: MIME_PPTX });
  }

  function getWholePresentationBytes(onProgress) {
    return new Promise((resolve, reject) => {
      Office.context.document.getFileAsync(
        Office.FileType.Compressed,
        { sliceSize: 4 * 1024 * 1024 },
        async (fileResult) => {
          if (fileResult.status !== Office.AsyncResultStatus.Succeeded) {
            reject(new Error(fileResult.error && fileResult.error.message));
            return;
          }

          const file = fileResult.value;
          try {
            const chunks = [];
            let totalLength = 0;
            for (let index = 0; index < file.sliceCount; index += 1) {
              const slice = await getFileSlice(file, index);
              const bytes = slice.data instanceof Uint8Array ? slice.data : new Uint8Array(slice.data);
              chunks.push(bytes);
              totalLength += bytes.length;
              onProgress((index + 1) / file.sliceCount);
            }

            const combined = new Uint8Array(totalLength);
            let offset = 0;
            chunks.forEach((chunk) => {
              combined.set(chunk, offset);
              offset += chunk.length;
            });
            resolve(combined);
          } catch (error) {
            reject(error);
          } finally {
            file.closeAsync();
          }
        }
      );
    });
  }

  function getFileSlice(file, index) {
    return new Promise((resolve, reject) => {
      file.getSliceAsync(index, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
        else reject(new Error(result.error && result.error.message));
      });
    });
  }

  async function extractSelectedSlides(sourceBytes, selectedIndexes, onProgress) {
    const zip = await JSZip.loadAsync(sourceBytes);
    onProgress(0.08);

    const presentationPath = "ppt/presentation.xml";
    const presentationRelsPath = "ppt/_rels/presentation.xml.rels";
    const presentationXml = await readZipText(zip, presentationPath);
    const presentationRelsXml = await readZipText(zip, presentationRelsPath);
    const presentationDoc = parseXml(presentationXml, presentationPath);
    const relsDoc = parseXml(presentationRelsXml, presentationRelsPath);

    const slideIdList = firstByLocalName(presentationDoc, "sldIdLst");
    if (!slideIdList) throw new Error("Slide list was not found in presentation.xml.");

    const slideIds = Array.from(slideIdList.children).filter((node) => node.localName === "sldId");
    const keepIndexes = new Set(selectedIndexes);
    if (selectedIndexes.some((index) => index < 0 || index >= slideIds.length)) {
      throw new Error("Selected slide indexes do not match the PPTX package.");
    }

    const keepRelationshipIds = new Set();
    slideIds.forEach((slideId, index) => {
      if (keepIndexes.has(index)) {
        keepRelationshipIds.add(slideId.getAttributeNS(REL_NS, "id") || slideId.getAttribute("r:id"));
      } else {
        slideId.remove();
      }
    });

    // 選択外スライドIDを参照し続けるプレゼンテーションレベル機能を除去します。
    removeElementsByLocalName(presentationDoc, "custShowLst");
    removeElementsByLocalName(presentationDoc, "sectionLst");

    Array.from(relsDoc.documentElement.children).forEach((relationship) => {
      const type = relationship.getAttribute("Type") || "";
      const id = relationship.getAttribute("Id");
      if ((type.endsWith("/slide") && !keepRelationshipIds.has(id)) || type.includes("/vbaProject")) {
        relationship.remove();
      }
    });

    zip.file(presentationPath, serializeXml(presentationDoc));
    zip.file(presentationRelsPath, serializeXml(relsDoc));
    await stripInvalidDigitalSignatures(zip);
    await updateExtendedProperties(zip, selectedIndexes.length);
    onProgress(0.25);

    // 選択外スライド、ノート、画像などをパッケージから到達不能にして削除します。
    const reachable = await collectReachableParts(zip);
    Object.keys(zip.files).forEach((name) => {
      if (!zip.files[name].dir && name !== "[Content_Types].xml" && !reachable.has(name)) {
        zip.remove(name);
      }
    });
    onProgress(0.52);

    await cleanContentTypes(zip);
    onProgress(0.62);

    return zip.generateAsync(
      {
        type: "blob",
        mimeType: MIME_PPTX,
        compression: "DEFLATE",
        compressionOptions: { level: 6 }
      },
      (metadata) => onProgress(0.62 + (metadata.percent / 100) * 0.38)
    );
  }

  async function collectReachableParts(zip) {
    const reachable = new Set(["[Content_Types].xml", "_rels/.rels"]);
    const visitedSources = new Set();
    const queue = [""];

    while (queue.length > 0) {
      const sourcePart = queue.shift();
      if (visitedSources.has(sourcePart)) continue;
      visitedSources.add(sourcePart);
      if (sourcePart) reachable.add(sourcePart);

      const relsPath = relationshipPartPath(sourcePart);
      const relsFile = zip.file(relsPath);
      if (!relsFile) continue;
      reachable.add(relsPath);

      const relsDoc = parseXml(await relsFile.async("text"), relsPath);
      Array.from(relsDoc.documentElement.children).forEach((relationship) => {
        if ((relationship.getAttribute("TargetMode") || "").toLowerCase() === "external") return;
        const target = relationship.getAttribute("Target");
        if (!target || /^[a-z][a-z0-9+.-]*:/i.test(target)) return;
        const targetPart = resolveRelationshipTarget(sourcePart, target);
        if (zip.file(targetPart) && !visitedSources.has(targetPart)) queue.push(targetPart);
      });
    }
    return reachable;
  }

  async function stripInvalidDigitalSignatures(zip) {
    const rootRelsPath = "_rels/.rels";
    const file = zip.file(rootRelsPath);
    if (!file) return;
    const doc = parseXml(await file.async("text"), rootRelsPath);
    Array.from(doc.documentElement.children).forEach((relationship) => {
      const type = relationship.getAttribute("Type") || "";
      if (type.includes("digital-signature")) relationship.remove();
    });
    zip.file(rootRelsPath, serializeXml(doc));
  }

  async function updateExtendedProperties(zip, selectedCount) {
    const path = "docProps/app.xml";
    const file = zip.file(path);
    if (!file) return;
    const doc = parseXml(await file.async("text"), path);
    const slides = firstByLocalName(doc, "Slides");
    if (slides) slides.textContent = String(selectedCount);
    removeElementsByLocalName(doc, "HeadingPairs");
    removeElementsByLocalName(doc, "TitlesOfParts");
    zip.file(path, serializeXml(doc));
  }

  async function cleanContentTypes(zip) {
    const path = "[Content_Types].xml";
    const doc = parseXml(await readZipText(zip, path), path);
    Array.from(doc.documentElement.children).forEach((entry) => {
      if (entry.localName !== "Override") return;
      const partName = normalizePartPath((entry.getAttribute("PartName") || "").replace(/^\//, ""));
      if (!zip.file(partName)) {
        entry.remove();
        return;
      }
      const contentType = entry.getAttribute("ContentType") || "";
      if (partName === "ppt/presentation.xml" && contentType.includes("macroEnabled")) {
        entry.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml");
      }
    });
    zip.file(path, serializeXml(doc));
  }

  function relationshipPartPath(sourcePart) {
    if (!sourcePart) return "_rels/.rels";
    const slash = sourcePart.lastIndexOf("/");
    const directory = slash >= 0 ? sourcePart.slice(0, slash + 1) : "";
    const fileName = slash >= 0 ? sourcePart.slice(slash + 1) : sourcePart;
    return `${directory}_rels/${fileName}.rels`;
  }

  function resolveRelationshipTarget(sourcePart, target) {
    let decodedTarget = target;
    try { decodedTarget = decodeURIComponent(target); } catch (_) { /* Keep the original URI. */ }
    if (decodedTarget.startsWith("/")) return normalizePartPath(decodedTarget.slice(1));
    const slash = sourcePart.lastIndexOf("/");
    const directory = slash >= 0 ? sourcePart.slice(0, slash + 1) : "";
    return normalizePartPath(directory + decodedTarget);
  }

  function normalizePartPath(path) {
    const output = [];
    path.split("/").forEach((part) => {
      if (!part || part === ".") return;
      if (part === "..") output.pop();
      else output.push(part);
    });
    return output.join("/");
  }

  function parseXml(text, label) {
    const doc = new DOMParser().parseFromString(text, "application/xml");
    if (doc.getElementsByTagName("parsererror").length > 0) {
      throw new Error(`Invalid XML: ${label}`);
    }
    return doc;
  }

  function serializeXml(doc) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${new XMLSerializer().serializeToString(doc.documentElement)}`;
  }

  function allByLocalName(root, localName) {
    return Array.from(root.getElementsByTagName("*")).filter((node) => node.localName === localName);
  }

  function firstByLocalName(root, localName) {
    return allByLocalName(root, localName)[0] || null;
  }

  function removeElementsByLocalName(root, localName) {
    allByLocalName(root, localName).forEach((node) => node.remove());
  }

  async function readZipText(zip, path) {
    const file = zip.file(path);
    if (!file) throw new Error(`Required PPTX part is missing: ${path}`);
    return file.async("text");
  }

  function base64ToBytes(value) {
    const base64 = stripDataUrl(value);
    const binary = atob(base64.replace(/\s/g, ""));
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
    return bytes;
  }

  function stripDataUrl(value) {
    return value.includes(",") ? value.slice(value.indexOf(",") + 1) : value;
  }

  function getPngDimensions(value) {
    const binary = atob(stripDataUrl(value).replace(/\s/g, "").slice(0, 32));
    if (binary.length < 24 || binary.slice(1, 4) !== "PNG") {
      throw new Error("PowerPoint returned an invalid PNG image.");
    }
    const byte = (index) => binary.charCodeAt(index) & 0xff;
    const unsigned32 = (index) => ((byte(index) * 0x1000000) + (byte(index + 1) << 16) + (byte(index + 2) << 8) + byte(index + 3));
    const width = unsigned32(16);
    const height = unsigned32(20);
    if (!width || !height) throw new Error("PNG dimensions are invalid.");
    return { width, height };
  }

  function getExportMode() {
    const selected = elements.modeInputs.find((input) => input.checked);
    return selected ? selected.value : "fast";
  }

  function makeFileName(title, count) {
    const base = String(title || "presentation")
      .replace(/\.(pptx|pptm)$/i, "")
      .replace(/[\\/:*?"<>|]/g, "-")
      .trim()
      .slice(0, 100) || "presentation";
    return `${base}-selected-${count}-slides.pptx`;
  }

  function setDownload(blob, fileName, count, mode) {
    if (state.objectUrl) URL.revokeObjectURL(state.objectUrl);
    state.objectUrl = URL.createObjectURL(blob);
    elements.downloadLink.href = state.objectUrl;
    elements.downloadLink.download = fileName;
    elements.downloadSummary.textContent = mode === "fast"
      ? `${count}枚を1つのPowerPointファイルに高速でまとめました。各ページは画像です。`
      : `${count}枚を1つの編集可能なPowerPointファイルにまとめました。`;
    elements.downloadCard.classList.remove("hidden");
  }

  function hideDownload() {
    elements.downloadCard.classList.add("hidden");
  }

  function setBusy(busy) {
    state.busy = busy;
    elements.exportButton.disabled = busy || !state.ready;
    elements.exportButton.classList.toggle("busy", busy);
    elements.exportButton.querySelector(".button-label").textContent = busy ? "処理中…" : "選択スライドを保存";
    elements.modeInputs.forEach((input) => { input.disabled = busy; });
    elements.progressWrap.classList.toggle("hidden", !busy);
    elements.progressWrap.setAttribute("aria-hidden", String(!busy));
  }

  function setProgress(percent) {
    elements.progressBar.style.width = `${Math.max(0, Math.min(100, percent))}%`;
  }

  function showStatus(message, kind) {
    elements.status.textContent = message;
    elements.status.className = `status ${kind}`;
  }

  class UserFacingError extends Error {}
})();
