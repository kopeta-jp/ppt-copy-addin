(function () {
  "use strict";

  const state = {
    ready: false,
    busy: false,
    pngBlob: null,
    objectUrl: null,
    fileName: "slide.png"
  };

  const elements = {};

  document.addEventListener("DOMContentLoaded", () => {
    elements.copy = document.getElementById("copy-slide");
    elements.status = document.getElementById("status");
    elements.fallback = document.getElementById("fallback");
    elements.preview = document.getElementById("preview");
    elements.retry = document.getElementById("retry-copy");
    elements.download = document.getElementById("download");

    elements.copy.addEventListener("click", copySelectedSlide);
    elements.retry.addEventListener("click", retryClipboardCopy);

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

    state.ready = true;
    elements.copy.disabled = false;
    showStatus("準備完了。スライドを1枚選んでボタンを押してください。", "success");
  }

  function copySelectedSlide() {
    if (!state.ready || state.busy) return;

    setBusy(true);
    hideFallback();
    showStatus("スライドをPNGに変換しています…", "neutral");

    // WebKitは、ユーザー操作から時間が経つとclipboard.writeを拒否することがあります。
    // そこでボタンイベント内で直ちにwriteを呼び、PNG生成をPromiseとして渡します。
    const renderPromise = renderSelectedSlide().then((result) => {
      setPreview(result.blob, result.fileName);
      return result;
    });

    let clipboardPromise;
    try {
      clipboardPromise = beginClipboardWrite(renderPromise.then((result) => result.blob));
    } catch (error) {
      clipboardPromise = Promise.reject(error);
    }

    Promise.all([renderPromise, clipboardPromise])
      .then(() => {
        showStatus("コピーしました。貼り付け先で ⌘V を押してください。", "success");
      })
      .catch(async (copyOrRenderError) => {
        try {
          await renderPromise;
          showFallback();
          showStatus("PNGは生成できましたが、直接コピーは許可されませんでした。下の方法を利用してください。", "warning");
          console.warn("Clipboard write failed", copyOrRenderError);
        } catch (renderError) {
          handleRenderError(renderError);
        }
      })
      .finally(() => setBusy(false));
  }

  async function renderSelectedSlide() {
    return PowerPoint.run(async (context) => {
      const slides = context.presentation.getSelectedSlides();
      slides.load("items");
      await context.sync();

      if (slides.items.length !== 1) {
        throw new UserFacingError(
          slides.items.length === 0
            ? "スライドが選択されていません。左側の一覧から1枚選択してください。"
            : "スライドを1枚だけ選択してください。"
        );
      }

      const imageResult = slides.items[0].getImageAsBase64({ width: 2560 });
      await context.sync();

      if (!imageResult.value) {
        throw new Error("PowerPoint returned an empty image.");
      }

      const blob = base64PngToBlob(imageResult.value);
      return {
        blob,
        fileName: makeFileName()
      };
    });
  }

  function beginClipboardWrite(pngBlobPromise) {
    if (!navigator.clipboard || typeof navigator.clipboard.write !== "function" || typeof ClipboardItem === "undefined") {
      throw new Error("Clipboard image API is unavailable in this WebView.");
    }

    const item = new ClipboardItem({ "image/png": pngBlobPromise });
    return navigator.clipboard.write([item]);
  }

  async function retryClipboardCopy() {
    if (!state.pngBlob) return;
    try {
      // 画像はすでに生成済みなので、この操作ではユーザー操作の直後に書き込めます。
      await beginClipboardWrite(Promise.resolve(state.pngBlob));
      showStatus("コピーしました。貼り付け先で ⌘V を押してください。", "success");
    } catch (error) {
      showFallback();
      showStatus("このPowerPointでは直接コピーできません。右クリックまたはPNG保存を利用してください。", "warning");
      console.warn("Clipboard retry failed", error);
    }
  }

  function base64PngToBlob(value) {
    const base64 = value.includes(",") ? value.slice(value.indexOf(",") + 1) : value;
    const binary = atob(base64.replace(/\s/g, ""));
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) {
      bytes[i] = binary.charCodeAt(i);
    }
    return new Blob([bytes], { type: "image/png" });
  }

  function setPreview(blob, fileName) {
    if (state.objectUrl) URL.revokeObjectURL(state.objectUrl);
    state.pngBlob = blob;
    state.fileName = fileName;
    state.objectUrl = URL.createObjectURL(blob);
    elements.preview.src = state.objectUrl;
    elements.download.href = state.objectUrl;
    elements.download.download = fileName;
  }

  function showFallback() {
    elements.fallback.classList.remove("hidden");
  }

  function hideFallback() {
    elements.fallback.classList.add("hidden");
  }

  function setBusy(busy) {
    state.busy = busy;
    elements.copy.disabled = busy || !state.ready;
    elements.copy.classList.toggle("busy", busy);
    elements.copy.querySelector(".button-label").textContent = busy ? "処理中…" : "画像としてコピー";
  }

  function showStatus(message, kind) {
    elements.status.textContent = message;
    elements.status.className = `status ${kind}`;
  }

  function handleRenderError(error) {
    const message = error instanceof UserFacingError
      ? error.message
      : "PNGの生成に失敗しました。スライドを選び直して、もう一度お試しください。";
    showStatus(message, "error");
    console.error(error);
  }

  function makeFileName() {
    const now = new Date();
    const pad = (number) => String(number).padStart(2, "0");
    return `slide-${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}-${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}.png`;
  }

  class UserFacingError extends Error {}
})();
