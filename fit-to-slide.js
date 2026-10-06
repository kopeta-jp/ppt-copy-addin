(function () {
  "use strict";

  const state = {
    ready: false,
    busy: false
  };

  const elements = {};
  const excludedTypes = new Set(["Line", "Placeholder", "Unsupported"]);

  document.addEventListener("DOMContentLoaded", () => {
    elements.fit = document.getElementById("fit-to-slide");
    elements.status = document.getElementById("status");
    elements.fit.addEventListener("click", resizeSelectedShapes);

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

    if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.10")) {
      showStatus("PowerPointApi 1.10以上が必要です。PowerPointを更新してください。", "error");
      return;
    }

    state.ready = true;
    elements.fit.disabled = false;
    showStatus("画像または図形を選択してください。複数選択にも対応しています。", "success");
  }

  async function resizeSelectedShapes() {
    if (!state.ready || state.busy) return;

    setBusy(true);
    showStatus("スライドサイズに合わせています…", "neutral");

    try {
      const result = await PowerPoint.run(async (context) => {
        const pageSetup = context.presentation.pageSetup;
        const shapes = context.presentation.getSelectedShapes();
        const shapeCount = shapes.getCount();

        pageSetup.load("slideWidth,slideHeight");
        shapes.load("items/type");
        await context.sync();

        if (shapeCount.value === 0) {
          throw new UserFacingError("画像または図形が選択されていません。スライド上のオブジェクトを選択してください。");
        }

        let changed = 0;
        let skipped = 0;

        shapes.items.forEach((shape) => {
          if (excludedTypes.has(shape.type)) {
            skipped += 1;
            return;
          }

          shape.left = 0;
          shape.top = 0;
          shape.width = pageSetup.slideWidth;
          shape.height = pageSetup.slideHeight;
          changed += 1;
        });

        if (changed === 0) {
          throw new UserFacingError("選択したオブジェクトは対象外です。画像または図形を選択してください。");
        }

        await context.sync();
        return {
          changed,
          skipped,
          width: pageSetup.slideWidth,
          height: pageSetup.slideHeight
        };
      });

      const skippedNote = result.skipped > 0 ? ` 対象外 ${result.skipped}個は変更していません。` : "";
      showStatus(
        `${result.changed}個をスライド全面に合わせました（${formatPoints(result.width)} × ${formatPoints(result.height)} pt）。${skippedNote}`,
        result.skipped > 0 ? "warning" : "success"
      );
    } catch (error) {
      const message = error instanceof UserFacingError
        ? error.message
        : "サイズ変更に失敗しました。オブジェクトを選び直して、もう一度お試しください。";
      showStatus(message, "error");
      console.error(error);
    } finally {
      setBusy(false);
    }
  }

  function formatPoints(value) {
    return Number.isInteger(value) ? String(value) : value.toFixed(1);
  }

  function setBusy(busy) {
    state.busy = busy;
    elements.fit.disabled = busy || !state.ready;
    elements.fit.classList.toggle("busy", busy);
    elements.fit.querySelector(".button-label").textContent = busy
      ? "処理中…"
      : "スライド全面に合わせる";
  }

  function showStatus(message, kind) {
    elements.status.textContent = message;
    elements.status.className = `status ${kind}`;
  }

  class UserFacingError extends Error {}
})();
