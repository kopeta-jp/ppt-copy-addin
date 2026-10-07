(function () {
  "use strict";

  const state = { ready: false, busy: false };
  const elements = {};
  const textShapeTypes = new Set(["TextBox", "GeometricShape", "Callout", "Placeholder", "Freeform"]);
  const fontProperties = [
    "allCaps",
    "bold",
    "color",
    "doubleStrikethrough",
    "italic",
    "name",
    "size",
    "smallCaps",
    "strikethrough",
    "subscript",
    "superscript",
    "underline"
  ];

  document.addEventListener("DOMContentLoaded", () => {
    elements.byLine = document.getElementById("split-by-line");
    elements.byBlank = document.getElementById("split-by-blank");
    elements.keepOriginal = document.getElementById("keep-original");
    elements.status = document.getElementById("status");

    elements.byLine.addEventListener("click", () => splitSelectedText("line"));
    elements.byBlank.addEventListener("click", () => splitSelectedText("blank"));

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
    setButtonsDisabled(false);
    showStatus("テキストボックスを1個選択してください。", "success");
  }

  async function splitSelectedText(mode) {
    if (!state.ready || state.busy) return;

    setBusy(true, mode);
    showStatus(mode === "line" ? "改行位置を確認しています…" : "空行位置を確認しています…", "neutral");

    try {
      const keepOriginal = elements.keepOriginal.checked;
      const result = await PowerPoint.run(async (context) => {
        const selectedShapes = context.presentation.getSelectedShapes();
        const shapeCount = selectedShapes.getCount();
        selectedShapes.load("items/type,items/left,items/top,items/width,items/height,items/rotation");
        await context.sync();

        if (shapeCount.value !== 1) {
          throw new UserFacingError(
            shapeCount.value === 0
              ? "テキストボックスが選択されていません。"
              : "テキストボックスを1個だけ選択してください。"
          );
        }

        const sourceShape = selectedShapes.items[0];
        if (!textShapeTypes.has(sourceShape.type)) {
          throw new UserFacingError("選択したオブジェクトから文字を取得できません。テキストボックスを選択してください。");
        }
        if (Math.abs(sourceShape.rotation || 0) > 0.1) {
          throw new UserFacingError("回転したテキストボックスには対応していません。回転を0°にしてから実行してください。");
        }

        const sourceFrame = sourceShape.textFrame;
        const sourceRange = sourceFrame.textRange;
        sourceFrame.load("hasText,leftMargin,rightMargin,topMargin,bottomMargin,wordWrap,verticalAlignment");
        sourceRange.load("text");
        await context.sync();

        if (!sourceFrame.hasText || sourceRange.text.trim().length === 0) {
          throw new UserFacingError("選択したテキストボックスに文字がありません。");
        }

        const plan = window.TextSplitCore.createSplitPlan(sourceRange.text, mode);
        if (plan.segments.length < 2) {
          throw new UserFacingError(
            mode === "line"
              ? "分割できる改行がありません。"
              : "分割できる空行がありません。文章の間に空の行を入れてください。"
          );
        }

        const sourceSegmentRanges = plan.segments.map((segment) => {
          const range = sourceRange.getSubstring(segment.start, segment.length);
          range.load("font,paragraphFormat/horizontalAlignment");
          return range;
        });
        await context.sync();

        const slide = sourceShape.getParentSlide();
        const rowHeight = sourceShape.height / plan.totalLines;
        const createdShapes = [];

        plan.segments.forEach((segment, index) => {
          const segmentRange = sourceSegmentRanges[index];
          const newShape = slide.shapes.addTextBox(segment.text, {
            left: sourceShape.left,
            top: sourceShape.top + (segment.startLine * rowHeight),
            width: sourceShape.width,
            height: Math.max(rowHeight * segment.lineSpan, 1)
          });

          newShape.name = `TEXT SPLIT ${index + 1}`;
          const newFrame = newShape.textFrame;
          newFrame.leftMargin = sourceFrame.leftMargin;
          newFrame.rightMargin = sourceFrame.rightMargin;
          newFrame.topMargin = 0;
          newFrame.bottomMargin = 0;
          newFrame.wordWrap = false;
          newFrame.verticalAlignment = "Top";
          newFrame.autoSizeSetting = "AutoSizeShapeToFitText";

          applyFont(segmentRange.font, newFrame.textRange.font);
          if (segmentRange.paragraphFormat.horizontalAlignment !== null) {
            newFrame.textRange.paragraphFormat.horizontalAlignment = segmentRange.paragraphFormat.horizontalAlignment;
          }
          createdShapes.push(newShape);
        });

        await context.sync();

        if (!keepOriginal) {
          sourceShape.delete();
          await context.sync();
        }

        return { count: createdShapes.length, kept: keepOriginal };
      });

      showStatus(
        `${result.count}個の独立したテキストボックスに分けました。${result.kept ? " 元のボックスも残しています。" : ""}`,
        "success"
      );
    } catch (error) {
      const message = error instanceof UserFacingError
        ? error.message
        : "テキストの分割に失敗しました。選択し直して、もう一度お試しください。";
      showStatus(message, "error");
      console.error(error);
    } finally {
      setBusy(false, mode);
    }
  }

  function applyFont(sourceFont, targetFont) {
    fontProperties.forEach((property) => {
      const value = sourceFont[property];
      if (value !== null && value !== undefined) targetFont[property] = value;
    });
  }

  function setBusy(busy, mode) {
    state.busy = busy;
    setButtonsDisabled(busy || !state.ready);
    elements.byLine.classList.toggle("busy", busy && mode === "line");
    elements.byBlank.classList.toggle("busy", busy && mode === "blank");
  }

  function setButtonsDisabled(disabled) {
    elements.byLine.disabled = disabled;
    elements.byBlank.disabled = disabled;
  }

  function showStatus(message, kind) {
    elements.status.textContent = message;
    elements.status.className = `status ${kind}`;
  }

  class UserFacingError extends Error {}
})();
