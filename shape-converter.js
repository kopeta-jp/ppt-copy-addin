(function () {
  "use strict";

  const state = { ready: false, busy: false };
  const elements = {};
  const maxRichTextCharacters = 800;
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
    elements.toRounded = document.getElementById("to-rounded");
    elements.toRectangle = document.getElementById("to-rectangle");
    elements.status = document.getElementById("status");

    elements.toRounded.addEventListener("click", () => convertSelectedShape("RoundRectangle"));
    elements.toRectangle.addEventListener("click", () => convertSelectedShape("Rectangle"));

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
    showStatus("変換する四角形を1個選択してください。", "success");
  }

  async function convertSelectedShape(targetType) {
    if (!state.ready || state.busy) return;
    setBusy(true, targetType);
    showStatus("図形と書式を読み取っています…", "neutral");

    try {
      const result = await PowerPoint.run(async (context) => {
        const selected = context.presentation.getSelectedShapes();
        const count = selected.getCount();
        selected.load("items/type,items/left,items/top,items/width,items/height,items/rotation,items/name,items/zOrderPosition,items/visible,items/altTextTitle,items/altTextDescription");
        await context.sync();

        if (count.value !== 1) {
          throw new UserFacingError(
            count.value === 0 ? "図形が選択されていません。" : "図形を1個だけ選択してください。"
          );
        }

        const source = selected.items[0];
        if (source.type !== "GeometricShape") {
          throw new UserFacingError("正方形、長方形、または角丸四角形を選択してください。");
        }

        const sourceFill = source.fill;
        const sourceLine = source.lineFormat;
        const sourceAdjustments = source.adjustments;
        const sourceFrame = source.getTextFrameOrNullObject();

        sourceFill.load("type,foregroundColor,transparency");
        sourceLine.load("visible,color,transparency,weight,dashStyle,style");
        sourceAdjustments.load("count");
        sourceFrame.load("hasText,leftMargin,rightMargin,topMargin,bottomMargin,wordWrap,verticalAlignment,autoSizeSetting");
        await context.sync();

        validateSource(sourceFill, sourceAdjustments.count, targetType);

        let sourceRange = null;
        let text = "";
        let fullFont = null;
        let fontRuns = [];
        let paragraphFormats = [];

        if (!sourceFrame.isNullObject && sourceFrame.hasText) {
          sourceRange = sourceFrame.textRange;
          sourceRange.load("text,font");
          await context.sync();
          text = sourceRange.text;
          fullFont = snapshotFont(sourceRange.font);

          if (hasMixedFont(fullFont) && text.length <= maxRichTextCharacters) {
            const characterRanges = [];
            for (let index = 0; index < text.length; index += 1) {
              const characterRange = sourceRange.getSubstring(index, 1);
              characterRange.load("font");
              characterRanges.push(characterRange);
            }
            await context.sync();
            fontRuns = window.ShapeConverterCore.compressRuns(
              characterRanges.map((range) => snapshotFont(range.font))
            );
          }

          const spans = window.ShapeConverterCore.paragraphSpans(text).filter((span) => span.length > 0);
          const paragraphRanges = spans.map((span) => {
            const range = sourceRange.getSubstring(span.start, span.length);
            range.load("paragraphFormat/horizontalAlignment,paragraphFormat/indentLevel,paragraphFormat/bulletFormat");
            return range;
          });
          await context.sync();
          paragraphFormats = paragraphRanges.map((range, index) => ({
            span: spans[index],
            horizontalAlignment: range.paragraphFormat.horizontalAlignment,
            indentLevel: range.paragraphFormat.indentLevel,
            bullet: {
              visible: range.paragraphFormat.bulletFormat.visible,
              type: range.paragraphFormat.bulletFormat.type,
              style: range.paragraphFormat.bulletFormat.style
            }
          }));
        }

        const slide = source.getParentSlide();
        const replacement = slide.shapes.addGeometricShape(targetType, {
          left: source.left,
          top: source.top,
          width: source.width,
          height: source.height
        });

        replacement.rotation = source.rotation;
        replacement.visible = source.visible;
        replacement.altTextTitle = source.altTextTitle;
        replacement.altTextDescription = source.altTextDescription;

        copyFill(sourceFill, replacement.fill);
        copyLine(sourceLine, replacement.lineFormat);

        if (!sourceFrame.isNullObject) {
          const targetFrame = replacement.textFrame;
          targetFrame.leftMargin = sourceFrame.leftMargin;
          targetFrame.rightMargin = sourceFrame.rightMargin;
          targetFrame.topMargin = sourceFrame.topMargin;
          targetFrame.bottomMargin = sourceFrame.bottomMargin;
          targetFrame.wordWrap = sourceFrame.wordWrap;
          targetFrame.verticalAlignment = sourceFrame.verticalAlignment;
          if (sourceFrame.autoSizeSetting !== "AutoSizeMixed") {
            targetFrame.autoSizeSetting = sourceFrame.autoSizeSetting;
          }

          if (sourceFrame.hasText) {
            targetFrame.textRange.text = text;
            applyFont(fullFont, targetFrame.textRange.font);
            fontRuns.forEach((run) => {
              applyFont(run.format, targetFrame.textRange.getSubstring(run.start, run.length).font);
            });
            paragraphFormats.forEach((format) => {
              const targetParagraph = targetFrame.textRange.getSubstring(format.span.start, format.span.length);
              applyParagraph(format, targetParagraph.paragraphFormat);
            });
          }
        }

        replacement.setZOrder("SendToBack");
        for (let index = 0; index < source.zOrderPosition; index += 1) {
          replacement.setZOrder("BringForward");
        }

        await context.sync();
        source.delete();
        replacement.name = source.name;
        await context.sync();

        return {
          targetType,
          simplifiedRichText: text.length > maxRichTextCharacters && hasMixedFont(fullFont)
        };
      });

      const label = result.targetType === "RoundRectangle" ? "角丸四角形" : "正方形／長方形";
      const note = result.simplifiedRichText
        ? " 文字数が多いため、混在する文字書式の一部は統一されています。"
        : "";
      showStatus(`${label}へ変換しました。${note}`, result.simplifiedRichText ? "warning" : "success");
    } catch (error) {
      const message = error instanceof UserFacingError
        ? error.message
        : "図形の変換に失敗しました。図形を選び直して、もう一度お試しください。";
      showStatus(message, "error");
      console.error(error);
    } finally {
      setBusy(false, targetType);
    }
  }

  function validateSource(fill, adjustmentCount, targetType) {
    if (fill.type !== "Solid" && fill.type !== "NoFill") {
      throw new UserFacingError("グラデーション、パターン、画像塗りの図形は書式を保持できないため変換しません。");
    }
    if (targetType === "RoundRectangle" && adjustmentCount !== 0) {
      throw new UserFacingError("標準の正方形または長方形を選択してください。");
    }
    if (targetType === "Rectangle" && adjustmentCount !== 1) {
      throw new UserFacingError("標準の角丸四角形を選択してください。");
    }
  }

  function copyFill(source, target) {
    if (source.type === "NoFill") {
      target.clear();
      return;
    }
    target.setSolidColor(source.foregroundColor);
    if (source.transparency !== null) target.transparency = source.transparency;
  }

  function copyLine(source, target) {
    if (source.visible !== null) target.visible = source.visible;
    if (!source.visible) return;
    if (source.color !== null) target.color = source.color;
    if (source.transparency !== null) target.transparency = source.transparency;
    if (source.weight !== null) target.weight = source.weight;
    if (source.dashStyle !== null) target.dashStyle = source.dashStyle;
    if (source.style !== null) target.style = source.style;
  }

  function snapshotFont(font) {
    const result = {};
    fontProperties.forEach((property) => { result[property] = font[property]; });
    return result;
  }

  function hasMixedFont(font) {
    if (!font) return false;
    return fontProperties.some((property) => font[property] === null);
  }

  function applyFont(source, target) {
    if (!source) return;
    fontProperties.forEach((property) => {
      const value = source[property];
      if (value !== null && value !== undefined) target[property] = value;
    });
  }

  function applyParagraph(source, target) {
    if (source.horizontalAlignment !== null) target.horizontalAlignment = source.horizontalAlignment;
    if (source.indentLevel !== null && source.indentLevel !== undefined) target.indentLevel = source.indentLevel;
    if (source.bullet.visible !== null) target.bulletFormat.visible = source.bullet.visible;
    if (source.bullet.type !== null && source.bullet.type !== "Unsupported") {
      target.bulletFormat.type = source.bullet.type;
    }
    if (source.bullet.style !== null && source.bullet.style !== "Unsupported") {
      target.bulletFormat.style = source.bullet.style;
    }
  }

  function setBusy(busy, targetType) {
    state.busy = busy;
    setButtonsDisabled(busy || !state.ready);
    elements.toRounded.classList.toggle("busy", busy && targetType === "RoundRectangle");
    elements.toRectangle.classList.toggle("busy", busy && targetType === "Rectangle");
  }

  function setButtonsDisabled(disabled) {
    elements.toRounded.disabled = disabled;
    elements.toRectangle.disabled = disabled;
  }

  function showStatus(message, kind) {
    elements.status.textContent = message;
    elements.status.className = `status ${kind}`;
  }

  class UserFacingError extends Error {}
})();
