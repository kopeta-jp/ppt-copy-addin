(function (root, factory) {
  "use strict";

  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  if (root) root.TextSplitCore = api;
})(typeof window !== "undefined" ? window : globalThis, function () {
  "use strict";

  function tokenizeLines(text) {
    const source = String(text || "");
    const lines = [];
    const breakPattern = /\r\n|\r|\n|\v/g;
    let lineStart = 0;
    let index = 0;
    let match;

    while ((match = breakPattern.exec(source)) !== null) {
      lines.push({
        index,
        start: lineStart,
        end: match.index,
        text: source.slice(lineStart, match.index)
      });
      index += 1;
      lineStart = match.index + match[0].length;
    }

    lines.push({
      index,
      start: lineStart,
      end: source.length,
      text: source.slice(lineStart)
    });

    return lines;
  }

  function isBlankLine(line) {
    return line.text.trim().length === 0;
  }

  function splitEveryLine(text, lines) {
    return lines
      .filter((line) => !isBlankLine(line))
      .map((line) => ({
        text: line.text,
        start: line.start,
        length: line.end - line.start,
        startLine: line.index,
        lineSpan: 1
      }));
  }

  function splitAtBlankLines(text, lines) {
    const segments = [];
    let block = [];

    function flushBlock() {
      if (block.length === 0) return;
      const first = block[0];
      const last = block[block.length - 1];
      segments.push({
        text: text.slice(first.start, last.end),
        start: first.start,
        length: last.end - first.start,
        startLine: first.index,
        lineSpan: last.index - first.index + 1
      });
      block = [];
    }

    lines.forEach((line) => {
      if (isBlankLine(line)) flushBlock();
      else block.push(line);
    });
    flushBlock();

    return segments;
  }

  function createSplitPlan(text, mode) {
    const source = String(text || "");
    const lines = tokenizeLines(source);
    const segments = mode === "line"
      ? splitEveryLine(source, lines)
      : splitAtBlankLines(source, lines);

    return {
      mode,
      totalLines: Math.max(lines.length, 1),
      segments
    };
  }

  return { createSplitPlan, tokenizeLines };
});
