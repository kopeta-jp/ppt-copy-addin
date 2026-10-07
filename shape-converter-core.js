(function (root, factory) {
  "use strict";
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  if (root) root.ShapeConverterCore = api;
})(typeof window !== "undefined" ? window : globalThis, function () {
  "use strict";

  function paragraphSpans(text) {
    const source = String(text || "");
    const spans = [];
    const pattern = /\r\n|\r|\n|\v/g;
    let start = 0;
    let match;

    while ((match = pattern.exec(source)) !== null) {
      spans.push({ start, length: match.index - start });
      start = match.index + match[0].length;
    }
    spans.push({ start, length: source.length - start });
    return spans;
  }

  function sameSnapshot(a, b) {
    if (!a || !b) return a === b;
    const keys = Object.keys(a);
    if (keys.length !== Object.keys(b).length) return false;
    return keys.every((key) => a[key] === b[key]);
  }

  function compressRuns(snapshots) {
    const runs = [];
    snapshots.forEach((snapshot, index) => {
      const previous = runs[runs.length - 1];
      if (previous && sameSnapshot(previous.format, snapshot)) {
        previous.length += 1;
      } else {
        runs.push({ start: index, length: 1, format: snapshot });
      }
    });
    return runs;
  }

  return { paragraphSpans, compressRuns, sameSnapshot };
});
