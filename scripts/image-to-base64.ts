/**
 * Inline a real base64 data URL into the primary image example of the
 * docx/pptx images docs (en + zh), so the example is self-contained and the
 * document exported from the docs site carries an actual picture.
 *
 * Only the FIRST `"data": "<Uint8Array>"` placeholder per file is replaced
 * (the basic/primary example). Any previously inlined data URL is reverted to
 * the placeholder first, so re-running with a different image swaps it cleanly.
 *
 * Usage: pnpm -C <root> exec tsx scripts/image-to-base64.ts
 */
import { readFileSync, writeFileSync } from "node:fs";

// Smallest project-owned raster image so the inlined docs stay small.
const IMAGE_PATH = "packages/pptx/demo/assets/test-poster.png";
const TARGETS = [
  "docs/content/en/04.docx/05.images.md",
  "docs/content/zh/04.docx/05.images.md",
  "docs/content/en/05.pptx/05.images.md",
  "docs/content/zh/05.pptx/05.images.md",
];

const PLACEHOLDER = '"data": "<Uint8Array>"';
// Matches a previously inlined data URL value (revert before re-inlining).
const DATA_URL_RE = /"data": "data:image\/[\w.+-]+;base64,[^"]*"/;

const bytes = readFileSync(IMAGE_PATH);
const dataUrl = `data:image/png;base64,${Buffer.from(bytes).toString("base64")}`;
const replacement = `"data": "${dataUrl}"`;

let updated = 0;
for (const file of TARGETS) {
  // Revert any previously inlined data URL back to the placeholder.
  let src = readFileSync(file, "utf8").replace(DATA_URL_RE, PLACEHOLDER);
  const index = src.indexOf(PLACEHOLDER);
  if (index === -1) {
    console.log(`SKIP (no placeholder): ${file}`);
    continue;
  }
  src = src.slice(0, index) + replacement + src.slice(index + PLACEHOLDER.length);
  writeFileSync(file, src);
  console.log(`Updated: ${file}`);
  updated++;
}
console.log(`Done. ${updated}/${TARGETS.length} files updated.`);
