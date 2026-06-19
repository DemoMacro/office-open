import {
  PPTX_NS,
  OoxmlMimeType,
  appendRelationship,
  applyCorePropertiesOverride,
  collectPlaceholderKeys,
  createReplacer,
  getNextRelationshipIndex,
  replaceHyperlinkPlaceholders,
  strFromU8,
  toJson,
  unzipSync,
  zipAndConvert,
} from "@office-open/core";
import type {
  BasePatchOptions,
  CorePropertiesOptions,
  OutputByType,
  OutputType,
} from "@office-open/core";
import { toUint8Array } from "@office-open/core";
import { js2xml, xml2js } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import { textRunDesc } from "@parts/descriptors/text";
import type { RunOptions } from "@shared/shape/paragraph/run";

import { PptxWriteContext } from "./context";

/** Reusable TextEncoder (stateless, safe to share). */
const encoder = new TextEncoder();

/**
 * Parts scanned for text replacement: slides, slide masters, slide layouts,
 * and notes slides. Replacing text in a master/layout propagates to every
 * dependent slide — intended for template branding.
 */
const TARGET_RE =
  /^ppt\/(?:slides\/slide\d+|slideMasters\/slideMaster\d+|slideLayouts\/slideLayout\d+|notesSlides\/notesSlide\d+)\.xml$/;
const HYPERLINK_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

/**
 * Inline run-level patch content. Reuses the generate vocabulary: a
 * {@link RunOptions} (or an array of them), or a plain string shorthand
 * for `{ text: "…" }`.
 */
export type Patch = RunOptions | RunOptions[] | string;

export interface PatchPresentationOptions<
  T extends OutputType = OutputType,
> extends BasePatchOptions<T> {
  /** Placeholder substitutions: `{{key}}` (per delimiters) → run content. */
  placeholders?: Readonly<Record<string, Patch>>;
  /** Literal find/replace: the find string → run content (no delimiters added). */
  findReplace?: Readonly<Record<string, Patch>>;
  /** Core-properties metadata override (merged over the existing docProps/core.xml). */
  coreProperties?: Partial<CorePropertiesOptions>;
  keepOriginalStyles?: boolean;
}

/** Write context for the current patch operation — accumulates hyperlinks. */
let currentPatchCtx: PptxWriteContext;

/**
 * Patch replacer. Serializes each run via {@link textRunDesc} so the full
 * RunProperties vocabulary (color, fill, outline, shadow, font, hyperlinks, …)
 * is supported — no hand-rolled `<a:r>` builder.
 */
const pptxReplacer = createReplacer({
  ns: PPTX_NS,
  formatChild: (child: unknown): Element[] => {
    const runOpts = (typeof child === "string" ? { text: child } : child) as RunOptions;
    const xmlStr = textRunDesc.stringify(runOpts, currentPatchCtx) ?? "<a:r/>";
    return [xml2js(xmlStr, { captureSpacesBetweenElements: true }).elements![0]];
  },
  preserveSpace: false,
});

/**
 * Normalize a user patch value into the replacer's `{ type, children }`
 * envelope. PPTX patches are always inline run-level (`type: "paragraph"`).
 */
const toReplacerPatch = (patch: Patch): { type: "paragraph"; children: RunOptions[] } =>
  typeof patch === "string"
    ? { type: "paragraph", children: [{ text: patch }] }
    : { type: "paragraph", children: Array.isArray(patch) ? patch : [patch] };

const createRelationshipFile = (): Element => ({
  declaration: {
    attributes: { encoding: "UTF-8", standalone: "yes", version: "1.0" },
  },
  elements: [
    {
      attributes: { xmlns: "http://schemas.openxmlformats.org/package/2006/relationships" },
      elements: [],
      name: "Relationships",
      type: "element",
    },
  ],
});

/**
 * Patches an existing .pptx presentation by replacing placeholders with run content.
 *
 * Patch content reuses the generate {@link RunOptions} vocabulary (serialized
 * via the same descriptor), so color, fonts, fills, and hyperlinks are all
 * supported. Hyperlinks introduced by patch runs are registered into each
 * slide's relationship part.
 *
 * @publicApi
 */
export const patchPresentation = async <T extends OutputType = OutputType>({
  outputType,
  data,
  placeholders,
  findReplace,
  coreProperties,
  keepOriginalStyles = true,
  placeholderDelimiters = { end: "}}", start: "{{" } as const,
}: PatchPresentationOptions<T>): Promise<OutputByType[T]> => {
  const { start, end } = placeholderDelimiters;
  if (!start.trim() || !end.trim()) {
    throw new Error("Both start and end delimiters must be non-empty strings.");
  }

  const zipContent = unzipSync(toUint8Array(data));

  const xmlMap = new Map<string, Element>();
  const binaryMap = new Map<string, Uint8Array>();

  // Separate XML files from binary files
  for (const [key, value] of Object.entries(zipContent)) {
    if (key.endsWith(".xml") || key.endsWith(".rels")) {
      xmlMap.set(key, toJson(strFromU8(value)));
    } else {
      binaryMap.set(key, value);
    }
  }

  // Build (find-text → patch envelope) entries. Placeholders wrap the key in
  // delimiters; findReplace uses the literal key. Both share the same engine.
  const entries: Array<{
    find: string;
    patch: { type: "paragraph"; children: RunOptions[] };
  }> = [];
  if (placeholders) {
    for (const [key, value] of Object.entries(placeholders)) {
      entries.push({ find: `${start}${key}${end}`, patch: toReplacerPatch(value) });
    }
  }
  if (findReplace) {
    for (const [key, value] of Object.entries(findReplace)) {
      entries.push({ find: key, patch: toReplacerPatch(value) });
    }
  }

  // Target parts: slides + masters + layouts + notes (deterministic order)
  const targetPaths = Object.keys(zipContent)
    .filter((k) => TARGET_RE.test(k))
    .sort();

  currentPatchCtx = new PptxWriteContext();
  const context = {};

  // Process text replacement on each target part
  for (const targetPath of targetPaths) {
    const json = xmlMap.get(targetPath);
    if (!json) continue;

    for (const { find, patch } of entries) {
      pptxReplacer({
        context,
        json,
        keepOriginalStyles,
        patch,
        patchText: find,
      });
    }
  }

  // Resolve hyperlink placeholders registered by patch runs into each part's rels
  const hlinkByKey = new Map(currentPatchCtx.hyperlinks.map((h) => [h.key, h]));
  for (const targetPath of targetPaths) {
    const targetEl = xmlMap.get(targetPath);
    if (!targetEl) continue;

    const targetXml = js2xml(targetEl);
    const hlinkKeys = collectPlaceholderKeys(targetXml, "hlink:");
    if (hlinkKeys.length === 0) continue;

    const targetHlinks = hlinkKeys
      .map((k) => hlinkByKey.get(k))
      .filter((h): h is NonNullable<typeof h> => h !== undefined);
    if (targetHlinks.length === 0) continue;

    // Generalized rels path: ppt/<dir>/_rels/<file>.rels
    const lastSlash = targetPath.lastIndexOf("/");
    const relsKey = `${targetPath.substring(0, lastSlash)}/_rels/${targetPath.substring(lastSlash + 1)}.rels`;
    const relsJson = xmlMap.get(relsKey) ?? createRelationshipFile();
    xmlMap.set(relsKey, relsJson);
    const offset = getNextRelationshipIndex(relsJson);

    const resolved = replaceHyperlinkPlaceholders(targetXml, targetHlinks, offset);
    xmlMap.set(targetPath, toJson(resolved));

    for (let i = 0; i < targetHlinks.length; i++) {
      appendRelationship(relsJson, offset + i, HYPERLINK_REL_TYPE, targetHlinks[i].url, "External");
    }
  }

  // Rebuild ZIP
  const files: Record<string, Uint8Array> = {};
  const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

  for (const [key, value] of xmlMap) {
    files[key] =
      key === "docProps/core.xml" && coreProperties
        ? encoder.encode(XML_DECL + applyCorePropertiesOverride(value, coreProperties))
        : encoder.encode(js2xml(value));
  }

  for (const [key, value] of binaryMap) {
    files[key] = value;
  }

  return await zipAndConvert(files, outputType, OoxmlMimeType.PPTX);
};
