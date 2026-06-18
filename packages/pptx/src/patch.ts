import {
  PPTX_NS,
  OoxmlMimeType,
  appendRelationship,
  collectPlaceholderKeys,
  createReplacer,
  getNextRelationshipIndex,
  replaceHyperlinkPlaceholders,
  strFromU8,
  toJson,
  unzipSync,
  zipAndConvert,
} from "@office-open/core";
import type { BasePatchOptions, OutputByType, OutputType } from "@office-open/core";
import { toUint8Array } from "@office-open/core";
import { js2xml, xml2js } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import { textRunDesc } from "@parts/descriptors/text";
import type { RunOptions } from "@shared/shape/paragraph/run";

import { PptxWriteContext } from "./context";

/** Reusable TextEncoder (stateless, safe to share). */
const encoder = new TextEncoder();

const SLIDE_RE = /^ppt\/slides\/slide(\d+)\.xml$/;
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
  patches: Readonly<Record<string, Patch>>;
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
  patches,
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

  // Collect slide paths sorted by number
  const slidePaths = Object.keys(zipContent)
    .filter((k) => SLIDE_RE.test(k))
    .sort((a, b) => {
      const na = parseInt(a.match(SLIDE_RE)![1], 10);
      const nb = parseInt(b.match(SLIDE_RE)![1], 10);
      return na - nb;
    });

  // Normalize patches once into the replacer's envelope
  const normalizedPatches: Record<string, { type: "paragraph"; children: RunOptions[] }> = {};
  for (const [key, value] of Object.entries(patches)) {
    normalizedPatches[key] = toReplacerPatch(value);
  }

  currentPatchCtx = new PptxWriteContext();
  const context = {};

  // Process text replacement on each slide
  for (const slidePath of slidePaths) {
    const json = xmlMap.get(slidePath);
    if (!json) continue;

    for (const [patchKey, patchValue] of Object.entries(normalizedPatches)) {
      const patchText = `${start}${patchKey}${end}`;
      pptxReplacer({
        context,
        json,
        keepOriginalStyles,
        patch: patchValue,
        patchText,
      });
    }
  }

  // Resolve hyperlink placeholders registered by patch runs into each slide's rels
  const hlinkByKey = new Map(currentPatchCtx.hyperlinks.map((h) => [h.key, h]));
  for (const slidePath of slidePaths) {
    const slideEl = xmlMap.get(slidePath);
    if (!slideEl) continue;

    const slideXml = js2xml(slideEl);
    const hlinkKeys = collectPlaceholderKeys(slideXml, "hlink:");
    if (hlinkKeys.length === 0) continue;

    const slideHlinks = hlinkKeys
      .map((k) => hlinkByKey.get(k))
      .filter((h): h is NonNullable<typeof h> => h !== undefined);
    if (slideHlinks.length === 0) continue;

    const relsKey = `ppt/slides/_rels/${slidePath.split("/").pop()}.rels`;
    const relsJson = xmlMap.get(relsKey) ?? createRelationshipFile();
    xmlMap.set(relsKey, relsJson);
    const offset = getNextRelationshipIndex(relsJson);

    const resolved = replaceHyperlinkPlaceholders(slideXml, slideHlinks, offset);
    xmlMap.set(slidePath, toJson(resolved));

    for (let i = 0; i < slideHlinks.length; i++) {
      appendRelationship(relsJson, offset + i, HYPERLINK_REL_TYPE, slideHlinks[i].url, "External");
    }
  }

  // Rebuild ZIP
  const files: Record<string, Uint8Array> = {};

  for (const [key, value] of xmlMap) {
    files[key] = encoder.encode(js2xml(value));
  }

  for (const [key, value] of binaryMap) {
    files[key] = value;
  }

  return await zipAndConvert(files, outputType, OoxmlMimeType.PPTX);
};
