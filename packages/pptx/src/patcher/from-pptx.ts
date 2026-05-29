import type { Context, XmlComponent } from "@file/xml-components";
import {
  PPTX_NS,
  Formatter,
  OoxmlMimeType,
  appendContentType,
  createReplacer,
  toJson,
  strFromU8,
  unzipSync,
  zipAndConvert,
} from "@office-open/core";
import type { OutputByType, OutputType } from "@office-open/core";
import { js2xml, xml } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import { textToUint8Array, toUint8Array } from "undio";

export type InputDataType = Buffer | string | number[] | Uint8Array | ArrayBuffer | Blob;

export const PatchType = {
  PARAGRAPH: "paragraph",
} as const;

export type PatchDocumentOutputType = OutputType;

interface ParagraphPatch {
  readonly type: typeof PatchType.PARAGRAPH;
  readonly children: readonly unknown[];
}

export type IPatch = ParagraphPatch;

export interface PatchPresentationOptions<
  T extends PatchDocumentOutputType = PatchDocumentOutputType,
> {
  readonly outputType: T;
  readonly data: InputDataType;
  readonly patches: Readonly<Record<string, IPatch>>;
  readonly keepOriginalStyles?: boolean;
  readonly placeholderDelimiters?: Readonly<{
    readonly start: string;
    readonly end: string;
  }>;
}

const formatter = new Formatter();

const pptxReplacer = createReplacer({
  ns: PPTX_NS,
  formatChild: (child: unknown, context: unknown): Element[] => {
    const jsonObj = toJson(xml(formatter.format(child as XmlComponent, context as Context)));
    return [jsonObj.elements![0]];
  },
  preserveSpace: false,
});

const IMAGE_CONTENT_TYPES: ReadonlyArray<readonly [string, string]> = [
  ["image/png", "png"],
  ["image/jpeg", "jpeg"],
  ["image/jpeg", "jpg"],
  ["image/bmp", "bmp"],
  ["image/gif", "gif"],
  ["image/svg+xml", "svg"],
  ["image/tiff", "tif"],
  ["image/tiff", "tiff"],
  ["image/x-emf", "emf"],
  ["image/x-wmf", "wmf"],
];

const SLIDE_RE = /^ppt\/slides\/slide(\d+)\.xml$/;

/**
 * Patches an existing .pptx presentation by replacing placeholders with new content.
 *
 * @publicApi
 */
export const patchPresentation = async <
  T extends PatchDocumentOutputType = PatchDocumentOutputType,
>({
  outputType,
  data,
  patches,
  keepOriginalStyles = true,
  placeholderDelimiters = { end: "}}", start: "{{" } as const,
}: PatchPresentationOptions<T>): Promise<OutputByType[T]> => {
  const zipContent = unzipSync(toUint8Array(data));

  const xmlMap = new Map<string, Element>();
  const binaryMap = new Map<string, Uint8Array>();
  let hasMedia = false;

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

  const { start, end } = placeholderDelimiters;
  if (!start.trim() || !end.trim()) {
    throw new Error("Both start and end delimiters must be non-empty strings.");
  }

  const context: Context = { stack: [] };

  // Process text replacement on each slide
  for (const slidePath of slidePaths) {
    const json = xmlMap.get(slidePath);
    if (!json) continue;

    for (const [patchKey, patchValue] of Object.entries(patches)) {
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

  // Handle image replacements in ppt/media/
  for (const [key, patch] of Object.entries(patches)) {
    if ("data" in patch && patch.data instanceof Uint8Array) {
      const mediaPath = `ppt/media/${key}`;
      if (binaryMap.has(mediaPath)) {
        binaryMap.set(mediaPath, patch.data);
        hasMedia = true;
      }
    }
  }

  // Add image content types if media was touched
  if (hasMedia) {
    const contentTypes = xmlMap.get("[Content_Types].xml");
    if (contentTypes) {
      for (const [contentType, extension] of IMAGE_CONTENT_TYPES) {
        appendContentType(contentTypes, contentType, extension);
      }
    }
  }

  // Rebuild ZIP
  const files: Record<string, Uint8Array> = {};

  for (const [key, value] of xmlMap) {
    files[key] = textToUint8Array(js2xml(value));
  }

  for (const [key, value] of binaryMap) {
    files[key] = value;
  }

  return await zipAndConvert(files, outputType, OoxmlMimeType.PPTX);
};
