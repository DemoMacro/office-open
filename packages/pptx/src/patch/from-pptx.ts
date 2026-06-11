import {
  PPTX_NS,
  OoxmlMimeType,
  appendContentType,
  createReplacer,
  toJson,
  strFromU8,
  unzipSync,
  zipAndConvert,
} from "@office-open/core";
import type { OutputByType, OutputType } from "@office-open/core";
import { toUint8Array } from "@office-open/core";
import { escapeXml, js2xml, xml2js } from "@office-open/xml";
import type { Element } from "@office-open/xml";

/** Reusable TextEncoder (stateless, safe to share). */
const encoder = new TextEncoder();

export type InputDataType = Buffer | string | number[] | Uint8Array | ArrayBuffer | Blob;

export const PatchType = {
  PARAGRAPH: "paragraph",
} as const;

export type PatchDocumentOutputType = OutputType;

interface ParagraphPatch {
  type: typeof PatchType.PARAGRAPH;
  children: unknown[];
}

export type IPatch = ParagraphPatch;

export interface PatchPresentationOptions<
  T extends PatchDocumentOutputType = PatchDocumentOutputType,
> {
  outputType: T;
  data: InputDataType;
  patches: Readonly<Record<string, IPatch>>;
  keepOriginalStyles?: boolean;
  placeholderDelimiters?: Readonly<{
    start: string;
    end: string;
  }>;
}

const pptxReplacer = createReplacer({
  ns: PPTX_NS,
  formatChild: (child: unknown): Element[] => {
    let xmlStr: string;
    if (typeof child === "string") {
      xmlStr = `<a:r><a:t>${escapeXml(child)}</a:t></a:r>`;
    } else if (typeof child === "object" && child !== null) {
      const obj = child as Record<string, unknown>;
      if ("text" in obj) {
        // Build a:r with optional run properties.
        // Per XSD CT_TextCharacterProperties: b/i/sz/u are attributes, latin is a child element.
        const attrs: string[] = [];
        if (obj.bold) attrs.push(' b="1"');
        if (obj.italic) attrs.push(' i="1"');
        if (obj.underline) attrs.push(` u="${escapeXml(String(obj.underline as string))}"`);
        if (obj.fontSize) attrs.push(` sz="${Number(obj.fontSize) * 100}"`);
        const attrStr = attrs.join("");
        const latinChild = obj.font ? `<a:latin typeface="${escapeXml(obj.font as string)}"/>` : "";
        const rPr = attrStr || latinChild ? `<a:rPr${attrStr}>${latinChild}</a:rPr>` : "";
        xmlStr = `<a:r>${rPr}<a:t>${escapeXml(String(obj.text))}</a:t></a:r>`;
      } else {
        // XSD requires a:r to have at least rPr or t child
        xmlStr = "<a:r><a:rPr/></a:r>";
      }
    } else {
      xmlStr = "<a:r><a:rPr/></a:r>";
    }
    const jsonObj = xml2js(xmlStr, { captureSpacesBetweenElements: true });
    return [jsonObj.elements![0]];
  },
  preserveSpace: false,
});

const IMAGE_CONTENT_TYPES: ReadonlyArray<[string, string]> = [
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

  const context = {};

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
    files[key] = encoder.encode(js2xml(value));
  }

  for (const [key, value] of binaryMap) {
    files[key] = value;
  }

  return await zipAndConvert(files, outputType, OoxmlMimeType.PPTX);
};
