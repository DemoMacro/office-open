import { TargetModeType } from "@office-open/core";
import {
  DOCX_NS,
  OoxmlMimeType,
  appendContentType,
  appendRelationship,
  applyCorePropertiesOverride,
  createReplacer,
  getNextRelationshipIndex,
  getReferencedMedia,
  replaceImagePlaceholders,
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
import { escapeXml, js2xml, xml2js } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import { DocumentAttributeNamespaces } from "@parts/document";
import {
  stringifyChildDispatch,
  stringifyParagraphInline,
  stringifyRunInline,
} from "@parts/inline";
import type { ParagraphChild } from "@parts/paragraph/paragraph";
import type { ParagraphOptions } from "@parts/paragraph/paragraph";
import type { RunOptions } from "@parts/paragraph/run/run";
import { tableDesc } from "@parts/table/descriptor";
import type { TableOptions } from "@parts/table/table";
import { Media } from "@shared/media";
import type { SectionChild } from "@shared/section";

import type { BodyContext } from "../context";
import type { ViewWrapper } from "../context";

/** Reusable TextEncoder (stateless, safe to share). */
const encoder = new TextEncoder();

/**
 * Document patching module for modifying existing .docx files.
 *
 * Uses compile-path stringifiers (zero class instantiation) to serialize
 * patch content — no Formatter, no XmlComponent instances.
 *
 * @module
 */

// ── Patch content stringification ──

/**
 * Lightweight BodyContext adapter for patch serialization.
 * Captures hyperlink and image relationships for post-processing.
 */
function createPatchContext(
  file: { media: Media },
  hyperlinkSink: Array<{ id: string; link: string }>,
): BodyContext {
  return {
    fileData: file as unknown as BodyContext["fileData"],
    file: file as unknown as BodyContext["file"],
    viewWrapper: {
      relationships: {
        addRelationship: (linkId: string, _type: string, target: string, _mode?: string) => {
          hyperlinkSink.push({ id: linkId, link: target });
        },
        relationshipCount: 0,
      },
    } as unknown as BodyContext["viewWrapper"],
    stringifyChild: () => "",
    addRelationship: () => "",
    addMedia: () => "",
  };
}

/**
 * Serialize a patch child (SectionChild / ParagraphChild / string) into XML
 * elements via the compile-path stringifiers. Shared by the placeholder
 * replacer and body-level `append`. Relies on the module-level
 * {@link currentPatchCtx} for relationship/media sinks.
 */
const formatChildElement = (child: unknown): Element[] => {
  let xmlStr: string;

  if (typeof child === "string") {
    // Plain string → simple run
    xmlStr = `<w:r><w:t xml:space="preserve">${escapeXml(child)}</w:t></w:r>`;
  } else if (typeof child === "object" && child !== null) {
    const obj = child as Record<string, unknown>;
    // SectionChild level (paragraph / table) — for DOCUMENT patches and append
    if ("paragraph" in obj) {
      xmlStr = stringifyParagraphInline(obj.paragraph as ParagraphOptions, currentPatchCtx);
    } else if ("table" in obj) {
      xmlStr = tableDesc.stringify(obj.table as TableOptions, currentPatchCtx) ?? "";
    } else {
      // ParagraphChild level — for PARAGRAPH patches
      // Try compile-path JSON child dispatch first
      const jr = stringifyChildDispatch(child as ParagraphChild, currentPatchCtx);
      if (jr !== undefined) {
        xmlStr = Array.isArray(jr) ? jr.join("") : jr;
      } else {
        // RunOptions (plain objects with text/children/bold/etc.)
        xmlStr = stringifyRunInline(child as RunOptions, currentPatchCtx);
      }
    }
  } else {
    xmlStr = "<w:r/>";
  }

  const jsonObj = xml2js(xmlStr, { captureSpacesBetweenElements: true });
  return [jsonObj.elements![0]];
};

const docxReplacer = createReplacer({
  ns: DOCX_NS,
  formatChild: formatChildElement,
});

/**
 * Splice block-level children into `<w:body>` before the trailing `<w:sectPr>`
 * (the final section properties). If the body has no trailing sectPr, append at
 * the end. Reuses {@link formatChildElement} so the same serialization and
 * relationship/media sinks apply as placeholder patches.
 */
const appendToBody = (root: Element, children: SectionChild[]): void => {
  const docEl = root.elements?.find((e) => e.name === "w:document");
  const body = docEl?.elements?.find((e) => e.name === "w:body");
  if (!body) return;
  const els = body.elements ?? (body.elements = []);
  const newEls: Element[] = [];
  for (const child of children) {
    newEls.push(...formatChildElement(child));
  }
  // Insert before the trailing <w:sectPr>; otherwise at the end.
  let insertAt = els.length;
  for (let i = els.length - 1; i >= 0; i--) {
    if (els[i].name === "w:sectPr") {
      insertAt = i;
      break;
    }
  }
  els.splice(insertAt, 0, ...newEls);
};

/** Current patch context — set per file in the main loop. */
let currentPatchCtx: BodyContext;

/**
 * A patch operation: replace a placeholder with either inline run-level
 * content (`type: "paragraph"`) or block-level content (`type: "document"`).
 *
 * The `type` is a bare string literal (no PatchType enum) — the core
 * replacer already discriminates on `patch.type`.
 */
export type Patch =
  | { type: "paragraph"; children: (string | RunOptions | ParagraphChild)[] }
  | { type: "document"; children: SectionChild[] };

interface ImageRelationshipAddition {
  key: string;
  mediaDatas: { fileName: string }[];
}

interface HyperlinkRelationshipAddition {
  key: string;
  hyperlink: { id: string; link: string };
}

export interface PatchDocumentOptions<
  T extends OutputType = OutputType,
> extends BasePatchOptions<T> {
  /** Placeholder substitutions: `{{key}}` (per delimiters) → patch content. */
  placeholders?: Readonly<Record<string, Patch>>;
  /** Literal find/replace: the find string → patch content (no delimiters added). */
  findReplace?: Readonly<Record<string, Patch>>;
  /** Core-properties metadata override (merged over the existing docProps/core.xml). */
  coreProperties?: Partial<CorePropertiesOptions>;
  keepOriginalStyles?: boolean;
  recursive?: boolean;
  /** Block-level content appended to the document body, before the final section break. */
  append?: SectionChild[];
}

const UTF16LE = new Uint8Array([0xff, 0xfe]);
const UTF16BE = new Uint8Array([0xfe, 0xff]);

const compareByteArrays = (a: Uint8Array, b: Uint8Array): boolean => {
  if (a.length !== b.length) {
    return false;
  }
  for (let i = 0; i < a.length; i++) {
    if (a[i] !== b[i]) {
      return false;
    }
  }
  return true;
};

/**
 * Patches an existing .docx document by replacing placeholders with new content.
 *
 * @publicApi
 */
export const patchDocument = async <T extends OutputType = OutputType>({
  outputType,
  data,
  placeholders,
  findReplace,
  coreProperties,
  append,
  keepOriginalStyles = true,
  placeholderDelimiters = { end: "}}", start: "{{" } as const,
  recursive = true,
}: PatchDocumentOptions<T>): Promise<OutputByType[T]> => {
  const zipContent = unzipSync(toUint8Array(data));
  const contexts = new Map<string, BodyContext>();
  const media = new Media();
  const file = { media } as BodyContext["file"];

  const map = new Map<string, Element>();

  const imageRelationshipAdditions: ImageRelationshipAddition[] = [];
  const hyperlinkRelationshipAdditions: HyperlinkRelationshipAddition[] = [];
  let hasMedia = false;

  const binaryContentMap = new Map<string, Uint8Array>();

  for (const [key, value] of Object.entries(zipContent)) {
    const startBytes = value.slice(0, 2);
    if (compareByteArrays(startBytes, UTF16LE) || compareByteArrays(startBytes, UTF16BE)) {
      binaryContentMap.set(key, value);
      continue;
    }

    if (!key.endsWith(".xml") && !key.endsWith(".rels")) {
      binaryContentMap.set(key, value);
      continue;
    }

    const json = toJson(strFromU8(value));

    if (key === "word/document.xml") {
      const document = json.elements?.find((i) => i.name === "w:document");
      if (document && document.attributes) {
        for (const ns of ["mc", "wp", "r", "w15", "m"] as const) {
          document.attributes[`xmlns:${ns}`] = DocumentAttributeNamespaces[ns];
        }
        document.attributes["mc:Ignorable"] =
          `${document.attributes["mc:Ignorable"] || ""} w15`.trim();
      }
    }

    if (key.startsWith("word/") && !key.endsWith(".xml.rels")) {
      const hyperlinkSink: Array<{ id: string; link: string }> = [];

      const context: BodyContext = {
        fileData: file,
        file,
        viewWrapper: {
          relationships: {
            addRelationship: (
              linkId: string,
              _: string,
              target: string,
              __: (typeof TargetModeType)[keyof typeof TargetModeType],
            ) => {
              hyperlinkRelationshipAdditions.push({
                hyperlink: {
                  id: linkId,
                  link: target,
                },
                key,
              });
            },
          },
        } as unknown as ViewWrapper,
        stringifyChild: () => "",
        addRelationship: () => "",
        addMedia: () => "",
      };
      contexts.set(key, context);

      if (!placeholderDelimiters?.start.trim() || !placeholderDelimiters?.end.trim()) {
        throw new Error("Both start and end delimiters must be non-empty strings.");
      }

      const { start, end } = placeholderDelimiters;

      // Create compile-path context for stringifying patch children
      const patchCtx = createPatchContext(file, hyperlinkSink);
      currentPatchCtx = patchCtx;

      // Build (find-text → patch) entries. Placeholders wrap the key in
      // delimiters; findReplace uses the literal key. Both share one engine.
      const entries: Array<{ find: string; patch: Patch }> = [];
      if (placeholders) {
        for (const [key, value] of Object.entries(placeholders)) {
          entries.push({ find: `${start}${key}${end}`, patch: value });
        }
      }
      if (findReplace) {
        for (const [key, value] of Object.entries(findReplace)) {
          entries.push({ find: key, patch: value });
        }
      }

      for (const { find: patchText, patch: patchValue } of entries) {
        while (true) {
          const { didFindOccurrence } = docxReplacer({
            context,
            json,
            keepOriginalStyles,
            patch: patchValue,
            patchText,
          });
          if (!recursive || !didFindOccurrence) {
            break;
          }
        }
      }

      // Append block-level children to the document body before the final sectPr.
      if (append && append.length > 0 && key === "word/document.xml") {
        appendToBody(json, append);
      }

      // Flush hyperlink relationships captured by the compile-path context
      for (const hl of hyperlinkSink) {
        hyperlinkRelationshipAdditions.push({
          hyperlink: { id: hl.id, link: hl.link },
          key,
        });
      }

      const mediaDatas = getReferencedMedia(JSON.stringify(json), media.array);
      if (mediaDatas.length > 0) {
        hasMedia = true;
        imageRelationshipAdditions.push({
          key,
          mediaDatas,
        });
      }
    }

    map.set(key, json);
  }

  for (const { key, mediaDatas } of imageRelationshipAdditions) {
    const relationshipKey = `word/_rels/${key.split("/").pop()}.rels`;
    const relationshipsJson = map.get(relationshipKey) ?? createRelationshipFile();
    map.set(relationshipKey, relationshipsJson);

    const index = getNextRelationshipIndex(relationshipsJson);
    const newJson = replaceImagePlaceholders(
      JSON.stringify(map.get(key)),
      mediaDatas,
      index,
      "plain",
    );
    map.set(key, JSON.parse(newJson) as Element);

    for (let i = 0; i < mediaDatas.length; i++) {
      const { fileName } = mediaDatas[i];
      appendRelationship(
        relationshipsJson,
        index + i,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        `media/${fileName}`,
      );
    }
  }

  for (const { key, hyperlink } of hyperlinkRelationshipAdditions) {
    const relationshipKey = `word/_rels/${key.split("/").pop()}.rels`;

    const relationshipsJson = map.get(relationshipKey) ?? createRelationshipFile();
    map.set(relationshipKey, relationshipsJson);

    appendRelationship(
      relationshipsJson,
      hyperlink.id,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
      hyperlink.link,
      TargetModeType.EXTERNAL,
    );
  }

  if (hasMedia) {
    const contentTypesJson = map.get("[Content_Types].xml");

    if (!contentTypesJson) {
      throw new Error("Could not find content types file");
    }

    appendContentType(contentTypesJson, "image/png", "png");
    appendContentType(contentTypesJson, "image/jpeg", "jpeg");
    appendContentType(contentTypesJson, "image/jpeg", "jpg");
    appendContentType(contentTypesJson, "image/bmp", "bmp");
    appendContentType(contentTypesJson, "image/gif", "gif");
    appendContentType(contentTypesJson, "image/svg+xml", "svg");
  }

  const files: Record<string, Uint8Array> = {};
  const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

  for (const [key, value] of map) {
    files[key] =
      key === "docProps/core.xml" && coreProperties
        ? encoder.encode(XML_DECL + applyCorePropertiesOverride(value, coreProperties))
        : encoder.encode(js2xml(value));
  }

  for (const [key, value] of binaryContentMap) {
    files[key] = value;
  }

  for (const { data: mediaData, fileName } of media.array) {
    files[`word/media/${fileName}`] =
      mediaData instanceof Uint8Array ? mediaData : new Uint8Array(mediaData);
  }

  return await zipAndConvert(files, outputType, OoxmlMimeType.DOCX);
};

const createRelationshipFile = (): Element => ({
  declaration: {
    attributes: {
      encoding: "UTF-8",
      standalone: "yes",
      version: "1.0",
    },
  },
  elements: [
    {
      attributes: {
        xmlns: "http://schemas.openxmlformats.org/package/2006/relationships",
      },
      elements: [],
      name: "Relationships",
      type: "element",
    },
  ],
});
