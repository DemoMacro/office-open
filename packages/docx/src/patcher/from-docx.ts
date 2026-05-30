import { DocumentAttributeNamespaces } from "@file/document";
import type { ViewWrapper } from "@file/document-wrapper";
import type { File } from "@file/file";
import type { FileChild } from "@file/file-child";
import { Media } from "@file/media";
import { ConcreteHyperlink, ExternalHyperlink } from "@file/paragraph";
import type { ParagraphChild } from "@file/paragraph";
import { TargetModeType } from "@file/relationships/relationship/relationship";
import type { Context, XmlComponent } from "@file/xml-components";
import {
  DOCX_NS,
  Formatter,
  OoxmlMimeType,
  appendContentType,
  appendRelationship,
  createReplacer,
  getNextRelationshipIndex,
  getReferencedMedia,
  replaceImagePlaceholders,
  strFromU8,
  toJson,
  unzipSync,
  zipAndConvert,
} from "@office-open/core";
import type { OutputByType, OutputType } from "@office-open/core";
import { js2xml, xml2js } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import { uniqueId } from "@util/convenience-functions";
import { toUint8Array } from "undio";

/** Reusable TextEncoder (stateless, safe to share). */
const encoder = new TextEncoder();

/**
 * Document patching module for modifying existing .docx files.
 *
 * @module
 */

const formatter = new Formatter();

const docxReplacer = createReplacer({
  ns: DOCX_NS,
  formatChild: (child: unknown, context: unknown): Element[] => {
    const xmlStr = formatter.formatToXml(child as XmlComponent, context as Context);
    const jsonObj = xml2js(xmlStr, { captureSpacesBetweenElements: true });
    return [jsonObj.elements![0]];
  },
});

/**
 * Supported input data types for document patching.
 */
export type InputDataType = Buffer | string | number[] | Uint8Array | ArrayBuffer | Blob;

/**
 * Patch type enumeration.
 *
 * @publicApi
 */
export const PatchType = {
  DOCUMENT: "file",
  PARAGRAPH: "paragraph",
} as const;

interface ParagraphPatch {
  readonly type: typeof PatchType.PARAGRAPH;
  readonly children: readonly ParagraphChild[];
}

interface FilePatch {
  readonly type: typeof PatchType.DOCUMENT;
  readonly children: readonly FileChild[];
}

interface ImageRelationshipAddition {
  readonly key: string;
  readonly mediaDatas: readonly { readonly fileName: string }[];
}

interface HyperlinkRelationshipAddition {
  readonly key: string;
  readonly hyperlink: { readonly id: string; readonly link: string };
}

export type IPatch = ParagraphPatch | FilePatch;

export type PatchDocumentOutputType = OutputType;

export interface PatchDocumentOptions<T extends PatchDocumentOutputType = PatchDocumentOutputType> {
  readonly outputType: T;
  readonly data: InputDataType;
  readonly patches: Readonly<Record<string, IPatch>>;
  readonly keepOriginalStyles?: boolean;
  readonly placeholderDelimiters?: Readonly<{
    readonly start: string;
    readonly end: string;
  }>;
  readonly recursive?: boolean;
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
export const patchDocument = async <T extends PatchDocumentOutputType = PatchDocumentOutputType>({
  outputType,
  data,
  patches,
  keepOriginalStyles = true,
  placeholderDelimiters = { end: "}}", start: "{{" } as const,
  recursive = true,
}: PatchDocumentOptions<T>): Promise<OutputByType[T]> => {
  const zipContent = unzipSync(toUint8Array(data));
  const contexts = new Map<string, Context>();
  const file = {
    media: new Media(),
  } as unknown as File;

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
      const context: Context = {
        fileData: file,
        file,
        stack: [],
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
      };
      contexts.set(key, context);

      if (!placeholderDelimiters?.start.trim() || !placeholderDelimiters?.end.trim()) {
        throw new Error("Both start and end delimiters must be non-empty strings.");
      }

      const { start, end } = placeholderDelimiters;

      for (const [patchKey, patchValue] of Object.entries(patches)) {
        const patchText = `${start}${patchKey}${end}`;
        while (true) {
          const { didFindOccurrence } = docxReplacer({
            context,
            json,
            keepOriginalStyles,
            patch: {
              ...patchValue,
              children: patchValue.children.map((element) => {
                if (element instanceof ExternalHyperlink) {
                  const concreteHyperlink = new ConcreteHyperlink(
                    element.options.children,
                    uniqueId(),
                    { tooltip: element.options.tooltip as string | undefined },
                  );
                  hyperlinkRelationshipAdditions.push({
                    hyperlink: {
                      id: concreteHyperlink.linkId,
                      link: element.options.link,
                    },
                    key,
                  });
                  return concreteHyperlink;
                } else {
                  return element;
                }
              }),
            } as any,
            patchText,
          });
          if (!recursive || !didFindOccurrence) {
            break;
          }
        }
      }

      const mediaDatas = getReferencedMedia(JSON.stringify(json), context.file.media.array);
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

  for (const [key, value] of map) {
    files[key] = encoder.encode(js2xml(value));
  }

  for (const [key, value] of binaryContentMap) {
    files[key] = value;
  }

  for (const { data: mediaData, fileName } of file.media.array) {
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
