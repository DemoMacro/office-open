/**
 * Microsoft classification-label metadata part types and descriptor.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, attrBool, children, escapeXml, stringifyElement } from "@office-open/xml";

const CLBL_NS = "http://schemas.microsoft.com/office/2020/mipLabelMetadata";
const CLASSIFICATION_LABEL_ATTRIBUTES = [
  "id",
  "enabled",
  "setDate",
  "method",
  "name",
  "siteId",
  "actionId",
  "contentBits",
  "removed",
] as const;

/** Classification label (CT_ClassificationLabel). */
export interface ClassificationLabelOptions {
  /** Label identifier (required). */
  id: string;
  /** True when the label is enabled (required). */
  enabled: boolean;
  /** Label assignment date (optional). */
  setDate?: string;
  /** Assignment method (required). */
  method: string;
  /** Display name (optional). */
  name?: string;
  /** Label site identifier as a braced GUID (required). */
  siteId: string;
  /** Action identifier (optional). */
  actionId?: string;
  /** Content-protection bit mask (optional, unsigned 32-bit integer). */
  contentBits?: number;
  /** True when the label was removed (required). */
  removed: boolean;
}

/** Extension element with an `xs:any` payload (CT_ClassificationExtension). */
export interface ClassificationExtensionOptions {
  /** Extension namespace URI (required). */
  uri: string;
  /** Foreign-namespace child XML declared by the extension registry. */
  innerXml?: string;
}

/** Options for docMetadata/LabelInfo.xml (CT_ClassificationLabelList). */
export interface ClassificationLabelListOptions {
  labels?: ClassificationLabelOptions[];
  extensions?: ClassificationExtensionOptions[];
}

/** Structured parse failure with the exact metadata-part location. */
export class ClassificationLabelsParseError extends Error {
  constructor(
    readonly part: string,
    readonly path: string,
    readonly name: string,
    readonly reason: string,
  ) {
    super(`${part}${path}: ${name}: ${reason}`);
  }
}

export const classificationLabelsDesc: CustomDescriptor<ClassificationLabelListOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const childrenXml = [
      ...(opts.labels ?? []).map((label) => {
        const attrs = [
          ` id="${escapeXml(label.id)}"`,
          ` enabled="${label.enabled ? 1 : 0}"`,
          ...(label.setDate !== undefined ? [` setDate="${escapeXml(label.setDate)}"`] : []),
          ` method="${escapeXml(label.method)}"`,
          ...(label.name !== undefined ? [` name="${escapeXml(label.name)}"`] : []),
          ` siteId="${escapeXml(label.siteId)}"`,
          ...(label.actionId !== undefined ? [` actionId="${escapeXml(label.actionId)}"`] : []),
          ...(label.contentBits !== undefined ? [` contentBits="${label.contentBits}"`] : []),
          ` removed="${label.removed ? 1 : 0}"`,
        ].join("");
        return `<clbl:label${attrs}/>`;
      }),
      ...(opts.extensions !== undefined
        ? [
            `<clbl:extLst>${(opts.extensions ?? [])
              .map(
                (extension) =>
                  `<clbl:ext uri="${escapeXml(extension.uri)}">${extension.innerXml ?? ""}</clbl:ext>`,
              )
              .join("")}</clbl:extLst>`,
          ]
        : []),
    ].join("");
    return `<clbl:labelList xmlns:clbl="${CLBL_NS}">${childrenXml}</clbl:labelList>`;
  },

  parse(el, _ctx) {
    const part = "docMetadata";
    for (const child of el.elements ?? []) {
      if (child.name !== "clbl:label" && child.name !== "clbl:extLst") {
        throw new ClassificationLabelsParseError(
          part,
          "/labelList",
          child.name ?? "unknown",
          "unexpected child of labelList",
        );
      }
    }
    const labels = children(el, "clbl:label").map((element) => {
      const path = "/labelList/label";
      const require = (name: (typeof CLASSIFICATION_LABEL_ATTRIBUTES)[number]): string => {
        const value = attr(element, name);
        if (value === undefined) {
          throw new ClassificationLabelsParseError(
            part,
            path,
            `@${name}`,
            "missing required value",
          );
        }
        return value;
      };
      const enabled = attrBool(element, "enabled");
      const removed = attrBool(element, "removed");
      if (enabled === undefined || removed === undefined) {
        throw new ClassificationLabelsParseError(
          part,
          path,
          "label",
          "enabled and removed must be boolean",
        );
      }
      for (const attribute of Object.keys(element.attributes ?? {})) {
        if (attribute === "xmlns:clbl" || attribute === "xmlns") continue;
        if (!CLASSIFICATION_LABEL_ATTRIBUTES.includes(attribute as never)) {
          throw new ClassificationLabelsParseError(
            part,
            path,
            `@${attribute}`,
            "unexpected attribute",
          );
        }
      }
      const contentBits = attr(element, "contentBits");
      return {
        id: require("id"),
        enabled,
        ...(attr(element, "setDate") !== undefined ? { setDate: attr(element, "setDate") } : {}),
        method: require("method"),
        ...(attr(element, "name") !== undefined ? { name: attr(element, "name") } : {}),
        siteId: require("siteId"),
        ...(attr(element, "actionId") !== undefined ? { actionId: attr(element, "actionId") } : {}),
        ...(contentBits !== undefined ? { contentBits: Number(contentBits) } : {}),
        removed,
      };
    });
    const extensionList = children(el, "clbl:extLst");
    if (extensionList.length > 1) {
      throw new ClassificationLabelsParseError(
        part,
        "/labelList/extLst",
        "extLst",
        "expected at most one extension list",
      );
    }
    const extensions = extensionList.flatMap((list) =>
      (list.elements ?? []).map((element) => {
        if (element.name !== "clbl:ext") {
          throw new ClassificationLabelsParseError(
            part,
            "/labelList/extLst",
            element.name ?? "unknown",
            "unexpected child of extLst",
          );
        }
        const uri = attr(element, "uri");
        if (uri === undefined) {
          throw new ClassificationLabelsParseError(
            part,
            "/labelList/extLst/ext",
            "@uri",
            "missing required value",
          );
        }
        for (const attribute of Object.keys(element.attributes ?? {})) {
          if (attribute === "xmlns:clbl" || attribute === "xmlns") continue;
          if (attribute !== "uri") {
            throw new ClassificationLabelsParseError(
              part,
              "/labelList/extLst/ext",
              `@${attribute}`,
              "unexpected attribute",
            );
          }
        }
        return {
          uri,
          innerXml: (element.elements ?? []).map(stringifyElement).join(""),
        };
      }),
    );
    const result: ClassificationLabelListOptions = {};
    result.labels = labels;
    if (extensionList.length > 0) result.extensions = extensions;
    return result;
  },
};
