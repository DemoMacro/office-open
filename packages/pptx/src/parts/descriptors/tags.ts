/**
 * User-defined tags part (p:tagLst) for PPTX.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, escapeXml } from "@office-open/xml";
import type { StringTagOptions, TagPartOptions } from "@parts/presentation";

import { PptxParseError } from "../../pptx-parse-error";

export const tagListDesc: CustomDescriptor<StringTagOptions[]> = {
  kind: "custom",

  stringify(tags) {
    const children = tags
      .map((tag) => `<p:tag name="${escapeXml(tag.name)}" val="${escapeXml(tag.val)}"/>`)
      .join("");
    return `<p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">${children}</p:tagLst>`;
  },

  parse(el) {
    return (el.elements ?? []).reduce<StringTagOptions[]>((tags, child) => {
      if (child.name !== "p:tag") {
        throw new PptxParseError(
          `Unsupported tags child "${child.name}"`,
          "p:tagLst",
          `p:tagLst/${child.name}`,
          "unsupported-element",
        );
      }
      const name = attr(child, "name");
      const val = attr(child, "val");
      if (name === undefined || val === undefined) {
        throw new PptxParseError(
          "Tag requires name and val",
          "p:tagLst",
          "p:tagLst/p:tag",
          "missing-required-attribute",
        );
      }
      tags.push({ name, val });
      return tags;
    }, []);
  },
};

export const tagPartDesc: CustomDescriptor<TagPartOptions> = {
  kind: "custom",

  stringify(part, context) {
    return tagListDesc.stringify(part.tags, context);
  },

  parse(el, context) {
    return { tags: tagListDesc.parse(el, context) };
  },
};
