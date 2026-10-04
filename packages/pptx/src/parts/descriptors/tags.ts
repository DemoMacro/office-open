/**
 * User-defined tags part (p:tagLst) for PPTX.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, escapeXml } from "@office-open/xml";
import type { StringTagOptions } from "@parts/presentation";

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
      const name = attr(child, "name");
      const val = attr(child, "val");
      if (name !== undefined && val !== undefined) tags.push({ name, val });
      return tags;
    }, []);
  },
};
