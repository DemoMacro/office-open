import { escapeRegex, formatId } from "@office-open/core";

export function replaceHyperlinkPlaceholders(
  xml: string,
  hyperlinks: readonly { readonly key: string }[],
  offset: number,
): string {
  let result = xml;
  hyperlinks.forEach((h, i) => {
    result = result.replace(
      new RegExp(`\\{hlink:${escapeRegex(h.key)}\\}`, "g"),
      formatId(offset, i, "rId"),
    );
  });
  return result;
}
