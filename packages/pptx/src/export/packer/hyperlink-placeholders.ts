import { formatId } from "@office-open/core";

export function replaceHyperlinkPlaceholders(
  xml: string,
  hyperlinks: readonly { readonly key: string }[],
  offset: number,
): string {
  let result = xml;
  hyperlinks.forEach((h, i) => {
    result = result.replaceAll(`{hlink:${h.key}}`, formatId(offset, i, "rId"));
  });
  return result;
}
