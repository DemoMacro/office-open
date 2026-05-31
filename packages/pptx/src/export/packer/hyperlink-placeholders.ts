import { formatId } from "@office-open/core";

export function replaceHyperlinkPlaceholders(
  xml: string,
  hyperlinks: readonly { readonly key: string }[],
  offset: number,
): string {
  let result = xml;
  for (let i = 0; i < hyperlinks.length; i++) {
    result = result.replaceAll(`{hlink:${hyperlinks[i].key}}`, formatId(offset, i, "rId"));
  }
  return result;
}
