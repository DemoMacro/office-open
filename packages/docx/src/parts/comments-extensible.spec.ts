import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { commentsExtensibleDesc, type CommentExtensibleOptions } from "./comments-extensible";

const COMMENTS_NS =
  'xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml"';

describe("commentsExtensibleDesc", () => {
  it("round-trips comment metadata and extension children", () => {
    const entries: CommentExtensibleOptions[] = [
      { durableId: "1EC7B1B1", dateUtc: "2026-01-01T00:00:00Z", intelligentPlaceholder: false },
      {
        durableId: "2B401A58",
        intelligentPlaceholder: true,
        ext: '<w16:ext w16:uri="example"><w16:value>example</w16:value></w16:ext>',
      },
    ];
    const xml = commentsExtensibleDesc.stringify(entries, {} as never)!;
    expect(xml).toContain('w16cex:dateUtc="2026-01-01T00:00:00Z"');
    expect(xml).toContain('w16cex:intelligentPlaceholder="0"');
    expect(xml).toContain(
      '<w16cex:extLst><w16:ext w16:uri="example"><w16:value>example</w16:value></w16:ext></w16cex:extLst>',
    );

    const parsed = commentsExtensibleDesc.parse(
      parseXml(
        xml.replace(
          /^<w16cex:commentsExtensible [^>]*>/,
          `<w16cex:commentsExtensible ${COMMENTS_NS}>`,
        ),
      ).elements![0]!,
      {} as never,
    );
    expect(parsed).toEqual(entries);
  });
});
