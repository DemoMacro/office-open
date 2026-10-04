import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { commentsIdsDesc, type CommentIdOptions } from "./comments-ids";

const COMMENTS_NS = 'xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"';

describe("commentsIdsDesc", () => {
  it("round-trips durable comment ids", () => {
    const entries: CommentIdOptions[] = [
      { paraId: "7FD6C115", durableId: "1EC7B1B1" },
      { paraId: "6CBA2F0C", durableId: "2B401A58" },
    ];
    const xml = commentsIdsDesc.stringify(entries, {} as never)!;
    expect(xml).toContain(
      '<w16cid:commentId w16cid:paraId="7FD6C115" w16cid:durableId="1EC7B1B1"/>',
    );

    const parsed = commentsIdsDesc.parse(
      parseXml(xml.replace(/^<w16cid:commentsIds [^>]*>/, `<w16cid:commentsIds ${COMMENTS_NS}>`))
        .elements![0]!,
      {} as never,
    );
    expect(parsed).toEqual(entries);
  });
});
