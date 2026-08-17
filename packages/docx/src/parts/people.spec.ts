import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { commentsExtendedDesc, type CommentExtendedOptions } from "./comments-extended";
import { peopleDesc, type CommentPersonOptions } from "./people";

const W15_NS = 'xmlns:w15="http://schemas.microsoft.com/office/word/2010/11/wordml"';

/** Parse a descriptor's output by swapping its root attributes for a minimal ns set. */
function reparse(xml: string, rootTag: string): ReturnType<typeof parseXml> {
  return parseXml(xml.replace(/^<w15:[\w]+ [^>]*>/, `<${rootTag} ${W15_NS}>`));
}

describe("peopleDesc", () => {
  it("round-trips author and contact", () => {
    const people: CommentPersonOptions[] = [
      { author: "Alice <alice@example.com>", contact: "alice@example.com" },
      { author: "Bob" },
    ];
    const xml = peopleDesc.stringify(people, {} as never)!;
    expect(xml).toContain('w15:author="Alice &lt;alice@example.com&gt;"');
    expect(xml).toContain('w15:contact="alice@example.com"');
    expect(xml).toContain('<w15:person w15:author="Bob"/>');

    const parsed = peopleDesc.parse(reparse(xml, "w15:people").elements![0]!, {} as never);
    expect(parsed).toEqual(people);
  });
});

describe("commentsExtendedDesc", () => {
  it("round-trips paraId, parent link, and done flag", () => {
    const entries: CommentExtendedOptions[] = [
      { paraId: "7FD6C115", done: false },
      { paraId: "6CBA2F0C", paraIdParent: "7FD6C115", done: true },
      { paraId: "3DABD945" },
    ];
    const xml = commentsExtendedDesc.stringify(entries, {} as never)!;
    expect(xml).toContain('<w15:commentEx w15:paraId="7FD6C115" w15:done="0"/>');
    expect(xml).toContain(
      '<w15:commentEx w15:paraId="6CBA2F0C" w15:paraIdParent="7FD6C115" w15:done="1"/>',
    );
    expect(xml).toContain('<w15:commentEx w15:paraId="3DABD945"/>');

    const parsed = commentsExtendedDesc.parse(
      reparse(xml, "w15:commentsEx").elements![0]!,
      {} as never,
    );
    expect(parsed).toEqual(entries);
  });
});
