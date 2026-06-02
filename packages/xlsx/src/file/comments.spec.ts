import { describe, expect, it } from "vite-plus/test";

import { Comments } from "./comments";

const context = { stack: [] };

describe("Comments", () => {
  it("generates comments XML with authors and commentList", () => {
    const c = new Comments([{ cell: "A1", author: "Alice", text: "Hello" }]);
    const xml = c.toXml(context);
    expect(xml).toContain("<comments");
    expect(xml).toContain("<authors>");
    expect(xml).toContain("<author>Alice</author>");
    expect(xml).toContain("<commentList>");
    expect(xml).toContain('ref="A1"');
    expect(xml).toContain('authorId="0"');
    expect(xml).toContain("<t>Hello</t>");
  });

  it("deduplicates authors", () => {
    const c = new Comments([
      { cell: "A1", author: "Alice", text: "First" },
      { cell: "B1", author: "Alice", text: "Second" },
    ]);
    const xml = c.toXml(context);
    expect(xml).toContain('authorId="0"');
    expect(xml).toContain('authorId="0"');
  });

  it("handles multiple authors", () => {
    const c = new Comments([
      { cell: "A1", author: "Alice", text: "Hi" },
      { cell: "B1", author: "Bob", text: "Hey" },
    ]);
    const xml = c.toXml(context);
    expect(xml).toContain('authorId="0"');
    expect(xml).toContain('authorId="1"');
  });

  it("escapes special characters in text", () => {
    const c = new Comments([{ cell: "A1", author: "Alice", text: "<b>bold</b>" }]);
    const xml = c.toXml(context);
    expect(xml).toContain("&lt;b&gt;bold&lt;/b&gt;");
  });
});
