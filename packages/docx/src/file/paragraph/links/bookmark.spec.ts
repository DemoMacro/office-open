import { beforeEach, describe, expect, it } from "vite-plus/test";

import { TextRun } from "../run";
import { Bookmark, BookmarkEnd, BookmarkStart } from "./bookmark";

describe("Bookmark", () => {
  let bookmark: Bookmark;

  beforeEach(() => {
    bookmark = new Bookmark({
      children: [new TextRun("Internal Link")],
      id: "anchor",
    });
  });

  it("should create a bookmark with three root elements", () => {
    expect(bookmark.start).to.be.instanceOf(BookmarkStart);
    expect(bookmark.end).to.be.instanceOf(BookmarkEnd);
    expect(bookmark.children).to.have.length(1);
  });
});
