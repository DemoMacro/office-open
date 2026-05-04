import { Formatter } from "@export/formatter";
import { assert, beforeEach, describe, expect, it } from "vite-plus/test";

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

    it("should create a bookmark with the correct attributes on the bookmark start element", () => {
        const tree = new Formatter().format(bookmark.start);
        const attrs = (tree["w:bookmarkStart"] as Record<string, unknown>)["_attr"] as Record<
            string,
            unknown
        >;
        assert.equal(attrs["w:name"], "anchor");
    });

    it("should create a bookmark with the correct text content", () => {
        const tree = new Formatter().format(new TextRun("Internal Link"));
        expect(JSON.stringify(tree)).to.include("Internal Link");
    });

    it("should create a bookmark with the correct attributes on the bookmark end element", () => {
        const tree = new Formatter().format(bookmark.end);
        const attrs = (tree["w:bookmarkEnd"] as Record<string, unknown>)["_attr"] as Record<
            string,
            unknown
        >;
        expect(attrs["w:id"]).to.be.a("number");
    });
});
