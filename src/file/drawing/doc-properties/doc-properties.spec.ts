import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { DocProperties } from "./doc-properties";

describe("DocProperties", () => {
    it("should create with name, description, and title", () => {
        const dp = new DocProperties({ description: "desc", id: "1", name: "test", title: "ttl" });
        const tree = new Formatter().format(dp);
        expect(tree).to.deep.equal({
            "wp:docPr": {
                _attr: {
                    descr: "desc",
                    id: "1",
                    name: "test",
                    title: "ttl",
                },
            },
        });
    });

    it("should omit description attribute when description is undefined", () => {
        const dp = new DocProperties({ id: "1", name: "test", title: "ttl" });
        const tree = new Formatter().format(dp);
        expect(tree["wp:docPr"]._attr).not.to.have.property("descr");
        expect(tree["wp:docPr"]._attr).to.have.property("title", "ttl");
    });

    it("should omit title attribute when title is undefined", () => {
        const dp = new DocProperties({ description: "desc", id: "1", name: "test" });
        const tree = new Formatter().format(dp);
        expect(tree["wp:docPr"]._attr).to.have.property("descr", "desc");
        expect(tree["wp:docPr"]._attr).not.to.have.property("title");
    });

    it("should omit both description and title when neither is provided", () => {
        const dp = new DocProperties({ id: "1", name: "test" });
        const tree = new Formatter().format(dp);
        expect(tree["wp:docPr"]._attr).not.to.have.property("descr");
        expect(tree["wp:docPr"]._attr).not.to.have.property("title");
    });
});
