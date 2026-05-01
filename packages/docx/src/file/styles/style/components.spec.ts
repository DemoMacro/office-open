import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import * as components from "./components";

describe("Style components", () => {
    it("Name#constructor", () => {
        const style = new components.Name("Style Name");
        const tree = new Formatter().format(style);
        expect(tree).to.deep.equal({ "w:name": { _attr: { "w:val": "Style Name" } } });
    });

    it("UiPriority#constructor", () => {
        const style = new components.UiPriority(123);
        const tree = new Formatter().format(style);
        expect(tree).to.deep.equal({ "w:uiPriority": { _attr: { "w:val": 123 } } });
    });
});
