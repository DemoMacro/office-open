import { describe, expect, it } from "vite-plus/test";

import type { IContext } from "./base";
import { InitializableXmlComponent } from "./initializable";

const emptyContext: IContext = { stack: [] };

describe("InitializableXmlComponent", () => {
    it("should create without init component", () => {
        class TestEl extends InitializableXmlComponent {
            public constructor(init?: InitializableXmlComponent) {
                super("w:p", init);
            }
        }

        const el = new TestEl();
        expect(el.prepForXml(emptyContext)).toEqual({ "w:p": {} });
    });

    it("should copy root from init component", () => {
        class TestEl extends InitializableXmlComponent {
            public constructor(init?: InitializableXmlComponent) {
                super("w:p", init);
            }
        }

        const source = new TestEl();
        (source as any).root.push("content");

        const copy = new TestEl(source);
        expect(copy.prepForXml(emptyContext)).toEqual({ "w:p": ["content"] });
    });
});
