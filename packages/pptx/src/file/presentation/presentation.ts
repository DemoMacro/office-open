import { BuilderElement, NextAttributeComponent, XmlComponent } from "@file/xml-components";

export interface IPresentationOptions {
    readonly slideWidth?: number;
    readonly slideHeight?: number;
    readonly slideIds: readonly number[];
}

/**
 * p:presentation — Root element of a PPTX file.
 */
export class Presentation extends XmlComponent {
    public constructor(options: IPresentationOptions) {
        super("p:presentation");
        this.root.push(
            new NextAttributeComponent({
                "xmlns:a": {
                    key: "xmlns:a",
                    value: "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                "xmlns:r": {
                    key: "xmlns:r",
                    value: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                },
                "xmlns:p": {
                    key: "xmlns:p",
                    value: "http://schemas.openxmlformats.org/presentationml/2006/main",
                },
            }),
        );

        this.root.push(
            new BuilderElement({
                name: "p:sldMasterIdLst",
                children: [
                    new BuilderElement({
                        name: "p:sldMasterId",
                        attributes: {
                            id: { key: "id", value: 2147483648 },
                            rId: { key: "r:id", value: "rId1" },
                        },
                    }),
                ],
            }),
        );

        this.root.push(
            new BuilderElement({
                name: "p:sldIdLst",
                children: options.slideIds.map(
                    (_id, i) =>
                        new BuilderElement({
                            name: "p:sldId",
                            attributes: {
                                id: { key: "id", value: 256 + i },
                                rId: { key: "r:id", value: `rId${i + 2}` },
                            },
                        }),
                ),
            }),
        );

        this.root.push(
            new BuilderElement({
                name: "p:sldSz",
                attributes: {
                    cx: { key: "cx", value: options.slideWidth ?? 9144000 },
                    cy: { key: "cy", value: options.slideHeight ?? 6858000 },
                    type: {
                        key: "type",
                        value:
                            (options.slideWidth ?? 9144000) === 9144000 &&
                            (options.slideHeight ?? 6858000) === 6858000
                                ? "screen4x3"
                                : undefined,
                    },
                },
            }),
        );

        this.root.push(
            new BuilderElement({
                name: "p:notesSz",
                attributes: {
                    cx: { key: "cx", value: 6858000 },
                    cy: { key: "cy", value: 9144000 },
                },
            }),
        );

        this.root.push(this.createDefaultTextStyle());
    }

    private createDefaultTextStyle(): BuilderElement {
        const defaultRunProps = (sz: number): BuilderElement =>
            new BuilderElement({
                name: "a:defRPr",
                attributes: {
                    sz: { key: "sz", value: sz * 100 },
                    kern: { key: "kern", value: 1200 },
                },
                children: [
                    new BuilderElement({
                        name: "a:solidFill",
                        children: [
                            new BuilderElement({
                                name: "a:schemeClr",
                                attributes: { val: { key: "val", value: "tx1" } },
                            }),
                        ],
                    }),
                    new BuilderElement({
                        name: "a:latin",
                        attributes: { typeface: { key: "typeface", value: "+mn-lt" } },
                    }),
                    new BuilderElement({
                        name: "a:ea",
                        attributes: { typeface: { key: "typeface", value: "+mn-ea" } },
                    }),
                    new BuilderElement({
                        name: "a:cs",
                        attributes: { typeface: { key: "typeface", value: "+mn-cs" } },
                    }),
                ],
            });

        const createLevel = (marL: number): BuilderElement =>
            new BuilderElement({
                name: `a:lvl${Math.floor(marL / 457200 + 1)}pPr`,
                attributes: {
                    marL: { key: "marL", value: marL },
                    algn: { key: "algn", value: "l" },
                    defTabSz: { key: "defTabSz", value: 914400 },
                    rtl: { key: "rtl", value: 0 },
                    eaLnBrk: { key: "eaLnBrk", value: 1 },
                    latinLnBrk: { key: "latinLnBrk", value: 0 },
                    hangingPunct: { key: "hangingPunct", value: 1 },
                },
                children: [defaultRunProps(18)],
            });

        return new BuilderElement({
            name: "p:defaultTextStyle",
            children: [
                new BuilderElement({
                    name: "a:defPPr",
                    children: [new BuilderElement({ name: "a:defRPr" })],
                }),
                createLevel(0),
                createLevel(457200),
                createLevel(914400),
                createLevel(1371600),
                createLevel(1828800),
                createLevel(2286000),
                createLevel(2743200),
                createLevel(3200400),
                createLevel(3657600),
            ],
        });
    }
}
