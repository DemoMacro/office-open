import { BuilderElement, XmlComponent } from "@file/xml-components";

/**
 * a:blipFill — Image fill with stretch mode.
 */
export class BlipFill extends XmlComponent {
    public constructor(fileName: string) {
        super("a:blipFill");
        this.root.push(
            new BuilderElement({
                name: "a:blip",
                attributes: {
                    "r:embed": { key: "r:embed", value: `{${fileName}}` },
                },
            }),
        );
        this.root.push(
            new BuilderElement({
                name: "a:stretch",
                children: [new BuilderElement({ name: "a:fillRect" })],
            }),
        );
    }
}
