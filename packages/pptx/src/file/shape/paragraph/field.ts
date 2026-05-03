import {
    NextAttributeComponent,
    BuilderElement,
    StringContainer,
    XmlComponent,
} from "@file/xml-components";

/**
 * a:fld — A field element for dynamic content in slide text.
 *
 * Used in slide master placeholders for slide numbers, dates, etc.
 */
export class Field extends XmlComponent {
    public constructor(fieldType: string, displayText: string) {
        super("a:fld");

        this.root.push(
            new NextAttributeComponent({
                id: { key: "id", value: `{${crypto.randomUUID()}}` },
                type: { key: "type", value: fieldType },
            }),
        );

        this.root.push(
            new BuilderElement({
                name: "a:rPr",
                attributes: {
                    lang: { key: "lang", value: "en-US" },
                    smtClean: { key: "smtClean", value: 0 },
                },
            }),
        );

        this.root.push(new StringContainer("a:t", displayText));
    }
}

/**
 * Slide number field — outputs current slide number.
 */
export class SlideNumberField extends Field {
    public constructor(displayText = "<#>") {
        super("slidenum", displayText);
    }
}

/**
 * Date/time field — outputs formatted date.
 */
export class DateTimeField extends Field {
    public constructor(displayText?: string, format?: string) {
        super(format ?? "datetimeFigureOut", displayText ?? "");
    }
}
