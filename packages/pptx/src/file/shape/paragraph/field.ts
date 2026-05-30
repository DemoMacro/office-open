import {
  buildAttrObject,
  BuilderElement,
  XmlComponent,
  stringContainerObj,
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
      buildAttrObject({
        id: `{${crypto.randomUUID()}}`,
        type: fieldType,
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

    this.root.push(stringContainerObj("a:t", displayText));
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
