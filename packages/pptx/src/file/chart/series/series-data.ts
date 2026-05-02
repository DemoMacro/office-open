import { EmptyElement, XmlComponent } from "@file/xml-components";
import { chartAttr, wrapEl } from "@file/xml-components";

export const createStrRef = (values: string | readonly string[]): XmlComponent => {
    const arr = typeof values === "string" ? [values] : values;
    return new StrRef(arr);
};

export const createNumRef = (values: readonly number[]): XmlComponent => new NumRef(values);

class StrRef extends XmlComponent {
    public constructor(values: readonly string[]) {
        super("c:strRef");
        this.root.push(new EmptyElement("c:f"));
        this.root.push(new StrCache(values));
    }
}

class StrCache extends XmlComponent {
    public constructor(values: readonly string[]) {
        super("c:strCache");
        this.root.push(wrapEl("c:ptCount", chartAttr({ val: values.length })));
        for (let i = 0; i < values.length; i++) {
            this.root.push(new StrPt(i, values[i]));
        }
    }
}

class StrPt extends XmlComponent {
    public constructor(index: number, value: string) {
        super("c:pt");
        this.root.push(chartAttr({ idx: index }));
        this.root.push(new StringValue("c:v", value));
    }
}

class NumRef extends XmlComponent {
    public constructor(values: readonly number[]) {
        super("c:numRef");
        this.root.push(new EmptyElement("c:f"));
        this.root.push(new NumCache(values));
    }
}

class NumCache extends XmlComponent {
    public constructor(values: readonly number[]) {
        super("c:numCache");
        this.root.push(wrapEl("c:ptCount", chartAttr({ val: values.length })));
        this.root.push(new FormatCode("General"));
        for (let i = 0; i < values.length; i++) {
            this.root.push(new NumPt(i, values[i]));
        }
    }
}

class NumPt extends XmlComponent {
    public constructor(index: number, value: number) {
        super("c:pt");
        this.root.push(chartAttr({ idx: index }));
        this.root.push(new StringValue("c:v", String(value)));
    }
}

class FormatCode extends XmlComponent {
    public constructor(code: string) {
        super("c:formatCode");
        this.root.push(code);
    }
}

class StringValue extends XmlComponent {
    public constructor(name: string, val: string) {
        super(name);
        this.root.push(val);
    }
}
