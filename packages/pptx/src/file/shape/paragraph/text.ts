import { StringContainer } from "@file/xml-components";

/**
 * a:t — Text content within a run.
 */
export class Text extends StringContainer {
    public constructor(value: string) {
        super("a:t", value);
    }
}
