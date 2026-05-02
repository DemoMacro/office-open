import { Relationships } from "@file/relationships/relationships";
import type { XmlComponent } from "@file/xml-components";

import { Presentation } from "./presentation";
import type { IPresentationOptions } from "./presentation";

export interface IViewWrapper {
    readonly View: XmlComponent;
    readonly Relationships: Relationships;
}

export class PresentationWrapper implements IViewWrapper {
    private readonly presentation: Presentation;
    private readonly relationships: Relationships;

    public constructor(options: IPresentationOptions) {
        this.presentation = new Presentation(options);
        this.relationships = new Relationships();
    }

    public get View(): Presentation {
        return this.presentation;
    }

    public get Relationships(): Relationships {
        return this.relationships;
    }
}
