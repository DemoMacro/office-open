import { Relationships } from "@file/relationships/relationships";
import type { BaseXmlComponent } from "@file/xml-components";

import { Presentation } from "./presentation";
import type { PresentationOptions } from "./presentation";

export interface ViewWrapper {
  readonly view: BaseXmlComponent;
  readonly relationships: Relationships;
}

export class PresentationWrapper implements ViewWrapper {
  private readonly presentation: Presentation;
  public readonly relationships: Relationships;

  public constructor(options: PresentationOptions) {
    this.presentation = new Presentation(options);
    this.relationships = new Relationships();
  }

  public get view(): Presentation {
    return this.presentation;
  }
}
