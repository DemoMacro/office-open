import { Relationships } from "@file/relationships/relationships";
import type { BaseXmlComponent } from "@file/xml-components";

import { Presentation } from "./presentation";
import type { PresentationOptions } from "./presentation";

export interface ViewWrapper {
  readonly View: BaseXmlComponent;
  readonly Relationships: Relationships;
}

export class PresentationWrapper implements ViewWrapper {
  private readonly presentation: Presentation;
  private readonly relationships: Relationships;

  public constructor(options: PresentationOptions) {
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
