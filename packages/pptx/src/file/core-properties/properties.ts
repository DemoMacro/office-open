/**
 * Core Properties module for PresentationML documents.
 *
 * Reference: ISO-IEC29500-4_2016 shared-documentPropertiesCore.xsd
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import { buildCorePropertiesXmlString } from "@office-open/core";

export interface CorePropertiesOptions {
  readonly title?: string;
  readonly subject?: string;
  readonly creator?: string;
  readonly keywords?: string;
  readonly description?: string;
  readonly lastModifiedBy?: string;
  readonly revision?: number;
}

export class CoreProperties extends BaseXmlComponent {
  private readonly options: CorePropertiesOptions;

  public constructor(options: CorePropertiesOptions) {
    super("cp:coreProperties");
    this.options = options;
  }

  public override toXml(_context: Context): string {
    return buildCorePropertiesXmlString(this.options);
  }
}
