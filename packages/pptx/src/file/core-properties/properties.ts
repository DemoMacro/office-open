/**
 * Core Properties module for PresentationML documents.
 *
 * Reference: ISO-IEC29500-4_2016 shared-documentPropertiesCore.xsd
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";
import { buildCorePropertiesXml } from "@office-open/core";

export interface ICorePropertiesOptions {
    readonly title?: string;
    readonly subject?: string;
    readonly creator?: string;
    readonly keywords?: string;
    readonly description?: string;
    readonly lastModifiedBy?: string;
    readonly revision?: number;
}

export class CoreProperties extends BaseXmlComponent {
    private readonly options: ICorePropertiesOptions;

    public constructor(options: ICorePropertiesOptions) {
        super("cp:coreProperties");
        this.options = options;
    }

    public override prepForXml(_context: IContext): IXmlableObject {
        return buildCorePropertiesXml(this.options);
    }
}
