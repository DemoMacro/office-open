import type { IMediaDataTransformation } from "@file/media";
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

import { Form } from "../pic/shape-properties/form";

export type GroupChild = XmlComponent;

export interface WpgGroupCoreOptions {
    readonly children: readonly GroupChild[];
}

export type WpgGroupOptions = WpgGroupCoreOptions & {
    readonly transformation: IMediaDataTransformation;
};

const createGroupProperties = (transform: IMediaDataTransformation): XmlComponent =>
    new BuilderElement({
        children: [new Form(transform)],
        name: "wpg:grpSpPr",
    });

const createNonVisualGroupProperties = (): XmlComponent =>
    new BuilderElement({
        name: "wpg:cNvGrpSpPr",
    });

export const createWpgGroup = (options: WpgGroupOptions): XmlComponent =>
    new BuilderElement({
        children: [
            createNonVisualGroupProperties(),
            createGroupProperties(options.transformation),
            ...options.children,
        ],
        name: "wpg:wgp",
    });
