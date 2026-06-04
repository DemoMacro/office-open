import type { DocPropertiesOptions } from "@file/drawing/doc-properties/doc-properties";
import type { Context } from "@file/xml-components";

import { Run } from ".";
import { Drawing } from "../../drawing";
import type { Floating } from "../../drawing";
import type {
  IGroupChildMediaData,
  IMediaData,
  MediaTransformation,
  WpgMediaData,
} from "../../media";
import { createTransformation } from "../../media";

export * from "@file/drawing/inline/graphic/graphic-data/wps/body-properties";

interface CoreGroupOptions {
  readonly children: readonly IGroupChildMediaData[];
  readonly transformation: MediaTransformation;
  readonly floating?: Floating;
  readonly altText?: DocPropertiesOptions;
}

/**
 * @publicApi
 */
export type IWpgGroupOptions = { readonly type: "wpg" } & CoreGroupOptions;

/**
 * @publicApi
 */
export class WpgGroupRun extends Run {
  private readonly wpgGroupData: WpgMediaData;
  private readonly mediaDatas: readonly IMediaData[];

  public constructor(options: IWpgGroupOptions) {
    super({});

    this.wpgGroupData = {
      children: options.children,
      transformation: createTransformation(options.transformation),
      type: options.type,
    };
    const drawing = new Drawing(this.wpgGroupData, {
      docProperties: options.altText,
      floating: options.floating,
    });

    this.mediaDatas = options.children
      .filter((child) => child.type !== "wps")
      .map((child) => child as IMediaData);

    this.extraChildren.push(drawing);
  }

  protected override registerResources(context: Context): void {
    for (const child of this.mediaDatas) {
      context.file.media.addImage(child.fileName, child);

      if (child.type === "svg") {
        context.file.media.addImage(child.fallback.fileName, child.fallback);
      }
    }
  }
}
