import type { LongHexNumber } from "@office-open/core";
import type { ChartExSpaceOptions } from "@office-open/core/chart";
import type { Floating } from "@parts/drawing";
import type { DrawingExtensionIds, GraphicFrameLocksOptions } from "@parts/drawing/descriptor";
import type { DocPropertiesOptions } from "@parts/drawing/doc-properties/doc-properties";
import type { RunPropertiesOptions } from "@parts/paragraph/run/properties";
import type { MediaTransformation } from "@shared/media";

export interface ChartExOptions extends ChartExSpaceOptions {
  transformation: MediaTransformation;
  floating?: Floating;
  altText?: DocPropertiesOptions;
  graphicFrameLocks?: GraphicFrameLocksOptions | null;
  extensionIds?: DrawingExtensionIds;
  runProperties?: RunPropertiesOptions;
  additionRsid?: LongHexNumber;
  deletionRsid?: LongHexNumber;
  runPropertiesRsid?: LongHexNumber;
  lastRenderedPageBreak?: boolean;
}
