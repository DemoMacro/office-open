import type { OfficeFormsOptions } from "@office-open/odf-schema";
import type { PresentationOptions, SlideOptions } from "@office-open/pptx";

export interface OdpPresentationOptions extends PresentationOptions {
  slides?: OdpSlideOptions[];
}
export interface OdpSlideOptions extends SlideOptions {
  forms?: OfficeFormsOptions;
}
