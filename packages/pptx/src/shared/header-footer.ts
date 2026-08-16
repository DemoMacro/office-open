/**
 * Slide header/footer: options plus the slide-level placeholder shapes that
 * realize them. CT_Slide has no p:hf child — PowerPoint controls per-slide
 * date/footer/slide-number visibility by instantiating dt/ftr/sldNum
 * placeholder shapes (inheriting position from layout/master) on the slide.
 *
 * @module
 */

import type { ShapeOptions } from "@shared/shape/shape";

export interface SlideHeaderFooterOptions {
  slideNumber?: boolean;
  dateTime?: boolean;
  footer?: string | boolean;
  header?: boolean;
}

// PowerPoint's canonical dt/ftr/sldNum idx triple (observed across its output).
const DT_IDX = 10;
const FTR_IDX = 11;
const SLD_NUM_IDX = 12;

/**
 * Build the dt/ftr/sldNum placeholder shapes a headerFooter option enables.
 * The shapes carry no xfrm (position is inherited from layout/master); the
 * date and slide-number bodies are a:fld so PowerPoint recomputes them.
 */
export function buildHeaderFooterShapes(opts: SlideHeaderFooterOptions): ShapeOptions[] {
  const shapes: ShapeOptions[] = [];

  if (opts.dateTime) {
    shapes.push({
      placeholder: "dt",
      placeholderIndex: DT_IDX,
      placeholderSize: "half",
      textBody: { paragraphs: [{ children: [{ type: "datetimeFigureOut", text: "1/1/1" }] }] },
    });
  }

  if (opts.footer) {
    shapes.push({
      placeholder: "ftr",
      placeholderIndex: FTR_IDX,
      placeholderSize: "quarter",
      // A false/absent footer stays inherited-hidden; a true footer with no
      // text is still emitted as an empty body so the layout footer shows.
      textBody: typeof opts.footer === "string" ? { text: opts.footer } : {},
    });
  }

  if (opts.slideNumber) {
    shapes.push({
      placeholder: "sldNum",
      placeholderIndex: SLD_NUM_IDX,
      placeholderSize: "quarter",
      textBody: { paragraphs: [{ children: [{ type: "slidenum", text: "1" }] }] },
    });
  }

  return shapes;
}
