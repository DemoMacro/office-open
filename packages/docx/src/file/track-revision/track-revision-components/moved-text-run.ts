/**
 * Moved text run module for track changes.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, moveFrom/moveTo (CT_RunTrackChange)
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";

import { createBreak } from "../../paragraph/run/break";
import { createBegin, createEnd, createSeparate } from "../../paragraph/run/field";
import { RunProperties } from "../../paragraph/run/properties";
import { PageNumber } from "../../paragraph/run/run";
import type { RunOptions } from "../../paragraph/run/run";
import { buildText } from "../../paragraph/run/run-components/text";
import type { ChangedAttributesProperties } from "../track-revision";

type IMovedRunOptions = RunOptions & ChangedAttributesProperties;

/**
 * Represents a move source run in a tracked changes document.
 *
 * Marks text that has been moved from this location. Content uses
 * `w:t` (standard text element).
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, moveFrom (CT_RunTrackChange)
 *
 * @example
 * ```typescript
 * new MovedFromTextRun({
 *   id: 1,
 *   author: "John",
 *   date: "2024-01-01T00:00:00Z",
 *   text: "Moved text",
 * });
 * ```
 */
export class MovedFromTextRun extends XmlComponent {
  public constructor(options: IMovedRunOptions) {
    super("w:moveFrom");
    this.root.push({
      _attr: { "w:author": options.author, "w:date": options.date, "w:id": options.id },
    });
    this.addChildElement(new MovedFromRunWrapper(options));
  }
}

/**
 * Represents a move destination run in a tracked changes document.
 *
 * Marks text that has been moved to this location. Content uses
 * standard `w:t` text element.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, moveTo (CT_RunTrackChange)
 *
 * @example
 * ```typescript
 * new MovedToTextRun({
 *   id: 1,
 *   author: "John",
 *   date: "2024-01-01T00:00:00Z",
 *   text: "Moved text",
 * });
 * ```
 */
export class MovedToTextRun extends XmlComponent {
  public constructor(options: IMovedRunOptions) {
    super("w:moveTo");
    this.root.push({
      _attr: { "w:author": options.author, "w:date": options.date, "w:id": options.id },
    });
    this.addChildElement(new MovedToRunWrapper(options));
  }
}

/**
 * Internal wrapper for the run element within a move source.
 * Uses `w:t` (standard text) — moveFrom is a move, not a deletion.
 *
 * @internal
 */
class MovedFromRunWrapper extends XmlComponent {
  public constructor(options: RunOptions) {
    super("w:r");
    this.root.push(new RunProperties(options));

    if (options.children) {
      for (const child of options.children) {
        if (typeof child === "string") {
          switch (child) {
            case PageNumber.CURRENT: {
              this.root.push(createBegin());
              this.root.push(buildText("PAGE"));
              this.root.push(createSeparate());
              this.root.push(createEnd());
              break;
            }
            case PageNumber.TOTAL_PAGES: {
              this.root.push(createBegin());
              this.root.push(buildText("NUMPAGES"));
              this.root.push(createSeparate());
              this.root.push(createEnd());
              break;
            }
            case PageNumber.TOTAL_PAGES_IN_SECTION: {
              this.root.push(createBegin());
              this.root.push(buildText("SECTIONPAGES"));
              this.root.push(createSeparate());
              this.root.push(createEnd());
              break;
            }
            default: {
              this.root.push(buildText(child));
              break;
            }
          }
          continue;
        }

        this.root.push(child);
      }
    } else if (options.text) {
      this.root.push(buildText(options.text));
    }

    if (options.break) {
      for (let i = 0; i < options.break; i++) {
        this.root.splice(1, 0, createBreak());
      }
    }
  }
}

/**
 * Internal wrapper for the run element within a move destination.
 * Uses standard `w:t` text element.
 *
 * @internal
 */
class MovedToRunWrapper extends XmlComponent {
  public constructor(options: RunOptions) {
    super("w:r");
    this.root.push(new RunProperties(options));

    if (options.children) {
      for (const child of options.children) {
        this.root.push(child);
      }
    } else if (options.text) {
      this.root.push(buildText(options.text));
    }

    if (options.break) {
      for (let i = 0; i < options.break; i++) {
        this.root.splice(1, 0, createBreak());
      }
    }
  }
}
