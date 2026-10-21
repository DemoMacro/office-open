import type { Element } from "@office-open/xml";

import { PatchError } from "./diagnostics";
import type { PlaceholderDelimiters } from "./types";
import { toJson } from "./xml-patch-utils";

/** Mutable OPC member view shared by format patch contexts. */
export class PatchContext {
  readonly xml = new Map<string, Element>();
  readonly binary = new Map<string, Uint8Array>();
  readonly text = new Map<string, string>();

  constructor(
    parts: Record<string, Uint8Array>,
    classify: (path: string, bytes: Uint8Array) => "binary" | "text" | "xml",
  ) {
    for (const [path, bytes] of Object.entries(parts)) {
      const kind = classify(path, bytes);
      if (kind === "binary") {
        this.binary.set(path, bytes);
        continue;
      }
      if (kind === "text") {
        this.text.set(path, new TextDecoder().decode(bytes));
      } else {
        this.xml.set(path, toJson(new TextDecoder().decode(bytes)));
      }
    }
  }

  part(path: string): Element | undefined {
    return this.xml.get(path);
  }

  setPart(path: string, element: Element): void {
    this.xml.set(path, element);
  }

  removePart(path: string): void {
    this.xml.delete(path);
    this.binary.delete(path);
    this.text.delete(path);
  }
}

/** Validate placeholder delimiters before patch work begins. */
export function assertPatchDelimiters(delimiters: PlaceholderDelimiters): void {
  if (!delimiters.start.trim() || !delimiters.end.trim()) {
    throw new PatchError({
      message: "Placeholder delimiters must be non-empty strings.",
      part: "[package]",
      path: "placeholderDelimiters",
      reason: "invalid-target",
    });
  }
}

/** Create a structured error for a format that does not support patching. */
export function unsupportedPatchFormat(format: string, family: string): PatchError {
  return new PatchError({
    message: `${format} does not support ${family} patching.`,
    operation: "patch",
    part: "[package]",
    path: "format",
    reason: "invalid-part",
  });
}
