/**
 * Patch detector for discovering placeholders in document templates.
 *
 * @module
 */
import { createTraverser, DOCX_NS, toJson } from "@office-open/core";
import { toUint8ArrayAsync } from "@office-open/core";
import type { DataType, PlaceholderDelimiters } from "@office-open/core";
import { strFromU8, unzipSync } from "@office-open/core/opc";

/**
 * Options for patch detection.
 *
 * @property data - The document template to scan for placeholder targets
 * @property placeholderDelimiters - Custom token delimiters
 */
export interface PatchDetectorOptions {
  data: DataType;
  placeholderDelimiters?: PlaceholderDelimiters;
}

/**
 * Detects placeholder targets present in a document template.
 *
 * Scans WordprocessingML paragraph text in `word/` for placeholders enclosed by
 * delimiters (default: {{placeholder}}). This helps discover the patch targets
 * a template expects before replacement.
 *
 * @param options - Patch detector configuration
 * @returns Array of placeholder target keys found in the document
 *
 * @example
 * ```typescript
 * const targets = await patchDetector({ data: templateBuffer });
 * // Returns: ["name", "date", "address"] if template contains {{name}}, {{date}}, {{address}}
 *
 * // Use detected targets to create ordered canonical commands
 * const commands = targets.map((placeholder) => ({
 *   type: "replace",
 *   at: { field: "sections", route: [{ index: 0 }], text: { placeholder } },
 *   value: getUserData(placeholder),
 * }));
 * ```
 */
export const patchDetector = async ({
  data,
  placeholderDelimiters = { start: "{{", end: "}}" },
}: PatchDetectorOptions): Promise<string[]> => {
  const zipContent = unzipSync(await toUint8ArrayAsync(data));
  const targets = new Set<string>();

  for (const [key, value] of Object.entries(zipContent)) {
    if (!key.endsWith(".xml") && !key.endsWith(".rels")) {
      continue;
    }
    if (key.startsWith("word/") && !key.endsWith(".xml.rels")) {
      const json = toJson(strFromU8(value));
      const { traverse } = createTraverser(DOCX_NS);
      for (const p of traverse(json)) {
        for (const target of findPatchKeys(p.text, placeholderDelimiters)) {
          targets.add(target);
        }
      }
    }
  }
  return [...targets];
};

/**
 * Extracts placeholder keys from text using regex pattern.
 *
 * @param text - Text to search for placeholder targets
 * @param delimiters - Token delimiters surrounding each target
 * @returns Target keys without delimiters
 */
const findPatchKeys = (text: string, delimiters: PlaceholderDelimiters): string[] => {
  const escapedStart = delimiters.start.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const escapedEnd = delimiters.end.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const pattern = new RegExp(`(?<=${escapedStart}).+?(?=${escapedEnd})`, "gs");
  return text.match(pattern) ?? [];
};
