/**
 * Replace relationship references in raw XML with `{fileName}` placeholders.
 *
 * Raw-XML passthrough paths (document background, mc:AlternateContent VML
 * fallback) carry VML/structured content verbatim. Their r:id / r:embed /
 * r:link references must be replaced with `{fileName}` placeholders and the
 * referenced media collected, so the compiler's placeholder pass registers the
 * media and resolves the placeholders into relationship ids. Otherwise the
 * carried source rIds dangle — they are not defined in the generated rels and
 * Word rejects the package as unreadable.
 */
import type { DocxReadContext } from "../context";
import type { BackgroundRawMediaOptions } from "../parts/document/document-background/document-background";
import { imageTypeFromPath } from "../parts/drawing/drawing-parse";

const REL_ATTR_RE = /\br:(id|embed|link)="([^"]+)"/g;

export function replaceRelsWithPlaceholders(
  xml: string,
  ctx: DocxReadContext,
  prefix: string,
): { rawXml: string; rawMedia: BackgroundRawMediaOptions[] } {
  const rawMedia: BackgroundRawMediaOptions[] = [];
  const rawXml = xml.replace(REL_ATTR_RE, (match, attrName: string, rId: string) => {
    const mediaPath = ctx.resolveRelationship(rId);
    const data = mediaPath ? ctx.getRaw(mediaPath) : undefined;
    if (!mediaPath || !data) return match;
    const type = imageTypeFromPath(mediaPath);
    const fileName = `${prefix}-${rId}.${type}`;
    if (!rawMedia.some((m) => m.fileName === fileName)) {
      rawMedia.push({ fileName, data, type });
    }
    return `r:${attrName}="{${fileName}}"`;
  });
  return { rawXml, rawMedia };
}
