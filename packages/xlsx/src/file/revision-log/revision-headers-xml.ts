/**
 * Revision Headers XML generator — produces xl/revisionHeaders.xml.
 *
 * Reference: OOXML transitional, sml.xsd, CT_RevisionHeaders
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import { attrs, escapeXml } from "@office-open/xml";

import type { RevisionHeadersOptions, UsersOptions } from "./revision-types";

// ── Component ──

export class RevisionHeadersXml extends BaseXmlComponent {
  private readonly opts: RevisionHeadersOptions;
  private readonly users?: UsersOptions;

  public constructor(options: RevisionHeadersOptions, users?: UsersOptions) {
    super("headers");
    this.opts = options;
    this.users = users;
  }

  public override toXml(_context: Context): string {
    const p: string[] = [
      '<headers xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
    ];

    // Root attributes
    const rootAttrs: Record<string, string | number | boolean | undefined> = {
      guid: this.opts.guid,
      lastGuid: this.opts.lastGuid,
      shared: this.opts.shared,
      history: this.opts.history,
      trackRevisions: this.opts.trackRevisions,
      revisionId: this.opts.revisionId,
      version: this.opts.version,
      keepChangeHistory: this.opts.keepChangeHistory,
      protected: this.opts.protected,
      preserveHistory: this.opts.preserveHistory,
    };
    p.push(`${attrs(rootAttrs)}>`);

    // header entries
    for (const h of this.opts.headers) {
      const headerAttrs: Record<string, string | number | undefined> = {
        guid: h.guid,
        dateTime: h.dateTime,
        maxSheetId: h.maxSheetId,
        userName: h.userName,
      };

      // Build header content
      const content: string[] = [];

      // sheetIdMap
      if (h.sheetIds && h.sheetIds.length > 0) {
        const idParts: string[] = [];
        for (const sid of h.sheetIds) {
          idParts.push(`<sheetId val="${sid.id}"/>`);
        }
        content.push(
          `<sheetIdMap${attrs({ count: h.sheetIds.length })}>${idParts.join("")}</sheetIdMap>`,
        );
      }

      p.push(
        `<header${attrs(headerAttrs)} r:id="${escapeXml(h.rId)}">${content.join("")}</header>`,
      );
    }

    p.push("</headers>");
    return p.join("");
  }

  /** Build users XML (CT_Users) — separate file xl/users.xml */
  public buildUsersXml(): string {
    if (!this.users?.users || this.users.users.length === 0) return "";
    const u = this.users.users;
    const parts: string[] = [
      '<users xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
      ` count="${u.length}">`,
    ];
    for (const user of u) {
      const uAttrs: Record<string, string | number | undefined> = {
        guid: user.guid,
        name: user.name,
        id: user.id,
        dateTime: user.dateTime,
      };
      parts.push(`<userInfo${attrs(uAttrs)}/>`);
    }
    parts.push("</users>");
    return parts.join("");
  }
}
