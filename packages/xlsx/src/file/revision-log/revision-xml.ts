/**
 * Revision Log XML generator — produces xl/revisions/revisionN.xml.
 *
 * Reference: OOXML transitional, sml.xsd, CT_Revisions
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import { attrs, escapeXml } from "@office-open/xml";

import type {
  RevisionEntry,
  RevisionRowColumnOptions,
  RevisionCellChangeOptions,
  RevisionMoveOptions,
  RevisionFormattingOptions,
  RevisionInsertSheetOptions,
  RevisionCommentOptions,
  RevisionDefinedNameOptions,
} from "./revision-types";

// ── Helpers ──

function buildRowColumn(opts: RevisionRowColumnOptions): string {
  const actionMap: Record<string, string> = {
    insertRow: "ir",
    insertCol: "ic",
    deleteRow: "dr",
    deleteCol: "dc",
  };
  const a: Record<string, string | number | boolean | undefined> = {
    rId: opts.rId,
    action: actionMap[opts.action],
    sId: opts.sheetIndex,
    edge: opts.edge ? 1 : undefined,
  };
  if (opts.action.includes("Row")) {
    a.row = opts.row;
  } else {
    a.col = opts.col;
  }
  return `<rrc${attrs(a)}/>`;
}

function buildCellChange(opts: RevisionCellChangeOptions): string {
  const a: Record<string, string | number | undefined> = {
    rId: opts.rId,
    sId: opts.sheetIndex,
  };

  const children: string[] = [];

  // Old value
  if (opts.oldValue !== undefined) {
    const oldType = opts.oldType ?? (typeof opts.oldValue === "number" ? "n" : "s");
    if (opts.formula) {
      children.push(`<f>${escapeXml(opts.formula)}</f>`);
    }
    children.push(
      `<oc${attrs({ t: oldType, vm: opts.numFmtId })}>` +
        `<v>${escapeXml(String(opts.oldValue))}</v></oc>`,
    );
  }

  // New value
  if (opts.newValue !== undefined) {
    const newType = opts.newType ?? (typeof opts.newValue === "number" ? "n" : "s");
    children.push(`<nc${attrs({ t: newType })}><v>${escapeXml(String(opts.newValue))}</v></nc>`);
  }

  return `<rcc${attrs({ ref: opts.ref, ...a })}>${children.join("")}</rcc>`;
}

function buildMove(opts: RevisionMoveOptions): string {
  const a: Record<string, string | number | undefined> = {
    rId: opts.rId,
    sId: opts.sheetIndex,
    source: opts.source,
    destination: opts.destination,
  };
  return `<rm${attrs(a)}/>`;
}

function buildFormatting(opts: RevisionFormattingOptions): string {
  const a: Record<string, string | number | undefined> = {
    rId: opts.rId,
    sId: opts.sheetIndex,
    ref: opts.ref,
    s: opts.s,
  };
  return `<rfmt${attrs(a)}/>`;
}

function buildInsertSheet(opts: RevisionInsertSheetOptions): string {
  const a: Record<string, string | number | undefined> = {
    rId: opts.rId,
    sId: opts.sheetIndex,
    name: opts.name,
  };
  return `<ris${attrs(a)}/>`;
}

function buildComment(opts: RevisionCommentOptions): string {
  const a: Record<string, string | number | undefined> = {
    rId: opts.rId,
    sId: opts.sheetIndex,
    ref: opts.ref,
    authorId: 0,
  };
  const children: string[] = [];
  if (opts.text) {
    children.push(`<t>${escapeXml(opts.text)}</t>`);
  }
  return `<rcmt${attrs(a)}>${children.join("")}</rcmt>`;
}

function buildDefinedName(opts: RevisionDefinedNameOptions): string {
  const a: Record<string, string | number | undefined> = {
    rId: opts.rId,
    name: opts.name,
    localSheetId: opts.localSheetId,
  };
  const children: string[] = [];
  if (opts.value) {
    children.push(`<formula>${escapeXml(opts.value)}</formula>`);
  }
  return `<rdn${attrs(a)}>${children.join("")}</rdn>`;
}

// ── Component ──

export class RevisionLogXml extends BaseXmlComponent {
  private readonly entries: readonly RevisionEntry[];

  public constructor(entries: readonly RevisionEntry[]) {
    super("revisions");
    this.entries = entries;
  }

  public override toXml(_context: Context): string {
    const p: string[] = [
      '<revisions xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    ];

    for (const entry of this.entries) {
      switch (entry.type) {
        case "rowColumn":
          p.push(buildRowColumn(entry.data));
          break;
        case "cellChange":
          p.push(buildCellChange(entry.data));
          break;
        case "move":
          p.push(buildMove(entry.data));
          break;
        case "formatting":
          p.push(buildFormatting(entry.data));
          break;
        case "insertSheet":
          p.push(buildInsertSheet(entry.data));
          break;
        case "comment":
          p.push(buildComment(entry.data));
          break;
        case "definedName":
          p.push(buildDefinedName(entry.data));
          break;
      }
    }

    p.push("</revisions>");
    return p.join("");
  }
}
