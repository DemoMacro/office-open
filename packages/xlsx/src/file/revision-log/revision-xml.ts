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
  RevisionAutoFormattingOptions,
  RevisionCustomViewOptions,
  RevisionSheetRenameOptions,
  RevisionQueryTableFieldOptions,
  RevisionConflictOptions,
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
  const a: Record<string, string | number | boolean | undefined> = {
    rId: opts.rId,
    sId: opts.sheetIndex,
    quotePrefix: opts.quotePrefix ? 1 : undefined,
    oldQuotePrefix: opts.oldQuotePrefix ? 1 : undefined,
    ph: opts.ph ? 1 : undefined,
    oldPh: opts.oldPh ? 1 : undefined,
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

  // ndxf/odxf — differential format (CT_Dxf placeholder)
  if (opts.xfDxf !== undefined) {
    children.push(`<ndxf><font/><numFmt/><fill/><border/><protection/></ndxf>`);
    children.push(`<odxf><font/><numFmt/><fill/><border/><protection/></odxf>`);
  }

  return `<rcc${attrs({ ref: opts.ref, ...a })}>${children.join("")}</rcc>`;
}

function buildMove(opts: RevisionMoveOptions): string {
  const a: Record<string, string | number | undefined> = {
    rId: opts.rId,
    sId: opts.sheetIndex,
    source: opts.source,
    destination: opts.destination,
    sourceSheetId: opts.sourceSheetId,
  };
  return `<rm${attrs(a)}/>`;
}

function buildFormatting(opts: RevisionFormattingOptions): string {
  const a: Record<string, string | number | undefined> = {
    rId: opts.rId,
    sId: opts.sheetIndex,
    ref: opts.ref,
    s: opts.s,
    xfDxf: opts.xfDxf,
  };
  return `<rfmt${attrs(a)}/>`;
}

function buildInsertSheet(opts: RevisionInsertSheetOptions): string {
  const a: Record<string, string | number | undefined> = {
    rId: opts.rId,
    sId: opts.sheetIndex,
    name: opts.name,
    sheetPosition: opts.sheetPosition,
  };
  return `<ris${attrs(a)}/>`;
}

function buildComment(opts: RevisionCommentOptions): string {
  const a: Record<string, string | number | boolean | undefined> = {
    rId: opts.rId,
    sId: opts.sheetIndex,
    ref: opts.ref,
    alwaysShow: opts.alwaysShow ? 1 : undefined,
    old: opts.old ? 1 : undefined,
    hiddenRow: opts.hiddenRow ? 1 : undefined,
    hiddenColumn: opts.hiddenColumn ? 1 : undefined,
    oldLength: opts.oldLength,
    newLength: opts.newLength,
  };
  const children: string[] = [];
  if (opts.text) {
    children.push(`<t>${escapeXml(opts.text)}</t>`);
  }
  if (opts.author) {
    children.push(`<author>${escapeXml(opts.author)}</author>`);
  }
  if (children.length > 0) {
    return `<rcmt${attrs(a)}>${children.join("")}</rcmt>`;
  }
  return `<rcmt${attrs(a)}/>`;
}

function buildDefinedName(opts: RevisionDefinedNameOptions): string {
  const a: Record<string, string | number | boolean | undefined> = {
    rId: opts.rId,
    name: opts.name,
    localSheetId: opts.localSheetId,
    customView: opts.customView ? 1 : undefined,
    function: opts["function"] ? 1 : undefined,
    oldFunction: opts.oldFunction ? 1 : undefined,
    functionGroupId: opts.functionGroupId,
    oldFunctionGroupId: opts.oldFunctionGroupId,
    shortcutKey: opts.shortcutKey,
    oldShortcutKey: opts.oldShortcutKey,
    oldHidden: opts.oldHidden ? 1 : undefined,
    customMenu: opts.customMenu,
    oldCustomMenu: opts.oldCustomMenu,
    oldDescription: opts.oldDescription,
    help: opts.help,
    oldHelp: opts.oldHelp,
    statusBar: opts.statusBar,
    oldStatusBar: opts.oldStatusBar,
  };
  const children: string[] = [];
  if (opts.value) {
    children.push(`<formula>${escapeXml(opts.value)}</formula>`);
  }
  if (opts.oldComment) {
    children.push(`<oldFormula>${escapeXml(opts.oldComment)}</oldFormula>`);
  }
  return `<rdn${attrs(a)}>${children.join("")}</rdn>`;
}

function buildAutoFormatting(opts: RevisionAutoFormattingOptions): string {
  const a: Record<string, string | number | undefined> = {
    rId: opts.rId,
    sId: opts.sheetIndex,
    ref: opts.ref,
  };
  return `<raf${attrs(a)}/>`;
}

function buildCustomView(opts: RevisionCustomViewOptions): string {
  const a: Record<string, string | number | undefined> = {
    rId: opts.rId,
    guid: opts.guid,
  };
  return `<rcv${attrs(a)}/>`;
}

function buildSheetRename(opts: RevisionSheetRenameOptions): string {
  const a: Record<string, string | number | undefined> = {
    rId: opts.rId,
    sId: opts.sheetIndex,
    oldName: opts.oldName,
    newName: opts.newName,
  };
  return `<rsnm${attrs(a)}/>`;
}

function buildQueryTableField(opts: RevisionQueryTableFieldOptions): string {
  const a: Record<string, string | number | undefined> = {
    rId: opts.rId,
    sId: opts.sheetIndex,
    fieldId: opts.fieldId,
  };
  return `<rqt${attrs(a)}/>`;
}

function buildConflict(opts: RevisionConflictOptions): string {
  const a: Record<string, string | number | undefined> = {
    rId: opts.rId,
    sId: opts.sheetIndex,
  };
  return `<rcft${attrs(a)}/>`;
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

    const localReviewedList: string[] = [];

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
        case "reviewed":
          localReviewedList.push(`<reviewed rId="${entry.data.rId}"/>`);
          p.push(`<reviewed rId="${entry.data.rId}"/>`);
          break;
        case "undo":
          p.push(`<undo rId="${entry.data.rId}"/>`);
          break;
        case "autoFormatting":
          p.push(buildAutoFormatting(entry.data));
          break;
        case "customView":
          p.push(buildCustomView(entry.data));
          break;
        case "sheetRename":
          p.push(buildSheetRename(entry.data));
          break;
        case "queryTableField":
          p.push(buildQueryTableField(entry.data));
          break;
        case "conflict":
          p.push(buildConflict(entry.data));
          break;
      }
    }

    // reviewedList — container for all reviewed entries (CT_ReviewedList)
    if (localReviewedList.length > 0) {
      p.push(`<reviewedList>${localReviewedList.join("")}</reviewedList>`);
    }

    p.push("</revisions>");
    return p.join("");
  }
}
