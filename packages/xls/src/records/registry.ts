import { XlsParseError } from "../errors";
import type { BiffRecord } from "./index";

export type BiffRecordScope = "globals" | "sheet" | "both";

export type BiffRecordDisposition =
  | { kind: "canonical"; owner: "record-reader" | "workbook-globals" | "worksheet" }
  | { kind: "structural-noop"; reason: string }
  | { kind: "unsupported" };

export interface BiffRegistryEntry {
  readonly code: number;
  readonly name: string;
  readonly scope: BiffRecordScope;
  readonly disposition: BiffRecordDisposition;
}

export type BiffRecordRegistry = ReadonlyMap<number, BiffRegistryEntry>;

const canonical = (
  code: number,
  name: string,
  scope: BiffRecordScope,
  owner: "record-reader" | "workbook-globals" | "worksheet",
): [number, BiffRegistryEntry] => [
  code,
  { code, name, scope, disposition: { kind: "canonical", owner } },
];

const noop = (
  code: number,
  name: string,
  scope: BiffRecordScope,
  reason: string,
): [number, BiffRegistryEntry] => [
  code,
  { code, name, scope, disposition: { kind: "structural-noop", reason } },
];

export const BIFF_RECORD_REGISTRY: BiffRecordRegistry = new Map([
  canonical(0x000a, "End Of File", "both", "record-reader"),
  canonical(0x0809, "Begin Of File", "both", "record-reader"),
  canonical(0x002f, "File Pass", "both", "record-reader"),
  canonical(0x003c, "Continue", "both", "record-reader"),

  canonical(0x00fc, "Shared String Table", "globals", "workbook-globals"),
  canonical(0x0085, "Bound Sheet", "globals", "workbook-globals"),
  canonical(0x0042, "Code Page", "globals", "workbook-globals"),
  canonical(0x0022, "Date Mode", "globals", "workbook-globals"),
  canonical(0x0031, "Font", "globals", "workbook-globals"),
  canonical(0x041e, "Number Format", "globals", "workbook-globals"),
  canonical(0x0092, "Palette", "globals", "workbook-globals"),
  canonical(0x01ae, "Supplemental Book", "globals", "workbook-globals"),
  canonical(0x0017, "Extern Sheet", "globals", "workbook-globals"),
  canonical(0x0018, "Defined Name", "globals", "workbook-globals"),
  canonical(0x00e0, "Extended Format", "globals", "workbook-globals"),

  canonical(0x0200, "Dimensions", "sheet", "worksheet"),
  canonical(0x0208, "Row", "sheet", "worksheet"),
  canonical(0x0001, "BIFF2 Blank", "sheet", "worksheet"),
  canonical(0x0201, "Blank", "sheet", "worksheet"),
  canonical(0x0005, "BIFF2 Boolean Or Error", "sheet", "worksheet"),
  canonical(0x0205, "Boolean Or Error", "sheet", "worksheet"),
  canonical(0x00be, "Multiple Blank", "sheet", "worksheet"),
  canonical(0x0003, "BIFF2 Number", "sheet", "worksheet"),
  canonical(0x0203, "Number", "sheet", "worksheet"),
  canonical(0x0002, "BIFF2 Integer", "sheet", "worksheet"),
  canonical(0x0004, "BIFF2 Label", "sheet", "worksheet"),
  canonical(0x0204, "Label", "sheet", "worksheet"),
  canonical(0x00d6, "Rich Text Label", "sheet", "worksheet"),
  canonical(0x00fd, "Shared String Label", "sheet", "worksheet"),
  canonical(0x027e, "Recursive Record Number", "sheet", "worksheet"),
  canonical(0x00bd, "Multiple Recursive Number", "sheet", "worksheet"),
  canonical(0x0006, "Formula", "sheet", "worksheet"),
  canonical(0x04bc, "Shared Formula", "sheet", "worksheet"),
  canonical(0x0221, "Array Formula", "sheet", "worksheet"),
  canonical(0x007d, "Column Information", "sheet", "worksheet"),
  canonical(0x00e5, "Merged Cells", "sheet", "worksheet"),
  canonical(0x01b8, "Hyperlink", "sheet", "worksheet"),
  canonical(0x005d, "Object", "sheet", "worksheet"),
  canonical(0x01b6, "Text Object", "sheet", "worksheet"),
  canonical(0x001c, "Note", "sheet", "worksheet"),
  canonical(0x00eb, "MSO Drawing Group", "both", "workbook-globals"),
  canonical(0x00ec, "MSO Drawing", "sheet", "worksheet"),
  canonical(0x01b0, "Conditional Format Range", "sheet", "worksheet"),
  canonical(0x01b1, "Conditional Format Rule", "sheet", "worksheet"),
  canonical(0x0014, "Header", "sheet", "worksheet"),
  canonical(0x0015, "Footer", "sheet", "worksheet"),
  canonical(0x0026, "Left Margin", "sheet", "worksheet"),
  canonical(0x0027, "Right Margin", "sheet", "worksheet"),
  canonical(0x0028, "Top Margin", "sheet", "worksheet"),
  canonical(0x0029, "Bottom Margin", "sheet", "worksheet"),
  canonical(0x00a1, "Page Setup", "sheet", "worksheet"),
  canonical(0x0007, "Cached Formula String", "sheet", "worksheet"),

  noop(0x005b, "File Sharing", "both", "file-sharing metadata is not projected"),
  noop(0x005c, "Write Access", "both", "author metadata is not projected"),
  noop(0x005e, "Uncalced", "both", "recalculation scope marker is not projected"),
  noop(0x000c, "Calculation Count", "both", "calculation metadata is not projected"),
  noop(0x000d, "Calculation Mode", "both", "calculation metadata is not projected"),
  noop(0x000e, "Precision", "both", "calculation metadata is not projected"),
  noop(0x000f, "Reference Mode", "both", "formula display metadata is not projected"),
  noop(0x0010, "Iteration Delta", "both", "calculation metadata is not projected"),
  noop(0x0011, "Iteration", "both", "calculation metadata is not projected"),
  noop(0x0012, "Workbook Protection", "both", "protection metadata is not projected"),
  noop(0x0013, "Password", "both", "password hash is not projected"),
  noop(0x0019, "Window Protection", "both", "protection metadata is not projected"),
  noop(0x0040, "Backup", "both", "backup metadata is not projected"),
  noop(0x008d, "Hidden Objects", "both", "legacy object visibility metadata is not projected"),
  noop(0x003d, "Window 1", "both", "workbook window state is not projected"),
  noop(0x0061, "Interface Header", "both", "transitional interface marker is not projected"),
  noop(0x00c1, "Menu Set", "both", "transitional menu metadata is not projected"),
  noop(0x008c, "Country", "both", "locale metadata is not projected"),
  noop(0x0293, "Style", "both", "built-in style metadata is not projected"),
  noop(0x01da, "Workbook Boolean", "both", "transitional workbook flag is not projected"),
  noop(0x0161, "Double Stream File", "both", "format-version marker is not projected"),
  noop(0x01af, "Revision Protection", "both", "protection metadata is not projected"),
  noop(0x01bc, "Revision Protection Password", "both", "password hash is not projected"),
  noop(0x005f, "Save Recalculation", "both", "calculation metadata is not projected"),
  noop(0x0099, "Standard Width", "both", "legacy default-column metadata is not projected"),
  noop(0x0080, "Gutters", "both", "row and column gutter metadata is not projected"),
  noop(0x0225, "Default Row Height", "both", "default row formatting is not projected"),
  noop(0x023e, "Window 2", "both", "sheet window state is not projected"),
  noop(0x001d, "Selection", "both", "editor selection state is not projected"),
  noop(0x00a4, "Sheet Properties", "both", "non-visual sheet metadata is not projected"),
  noop(0x0041, "Pane", "both", "editor split state is not projected"),
]);

export function biffRecordContext(
  record: BiffRecord,
  reason: string,
  extra: { readonly property?: string } = {},
) {
  const entry = BIFF_RECORD_REGISTRY.get(record.code);
  return {
    part: "stream" as const,
    recordType: record.code,
    recordName: entry?.name ?? `0x${record.code.toString(16).padStart(4, "0")}`,
    offset: record.start,
    length: record.body.byteLength,
    byteRange: [record.start + 4, record.start + 4 + record.body.byteLength] as const,
    reason,
    ...extra,
  };
}

export function assertRegisteredBiffRecord(record: BiffRecord): BiffRegistryEntry | undefined {
  const entry = BIFF_RECORD_REGISTRY.get(record.code);
  if (!entry) {
    // Real-world workbooks (all of which Excel opens) carry records outside
    // any public registry. Skipping with a structured warning keeps them
    // readable; explicit registry entries still control strict rejection.
    console.warn(
      `xls parse diagnostic: skipped unknown record (${JSON.stringify({
        ...biffRecordContext(record, "unknown-record"),
      })})`,
    );
    return undefined;
  }
  if (entry.disposition.kind === "unsupported") {
    throw new XlsParseError(
      `Unsupported legacy XLS record ${entry.name}`,
      biffRecordContext(record, "unsupported-record"),
    );
  }
  return entry;
}
