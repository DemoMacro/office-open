import type { PivotXmlElementOptions } from "../pivot/pivot-xml";

export interface PivotCacheRecordEntryOptions {
  /** Cell metadata attributes shared by cached values. */
  unformatted?: boolean;
  formula?: boolean;
  caption?: string;
  captionLength?: number;
  /** Serialized tpls children (CT_Tuples), in document order. */
  tuples?: PivotXmlElementOptions[];
  /** Serialized x children (CT_X), in document order. */
  indexes?: PivotXmlElementOptions[];
}

export interface PivotCacheRecordValue extends PivotCacheRecordEntryOptions {
  /** Cached value; number for n, string for other value types. */
  value?: string | number | boolean;
}

export interface MissingPivotCacheRecord extends PivotCacheRecordValue {
  /** Value kind discriminator: Missing value (CT_Missing). */
  type: "missing";
}

export interface NumberPivotCacheRecord extends PivotCacheRecordValue {
  /** Value kind discriminator: Numeric value (CT_Number). */
  type: "number";
}

export interface BooleanPivotCacheRecord extends PivotCacheRecordValue {
  /** Value kind discriminator: Boolean value (CT_Boolean). */
  type: "boolean";
}

export interface ErrorPivotCacheRecord extends PivotCacheRecordValue {
  /** Value kind discriminator: Cached formula error (CT_Error). */
  type: "error";
}

export interface StringPivotCacheRecord extends PivotCacheRecordValue {
  /** Value kind discriminator: String value (CT_String). */
  type: "string";
}

export interface DatePivotCacheRecord extends PivotCacheRecordValue {
  /** Value kind discriminator: ISO date-time value (CT_DateTime). */
  type: "date";
}

export interface SharedItemPivotCacheRecord extends PivotCacheRecordValue {
  /** Value kind discriminator: Shared-items index (CT_Index). */
  type: "sharedItem";
}

export type PivotCacheRecordEntry =
  | MissingPivotCacheRecord
  | NumberPivotCacheRecord
  | BooleanPivotCacheRecord
  | ErrorPivotCacheRecord
  | StringPivotCacheRecord
  | DatePivotCacheRecord
  | SharedItemPivotCacheRecord;

export interface PivotCacheRecordOptions {
  /** Record-level extension list (CT_ExtensionList). */
  extLst?: PivotXmlElementOptions;
  /** Cached cells (CT_Record choice), in document order. */
  entries: PivotCacheRecordEntry[];
}

export interface PivotCacheRecordsOptions {
  /** Exact source XML captured by parse — round-trip only, do not hand-author. */
  source?: string;
  /** XML attribute @count (record count). */
  count?: number;
  /** Cached records (CT_Record), in document order. */
  records: PivotCacheRecordOptions[];
  /** Root extension list (CT_ExtensionList). */
  extLst?: PivotXmlElementOptions;
}
