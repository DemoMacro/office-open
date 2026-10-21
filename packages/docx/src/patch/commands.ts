import type {
  PatchNestedValue,
  PatchField,
  PatchNodeAddress,
  PatchAddress,
  PatchPosition,
  PatchTextAddress,
  PatchPropertyPath,
  PatchOptions,
} from "@office-open/core";
import type { OoxmlPackageVariant, OutputType } from "@office-open/core";
import type { DocumentOptions } from "@parts/core-properties";

/** Canonical root fields addressable by a DOCX patch command. */
export type DocumentPatchField = PatchField<DocumentOptions>;

type DocumentPatchCollectionField = {
  [Field in DocumentPatchField]: NonNullable<DocumentOptions[Field]> extends readonly unknown[]
    ? Field
    : never;
}[DocumentPatchField];

/** Canonical item that can be inserted into a DOCX collection field. */
export type DocumentPatchInsertValue = PatchNestedValue<
  DocumentOptions[DocumentPatchCollectionField]
>;

/** Canonical value assignable to an addressed DOCX root or nested node. */
export type DocumentPatchValue =
  | DocumentOptions[DocumentPatchField]
  | PatchNestedValue<DocumentOptions[DocumentPatchField]>;

/** Canonical partial value accepted by a DOCX update command. */
export type DocumentPatchUpdateValue = DocumentPatchValue | DocumentPatchInsertValue;

/** One ordered canonical DOCX edit. */
export type DocumentPatchCommand =
  | {
      type: "insert";
      at: PatchAddress<DocumentPatchField>;
      position?: PatchPosition;
      value: readonly DocumentPatchInsertValue[];
    }
  | { type: "replace"; at: PatchNodeAddress<DocumentPatchField>; value: DocumentPatchValue }
  | {
      type: "replace";
      at: PatchTextAddress<DocumentPatchField>;
      value: string;
    }
  | {
      type: "update";
      at: PatchNodeAddress<DocumentPatchField>;
      set?: DocumentPatchUpdateValue;
      unset?: PatchPropertyPath;
    }
  | { type: "remove"; at: PatchNodeAddress<DocumentPatchField> }
  | { type: "remove"; at: PatchTextAddress<DocumentPatchField> }
  | { type: "move"; at: PatchNodeAddress<DocumentPatchField>; to: PatchPosition };

/**
 * Options for patching an existing DOCX package through its canonical model.
 *
 * @property packageVariant - Output package variant; required when patching a macro/template source into the same variant.
 */
export interface DocumentPatchOptions<
  T extends OutputType = "uint8array",
> extends PatchOptions<DocumentPatchCommand> {
  outputType?: T;
  packageVariant?: OoxmlPackageVariant;
}
