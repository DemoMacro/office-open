/** Common identity and interaction properties for a form control. */
export interface FormControlCommonOptions {
  /** Stable control identifier. */
  id?: string;
  /** User-visible control name. */
  name?: string;
  /** Whether keyboard interaction is disabled. */
  disabled?: boolean;
  /** Tab order position. */
  tabIndex?: number;
  /** Whether the container requests initial keyboard focus. */
  automaticFocus?: boolean;
  /** Whether controls are shown in design mode rather than active mode. */
  designMode?: boolean;
}

/** Text-entry form control. */
export interface TextControlOptions extends FormControlCommonOptions {
  /** Text-entry control discriminant. */
  control: "text";
  /** Current text value. */
  value?: string;
  /** Maximum accepted character count. */
  maxLength?: number;
  /** Declared stored value type for the text control. */
  valueType?: "float" | "string" | "date" | "time";
}

/** Checkbox form control. */
export interface CheckBoxControlOptions extends FormControlCommonOptions {
  /** Checkbox control discriminant. */
  control: "checkBox";
  /** Current checked state. */
  checked?: boolean;
}

/** Dropdown form control. */
export interface DropDownListControlOptions extends FormControlCommonOptions {
  /** Dropdown control discriminant. */
  control: "dropDownList";
  /** Selectable labels in display order. */
  entries: string[];
  /** Zero-based selected entry index. */
  selectedIndex?: number;
}

export type FormControlOptions =
  | TextControlOptions
  | CheckBoxControlOptions
  | DropDownListControlOptions;

/** A named container holding one or more form controls. */
export interface FormContainerOptions {
  /** Container name shown to form consumers. */
  name?: string;
  /** Whether the container requests initial keyboard focus. */
  automaticFocus?: boolean;
  /** Whether controls are shown in design mode rather than active mode. */
  designMode?: boolean;
  /** Controls contained by this form. */
  controls: FormControlOptions[];
}
