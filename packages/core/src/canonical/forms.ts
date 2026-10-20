export type FormTriState = "unchecked" | "checked" | "unknown";
export type FormListSourceType =
  | "table"
  | "query"
  | "sql"
  | "sqlPassThrough"
  | "valueList"
  | "tableFields";
export type FormButtonType = "submit" | "reset" | "push" | "url";
export type FormNavigationMode = "none" | "current" | "parent";
export type FormTabCycle = "records" | "current" | "page";
export type FormCommandType = "table" | "query" | "command";
export type FormListLinkageType = "selection" | "selectionIndices";
export type FormVisualEffect = "flat" | "3d";
export type FormImagePosition = "start" | "center" | "end" | "top" | "bottom";
export type FormImageAlign = "start" | "center" | "end";
export type FormOrientation = "horizontal" | "vertical";

export interface FormFloatValueOptions {
  type: "float";
  value: number;
}
export interface FormPercentageValueOptions {
  type: "percentage";
  value: number;
}
export interface FormCurrencyValueOptions {
  type: "currency";
  value: number;
  currency?: string;
}
export interface FormDateValueOptions {
  type: "date";
  value: string;
}
export interface FormTimeValueOptions {
  type: "time";
  value: string;
}
export interface FormBooleanValueOptions {
  type: "boolean";
  value: boolean;
}
export interface FormStringValueOptions {
  type: "string";
  value?: string;
}
export interface FormVoidValueOptions {
  type: "void";
}
export type FormPropertyValueOptions =
  | FormFloatValueOptions
  | FormPercentageValueOptions
  | FormCurrencyValueOptions
  | FormDateValueOptions
  | FormTimeValueOptions
  | FormBooleanValueOptions
  | FormStringValueOptions
  | FormVoidValueOptions;

export interface FormPropertyOptions {
  name: string;
  value: FormPropertyValueOptions;
}

export type FormFloatListValueOptions = { type: "float"; values: number[] };
export type FormPercentageListValueOptions = { type: "percentage"; values: number[] };
export type FormCurrencyListValueOptions = {
  type: "currency";
  currency?: string;
  values: number[];
};
export type FormDateListValueOptions = { type: "date"; values: string[] };
export type FormTimeListValueOptions = { type: "time"; values: string[] };
export type FormBooleanListValueOptions = { type: "boolean"; values: boolean[] };
export type FormStringListValueOptions = { type: "string"; values: string[] };
export type FormListPropertyValueOptions =
  | FormFloatListValueOptions
  | FormPercentageListValueOptions
  | FormCurrencyListValueOptions
  | FormDateListValueOptions
  | FormTimeListValueOptions
  | FormBooleanListValueOptions
  | FormStringListValueOptions;

export interface FormListPropertyOptions {
  name: string;
  value: FormListPropertyValueOptions;
}

export interface FormMacroEventListenerOptions {
  target: "macro";
  eventName: string;
  language: string;
  macroName: string;
}
export interface FormUrlEventListenerOptions {
  target: "url";
  eventName: string;
  language: string;
  href: string;
  actuate?: "onRequest";
}
export type FormEventListenerOptions = FormMacroEventListenerOptions | FormUrlEventListenerOptions;

export interface XFormsElementOptions {
  name: string;
  attributes: Record<string, string>;
  children?: Array<XFormsElementOptions | string>;
}

export interface XFormsModelOptions {
  attributes: Record<string, string>;
  children?: XFormsElementOptions[];
}

export interface FormControlBaseOptions {
  xmlId?: string;
  id?: string;
  name?: string;
  controlImplementation?: string;
  bind?: string;
  inputRequired?: boolean;
  properties?: Array<FormPropertyOptions | FormListPropertyOptions>;
  events?: FormEventListenerOptions[];
}

export interface FormInputControlOptions extends FormControlBaseOptions {
  currentValue?: string;
  disabled?: boolean;
  maxLength?: number;
  printable?: boolean;
  readonly?: boolean;
  tabIndex?: number;
  tabStop?: boolean;
  title?: string;
  value?: string;
  convertEmptyToNull?: boolean;
  dataField?: string;
  linkedCell?: string;
}

export interface TextControlOptions extends FormInputControlOptions {
  control: "text";
}
export interface TextAreaControlOptions extends FormInputControlOptions {
  control: "textArea";
  paragraphs?: string[];
}
export interface FormattedTextControlOptions extends FormInputControlOptions {
  control: "formattedText";
  spinButton?: boolean;
  repeat?: boolean;
  delayForRepeat?: string;
  minValue?: string;
  maxValue?: string;
  validation?: boolean;
}
export interface NumberControlOptions extends FormControlBaseOptions {
  control: "number";
  value?: number;
  currentValue?: number;
  minValue?: number;
  maxValue?: number;
  disabled?: boolean;
  maxLength?: number;
  printable?: boolean;
  readonly?: boolean;
  tabIndex?: number;
  tabStop?: boolean;
  title?: string;
  convertEmptyToNull?: boolean;
  dataField?: string;
  linkedCell?: string;
  spinButton?: boolean;
  repeat?: boolean;
  delayForRepeat?: string;
  validation?: boolean;
}
export interface FormTemporalControlOptions extends Omit<
  NumberControlOptions,
  "control" | "value" | "currentValue" | "minValue" | "maxValue"
> {
  value?: string;
  currentValue?: string;
  minValue?: string;
  maxValue?: string;
}
export interface DateControlOptions extends FormTemporalControlOptions {
  control: "date";
}
export interface TimeControlOptions extends FormTemporalControlOptions {
  control: "time";
}
export interface ComboBoxItemOptions {
  label?: string;
  labelAttribute?: string;
}
export interface ComboBoxControlOptions extends FormInputControlOptions {
  control: "comboBox";
  dropdown?: boolean;
  size?: number;
  listSource?: string;
  listSourceType?: FormListSourceType;
  xformsListSource?: string;
  sourceCellRange?: string;
  autoComplete?: boolean;
  items: ComboBoxItemOptions[];
}
export interface ListBoxOptionOptions {
  currentSelected?: boolean;
  selected?: boolean;
  labelAttribute?: string;
  label?: string;
  value?: string;
}
export interface ListBoxControlOptions extends FormControlBaseOptions {
  control: "listBox";
  disabled?: boolean;
  dropdown?: boolean;
  printable?: boolean;
  size?: number;
  tabIndex?: number;
  tabStop?: boolean;
  title?: string;
  boundColumn?: string;
  dataField?: string;
  listSource?: string;
  listSourceType?: FormListSourceType;
  linkedCell?: string;
  listLinkageType?: FormListLinkageType;
  sourceCellRange?: string;
  multiple?: boolean;
  xformsListSource?: string;
  options: ListBoxOptionOptions[];
}
export interface CheckboxControlOptions extends FormControlBaseOptions {
  control: "checkbox";
  disabled?: boolean;
  label?: string;
  printable?: boolean;
  tabIndex?: number;
  tabStop?: boolean;
  title?: string;
  value?: string;
  dataField?: string;
  visualEffect?: FormVisualEffect;
  imagePosition?: FormImagePosition;
  imageAlign?: FormImageAlign;
  linkedCell?: string;
  currentState?: FormTriState;
  triState?: boolean;
  state?: FormTriState;
}
export interface PasswordControlOptions extends FormControlBaseOptions {
  control: "password";
  disabled?: boolean;
  maxLength?: number;
  printable?: boolean;
  tabIndex?: number;
  tabStop?: boolean;
  title?: string;
  value?: string;
  convertEmptyToNull?: boolean;
  linkedCell?: string;
  echoChar?: string;
}
export type FileControlOptions = FormInputControlOptions & { control: "file" };
export interface FixedTextControlOptions extends FormControlBaseOptions {
  control: "fixedText";
  for?: string;
  disabled?: boolean;
  label?: string;
  printable?: boolean;
  title?: string;
  multiLine?: boolean;
}
export interface ButtonControlOptions extends FormControlBaseOptions {
  control: "button";
  buttonType?: FormButtonType;
  disabled?: boolean;
  label?: string;
  imageData?: string;
  printable?: boolean;
  tabIndex?: number;
  tabStop?: boolean;
  targetFrame?: string;
  href?: string;
  title?: string;
  value?: string;
  imagePosition?: FormImagePosition;
  imageAlign?: FormImageAlign;
  repeat?: boolean;
  delayForRepeat?: string;
  defaultButton?: boolean;
  toggle?: boolean;
  focusOnClick?: boolean;
  xformsSubmission?: string;
}
export interface ImageControlOptions extends FormControlBaseOptions {
  control: "image";
  buttonType?: FormButtonType;
  disabled?: boolean;
  imageData?: string;
  printable?: boolean;
  tabIndex?: number;
  tabStop?: boolean;
  targetFrame?: string;
  href?: string;
  title?: string;
  value?: string;
}
export interface RadioControlOptions extends FormControlBaseOptions {
  control: "radio";
  groupName?: string;
  currentSelected?: boolean;
  disabled?: boolean;
  label?: string;
  printable?: boolean;
  selected?: boolean;
  tabIndex?: number;
  tabStop?: boolean;
  title?: string;
  value?: string;
  dataField?: string;
  visualEffect?: FormVisualEffect;
  imagePosition?: FormImagePosition;
  imageAlign?: FormImageAlign;
  linkedCell?: string;
}
export interface FrameControlOptions extends FormControlBaseOptions {
  control: "frame";
  disabled?: boolean;
  for?: string;
  label?: string;
  printable?: boolean;
  title?: string;
}
export interface ImageFrameControlOptions extends FormControlBaseOptions {
  control: "imageFrame";
  label?: string;
  disabled?: boolean;
  imageData?: string;
  printable?: boolean;
  readonly?: boolean;
  title?: string;
  dataField?: string;
}
export interface HiddenControlOptions extends FormControlBaseOptions {
  control: "hidden";
  value?: string;
}
export type GridFormControlOptions =
  | TextControlOptions
  | TextAreaControlOptions
  | FormattedTextControlOptions
  | NumberControlOptions
  | DateControlOptions
  | TimeControlOptions
  | ComboBoxControlOptions
  | ListBoxControlOptions
  | CheckboxControlOptions;
export interface GridColumnOptions {
  name?: string;
  controlImplementation?: string;
  label?: string;
  textStyleName?: string;
  controls: GridFormControlOptions[];
}
export interface GridControlOptions extends FormControlBaseOptions {
  control: "grid";
  disabled?: boolean;
  printable?: boolean;
  tabIndex?: number;
  tabStop?: boolean;
  title?: string;
  columns: GridColumnOptions[];
}
export interface ValueRangeControlOptions extends FormControlBaseOptions {
  control: "valueRange";
  disabled?: boolean;
  printable?: boolean;
  tabIndex?: number;
  tabStop?: boolean;
  title?: string;
  value?: string;
  linkedCell?: string;
  repeat?: boolean;
  delayForRepeat?: string;
  maxValue?: number;
  minValue?: number;
  stepSize?: number;
  pageStepSize?: number;
  orientation?: FormOrientation;
}
export interface GenericControlOptions extends FormControlBaseOptions {
  control: "genericControl";
}

export type FormControlOptions =
  | GridFormControlOptions
  | PasswordControlOptions
  | FileControlOptions
  | FixedTextControlOptions
  | ButtonControlOptions
  | ImageControlOptions
  | RadioControlOptions
  | FrameControlOptions
  | ImageFrameControlOptions
  | HiddenControlOptions
  | GridControlOptions
  | ValueRangeControlOptions
  | GenericControlOptions;

export interface FormConnectionResourceOptions {
  href: string;
}
export interface FormContainerOptions {
  name?: string;
  controlImplementation?: string;
  href?: string;
  targetFrame?: string;
  method?: string;
  enctype?: string;
  allowDeletes?: boolean;
  allowInserts?: boolean;
  allowUpdates?: boolean;
  applyFilter?: boolean;
  commandType?: FormCommandType;
  command?: string;
  datasource?: string;
  masterFields?: string;
  detailFields?: string;
  escapeProcessing?: boolean;
  filter?: string;
  ignoreResult?: boolean;
  navigationMode?: FormNavigationMode;
  order?: string;
  tabCycle?: FormTabCycle;
  properties?: Array<FormPropertyOptions | FormListPropertyOptions>;
  events?: FormEventListenerOptions[];
  controls: FormControlOptions[];
  forms?: FormContainerOptions[];
  connectionResource?: FormConnectionResourceOptions;
}

export interface OfficeFormsOptions {
  automaticFocus?: boolean;
  designMode?: boolean;
  containers: FormContainerOptions[];
  models?: XFormsModelOptions[];
}
