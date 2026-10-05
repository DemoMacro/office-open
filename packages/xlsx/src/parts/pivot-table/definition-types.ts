import type { PivotXmlElementOptions } from "../pivot/pivot-xml";

export interface PivotTableDefinitionAttributes {
  /** XML attribute @name. */
  name?: string;
  /** XML attribute @cacheId. */
  cacheId?: number;
  /** XML attribute @dataOnRows. */
  dataOnRows?: boolean;
  /** XML attribute @dataPosition. */
  dataPosition?: number;
  /** XML attribute @dataCaption. */
  dataCaption?: string;
  /** XML attribute @grandTotalCaption. */
  grandTotalCaption?: string;
  /** XML attribute @errorCaption. */
  errorCaption?: string;
  /** XML attribute @showError. */
  showError?: boolean;
  /** XML attribute @missingCaption. */
  missingCaption?: string;
  /** XML attribute @showMissing. */
  showMissing?: boolean;
  /** XML attribute @pageStyle. */
  pageStyle?: string;
  /** XML attribute @pivotTableStyle. */
  pivotTableStyle?: string;
  /** XML attribute @vacatedStyle. */
  vacatedStyle?: string;
  /** XML attribute @tag. */
  tag?: string;
  /** XML attribute @updatedVersion. */
  updatedVersion?: number;
  /** XML attribute @minRefreshableVersion. */
  minRefreshableVersion?: number;
  /** XML attribute @asteriskTotals. */
  asteriskTotals?: boolean;
  /** XML attribute @showItems. */
  showItems?: boolean;
  /** XML attribute @editData. */
  editData?: boolean;
  /** XML attribute @disableFieldList. */
  disableFieldList?: boolean;
  /** XML attribute @showCalcMbrs. */
  showCalcMbrs?: boolean;
  /** XML attribute @visualTotals. */
  visualTotals?: boolean;
  /** XML attribute @showMultipleLabel. */
  showMultipleLabel?: boolean;
  /** XML attribute @showDataDropDown. */
  showDataDropDown?: boolean;
  /** XML attribute @showDrill. */
  showDrill?: boolean;
  /** XML attribute @printDrill. */
  printDrill?: boolean;
  /** XML attribute @showMemberPropertyTips. */
  showMemberPropertyTips?: boolean;
  /** XML attribute @showDataTips. */
  showDataTips?: boolean;
  /** XML attribute @enableWizard. */
  enableWizard?: boolean;
  /** XML attribute @enableDrill. */
  enableDrill?: boolean;
  /** XML attribute @enableFieldProperties. */
  enableFieldProperties?: boolean;
  /** XML attribute @preserveFormatting. */
  preserveFormatting?: boolean;
  /** XML attribute @useAutoFormatting. */
  useAutoFormatting?: boolean;
  /** XML attribute @pageWrap. */
  pageWrap?: number;
  /** XML attribute @pageOverThenDown. */
  pageOverThenDown?: boolean;
  /** XML attribute @subtotalHiddenItems. */
  subtotalHiddenItems?: boolean;
  /** XML attribute @rowGrandTotals. */
  rowGrandTotals?: boolean;
  /** XML attribute @colGrandTotals. */
  colGrandTotals?: boolean;
  /** XML attribute @fieldPrintTitles. */
  fieldPrintTitles?: boolean;
  /** XML attribute @itemPrintTitles. */
  itemPrintTitles?: boolean;
  /** XML attribute @mergeItem. */
  mergeItem?: boolean;
  /** XML attribute @showDropZones. */
  showDropZones?: boolean;
  /** XML attribute @createdVersion. */
  createdVersion?: number;
  /** XML attribute @indent. */
  indent?: number;
  /** XML attribute @showEmptyRow. */
  showEmptyRow?: boolean;
  /** XML attribute @showEmptyCol. */
  showEmptyCol?: boolean;
  /** XML attribute @showHeaders. */
  showHeaders?: boolean;
  /** XML attribute @compact. */
  compact?: boolean;
  /** XML attribute @outline. */
  outline?: boolean;
  /** XML attribute @outlineData. */
  outlineData?: boolean;
  /** XML attribute @compactData. */
  compactData?: boolean;
  /** XML attribute @published. */
  published?: boolean;
  /** XML attribute @gridDropZones. */
  gridDropZones?: boolean;
  /** XML attribute @immersive. */
  immersive?: boolean;
  /** XML attribute @multipleFieldFilters. */
  multipleFieldFilters?: boolean;
  /** XML attribute @chartFormat. */
  chartFormat?: number;
  /** XML attribute @rowHeaderCaption. */
  rowHeaderCaption?: string;
  /** XML attribute @colHeaderCaption. */
  colHeaderCaption?: string;
  /** XML attribute @fieldListSortAscending. */
  fieldListSortAscending?: boolean;
  /** XML attribute @mdxSubqueries. */
  mdxSubqueries?: boolean;
  /** XML attribute @customListSort. */
  customListSort?: boolean;
}

export interface PivotTableDefinitionOptions extends PivotTableDefinitionAttributes {
  /** Exact source XML captured by parse — round-trip only, do not hand-author. */
  source?: string;
  /** Source pivotTable part path — filled by parse, round-trip only. */
  sourcePath?: string;
  /** Source cache relationship id — filled by parse, round-trip only. */
  cacheRelationshipId?: string;
  /** Source cache relationship target — filled by parse, round-trip only. */
  cacheRelationshipTarget?: string;
  /** Source cache relationship type — filled by parse, round-trip only. */
  cacheRelationshipType?: string;
  /** Serialized location child element (complete CT subtree). */
  location?: PivotXmlElementOptions;
  /** Serialized pivotFields child element (complete CT subtree). */
  pivotFields?: PivotXmlElementOptions;
  /** Serialized rowFields child element (complete CT subtree). */
  rowFields?: PivotXmlElementOptions;
  /** Serialized rowItems child element (complete CT subtree). */
  rowItems?: PivotXmlElementOptions;
  /** Serialized colFields child element (complete CT subtree). */
  colFields?: PivotXmlElementOptions;
  /** Serialized colItems child element (complete CT subtree). */
  colItems?: PivotXmlElementOptions;
  /** Serialized pageFields child element (complete CT subtree). */
  pageFields?: PivotXmlElementOptions;
  /** Serialized dataFields child element (complete CT subtree). */
  dataFields?: PivotXmlElementOptions;
  /** Serialized formats child element (complete CT subtree). */
  formats?: PivotXmlElementOptions;
  /** Serialized conditionalFormats child element (complete CT subtree). */
  conditionalFormats?: PivotXmlElementOptions;
  /** Serialized chartFormats child element (complete CT subtree). */
  chartFormats?: PivotXmlElementOptions;
  /** Serialized pivotHierarchies child element (complete CT subtree). */
  pivotHierarchies?: PivotXmlElementOptions;
  /** Serialized pivotTableStyleInfo child element (complete CT subtree). */
  pivotTableStyleInfo?: PivotXmlElementOptions;
  /** Serialized filters child element (complete CT subtree). */
  filters?: PivotXmlElementOptions;
  /** Serialized rowHierarchiesUsage child element (complete CT subtree). */
  rowHierarchiesUsage?: PivotXmlElementOptions;
  /** Serialized colHierarchiesUsage child element (complete CT subtree). */
  colHierarchiesUsage?: PivotXmlElementOptions;
  /** Serialized extLst child element (complete CT subtree). */
  extLst?: PivotXmlElementOptions;
}
