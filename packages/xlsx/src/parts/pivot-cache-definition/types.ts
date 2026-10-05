import type {
  CacheFieldExtraAttrs,
  CacheHierarchyOptions,
  ConsolidationOptions,
  FieldGroupOptions,
  KpiOptions,
  MeasureDimensionMapOptions,
  MeasureGroupOptions,
  OLAPPropertiesOptions,
  PivotDimensionOptions,
  QueryCacheEntryOptions,
  ServerFormatOptions,
  SetOptions,
  TupleCacheEntryOptions,
} from "../pivot/pivot-utils";
import type { CalculatedItemOptions, CalculatedMemberOptions } from "../pivot/pivot-utils";
import type { PivotXmlElementOptions } from "../pivot/pivot-xml";

export interface PivotCacheDefinitionAttributes {
  /** Cache records relationship id (`@r:id`). */
  recordsRelationshipId?: string;
  /** XML attribute @invalid. */
  invalid?: boolean;
  /** XML attribute @saveData. */
  saveData?: boolean;
  /** XML attribute @refreshOnLoad. */
  refreshOnLoad?: boolean;
  /** XML attribute @optimizeMemory. */
  optimizeMemory?: boolean;
  /** XML attribute @enableRefresh. */
  enableRefresh?: boolean;
  /** XML attribute @refreshedBy. */
  refreshedBy?: string;
  /** XML attribute @refreshedDate. */
  refreshedDate?: number;
  /** XML attribute @refreshedDateIso. */
  refreshedDateIso?: string;
  /** XML attribute @backgroundQuery. */
  backgroundQuery?: boolean;
  /** XML attribute @missingItemsLimit. */
  missingItemsLimit?: number;
  /** XML attribute @createdVersion. */
  createdVersion?: number;
  /** XML attribute @refreshedVersion. */
  refreshedVersion?: number;
  /** XML attribute @minRefreshableVersion. */
  minRefreshableVersion?: number;
  /** XML attribute @recordCount. */
  recordCount?: number;
  /** XML attribute @upgradeOnRefresh. */
  upgradeOnRefresh?: boolean;
  /** XML attribute @tupleCache. */
  tupleCache?: boolean;
  /** XML attribute @supportSubquery. */
  supportSubquery?: boolean;
  /** XML attribute @supportAdvancedDrill. */
  supportAdvancedDrill?: boolean;
}

export interface PivotSharedItemsAttributes {
  /** XML attribute @containsSemiMixedTypes. */
  containsSemiMixedTypes?: boolean;
  /** XML attribute @containsNonDate. */
  containsNonDate?: boolean;
  /** XML attribute @containsDate. */
  containsDate?: boolean;
  /** XML attribute @containsString. */
  containsString?: boolean;
  /** XML attribute @containsBlank. */
  containsBlank?: boolean;
  /** XML attribute @containsMixedTypes. */
  containsMixedTypes?: boolean;
  /** XML attribute @containsNumber. */
  containsNumber?: boolean;
  /** XML attribute @containsInteger. */
  containsInteger?: boolean;
  /** XML attribute @minValue. */
  minValue?: number;
  /** XML attribute @maxValue. */
  maxValue?: number;
  /** XML attribute @minDate. */
  minDate?: string;
  /** XML attribute @maxDate. */
  maxDate?: string;
  /** XML attribute @count. */
  count?: number;
  /** XML attribute @longText. */
  longText?: boolean;
}

export interface PivotCacheFieldAttributes {
  /** XML attribute @name. */
  name?: string;
  /** XML attribute @caption. */
  caption?: string;
  /** XML attribute @propertyName. */
  propertyName?: string;
  /** XML attribute @serverField. */
  serverField?: boolean;
  /** XML attribute @uniqueList. */
  uniqueList?: boolean;
  /** XML attribute @numFmtId. */
  numFmtId?: number;
  /** XML attribute @formula. */
  formula?: string;
  /** XML attribute @sqlType. */
  sqlType?: number;
  /** XML attribute @hierarchy. */
  hierarchy?: number;
  /** XML attribute @level. */
  level?: number;
  /** XML attribute @databaseField. */
  databaseField?: boolean;
  /** XML attribute @mappingCount. */
  mappingCount?: number;
  /** XML attribute @memberPropertyField. */
  memberPropertyField?: boolean;
}

export interface PivotCacheFieldOptions extends PivotCacheFieldAttributes {
  /** Typed sharedItems attributes; child values stay on sharedItems. */
  sharedItemsAttributes?: PivotSharedItemsAttributes;
  /** Serialized sharedItems child (CT_SharedItems). */
  sharedItems?: PivotXmlElementOptions;
  /** Serialized fieldGroup child (CT_FieldGroup). */
  fieldGroup?: PivotXmlElementOptions;
  /** Serialized mpMap children (CT_X), in document order. */
  mpMaps?: PivotXmlElementOptions[];
  /** Serialized extLst child (CT_ExtensionList). */
  extLst?: PivotXmlElementOptions;
}

export interface PivotCacheFieldsOptions {
  /** XML attribute @count. */
  count?: number;
  /** Cache fields (CT_CacheField), in document order. */
  fields: PivotCacheFieldOptions[];
}

export interface PivotCacheDefinitionOptions extends PivotCacheDefinitionAttributes {
  /** Exact source XML captured by parse — round-trip only, do not hand-author. */
  source?: string;
  /** Source records relationship id from .rels — round-trip only. */
  recordsRelationshipSourceId?: string;
  /** Source records relationship target — filled by parse, round-trip only. */
  recordsRelationshipTarget?: string;
  /** Source .rels relationship ids in document order — round-trip only. */
  relationshipOrder?: string[];
  /** Additional source .rels relationships, excluding records — round-trip only. */
  externalRelationships?: {
    id: string;
    type: string;
    target: string;
    targetMode?: "External";
  }[];
  /** Serialized cacheSource child (CT_CacheSource). */
  cacheSource?: PivotXmlElementOptions;
  /** Parsed cacheFields container (CT_CacheFields). */
  cacheFields?: PivotCacheFieldsOptions;
  /** Serialized cacheHierarchies child (CT_CacheHierarchies). */
  cacheHierarchies?: CacheHierarchyOptions[];
  /** Serialized kpis child (CT_PCDKPIs). */
  kpis?: KpiOptions[];
  /** Serialized tupleCache child (CT_TupleCache). */
  tupleCacheData?: PivotXmlElementOptions;
  /** Calculated items (CT_CalculatedItems). */
  calculatedItems?: CalculatedItemOptions[];
  /** Calculated members (CT_CalculatedMembers). */
  calculatedMembers?: CalculatedMemberOptions[];
  /** Serialized dimensions child (CT_Dimensions). */
  dimensions?: PivotDimensionOptions[];
  /** Serialized measureGroups child (CT_MeasureGroups). */
  measureGroups?: MeasureGroupOptions[];
  /** Serialized maps child (CT_MeasureDimensionMaps). */
  maps?: PivotXmlElementOptions;
  /** Serialized extLst child (CT_ExtensionList). */
  extLst?: PivotXmlElementOptions;
  /** Legacy OLAP properties merged into cacheSource/extLst on source output. */
  olap?: OLAPPropertiesOptions;
  /** Consolidation source (alternative to worksheetSource). */
  consolidation?: ConsolidationOptions;
  /** Tuple-cache entries (CT_PCDSDTCEntries). */
  entries?: TupleCacheEntryOptions[];
  /** Tuple-cache sets (CT_Sets). */
  sets?: SetOptions[];
  /** Tuple-cache server formats (CT_ServerFormats). */
  serverFormats?: ServerFormatOptions[];
  /** Tuple-cache query entries (CT_QueryCache). */
  queryCache?: QueryCacheEntryOptions[];
  /** Field groups keyed by source field index. */
  fieldGroups?: ReadonlyMap<number, FieldGroupOptions>;
  /** Member-property maps (mpMap). */
  mpMaps?: { fieldIndex: number; x: number }[];
  /** Measure-dimension maps (CT_MeasureDimensionMaps). */
  measureDimensionMaps?: MeasureDimensionMapOptions[];
  /** Additional cache-field attributes keyed by source field index. */
  cacheFieldOverrides?: ReadonlyMap<number, CacheFieldExtraAttrs>;
}
