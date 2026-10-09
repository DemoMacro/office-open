/**
 * XLSX Compiler — compiles WorkbookOptions into a Zippable structure.
 *
 * Accepts pure JSON WorkbookOptions — no intermediate File class needed.
 * Uses XlsxWriteContext for shared state (strings, styles, media, charts).
 *
 * @module
 */

import {
  RELATIONSHIP_TYPES,
  Relationships,
  TargetModeType,
  appPropertiesDesc,
  buildCorePropertiesXmlString,
  buildRootRelationships,
  compileMapping,
  dropDanglingPassthroughRels,
  finalizeContentTypes,
  type OoxmlPackageVariant,
  type PassthroughRelationship,
  type RelationshipType,
  type ReproducibleScope,
  partPathToRelsPath,
  resolveRelationshipTarget,
  ooxmlPackageFormatInfo,
  resolveMainPartContentType,
  resolverFromRegistry,
  XLSX_PARTS,
  customPropertiesDesc,
  toUint8Array,
  IMAGE_MEDIA_CONTENT_TYPES,
  type XmlifyedFile,
  type Zippable,
} from "@office-open/core";
import { buildUserShapesData, chartSpaceDesc } from "@office-open/core/chart";
import type { WriteContext } from "@office-open/core/descriptor";
import { buildThemeXml, themeOverrideDesc } from "@office-open/core/theme";
import { escapeXml, OOXML_XML_DECLARATION } from "@office-open/xml";
import { activeXControlDesc, type ActiveXControlOptions } from "@parts/active-x-control";
import type { CalcCell } from "@parts/calc-chain";
import { calcChainDesc } from "@parts/calc-chain";
import { chartsheetDesc, type ChartsheetOptions } from "@parts/chartsheet";
import { classificationLabelsDesc } from "@parts/classification-labels";
import { commentsDesc, vmlNotesDesc } from "@parts/comments";
import { connectionsDesc } from "@parts/connection";
import { controlPropertiesDesc } from "@parts/control-properties";
import { dialogsheetDesc, type DialogsheetOptions } from "@parts/dialogsheet";
import { A_NS, R_NS, XDR_NS, graphicFrameXml, wrapAnchor } from "@parts/drawing/stringify";
import { externalLinkDesc } from "@parts/external-link";
import type { WorkbookOptions } from "@parts/file";
import { metadataDesc } from "@parts/metadata";
import type { MetadataOptions } from "@parts/metadata";
import { personsDesc } from "@parts/persons";
import { queryTableDesc } from "@parts/query-table";
import { revisionHeadersDesc, revisionLogDesc, usersDesc } from "@parts/revision-log";
import {
  richValueDataDesc,
  richValueRelsDesc,
  richValueStructuresDesc,
  richValueTypesInfoDesc,
} from "@parts/rich-data";
import { sharedStringsDesc } from "@parts/shared-strings";
import { stylesDesc } from "@parts/styles";
import { tableDesc } from "@parts/table";
import { createThemeXml } from "@parts/theme";
import { buildVolTypesXml } from "@parts/vol-types";
import { webExtensionPartDesc } from "@parts/web-extension";
import type { TablePartReference, SheetDefinition } from "@parts/workbook";
import { workbookDesc, buildTablePartsXml, buildExternalReferencesXml } from "@parts/workbook";
import {
  buildWorksheetXml,
  editSheetTailMarker,
  stripWorksheetPlaceholders,
  type WorksheetContext,
} from "@parts/worksheet";
import type { WorksheetOptions } from "@parts/worksheet";
import { mapInfoDesc, singleXmlCellsDesc } from "@parts/xml-mapping";
import { columnToLetter } from "@util/index";

import {
  bindMediaPlaceholders,
  compileSheetDrawing,
  preserveChartDecimalAttributes,
} from "./compile/sheet-drawing";
import {
  compileDefinitionPivotCaches,
  compileSheetPivots,
  renderPivotSheetData,
} from "./compile/sheet-pivots";
import { XlsxWriteContext } from "./context";

const XML_DECL = OOXML_XML_DECLARATION;

const IMAGE_REL = RELATIONSHIP_TYPES.image;

/** XLSX part path → content type, derived from the part registry. Matches
 * actual file paths, so the dense/sequential xlsx part naming is handled. */
const XLSX_CONTENT_TYPE_RESOLVER = resolverFromRegistry(XLSX_PARTS);

/** Chart part → user-shapes part relationship (c:userShapes bridge). */
const CHART_USER_SHAPES_REL = RELATIONSHIP_TYPES.chartUserShapes;
const VOLATILE_DEPENDENCIES_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/volatileDependencies" as RelationshipType;
const LEGACY_RELATIONSHIPS_NAMESPACE = "http://schemas.microsoft.com/package/2005/06/relationships";
const LEGACY_WORKBOOK_RELATIONSHIP_TYPES = new Map<string, string>([
  [
    "http://schemas.microsoft.com/office/2006/relationships/xlWorksheet",
    RELATIONSHIP_TYPES.worksheet,
  ],
  [
    "http://schemas.microsoft.com/office/2006/relationships/xlSharedStrings",
    RELATIONSHIP_TYPES.sharedStrings,
  ],
  ["http://schemas.microsoft.com/office/2006/relationships/xlStyles", RELATIONSHIP_TYPES.styles],
  [
    "http://schemas.microsoft.com/office/2006/relationships/xlCalcChain",
    RELATIONSHIP_TYPES.calcChain,
  ],
  ["http://schemas.microsoft.com/office/2006/relationships/theme", RELATIONSHIP_TYPES.theme],
  [
    "http://schemas.microsoft.com/office/2006/relationships/xlVolatileDependencies",
    VOLATILE_DEPENDENCIES_REL,
  ],
]);

function withPartRelationships(
  ctx: XlsxWriteContext,
  relationships: Relationships,
  ownerPath: string,
): WriteContext {
  const partCtx = Object.create(ctx);
  partCtx.addRelationship = (type: RelationshipType, target: string, mode?: string) => {
    const source = ownerPath.split("/").slice(0, -1);
    const targetPath = target.split("/").slice(0, -1);
    let common = 0;
    while (
      common < source.length &&
      common < targetPath.length &&
      source[common] === targetPath[common]
    )
      common++;
    const relativeTarget = [
      ...Array.from({ length: source.length - common }, () => ".."),
      ...target.split("/").slice(common),
    ].join("/");
    return `rId${relationships.add(
      type,
      target.startsWith("/") ? target : relativeTarget,
      mode as "External" | undefined,
    )}`;
  };
  partCtx.addMedia = ctx.addMedia.bind(ctx);
  return partCtx;
}
/** Extension → MIME for image and VML Default entries. Declared only for
 * extensions actually present in the package. VML backs legacy comment
 * anchors (xl/drawings/vmlDrawing${i}.vml). */
const XLSX_MEDIA_CONTENT_TYPES: Record<string, string> = {
  ...IMAGE_MEDIA_CONTENT_TYPES,
  vml: "application/vnd.openxmlformats-officedocument.vmlDrawing",
};

/** UTF-8 encoder for serializing the derived [Content_Types].xml into the package. */
const encoder = new TextEncoder();

/**
 * Compile workbook options into a Zippable structure.
 */
export function compileWorkbook(
  options: WorkbookOptions,
  overrides: XmlifyedFile[] = [],
  mediaLevel: number = 0,
  reproducible?: ReproducibleScope,
  packageVariant: OoxmlPackageVariant = "standard",
): Zippable {
  const ctx = new XlsxWriteContext();
  ctx.reproducible = reproducible;
  // Round-trip packages may keep the workbook (and its satellites) outside
  // xl/ — minimal hand-built packages often live at the archive root.
  const workbookPath = options.workbookPath ?? "xl/workbook.xml";
  const stylesPath = options.stylesPath ?? "xl/styles.xml";
  const sharedStringsPath = options.sharedStringsPath ?? "xl/sharedStrings.xml";
  const relationshipNamespace =
    options.relationshipNamespace === "microsoft2005" ? LEGACY_RELATIONSHIPS_NAMESPACE : undefined;
  ctx.workbookRels = new Relationships(workbookPath, relationshipNamespace);
  const mapping: Record<string, { data: string; path: string }> = {};
  for (const rel of options.passthroughRelationships ?? []) {
    if (rel.source === workbookPath) ctx.workbookRels.claimSourceRel(rel);
    const canonicalType = LEGACY_WORKBOOK_RELATIONSHIP_TYPES.get(rel.relationshipType);
    if (canonicalType && ctx.workbookRels.hasId(rel.rId))
      ctx.workbookRels.retypeRelationship(rel.rId, canonicalType);
  }
  const addWorkbookRelationship = (
    type: RelationshipType,
    target: string,
    targetMode?: "External",
  ): string => {
    const existing = ctx.workbookRels.idOf(type, target);
    if (existing !== undefined) return existing;
    const rid = ctx.workbookRels.nextRelationshipId;
    ctx.workbookRels.addRelationship(rid, type, target, targetMode);
    return `rId${rid}`;
  };

  // Seed the shared string table from parsed entries so round-tripped cells
  // keep their si indices and rich-text structure (identity dedup in
  // registerRich resolves the same entry object back to its source index).
  if (options.sharedStrings) ctx.sharedStrings.loadEntries(options.sharedStrings);
  if (options.sharedStringsCount !== undefined)
    ctx.sharedStrings.setSourceCount(options.sharedStringsCount);
  if (options.sharedStringsUniqueCount !== undefined)
    ctx.sharedStrings.setSourceUniqueCount(options.sharedStringsUniqueCount);

  const chartsheetConfigs = options.chartsheets ?? [];
  const dialogsheetConfigs = options.dialogsheets ?? [];
  const hasCustomProperties = !!options.customProperties && options.customProperties.length > 0;
  const includeCustomProperties = hasCustomProperties || options.customPropertiesDeclared === true;
  const isRoundTrip = options.contentTypes !== undefined;

  // Fresh authoring pairs an explicit row height with customHeight="1" (Excel
  // semantics); round-trips keep the source attributes verbatim.
  const worksheetConfigs = isRoundTrip
    ? (options.worksheets ?? [])
    : (options.worksheets ?? []).map((ws) => ({
        ...ws,
        ...(ws.pageMargins
          ? {}
          : {
              pageMargins: {
                left: 0.75,
                right: 0.75,
                top: 1,
                bottom: 1,
                header: 0.5,
                footer: 0.5,
              },
            }),
        rows: ws.rows?.map((row) =>
          row.height !== undefined && row.customHeight === undefined
            ? { ...row, customHeight: true }
            : row,
        ),
      }));

  // Core/app properties are required in fresh packages, but preserve source
  // packages that omit these optional OPC metadata parts.
  const hasMetadataOverride = (partName: string): boolean =>
    options.contentTypes?.overrides.some(
      (entry) => entry.partName.toLowerCase() === partName.toLowerCase(),
    ) === true;
  const includeCoreProperties =
    !isRoundTrip ||
    options.corePropertiesPath !== undefined ||
    hasMetadataOverride("/docProps/core.xml");
  const includeAppProperties =
    !isRoundTrip ||
    options.appPropertiesPath !== undefined ||
    hasMetadataOverride("/docProps/app.xml") ||
    options.appProperties !== undefined;

  if (includeCoreProperties) {
    mapping["Properties"] = {
      data: XML_DECL + buildCorePropertiesXmlString(options, reproducible),
      path: options.corePropertiesPath ?? "docProps/core.xml",
    };
  }

  if (includeAppProperties) {
    mapping["AppProperties"] = {
      data: XML_DECL + (appPropertiesDesc.stringify(options.appProperties ?? {}, ctx) ?? ""),
      path: options.appPropertiesPath ?? "docProps/app.xml",
    };
  }

  // Custom properties (optional part; only emitted when present)
  if (includeCustomProperties) {
    mapping["CustomProperties"] = {
      data:
        XML_DECL +
        (customPropertiesDesc.stringify({ properties: options.customProperties ?? [] }, ctx) ?? ""),
      path: options.customPropertiesPath ?? "docProps/custom.xml",
    };
  }

  if (options.classificationLabels) {
    const labelsPath = options.classificationLabelsPath ?? "docMetadata/LabelInfo.xml";
    mapping["ClassificationLabels"] = {
      data:
        XML_DECL + (classificationLabelsDesc.stringify(options.classificationLabels, ctx) ?? ""),
      path: labelsPath,
    };
  }

  if (options.persons) {
    const personsPath = options.personsPath ?? "xl/persons/person.xml";
    mapping["Persons"] = {
      data: XML_DECL + (personsDesc.stringify(options.persons, ctx) ?? ""),
      path: personsPath,
    };
    const target = personsPath.replace(/^xl\//, "");
    const personsRel = RELATIONSHIP_TYPES.threadedCommentPersons;
    if (ctx.workbookRels.idOf(personsRel, target) === undefined) {
      ctx.workbookRels.addRelationship(ctx.workbookRels.nextRelationshipId, personsRel, target);
    }
  }

  // File-level relationships (_rels/.rels)
  const fileRels = buildRootRelationships(
    workbookPath,
    includeCustomProperties,
    options.passthroughRelationships,
    {
      includeCoreProperties,
      includeAppProperties,
      corePropertiesPath: options.corePropertiesPath,
      corePropertiesType: options.corePropertiesRelationshipType,
      appPropertiesPath: options.appPropertiesPath,
      customPropertiesPath: options.customPropertiesPath,
      namespace: relationshipNamespace,
      additionalRelationships: options.classificationLabels
        ? [
            {
              relationshipType:
                "http://schemas.microsoft.com/office/2020/02/relationships/classificationlabels",
              target: options.classificationLabelsPath ?? "docMetadata/LabelInfo.xml",
            },
          ]
        : [],
    },
  );
  mapping["FileRelationships"] = {
    data: XML_DECL + fileRels.serialize(),
    path: "_rels/.rels",
  };

  // Register predefined DXFs before worksheets use styles. options.dxfs === []
  // means the source declared an empty <dxfs/> container — kept as-is.
  if (options.dxfs !== undefined) ctx.styles.setDxfs(options.dxfs);

  // Adopt the parsed style table wholesale (the SDK's Stylesheet model): fonts/
  // fills/borders/cellXfs/numFmts replace the fresh-file defaults, so raw cell
  // style indices carried by parsed rows resolve exactly as in the source.
  // fonts !== undefined is the "source styles.xml was parsed" signal (parse
  // fills all table sections with [] when the part exists, even bare).
  if (options.fonts !== undefined) {
    ctx.styles.adopt({
      fonts: options.fonts,
      fontsContainer: options.fontsContainer,
      fills: options.fills ?? [],
      borders: options.borders ?? [],
      cellXfs: options.cellXfs ?? [],
      numFmts: options.numFmts,
    });
  }

  // Re-apply parsed style sections so a declarative round-trip preserves
  // them (colors, table/cell styles, extensions) alongside DXFs.
  if (options.colors) ctx.styles.setColors(options.colors);
  if (options.tableStyles) ctx.styles.setTableStyles(options.tableStyles);
  if (options.cellStyles !== undefined) ctx.styles.setCustomCellStyles(options.cellStyles);
  if (options.cellStyleXfs !== undefined) ctx.styles.setCellStyleXfs(options.cellStyleXfs);
  if (options.styleExtensions) ctx.styles.setExtensions(options.styleExtensions);

  // Build workbook relationships
  const includeStyles = !isRoundTrip || options.fonts !== undefined;
  const includeTheme = !isRoundTrip || options.theme !== undefined;
  // Round-trip worksheets keep their source part names (sheet26.xml); fresh
  // authoring assigns sheet1..N. Relationship targets must match whichever
  // naming the compiled package actually uses.
  const worksheetRelTargets = worksheetConfigs.map((ws, i) => {
    const sourcePath = (ws as WorksheetOptions).sourcePath;
    return sourcePath ? sourcePath.replace(/^xl\//, "") : `worksheets/sheet${i + 1}.xml`;
  });
  const sheetRelationshipIds = buildWorkbookRelationships(
    ctx.workbookRels,
    worksheetRelTargets,
    chartsheetConfigs.length,
    dialogsheetConfigs.map((config, index) =>
      (config.sourcePath ?? `xl/dialogSheets/sheet${index + 1}.xml`).replace(/^xl\//, ""),
    ),
    includeStyles,
    includeTheme,
    options.themePath ? options.themePath.replace(/^xl\//, "") : "theme/theme1.xml",
  );

  // Build sheet definitions for workbook XML. An explicit sheetId wins; the
  // fallback counter skips past every id handed out so ids stay unique even
  // when options mix explicit and generated values.
  let sheets: SheetDefinition[] = [];
  let sheetId = 1;
  const nextSheetId = (explicit: number | undefined): number => {
    const id = explicit ?? sheetId;
    if (id >= sheetId) sheetId = id + 1;
    return id;
  };
  let sheetRelationshipIndex = 0;
  const sheetIdentity = (sheet: SheetDefinition): number => sheet.sheetId ?? sheet.tabId ?? 0;
  for (const ws of worksheetConfigs) {
    sheets.push({
      name: ws.name ?? `Sheet${sheetId}`,
      sheetId: nextSheetId(ws.sheetId),
      state: ws.state,
      rId: sheetRelationshipIds[sheetRelationshipIndex++] ?? `rId${sheetRelationshipIndex}`,
    });
  }
  for (const cs of chartsheetConfigs) {
    sheets.push({
      name: cs.name ?? `Chart${sheetId}`,
      sheetId: nextSheetId(cs.sheetId),
      state: cs.state,
      rId: sheetRelationshipIds[sheetRelationshipIndex++] ?? `rId${sheetRelationshipIndex}`,
    });
  }
  for (const ds of dialogsheetConfigs) {
    sheets.push({
      name: ds.name ?? `Dialog${sheetId}`,
      sheetId: nextSheetId(ds.sheetId),
      state: ds.state,
      rId: sheetRelationshipIds[sheetRelationshipIndex++] ?? `rId${sheetRelationshipIndex}`,
    });
  }
  if (options.sheetDefinitions?.length === sheets.length) {
    const knownDefinitions =
      options.sheetDefinitions?.length === sheets.length &&
      options.sheetDefinitions.every((definition) =>
        sheets.some((sheet) => sheetIdentity(sheet) === sheetIdentity(definition)),
      );
    if (knownDefinitions && options.sheetDefinitions) sheets = [...options.sheetDefinitions];
  }
  // Sheets whose parts the model does not rebuild (Excel 4 macro sheets, …)
  // still need their <sheet> entry: the part travels as passthrough at its
  // source path, so re-emit the entry wired to the source relationship id.
  const unmatchedDefinitions = (options.sheetDefinitions ?? []).filter(
    (definition) => !sheets.some((sheet) => sheetIdentity(sheet) === sheetIdentity(definition)),
  );
  if (unmatchedDefinitions.length > 0) {
    for (const definition of unmatchedDefinitions) {
      const passthrough = (options.passthroughRelationships ?? []).find(
        (rel) => rel.source === workbookPath && rel.rId === definition.rId,
      );
      if (definition.rId === "") {
        sheets.push({
          name: definition.name,
          sheetId: definition.sheetId,
          tabId: definition.tabId,
          state: definition.state,
          rId: definition.rId,
        });
        continue;
      }
      if (!passthrough) continue;
      const numericId = /^rId(\d+)$/.exec(definition.rId)?.[1];
      if (numericId === undefined) continue;
      const canonicalType = LEGACY_WORKBOOK_RELATIONSHIP_TYPES.get(passthrough.relationshipType);
      if (
        !ctx.workbookRels.hasRelationship(
          canonicalType ?? passthrough.relationshipType,
          passthrough.target,
        )
      ) {
        ctx.workbookRels.addRelationship(
          Number(numericId),
          passthrough.relationshipType as RelationshipType,
          passthrough.target,
        );
      }
      sheets.push({
        name: definition.name,
        sheetId: definition.sheetId,
        tabId: definition.tabId,
        state: definition.state,
        rId: definition.rId,
      });
    }
    const definitionOrder = new Map(
      (options.sheetDefinitions ?? []).map((definition, index) => [
        sheetIdentity(definition),
        index,
      ]),
    );
    sheets.sort(
      (left, right) =>
        (definitionOrder.get(sheetIdentity(left)) ?? Number.POSITIVE_INFINITY) -
        (definitionOrder.get(sheetIdentity(right)) ?? Number.POSITIVE_INFINITY),
    );
  }

  const wsContext: WorksheetContext = { sharedStrings: ctx.sharedStrings, styles: ctx.styles };
  const state: WorksheetCompileState = {
    globalMediaIdx: 0,
    globalChartIdx: 0,
    globalPivotIdx: 0,
    globalPivotCacheIdx: 0,
    globalTableIdx: 0,
    globalQueryTableIdx: 0,
    globalSingleXmlCellsIdx: 0,
    globalControlPropertiesIdx: 0,
    globalActiveXIdx: 0,
    pivotCacheDataMap: new Map<string, { cacheId: number; cacheIdx: number }>(),
    pivotCachePathById: new Map<number, string>(),
    definedPivotCacheCount: 0,
    calcCells: [],
    allTableParts: [],
    chartPaths: new Map(),
    nextDrawingIndex: 1,
  };
  compileDefinitionPivotCaches(options, ctx, mapping, state);
  for (const [i, wsOpts] of worksheetConfigs.entries()) {
    compileWorksheetPart(
      wsOpts,
      i,
      worksheetConfigs,
      ctx,
      mapping,
      wsContext,
      state,
      options.passthroughRelationships,
    );
  }

  compileChartsheets(chartsheetConfigs, ctx, mapping, state, options.passthroughRelationships);
  compileDialogsheets(dialogsheetConfigs, ctx, mapping, options.passthroughRelationships);
  // Workbook XML (via descriptor)
  const freshWorkbookDefaults = {
    fileVersion:
      options.fileVersion ??
      (isRoundTrip
        ? undefined
        : {
            appName: "xl",
            lastEdited: 7,
            lowestEdited: 6,
            rupBuild: 29929,
          }),
    properties: options.properties ?? (isRoundTrip ? undefined : {}),
    bookView:
      options.bookView ??
      (isRoundTrip
        ? undefined
        : {
            xWindow: 0,
            yWindow: 0,
            windowWidth: 28800,
            windowHeight: 12300,
          }),
  };
  let wbXml =
    workbookDesc.stringify(
      {
        ...freshWorkbookDefaults,
        sheets,
        conformance: options.conformance,
        ...(options.legacyChildOrder ? { legacyChildOrder: true } : {}),
        pivotCaches: ctx.pivotCacheRefs,
        protection: options.workbookProtection,
        customViews: options.customWorkbookViews,
        fileRecovery: options.fileRecovery,
        functionGroups: options.functionGroups,
        webPublishing: options.webPublishing,
        fileSharing: options.fileSharing,
        webPublishObjects: options.webPublishObjects,
        definedNames: options.definedNames,
        calculation: options.calculation,
        oleSize: options.oleSize,
        ...(options.absPath !== undefined ? { absPath: options.absPath } : {}),
        ...(options.absPathLegacyPrefix ? { absPathLegacyPrefix: true } : {}),
        ...(options.revisionPtr ? { revisionPtr: options.revisionPtr } : {}),
        ...(options.extensions ? { extensions: options.extensions } : {}),
      },
      ctx,
    ) ?? "";

  // Connections — xl/connections.xml (single part, workbook-level relationship)
  if (options.connections && options.connections.length > 0) {
    addWorkbookRelationship(RELATIONSHIP_TYPES.connections, "connections.xml");
    mapping["Connections"] = {
      data: XML_DECL + connectionsDesc.stringify({ connections: options.connections }, ctx),
      path: "xl/connections.xml",
    };
  }

  // Metadata — xl/metadata.xml (single part, workbook-level relationship)
  if (options.metadata && hasMetadataContent(options.metadata)) {
    addWorkbookRelationship(RELATIONSHIP_TYPES.sheetMetadata, "metadata.xml");
    mapping["Metadata"] = {
      data: XML_DECL + metadataDesc.stringify(options.metadata, ctx),
      path: "xl/metadata.xml",
    };
  }

  // XML mappings — xl/xmlMaps.xml (single part, workbook-level relationship)
  if (options.xmlMaps) {
    addWorkbookRelationship(RELATIONSHIP_TYPES.xmlMaps, "xmlMaps.xml");
    mapping["XmlMaps"] = {
      data: XML_DECL + mapInfoDesc.stringify(options.xmlMaps, ctx),
      path: "xl/xmlMaps.xml",
    };
  }

  // Volatile function types — xl/volTypes.xml (single part, workbook-level
  // relationship; sml.xsd declares volTypes as a part root, never a workbook child)
  if (options.volTypes !== undefined) {
    const volTypesPath = options.volTypesPath ?? "volTypes.xml";
    addWorkbookRelationship(
      volTypesPath === "volatileDependencies.xml"
        ? VOLATILE_DEPENDENCIES_REL
        : RELATIONSHIP_TYPES.volTypes,
      volTypesPath,
    );
    mapping["VolTypes"] = {
      data: XML_DECL + buildVolTypesXml(options.volTypes, options.volTypesCount),
      path: `xl/${volTypesPath}`,
    };
  }

  // External links — generate XML files and inject externalReferences into workbook
  const extLinks = options.externalLinks ?? [];
  if (extLinks.length > 0) {
    const extRefs: { rId: string }[] = [];
    for (let ei = 0; ei < extLinks.length; ei++) {
      const elIdx = ei + 1;
      // Create the rels file for this external link
      const elOpts = extLinks[ei];
      if (!elOpts) continue;
      const externalLinkPartPath = elOpts.sourcePath ?? `xl/externalLinks/externalLink${elIdx}.xml`;
      const externalLinkRelTarget = externalLinkPartPath.replace(/^xl\//, "");
      const externalLinkRid = addWorkbookRelationship(
        RELATIONSHIP_TYPES.externalLink,
        externalLinkRelTarget,
      );
      let bookRId: string | undefined;
      let oleRId: string | undefined;
      const elRels = new Relationships();
      if (elOpts.externalBook?.target) {
        bookRId = elOpts.bookRId ?? "rId1";
        elRels.addRelationship(
          bookRId,
          elOpts.externalBook.targetTypeRaw ?? RELATIONSHIP_TYPES.externalLinkPath,
          elOpts.externalBook.target,
          TargetModeType.EXTERNAL,
        );
      }
      if (elOpts.oleLink && elOpts.oleTarget) {
        oleRId = elOpts.oleRId ?? "rId1";
        elRels.addRelationship(
          oleRId,
          (elOpts.oleTargetTypeRaw ?? RELATIONSHIP_TYPES.oleObject) as RelationshipType,
          elOpts.oleTarget,
          TargetModeType.EXTERNAL,
        );
      }
      if (elRels.relationshipCount > 0) {
        mapping[`ExternalLinkRels${elIdx}`] = {
          data: XML_DECL + elRels.serialize(),
          path: partPathToRelsPath(externalLinkPartPath),
        };
      }

      // Generate the external link XML
      mapping[`ExternalLink${elIdx}`] = {
        data: XML_DECL + externalLinkDesc.stringify({ ...elOpts, bookRId, oleRId }, ctx),
        path: externalLinkPartPath,
      };

      extRefs.push({ rId: externalLinkRid });
    }

    // Inject externalReferences into workbook XML
    const extRefsXml = buildExternalReferencesXml(extRefs);
    wbXml = wbXml.replace("<!--EXTERNAL_REFS-->", extRefsXml);
  } else {
    wbXml = wbXml.replace("<!--EXTERNAL_REFS-->", "");
  }

  mapping["Workbook"] = {
    data: XML_DECL + wbXml,
    path: workbookPath,
  };

  // Shared Strings — AFTER worksheets so all strings are collected
  if (ctx.sharedStrings.count > 0 || options.sharedStringsDeclared === true) {
    addWorkbookRelationship(RELATIONSHIP_TYPES.sharedStrings, "sharedStrings.xml");
    const ssXml = sharedStringsDesc.stringify(ctx.sharedStrings.toDescriptorOptions(), ctx);
    mapping["SharedStrings"] = {
      data: XML_DECL + ssXml,
      path: sharedStringsPath,
    };
  }

  // Styles and theme are optional OPC parts. Fresh documents get Office's
  // defaults; a parsed round-trip follows the source and omits absent parts.
  if (includeStyles) {
    const stylesXml = stylesDesc.stringify({ styles: ctx.styles }, ctx);
    mapping["Styles"] = {
      data: XML_DECL + stylesXml,
      path: stylesPath,
    };
  }

  if (includeTheme) {
    const themeRels = new Relationships();
    const themePath = options.themePath ?? "xl/theme/theme1.xml";
    const themeXml = options.theme
      ? bindMediaPlaceholders(
          buildThemeXml(options.theme, withPartRelationships(ctx, themeRels, themePath)),
          ctx.media,
          themeRels,
        )
      : createThemeXml();
    mapping["Theme"] = {
      data: XML_DECL + themeXml,
      path: themePath,
    };
    if (themeRels.relationshipCount > 0) {
      mapping["ThemeRels"] = {
        data: XML_DECL + themeRels.serialize(),
        path: partPathToRelsPath(themePath),
      };
    }
  }

  // Charts — AFTER worksheets so charts are registered
  for (const [i, chartData] of ctx.charts.array.entries()) {
    const chartPath = state.chartPaths.get(chartData.key) ?? `xl/charts/chart${i + 1}.xml`;
    const chartRels = new Relationships();
    const chartOptions = ctx.chartOptions.get(chartData.key);
    const externalData = chartOptions?.externalData;
    if (externalData && chartOptions?.externalDataRelationshipType) {
      chartRels.addRelationship(
        externalData.relationshipId,
        chartOptions.externalDataRelationshipType as RelationshipType,
        chartOptions.externalDataRelationshipTarget ?? "../externalLinks/externalLink1.xml",
      );
    }
    const chartXml = bindMediaPlaceholders(chartData.chartSpaceXml, ctx.media, chartRels);
    mapping[`Chart${i}`] = {
      data: XML_DECL + chartXml,
      path: chartPath,
    };
    // User-shapes part behind c:userShapes: the chart part's own rels entry
    // plus the body part (chartUserShapes relationship, same directory).
    if (chartData.userShapes) {
      const rid = chartData.userShapes.relationshipId;
      const userShapesPath = chartData.userShapes.path ?? `xl/charts/userShapes${i + 1}.xml`;
      const userShapesTarget = userShapesPath.startsWith("xl/charts/")
        ? userShapesPath.slice("xl/charts/".length)
        : `../${userShapesPath.replace(/^xl\//, "")}`;
      mapping[`ChartUserShapes${i}`] = {
        data: XML_DECL + chartData.userShapes.xml,
        path: userShapesPath,
      };
      chartRels.addRelationship(rid, CHART_USER_SHAPES_REL, userShapesTarget);
    }
    if (chartOptions?.externalLink && chartOptions.externalLinkPath) {
      const externalLinkPath = chartOptions.externalLinkPath;
      mapping[`ChartExternalLink${i}`] = {
        data: XML_DECL + (externalLinkDesc.stringify(chartOptions.externalLink, ctx) ?? ""),
        path: externalLinkPath,
      };
      if (chartOptions.externalLink.externalBook?.target) {
        const externalLinkRels = new Relationships(externalLinkPath);
        const preferredRid = /^rId(\d+)$/.exec(chartOptions.externalLink.bookRId ?? "")?.[1];
        externalLinkRels.addRelationship(
          preferredRid ? Number(preferredRid) : externalLinkRels.nextRelationshipId,
          RELATIONSHIP_TYPES.externalLinkPath,
          chartOptions.externalLink.externalBook.target,
          TargetModeType.EXTERNAL,
        );
        mapping[`ChartExternalLinkRels${i}`] = {
          data: XML_DECL + externalLinkRels.serialize(),
          path: partPathToRelsPath(externalLinkPath),
        };
      }
    }
    if (chartOptions?.themeOverridePath) {
      const themeOverridePath = chartOptions.themeOverridePath;
      mapping[`ChartThemeOverride${i}`] = {
        data: XML_DECL + (themeOverrideDesc.stringify(chartOptions.themeOverride ?? {}, ctx) ?? ""),
        path: themeOverridePath,
      };
      const preferredId = /^rId(\d+)$/.exec(
        (options.passthroughRelationships ?? []).find(
          (rel) =>
            rel.source === chartPath &&
            rel.relationshipType.endsWith("/themeOverride") &&
            resolveRelationshipTarget(
              rel.source.replace(/\/_rels\/[^/]+\.rels$/, ""),
              rel.target,
            ) === themeOverridePath,
        )?.rId ?? "",
      )?.[1];
      chartRels.addRelationship(
        preferredId ? Number(preferredId) : chartRels.nextRelationshipId,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride" as RelationshipType,
        `../${themeOverridePath.replace(/^xl\//, "")}`,
      );
    }
    // Chart media/userShapes/externalData are modeled; lesser-known companion
    // parts still travel with the chart's source relationship topology.
    for (const sourceRel of options.passthroughRelationships ?? []) {
      if (sourceRel.source !== chartPath) continue;
      if (chartRels.hasRelationship(sourceRel.relationshipType, sourceRel.target)) continue;
      const preferred = /^rId\d+$/.exec(sourceRel.rId)?.[0];
      if (preferred && !chartRels.hasId(preferred)) {
        chartRels.addRelationship(
          preferred,
          sourceRel.relationshipType as RelationshipType,
          sourceRel.target,
          sourceRel.targetMode,
        );
        continue;
      }
      chartRels.addRelationship(
        chartRels.nextRelationshipId,
        sourceRel.relationshipType as RelationshipType,
        sourceRel.target,
        sourceRel.targetMode,
      );
    }
    if (chartRels.relationshipCount > 0) {
      chartData.relsXml = chartRels.serialize();
    }
    if (chartData.relsXml) {
      mapping[`ChartMediaRels${i}`] = {
        data: XML_DECL + chartData.relsXml,
        path: partPathToRelsPath(chartPath),
      };
    }
  }

  // Calculation chain — round-trips preserve the parsed chain verbatim (the
  // chain encodes Excel's own evaluation order, not something derivable);
  // fresh authoring rebuilds one from formula cells. A source whose workbook
  // rels reference calcChain but ship no part (repair-style files) keeps the
  // part absent — Excel tolerates the dangling reference exactly as received.
  const calcChainCells = options.calcChain === false ? [] : (options.calcChain ?? state.calcCells);
  const srcReferencesCalcChain = (options.passthroughRelationships ?? []).some(
    (r) => r.source === workbookPath && r.relationshipType.endsWith("/calcChain"),
  );
  if (calcChainCells.length > 0 && !(srcReferencesCalcChain && options.calcChain === undefined)) {
    mapping["CalcChain"] = {
      data: calcChainDesc.stringify({ cells: calcChainCells }, ctx) ?? "",
      path: "xl/calcChain.xml",
    };
    addWorkbookRelationship(RELATIONSHIP_TYPES.calcChain, "calcChain.xml");
  }

  if (options.revisionLog) {
    compileRevisionLogs(options.revisionLog, ctx, mapping);
  }

  if (options.richData) {
    const richData = options.richData;
    const richDataDefaults = {
      data: "xl/richData/rdrichvalue.xml",
      structures: "xl/richData/rdrichvaluestructure.xml",
      types: "xl/richData/rdRichValueTypes.xml",
      relationships: "xl/richData/richValueRel.xml",
    };
    const richDataWorkbookRel = (
      type: string,
      target: string,
      preferredId: string | undefined,
    ): void => {
      if (ctx.workbookRels.idOf(type as RelationshipType, target) !== undefined) return;
      const preferred = /^rId(\d+)$/.exec(preferredId ?? "")?.[1];
      if (preferred !== undefined && !ctx.workbookRels.hasId(`rId${preferred}`)) {
        ctx.workbookRels.addRelationship(
          Number(preferred),
          type as RelationshipType,
          target.replace(/^xl\//, ""),
        );
        return;
      }
      addWorkbookRelationship(type as RelationshipType, target.replace(/^xl\//, ""));
    };
    if (richData.data) {
      const path = richData.dataPath ?? richDataDefaults.data;
      mapping["RichValueData"] = {
        data: XML_DECL + (richValueDataDesc.stringify(richData.data, ctx) ?? ""),
        path,
      };
      richDataWorkbookRel(
        "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue",
        path,
        richData.dataRelationshipId,
      );
    }
    if (richData.structures) {
      const path = richData.structuresPath ?? richDataDefaults.structures;
      mapping["RichValueStructures"] = {
        data: XML_DECL + (richValueStructuresDesc.stringify(richData.structures, ctx) ?? ""),
        path,
      };
      richDataWorkbookRel(
        "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure",
        path,
        richData.structuresRelationshipId,
      );
    }
    if (richData.types) {
      const path = richData.typesPath ?? richDataDefaults.types;
      mapping["RichValueTypes"] = {
        data: XML_DECL + (richValueTypesInfoDesc.stringify(richData.types, ctx) ?? ""),
        path,
      };
      richDataWorkbookRel(
        "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes",
        path,
        richData.typesRelationshipId,
      );
    }
    if (richData.relationships) {
      const path = richData.relationshipsPath ?? richDataDefaults.relationships;
      mapping["RichValueRelationships"] = {
        data: XML_DECL + (richValueRelsDesc.stringify(richData.relationships, ctx) ?? ""),
        path,
      };
      richDataWorkbookRel(
        "http://schemas.microsoft.com/office/2022/10/relationships/richValueRel",
        path,
        richData.relationshipsRelationshipId,
      );
      if (richData.partRelationships?.length) {
        const richValueRels = new Relationships(path);
        for (const relationship of richData.partRelationships) {
          const preferred = /^rId(\d+)$/.exec(relationship.rId)?.[1];
          richValueRels.addRelationship(
            preferred ? Number(preferred) : richValueRels.nextRelationshipId,
            relationship.relationshipType as RelationshipType,
            relationship.target,
            relationship.targetMode,
          );
        }
        mapping["RichValueRelationshipRels"] = {
          data: XML_DECL + richValueRels.serialize(),
          path: partPathToRelsPath(path),
        };
      }
    }
  }

  // Workbook relationships — serialized after calcChain/revision register their
  // targets, so every workbook-level relationship lands in workbook.xml.rels.
  // Passthrough relationships (round-trip) follow: the source workbook.xml.rels
  // referenced parts the model carries verbatim (externalLinks, pivotCaches, …).
  // Re-emitted as written — targets are passthrough paths that never move.
  for (const rel of options.passthroughRelationships ?? []) {
    if (rel.source !== workbookPath) continue;
    if (relationshipNamespace) continue;
    if (ctx.workbookRels.hasRelationship(rel.relationshipType, rel.target)) continue;
    ctx.workbookRels.add(rel.relationshipType as RelationshipType, rel.target);
  }
  for (const rel of options.passthroughRelationships ?? []) {
    if (rel.source === workbookPath && ctx.workbookRels.hasId(rel.rId))
      ctx.workbookRels.retypeRelationship(rel.rId, rel.relationshipType);
  }
  mapping["WorkbookRelationships"] = {
    data: XML_DECL + ctx.workbookRels.serialize(),
    path: partPathToRelsPath(workbookPath),
  };

  // Convert mapping to Zippable
  const mediaFiles: Array<{ data: Uint8Array; path: string }> = [];
  for (const img of ctx.media.array) {
    mediaFiles.push({ data: img.data, path: `xl/media/${img.fileName}` });
  }

  const files = compileMapping(mapping, overrides, mediaFiles, mediaLevel);
  // Derive [Content_Types].xml from the actual parts written — the file set is
  // the single source of truth, so content-type declarations cannot drift from
  // what is written. Sparse/index-based naming is handled naturally. Raw
  // passthrough parts (drawings, VML, external links, unknown extensions, …)
  // copy in first: the compiler output above wins at the same path, so only
  // what the model missed actually passes through.
  files["[Content_Types].xml"] = encoder.encode(
    finalizeContentTypes(
      files,
      {
        resolve: XLSX_CONTENT_TYPE_RESOLVER,
        mediaContentTypes: XLSX_MEDIA_CONTENT_TYPES,
        // Round-trip: the source declaration table is the base; derived entries
        // only fill what surviving source entries leave uncovered or mistyped.
        source: options.contentTypes,
        rawParts: options.rawParts,
        omitDerivedOverrides: options.contentTypes ? ["docProps/app.xml"] : undefined,
        forcedOverrides: [
          ...(options.contentTypes?.overrides ?? [])
            .filter(
              (override) =>
                override.contentType ===
                "application/vnd.openxmlformats-officedocument.spreadsheetml.revisionLog+xml",
            )
            .map((override) => ({
              path: override.partName.replace(/^\//, ""),
              contentType: override.contentType,
            })),
          {
            path: workbookPath,
            contentType: resolveMainPartContentType(
              "spreadsheet",
              options.contentTypes,
              workbookPath,
              ooxmlPackageFormatInfo("spreadsheet", packageVariant).mainContentType,
            ),
          },
        ],
      },
      ctx,
    ),
  );
  // Guard: drop passthrough rels whose target part never made it into the
  // package (hand-authored input) — Office refuses to open dangling rels.
  dropDanglingPassthroughRels(
    files,
    options.passthroughRelationships,
    (rel) => rel.source === workbookPath && rel.relationshipType.endsWith("/calcChain"),
  );
  return files;
}

/** Cross-worksheet compile state threaded through the worksheet compile phases. */
export interface WorksheetCompileState {
  globalMediaIdx: number;
  globalChartIdx: number;
  globalPivotIdx: number;
  globalPivotCacheIdx: number;
  globalTableIdx: number;
  globalQueryTableIdx: number;
  globalSingleXmlCellsIdx: number;
  globalControlPropertiesIdx: number;
  globalActiveXIdx: number;
  pivotCacheDataMap: Map<string, { cacheId: number; cacheIdx: number }>;
  pivotCachePathById: Map<number, string>;
  definedPivotCacheCount: number;
  calcCells: CalcCell[];
  allTableParts: TablePartReference[];
  chartPaths: Map<string, string>;
  nextDrawingIndex: number;
}

/**
 * Compile one worksheet: sheet XML, calcChain cells, drawing/media,
 * comments + VML, background, pivots, tables, query tables and
 * single-cell XML tables, with their worksheet-level relationships.
 * Rel registration order is significant (rIds are assigned sequentially).
 */
function compileWorksheetPart(
  wsOpts: WorksheetOptions,
  i: number,
  worksheetConfigs: WorksheetOptions[],
  ctx: XlsxWriteContext,
  mapping: Record<string, { data: string; path: string }>,
  wsContext: WorksheetContext,
  state: WorksheetCompileState,
  passthroughRelationships?: readonly PassthroughRelationship[],
): void {
  const imgOpts = wsOpts.images ?? [];
  const chartOpts = wsOpts.charts ?? [];
  const smartArtOpts = wsOpts.smartArts ?? [];
  const webExtensionOpts = wsOpts.webExtensions ?? [];
  const shapeOpts = wsOpts.shapes ?? [];
  const connectorOpts = wsOpts.connectors ?? [];
  const groupOpts = wsOpts.groups ?? [];
  const contentPartOpts = wsOpts.contentParts ?? [];
  const hlOpts = wsOpts.hyperlinks ?? [];
  const sheetName = wsOpts.name ?? `Sheet${i + 1}`;

  // Collect formula cells for calcChain. calcChain's i attribute is the
  // workbook sheetId (CT_Sheet @sheetId), not the sheet's position — the
  // fallback matches the all-generated case where both coincide.
  const sheetIdx = wsOpts.sheetId ?? i + 1;
  const wsRows = wsOpts.rows ?? [];
  for (let ri = 0; ri < wsRows.length; ri++) {
    const rowOpts = wsRows[ri]!;
    const rowNumber = rowOpts.rowNumber ?? ri + 1;
    const cells = rowOpts.cells;
    if (!cells) continue;
    for (let ci = 0; ci < cells.length; ci++) {
      const cell = cells[ci]!;
      if (!cell.formula) continue;
      const ref = cell.reference ?? columnToLetter(ci + 1) + rowNumber;
      state.calcCells.push({
        reference: ref,
        sheetIndex: sheetIdx,
        array: typeof cell.formula === "object" && cell.formula.type === "array",
      });
    }
  }

  const hasMedia =
    wsOpts.drawingPresent === true ||
    imgOpts.length > 0 ||
    chartOpts.length > 0 ||
    smartArtOpts.length > 0 ||
    webExtensionOpts.length > 0 ||
    shapeOpts.length > 0 ||
    connectorOpts.length > 0 ||
    groupOpts.length > 0 ||
    contentPartOpts.length > 0;
  const hasExternalHyperlinks = hlOpts.some((h) => h.url !== undefined);
  const hasDataConsolidateRelationships = wsOpts.dataConsolidate?.refs?.some(
    (ref) => typeof ref !== "string" && (ref.rId !== undefined || ref.target !== undefined),
  );
  const commentOpts = wsOpts.comments ?? [];
  const hasComments = commentOpts.length > 0;
  const pivotOpts = wsOpts.pivotTables ?? [];
  const hasPivots = pivotOpts.length > 0;
  const tableOpts = wsOpts.tables ?? [];
  const hasTables = tableOpts.length > 0;
  const queryTableOpts = wsOpts.queryTables ?? [];
  const hasQueryTables = queryTableOpts.length > 0;
  const singleXmlCellOpts = wsOpts.singleXmlCells ?? [];
  const bgImg = wsOpts.backgroundImage;
  const controlOpts = wsOpts.controls ?? [];

  // Worksheet-level relationships
  const wsPath = wsOpts.sourcePath ?? `xl/worksheets/sheet${i + 1}.xml`;
  const sourceWorksheetRels = (passthroughRelationships ?? []).filter(
    (rel) => rel.source === wsPath,
  );
  let wsRels: Relationships | undefined;

  if (
    hasMedia ||
    hasExternalHyperlinks ||
    hasComments ||
    hasPivots ||
    hasTables ||
    hasQueryTables ||
    singleXmlCellOpts.length > 0 ||
    controlOpts.length > 0 ||
    bgImg ||
    hasDataConsolidateRelationships ||
    sourceWorksheetRels.length > 0
  ) {
    wsRels = new Relationships();
    for (const rel of sourceWorksheetRels) {
      if (rel.relationshipType.endsWith("/externalLinkPath")) continue;
      wsRels.claimSourceRel(rel);
    }
  }

  const addWorksheetRelationship = (
    type: RelationshipType,
    target: string,
    targetMode?: "External",
  ): string => {
    const existing = wsRels!.idOf(type, target);
    if (existing !== undefined) return existing;
    return `rId${wsRels!.add(type, target, targetMode)}`;
  };
  const registerHyperlinkRelationship = (hl: { url?: string; relationshipId?: string }): string => {
    if (hl.url === undefined) throw new Error("internal hyperlink has no relationship");
    const preferred = /^rId\d+$/.exec(hl.relationshipId ?? "")?.[0];
    if (preferred && !wsRels!.hasId(preferred)) {
      wsRels!.addRelationship(
        preferred,
        RELATIONSHIP_TYPES.hyperlink,
        hl.url,
        TargetModeType.EXTERNAL,
      );
      return preferred;
    }
    const existing = wsRels!.idOf(RELATIONSHIP_TYPES.hyperlink, hl.url);
    if (existing !== undefined) return existing;
    return `rId${wsRels!.add(RELATIONSHIP_TYPES.hyperlink, hl.url, TargetModeType.EXTERNAL)}`;
  };
  const sourceRelationshipPath = (rel: { target: string } | undefined, fallback: string): string =>
    rel ? resolveRelationshipTarget(wsPath, rel.target) : fallback;

  const sourceControlRels = sourceWorksheetRels.filter(
    (rel) =>
      rel.relationshipType.endsWith("/ctrlProp") ||
      rel.relationshipType.endsWith("/control") ||
      rel.relationshipType.endsWith("/controls"),
  );
  let freshFormControlIndex = state.globalControlPropertiesIdx;
  let freshActiveXIndex = state.globalActiveXIdx;
  const wiredControls = controlOpts.map((control, controlIndex) => {
    const sourceRel = control.rId
      ? sourceControlRels.find((rel) => rel.rId === control.rId)
      : sourceControlRels[controlIndex];
    const isForm = control.kind === "form";
    const fallbackIndex = isForm ? ++freshFormControlIndex : ++freshActiveXIndex;
    const relationshipType =
      sourceRel?.relationshipType ??
      (isForm ? RELATIONSHIP_TYPES.ctrlProps : RELATIONSHIP_TYPES.controls);
    const fallbackTarget = isForm
      ? `../ctrlProps/ctrlProp${fallbackIndex}.xml`
      : `../activeX/activeX${fallbackIndex}.xml`;
    const target = sourceRel?.target ?? fallbackTarget;
    return {
      ...control,
      rId: addWorksheetRelationship(relationshipType as RelationshipType, target),
      sourceTarget: target,
    };
  });

  // Round-trip drawing/legacyDrawing references. The referenced part passes
  // through verbatim when its anchors do not map onto options (e.g. OLE
  // object shape representations). With untouched (passthrough) worksheet
  // rels the original id stays valid; when the rels were rebuilt because the
  // same sheet carries comments/tables/…, re-register the passthrough
  // relationship — keeping the source id when it is free — so the reference
  // stays resolvable instead of dangling.
  const resolvePassthroughRid = (typeFragment: string, originalRid: string): string => {
    if (!wsRels) return originalRid;
    const rel = (passthroughRelationships ?? []).find(
      (r) =>
        r.source === wsPath && r.rId === originalRid && r.relationshipType.endsWith(typeFragment),
    );
    if (!rel) return originalRid;
    const existing = wsRels.idOf(rel.relationshipType, rel.target);
    if (existing) return existing;
    if (wsRels.hasId(originalRid)) {
      const n = wsRels.add(rel.relationshipType as RelationshipType, rel.target);
      return `rId${n}`;
    }
    wsRels.addRelationship(originalRid, rel.relationshipType as RelationshipType, rel.target);
    return originalRid;
  };
  // An OLE embedding's relationship is either the classic oleObject type or
  // the native-package type; the r:id is the same either way.
  const resolveEmbeddingRid = (originalRid: string): string => {
    const rid = resolvePassthroughRid("/oleObject", originalRid);
    return rid === originalRid ? resolvePassthroughRid("/package", originalRid) : rid;
  };

  // Every source r:id the sheet XML emits verbatim (OLE objects, controls,
  // printer settings, header/footer VML, custom parts) resolves here — before
  // stringify, so the emitted id is final and no post-stringify patching is
  // needed. The shallow copies keep the caller's options tree untouched.
  let xmlOpts = wsOpts;
  if (hasExternalHyperlinks) {
    xmlOpts = {
      ...xmlOpts,
      hyperlinks: hlOpts.map((hl) =>
        hl.url === undefined ? hl : { ...hl, relationshipId: registerHyperlinkRelationship(hl) },
      ),
    };
  }
  if (hasDataConsolidateRelationships && wsRels) {
    const dataConsolidate = {
      ...wsOpts.dataConsolidate!,
      refs: wsOpts.dataConsolidate!.refs?.map((input) => {
        if (typeof input === "string") return input;
        if (input.target === undefined || input.rId === undefined) return input;
        let rId = wsRels.idOf(RELATIONSHIP_TYPES.externalLinkPath, input.target);
        if (rId === undefined) {
          const preferred = /^rId(\d+)$/.exec(input.rId)?.[1];
          if (preferred && !wsRels.hasId(input.rId)) {
            wsRels.addRelationship(
              Number(preferred),
              RELATIONSHIP_TYPES.externalLinkPath,
              input.target,
              TargetModeType.EXTERNAL,
            );
            rId = input.rId;
          } else {
            rId = `rId${wsRels.add(
              RELATIONSHIP_TYPES.externalLinkPath,
              input.target,
              TargetModeType.EXTERNAL,
            )}`;
          }
        }
        return { ...input, rId };
      }),
    };
    xmlOpts = { ...wsOpts, dataConsolidate };
  }
  if (wsRels && (passthroughRelationships?.length ?? 0) > 0) {
    const oleObjects = wsOpts.oleObjects?.map((ole) => ({
      ...ole,
      rId: ole.rId ? resolveEmbeddingRid(ole.rId) : ole.rId,
      properties:
        ole.properties?.iconRid !== undefined
          ? {
              ...ole.properties,
              iconRid: resolvePassthroughRid("/image", ole.properties.iconRid),
            }
          : ole.properties,
    }));
    const projectedControls = wiredControls.map((control) => ({
      ...control,
      properties:
        control.properties?.iconRid !== undefined
          ? {
              ...control.properties,
              iconRid: resolvePassthroughRid("/image", control.properties.iconRid),
            }
          : control.properties,
    }));
    const pageSetup = wsOpts.pageSetup?.printerSettingsRId
      ? {
          ...wsOpts.pageSetup,
          printerSettingsRId: resolvePassthroughRid(
            "/printerSettings",
            wsOpts.pageSetup.printerSettingsRId,
          ),
        }
      : wsOpts.pageSetup;
    const pivotSelection = wsOpts.pivotSelection?.rId
      ? {
          ...wsOpts.pivotSelection,
          rId: resolvePassthroughRid("/pivotTable", wsOpts.pivotSelection.rId),
        }
      : wsOpts.pivotSelection;
    const legacyDrawingHF = wsOpts.legacyDrawingHF
      ? resolvePassthroughRid("/vmlDrawing", wsOpts.legacyDrawingHF)
      : wsOpts.legacyDrawingHF;
    const customProperties = wsOpts.customProperties?.map((cp) => ({
      ...cp,
      rId: resolvePassthroughRid("/customXml", cp.rId),
    }));
    xmlOpts = {
      ...xmlOpts,
      ...(oleObjects ? { oleObjects } : {}),
      ...(projectedControls ? { controls: projectedControls } : {}),
      ...(pageSetup !== wsOpts.pageSetup ? { pageSetup } : {}),
      ...(pivotSelection !== wsOpts.pivotSelection ? { pivotSelection } : {}),
      ...(legacyDrawingHF !== wsOpts.legacyDrawingHF ? { legacyDrawingHF } : {}),
      ...(customProperties ? { customProperties } : {}),
    };
  }

  // Worksheet uses buildWorksheetXml fast path (zero-allocation string concat)
  xmlOpts = { ...xmlOpts, controls: wiredControls };

  let sheetXml = buildWorksheetXml(xmlOpts, wsContext);

  if (hasMedia) {
    sheetXml = compileSheetDrawing(
      wsOpts,
      i,
      sheetXml,
      ctx,
      mapping,
      state,
      wsRels!,
      passthroughRelationships,
      wsPath,
    );
  }

  // Comments
  if (hasComments) {
    const commentsIdx = i + 1;
    const commentsRel = sourceWorksheetRels.find((rel) =>
      rel.relationshipType.endsWith("/comments"),
    );
    const vmlRel = sourceWorksheetRels.find((rel) => rel.relationshipType.endsWith("/vmlDrawing"));
    const commentsPath = sourceRelationshipPath(commentsRel, `xl/comments${commentsIdx}.xml`);
    const vmlPath = sourceRelationshipPath(vmlRel, `xl/drawings/vmlDrawing${commentsIdx}.vml`);

    // Comments XML (via descriptor)
    const commentsXml = commentsDesc.stringify({ comments: commentOpts }, ctx);
    mapping[`Comments${i}`] = {
      data: XML_DECL + commentsXml,
      path: commentsPath,
    };

    // VML drawing (via descriptor)
    const vmlXml = vmlNotesDesc.stringify(
      {
        comments: commentOpts,
        ...(wsOpts.commentsVmlLayout ? { layout: wsOpts.commentsVmlLayout } : {}),
        ...(wsOpts.commentsVmlShapeType ? { shapeType: wsOpts.commentsVmlShapeType } : {}),
      },
      ctx,
    );
    mapping[`VmlDrawing${i}`] = {
      data:
        (wsOpts.commentsVmlBom ? "\uFEFF" : "") + (wsOpts.commentsVmlSource ?? XML_DECL + vmlXml),
      path: vmlPath,
    };

    // Worksheet rels: comments → comments XML, legacyDrawing → VML file
    addWorksheetRelationship(
      RELATIONSHIP_TYPES.comments,
      commentsRel?.target ?? `../comments${commentsIdx}.xml`,
    );

    const vmlRid = addWorksheetRelationship(
      RELATIONSHIP_TYPES.vmlDrawing,
      vmlRel?.target ?? `../drawings/vmlDrawing${commentsIdx}.vml`,
    );

    // Insert legacyDrawing reference at its CT_Worksheet sequence position.
    sheetXml = editSheetTailMarker(
      sheetXml,
      "<!--LEGACY_DRAWING-->",
      `<legacyDrawing r:id="${vmlRid}"/>`,
    );
  }

  // Background picture
  if (bgImg) {
    const ext = bgImg.type === "jpg" ? "jpeg" : bgImg.type;
    const rawBytes = toUint8Array(bgImg.data, { encoding: "base64" });
    const backgroundRel = sourceWorksheetRels.find((rel) =>
      rel.relationshipType.endsWith("/image"),
    );
    const backgroundPath = sourceRelationshipPath(backgroundRel, "");
    const entry = ctx.media.addMedia(
      rawBytes,
      ext,
      (fileName) => ({
        fileName,
        type: ext,
        data: rawBytes,
        width: 0,
        height: 0,
      }),
      backgroundPath ? backgroundPath.split("/").pop() : undefined,
    );
    state.globalMediaIdx++;
    const bgRid = addWorksheetRelationship(
      IMAGE_REL,
      backgroundRel?.target ?? `../media/${entry.fileName}`,
    );
    sheetXml = editSheetTailMarker(
      sheetXml,
      "<!--BACKGROUND_PICTURE-->",
      `<picture r:id="${bgRid}"/>`,
    );
  }

  if (wsOpts.drawingRid) {
    const rid = escapeXml(resolvePassthroughRid("/drawing", wsOpts.drawingRid));
    sheetXml = editSheetTailMarker(sheetXml, "<!--DRAWING-->", `<drawing r:id="${rid}"/>`);
  }
  if (wsOpts.legacyDrawingRid) {
    const rid = escapeXml(resolvePassthroughRid("/vmlDrawing", wsOpts.legacyDrawingRid));
    sheetXml = editSheetTailMarker(
      sheetXml,
      "<!--LEGACY_DRAWING-->",
      `<legacyDrawing r:id="${rid}"/>`,
    );
  }

  // Form-control and ActiveX metadata parts. ActiveX binary contents remain
  // opaque rawParts; only ax:ocx and the relationship to that binary are built.
  for (const control of wiredControls) {
    const isForm = control.kind === "form";
    if (isForm) {
      state.globalControlPropertiesIdx++;
      const partPath = resolveRelationshipTarget(
        wsPath,
        control.sourceTarget ?? `../ctrlProps/ctrlProp${state.globalControlPropertiesIdx}.xml`,
      );
      mapping[`ControlProperties${state.globalControlPropertiesIdx}`] = {
        data: XML_DECL + controlPropertiesDesc.stringify(control.formControlProperties, ctx),
        path: partPath,
      };
    } else {
      state.globalActiveXIdx++;
      const activeXPath = resolveRelationshipTarget(
        wsPath,
        control.sourceTarget ?? `../activeX/activeX${state.globalActiveXIdx}.xml`,
      );
      const activeXRels = new Relationships(activeXPath);
      let binaryRelationshipId: string | undefined;
      if (control.activeXControl.binaryPath !== undefined) {
        const ownerDir = activeXPath.split("/").slice(0, -1);
        const binarySegments = control.activeXControl.binaryPath.split("/");
        let common = 0;
        while (
          common < ownerDir.length &&
          common < binarySegments.length - 1 &&
          ownerDir[common] === binarySegments[common]
        ) {
          common++;
        }
        const upSegments = ownerDir.slice(common).map(() => "..");
        const downSegments = binarySegments.slice(common);
        const target = [...upSegments, ...downSegments].join("/") || ".";
        binaryRelationshipId = `rId${activeXRels.add(
          RELATIONSHIP_TYPES.activeXControlBinary,
          target,
        )}`;
      }
      mapping[`ActiveXControl${state.globalActiveXIdx}`] = {
        data:
          XML_DECL +
          activeXControlDesc.stringify(
            {
              ...control.activeXControl,
              ...(binaryRelationshipId ? { relationshipId: binaryRelationshipId } : {}),
            } as ActiveXControlOptions,
            ctx,
          ),
        path: activeXPath,
      };
      mapping[`ActiveXControlRels${state.globalActiveXIdx}`] = {
        data: XML_DECL + activeXRels.serialize(),
        path: partPathToRelsPath(activeXPath),
      };
    }
  }

  // Pivot tables
  if (hasPivots) {
    compileSheetPivots(
      wsOpts,
      worksheetConfigs,
      ctx,
      mapping,
      state,
      wsRels!,
      passthroughRelationships,
      wsPath,
      sheetName,
    );
  }

  // Tables (list objects)
  const wsTableParts: TablePartReference[] = [];
  if (hasTables) {
    const sourceTableRels = sourceWorksheetRels.filter((rel) =>
      rel.relationshipType.endsWith("/table"),
    );
    for (const [tableIndex, tbl] of tableOpts.entries()) {
      state.globalTableIdx++;
      const tableIdx = state.globalTableIdx;

      // A table without columns cannot form valid tableColumns XML — skip
      // instead of emitting a broken part (defensive; parse filters these).
      if (!tbl.columns?.length) continue;
      const tableRel = sourceTableRels[tableIndex];
      const tablePath = sourceRelationshipPath(tableRel, `xl/tables/table${tableIdx}.xml`);

      // Generate table XML
      const tableXmlStr = XML_DECL + tableDesc.stringify({ ...tbl, id: tbl.id ?? tableIdx }, ctx);
      mapping[`Table${tableIdx}`] = {
        data: tableXmlStr,
        path: tablePath,
      };

      // Worksheet rels → table
      const tblRid = addWorksheetRelationship(
        RELATIONSHIP_TYPES.table,
        tableRel?.target ?? `../tables/table${tableIdx}.xml`,
      );

      wsTableParts.push({ rId: tblRid, order: tbl.tablePartOrder });
      state.allTableParts.push({ rId: tblRid, order: tbl.tablePartOrder });
    }
  }

  // Query tables
  if (hasQueryTables) {
    const allSourceQueryTableRels = (passthroughRelationships ?? []).filter((rel) =>
      rel.relationshipType.endsWith("/queryTable"),
    );
    const worksheetSourceQueryTableRels = sourceWorksheetRels.filter((rel) =>
      rel.relationshipType.endsWith("/queryTable"),
    );
    const tableOwnedQueryTables = (wsOpts.tables ?? []).filter((table) => table.queryTablePath);
    const tableOwnedSourcePaths = new Set(
      tableOwnedQueryTables.map((table) => table.queryTablePath!),
    );
    let worksheetQueryTableIdx = 0;
    for (const [queryTableIndex, qt] of queryTableOpts.entries()) {
      state.globalQueryTableIdx++;
      const relsBaseDir = (relsPath: string): string => {
        // "xl/tables/_rels/table1.xml.rels" → "xl/tables/table1.xml"
        const parent = relsPath.replace(/\/_rels\/[^/]+\.rels$/, "");
        const fileName = /_rels\/([^/]+)\.rels$/.exec(relsPath)?.[1] ?? "";
        return `${parent}/${fileName}`;
      };
      const sourceRelationship = qt.sourcePath
        ? allSourceQueryTableRels.find(
            (rel) =>
              resolveRelationshipTarget(relsBaseDir(rel.source), rel.target) === qt.sourcePath,
          )
        : undefined;
      const queryTablePath =
        qt.sourcePath ??
        (sourceRelationship
          ? resolveRelationshipTarget(
              relsBaseDir(sourceRelationship.source),
              sourceRelationship.target,
            )
          : `xl/queryTables/queryTable${state.globalQueryTableIdx}.xml`);
      mapping[`QueryTable${state.globalQueryTableIdx}`] = {
        data: XML_DECL + queryTableDesc.stringify(qt, ctx),
        path: queryTablePath,
      };
      const target =
        sourceRelationship?.target ?? `../queryTables/queryTable${state.globalQueryTableIdx}.xml`;
      if (qt.sourcePath && tableOwnedSourcePaths.has(qt.sourcePath)) {
        // Query table is table-owned: emit a table .rels that references it.
        const ownerTable = tableOwnedQueryTables.find(
          (table) => table.queryTablePath === qt.sourcePath,
        );
        const tableSourcePath = ownerTable?.sourcePath ?? "";
        const tableRelsPath = tableSourcePath ? partPathToRelsPath(tableSourcePath) : "";
        if (tableRelsPath) {
          const tableRelationships = new Relationships(tableRelsPath);
          const relTarget = `../${queryTablePath.replace(/^xl\//, "")}`;
          tableRelationships.add(RELATIONSHIP_TYPES.queryTable, relTarget);
          if (sourceRelationship) {
            const preferredId = /^rId(\d+)$/.exec(sourceRelationship.rId)?.[1];
            if (preferredId) {
              tableRelationships.renameEntryByType(
                RELATIONSHIP_TYPES.queryTable,
                Number(preferredId),
              );
            }
          }
          mapping[`QueryTableOwnedRels${state.globalQueryTableIdx}`] = {
            data: XML_DECL + tableRelationships.serialize(),
            path: tableRelsPath,
          };
        }
        // Query table relationship is table-owned — do not add a worksheet rel.
      } else if (sourceRelationship?.source === wsPath) {
        addWorksheetRelationship(RELATIONSHIP_TYPES.queryTable, target);
        worksheetQueryTableIdx++;
      } else if (worksheetSourceQueryTableRels[queryTableIndex]) {
        addWorksheetRelationship(
          RELATIONSHIP_TYPES.queryTable,
          worksheetSourceQueryTableRels[queryTableIndex].target,
        );
        worksheetQueryTableIdx++;
      }
      // No worksheet-level source relationship → table .rels already owns it
      // through passthrough; do not add a worksheet rel.
    }
  }

  // Single-cell XML tables
  if (singleXmlCellOpts.length > 0) {
    const sourceSingleXmlCellRels = sourceWorksheetRels.filter((rel) =>
      rel.relationshipType.endsWith("/tableSingleCells"),
    );
    state.globalSingleXmlCellsIdx++;
    const singleXmlCellRel = sourceSingleXmlCellRels[0];
    mapping[`TableSingleCells${state.globalSingleXmlCellsIdx}`] = {
      data: XML_DECL + singleXmlCellsDesc.stringify({ cells: singleXmlCellOpts }, ctx),
      path: sourceRelationshipPath(
        singleXmlCellRel,
        `xl/tables/tableSingleCells${state.globalSingleXmlCellsIdx}.xml`,
      ),
    };
    addWorksheetRelationship(
      RELATIONSHIP_TYPES.tableSingleCells,
      singleXmlCellRel?.target ?? `../tables/tableSingleCells${state.globalSingleXmlCellsIdx}.xml`,
    );
  }

  // WebExtension parts — typed model output with snapshot image rels
  const emittedWebExtensionParts = new Set<string>();
  for (const we of webExtensionOpts) {
    const wePart = we.part;
    if (!wePart || !wePart.sourcePath || emittedWebExtensionParts.has(wePart.sourcePath)) {
      continue;
    }
    emittedWebExtensionParts.add(wePart.sourcePath);
    const wePath = wePart.sourcePath;
    // Emit the webextension's own rels only when the source part carried an
    // image relationship (we:snapshot with r:embed). The part XML is always
    // rebuilt from the typed model.
    if (wePart.snapshotRId !== undefined) {
      const weRels = new Relationships(wePath);
      if (we.snapshotSourcePath) {
        const weImageTarget = `../${we.snapshotSourcePath.replace(/^xl\//, "")}`;
        const sourceImageRel = sourceWorksheetRels.find(
          (rel) =>
            rel.relationshipType.endsWith("/image") &&
            resolveRelationshipTarget(wePath, rel.target) === we.snapshotSourcePath,
        );
        const preferredId = /^rId(\d+)$/.exec(sourceImageRel?.rId ?? wePart.snapshotRId)?.[1];
        weRels.addRelationship(
          preferredId ? Number(preferredId) : weRels.nextRelationshipId,
          RELATIONSHIP_TYPES.image,
          weImageTarget,
        );
      }
      mapping[`WebExtensionRels${emittedWebExtensionParts.size}`] = {
        data: XML_DECL + weRels.serialize(),
        path: partPathToRelsPath(wePath),
      };
    }
    mapping[`WebExtension${emittedWebExtensionParts.size}`] = {
      data: XML_DECL + webExtensionPartDesc.stringify(wePart, ctx),
      path: wePath,
    };
  }

  // Pre-render pivot table data into sheetData
  if (hasPivots) {
    const rendered = renderPivotSheetData(
      pivotOpts.filter((pivot) => pivot.mode !== "definition"),
      worksheetConfigs,
      ctx.sharedStrings,
      sheetName,
    );
    if (rendered.sheetData.length > 0) {
      // Replace empty <sheetData/> or <sheetData></sheetData> with rendered data
      sheetXml = sheetXml.replace(/<sheetData\/>|<sheetData><\/sheetData>/, rendered.sheetData);
      // Inject <dimension> before <sheetViews> (XSD sequence order: dimension before sheetViews)
      if (!sheetXml.includes("<dimension")) {
        sheetXml = sheetXml.replace(
          "<sheetViews",
          `<dimension ref="${rendered.dimensionRef}"/><sheetViews`,
        );
      }
    }
  }

  // Insert tableParts at their CT_Worksheet sequence position
  if (wsTableParts.length > 0) {
    wsTableParts.sort(
      (left, right) =>
        (Number(/^rId(\d+)$/.exec(left.rId)?.[1] ?? 0) || 0) -
        (Number(/^rId(\d+)$/.exec(right.rId)?.[1] ?? 0) || 0),
    );
    sheetXml = editSheetTailMarker(
      sheetXml,
      "<!--TABLE_PARTS-->",
      buildTablePartsXml(wsTableParts),
    );
  } else if (wsOpts.preserveEmptyTableParts) {
    sheetXml = editSheetTailMarker(sheetXml, "<!--TABLE_PARTS-->", '<tableParts count="0"/>');
  }
  // Strip placeholders the compiler did not replace (drawing, legacyDrawing,
  // tableParts) — their existence depends on relationships owned here.
  sheetXml = stripWorksheetPlaceholders(sheetXml);

  // Round-trip: re-emit sheet relationships the model did not absorb
  // (printerSettings above all). Rebuilding the rels for tables/comments/…
  // must not drop part associations the worksheet XML never references by
  // r:id. Targets are passthrough paths that never move, so kind+target
  // identifies an instance the model already registered (drawing above).
  if (wsRels) {
    // Model-unabsorbed worksheet rels — claim semantics keep the source id
    // (reserved up front, so it is free unless the model genuinely took it)
    // instead of renumbering onto an id the sheet XML still references.
    for (const rel of passthroughRelationships ?? []) {
      if (rel.source !== wsPath) continue;
      wsRels.claimSourceRel(rel);
    }
  }

  // Write worksheet rels if needed
  if (wsRels) {
    mapping[`WorksheetRels${i}`] = {
      data: XML_DECL + wsRels.serialize(),
      path: partPathToRelsPath(wsPath),
    };
  }

  mapping[`Worksheet${i}`] = {
    data: sheetXml,
    path: wsPath,
  };
}

/** Full-page default frame for a fresh chartsheet chart (≈9.70×6.35 in). */
const CHARTSHEET_DEFAULT_CX = 9308969;
const CHARTSHEET_DEFAULT_CY = 6096000;

/** Compile all chartsheets: chart part, chartsheet/drawing XML and rels. */
function compileChartsheets(
  chartsheetConfigs: ChartsheetOptions[],
  ctx: XlsxWriteContext,
  mapping: Record<string, { data: string; path: string }>,
  state: WorksheetCompileState,
  passthroughRelationships?: readonly PassthroughRelationship[],
): void {
  // Chartsheets — chart-only sheets
  for (const [i, csOpts] of chartsheetConfigs.entries()) {
    // Register chart in the charts collection. Skip a chartsheet whose chart
    // could not be resolved (missing drawing/chart part in a broken source) —
    // safer than crashing the whole workbook compile.
    const chartDef = csOpts.chart;
    if (!chartDef) continue;
    const csChartGlobalIdx = ctx.charts.array.length;
    const csChartKey = `cs_chart_${csChartGlobalIdx}`;
    const sourceDrawingRel = (passthroughRelationships ?? []).find(
      (rel) =>
        rel.source === `xl/chartsheets/sheet${i + 1}.xml` &&
        rel.relationshipType === RELATIONSHIP_TYPES.drawing,
    );
    const drawingPath = csOpts.sourceDrawingPath ?? `xl/drawings/drawing${i + 1}.xml`;
    const drawingIndex = Number(
      drawingPath.match(/drawing(\d+)\.xml$/)?.[1] ?? state.nextDrawingIndex++,
    );
    const chartPath = csOpts.sourceChartPath ?? `xl/charts/chart${csChartGlobalIdx + 1}.xml`;
    state.chartPaths.set(csChartKey, chartPath);
    ctx.chartOptions.set(csChartKey, {
      externalData: chartDef.externalData,
      externalLink: csOpts.externalLink,
      externalLinkPath: csOpts.externalLinkPath,
      externalDataRelationshipType: csOpts.externalDataRelationshipType,
      externalDataRelationshipTarget: csOpts.externalDataRelationshipTarget,
    });
    const csChartRels = new Relationships();
    const csChartXml = bindMediaPlaceholders(
      preserveChartDecimalAttributes(chartSpaceDesc.stringify(chartDef, ctx) ?? ""),
      ctx.media,
      csChartRels,
    );
    const csUserShapes = chartDef.userShapes
      ? {
          ...buildUserShapesData(chartDef.userShapes),
          ...(chartDef.userShapes.path ? { path: chartDef.userShapes.path } : {}),
        }
      : undefined;
    ctx.charts.addChart(csChartKey, {
      key: csChartKey,
      chartSpaceXml: csChartXml,
      ...(csUserShapes ? { userShapes: csUserShapes } : {}),
      ...(csChartRels.relationshipCount > 0 ? { relsXml: csChartRels.serialize() } : {}),
    });

    // Chartsheet relationships: drawing (required)
    const csRels = new Relationships();
    const csDrawingIdx = drawingIndex;
    const drawingTarget = sourceDrawingRel?.target ?? `../drawings/drawing${csDrawingIdx}.xml`;
    csRels.addRelationship(
      csOpts.sourceDrawingRelationshipId ?? 1,
      RELATIONSHIP_TYPES.drawing,
      drawingTarget,
    );
    const drawingRId = csRels.idOf(RELATIONSHIP_TYPES.drawing, drawingTarget) ?? "rId1";

    // Round-trip: re-emit chartsheet relationships the model did not absorb
    // (printerSettings above all) — same contract as worksheet rels.
    for (const rel of passthroughRelationships ?? []) {
      if (rel.source !== `xl/chartsheets/sheet${i + 1}.xml`) continue;
      if (csRels.hasRelationship(rel.relationshipType, rel.target)) continue;
      csRels.addRelationship(
        csRels.nextRelationshipId,
        rel.relationshipType as RelationshipType,
        rel.target,
      );
    }

    // Drawing rels: chart reference
    const csDrawingRels = new Relationships();
    const sourceChartRel = (passthroughRelationships ?? []).find(
      (rel) => rel.source === drawingPath && rel.relationshipType === RELATIONSHIP_TYPES.chart,
    );
    csDrawingRels.addRelationship(
      1,
      RELATIONSHIP_TYPES.chart,
      sourceChartRel?.target ?? `../charts/chart${csChartGlobalIdx + 1}.xml`,
    );

    // Drawing XML with chart anchor — reuses the parts/drawing anchor and
    // graphicFrame builders (no hand-rolled xdr emitter). Round-tripped
    // sheets keep the source anchor geometry (the rendered chart size —
    // Excel writes it back verbatim on save); fresh sheets anchor at origin
    // with the full-page frame. The graphicFrame xfrm stays 0×0, Excel's own
    // chartsheet form: it sizes from the anchor ext and zeroes xfrm on save.
    const frame = graphicFrameXml(
      csOpts.shapeId ?? 1,
      { ...csOpts, name: csOpts.chartName },
      `Chart ${i + 1}`,
      "rId1",
      0,
      0,
      ctx,
      {
        frameLocks: csOpts.frameLocks,
        macro: csOpts.macro,
      },
    );
    const anchor = wrapAnchor(
      {
        anchorType: "absolute",
        col: 1,
        row: 1,
        absoluteX: csOpts.absoluteX ?? 0,
        absoluteY: csOpts.absoluteY ?? 0,
        extentCx: csOpts.extentCx ?? CHARTSHEET_DEFAULT_CX,
        extentCy: csOpts.extentCy ?? CHARTSHEET_DEFAULT_CY,
      },
      `${frame}<xdr:clientData/>`,
    );
    const csDrawingXml = `<xdr:wsDr xmlns:xdr="${XDR_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">${anchor}</xdr:wsDr>`;

    mapping[`ChartsheetDrawing${i}`] = {
      data: csDrawingXml,
      path: `xl/drawings/drawing${csDrawingIdx}.xml`,
    };
    mapping[`ChartsheetDrawingRels${i}`] = {
      data: XML_DECL + csDrawingRels.serialize(),
      path: `xl/drawings/_rels/drawing${csDrawingIdx}.xml.rels`,
    };

    mapping[`ChartsheetRels${i}`] = {
      data: XML_DECL + csRels.serialize(),
      path: `xl/chartsheets/_rels/sheet${i + 1}.xml.rels`,
    };

    // pageSetup r:id → printerSettings: the rebuilt rels renumber every id
    // (drawing holds rId1), so remap the source id onto the re-emitted
    // relationship — otherwise the passthrough .bin stays orphaned and Excel
    // discards it on the next save.
    let csPageSetup = csOpts.pageSetup;
    if (csPageSetup?.printerSettingsRId) {
      const srcRid = csPageSetup.printerSettingsRId;
      const rel = (passthroughRelationships ?? []).find(
        (r) =>
          r.source === `xl/chartsheets/sheet${i + 1}.xml` &&
          r.rId === srcRid &&
          r.relationshipType.endsWith("/printerSettings"),
      );
      const rid = rel ? (csRels.idOf(rel.relationshipType, rel.target) ?? srcRid) : srcRid;
      if (rid !== srcRid) csPageSetup = { ...csPageSetup, printerSettingsRId: rid };
    }
    mapping[`Chartsheet${i}`] = {
      data:
        XML_DECL + chartsheetDesc.stringify({ ...csOpts, drawingRId, pageSetup: csPageSetup }, ctx),
      path: `xl/chartsheets/sheet${i + 1}.xml`,
    };
  }
}

/** Compile all dialog sheets (legacy Excel 5.0 dialog sheets). */
function compileDialogsheets(
  dialogsheetConfigs: DialogsheetOptions[],
  ctx: XlsxWriteContext,
  mapping: Record<string, { data: string; path: string }>,
  passthroughRelationships?: readonly PassthroughRelationship[],
): void {
  // Dialogsheets — legacy Excel 5.0 dialog sheets. Legacy drawing relationships
  // are rebuilt from the captured source relationship topology.
  for (const [i, dsOpts] of dialogsheetConfigs.entries()) {
    const path = dsOpts.sourcePath ?? `xl/dialogSheets/sheet${i + 1}.xml`;
    mapping[`Dialogsheet${i}`] = {
      data: XML_DECL + dialogsheetDesc.stringify(dsOpts, ctx),
      path,
    };
    const sourceRels = (passthroughRelationships ?? []).filter(
      (relationship) => relationship.source === path,
    );
    if (sourceRels.length === 0) continue;
    const relationships = new Relationships(path);
    for (const relationship of sourceRels) {
      relationships.add(
        relationship.relationshipType as RelationshipType,
        relationship.target,
        relationship.targetMode,
      );
    }
    for (const relationship of sourceRels) {
      const preferredId = /^rId(\d+)$/.exec(relationship.rId)?.[1];
      if (preferredId) {
        relationships.renameEntryByType(relationship.relationshipType, Number(preferredId));
      }
    }
    mapping[`DialogsheetRels${i}`] = {
      data: XML_DECL + relationships.serialize(),
      path: partPathToRelsPath(path),
    };
  }
}

/**
 * Compile shared-workbook revisions (xl/revisionHeaders.xml +
 * xl/revisions/revisionN.xml + xl/users.xml). CT_Workbook has no revision
 * element — parts are discovered via workbook.xml.rels + [Content_Types].
 */
function compileRevisionLogs(
  rl: NonNullable<WorkbookOptions["revisionLog"]>,
  ctx: XlsxWriteContext,
  mapping: Record<string, { data: string; path: string }>,
): void {
  const REV_HEADERS_REL = RELATIONSHIP_TYPES.revisionHeaders;
  const REV_LOG_REL = RELATIONSHIP_TYPES.revisionLog;
  const USERS_REL = rl.usersRelationshipType ?? RELATIONSHIP_TYPES.users;
  const relativeTarget = (from: string, to: string): string => {
    const fromDirs = from.split("/").slice(0, -1);
    const toDirs = to.split("/");
    let common = 0;
    while (
      common < fromDirs.length &&
      common < toDirs.length &&
      fromDirs[common] === toDirs[common]
    ) {
      common++;
    }
    return "../".repeat(fromDirs.length - common) + toDirs.slice(common).join("/");
  };

  // xl/revisionHeaders.xml — target of an implicit relationship from the workbook.
  // Round-trip writes the part at its source path; fresh authoring uses the
  // conventional xl/revisionHeaders.xml.
  const revHeadersPath = rl.headersPath ?? "revisionHeaders.xml";
  mapping["RevisionHeaders"] = {
    data: XML_DECL + (revisionHeadersDesc.stringify(rl.headers, ctx) ?? ""),
    path: `xl/${revHeadersPath}`,
  };
  // The passthrough rel capture may have claimed the source relationship
  // already — dedupe so the rebuilt rels keeps one entry.
  if (ctx.workbookRels.idOf(REV_HEADERS_REL, revHeadersPath) === undefined) {
    ctx.workbookRels.addRelationship(
      ctx.workbookRels.nextRelationshipId,
      REV_HEADERS_REL,
      revHeadersPath,
    );
  }

  // One revision log per header entry, plus revisionHeaders.xml.rels pointing to each.
  const revHeadersRels = new Relationships();
  for (const [i, log] of rl.logs.entries()) {
    const logPath = log.path ?? `revisions/revision${i + 1}.xml`;
    mapping[`RevisionLog${i}`] = {
      data: XML_DECL + (revisionLogDesc.stringify(log, ctx) ?? ""),
      path: `xl/${logPath}`,
    };
    const preferredRid = /^rId(\d+)$/.exec(rl.headers.headers[i]?.rId ?? "")?.[1];
    const rid =
      preferredRid && !revHeadersRels.hasId(`rId${preferredRid}`)
        ? Number(preferredRid)
        : revHeadersRels.nextRelationshipId;
    revHeadersRels.addRelationship(
      rid,
      REV_LOG_REL,
      log.relationshipTarget ?? relativeTarget(revHeadersPath, logPath),
    );
  }
  mapping["RevisionHeadersRels"] = {
    data: XML_DECL + revHeadersRels.serialize(),
    path: partPathToRelsPath(`xl/${revHeadersPath}`),
  };

  // xl/users.xml (optional) — honors the source path on round-trip.
  if (rl.users) {
    const usersXml = usersDesc.stringify(rl.users, ctx);
    if (usersXml) {
      const usersPath = rl.usersPath ?? "users.xml";
      mapping["Users"] = { data: XML_DECL + usersXml, path: `xl/${usersPath}` };
      if (ctx.workbookRels.idOf(USERS_REL, usersPath) === undefined) {
        ctx.workbookRels.addRelationship(ctx.workbookRels.nextRelationshipId, USERS_REL, usersPath);
      }
    }
  }
}

// ── Pure helper functions ──

function buildWorkbookRelationships(
  rels: Relationships,
  worksheetTargets: readonly string[],
  csCount: number,
  dialogsheetTargets: readonly string[] = [],
  includeStyles = true,
  includeTheme = true,
  themeTarget = "theme/theme1.xml",
): string[] {
  const add = (type: RelationshipType, target: string): string => {
    const existing = rels.idOf(type, target);
    if (existing !== undefined) return existing;
    const rid = rels.nextRelationshipId;
    rels.addRelationship(rid, type, target);
    return `rId${rid}`;
  };
  const ids: string[] = [];
  for (const target of worksheetTargets) {
    ids.push(add(RELATIONSHIP_TYPES.worksheet, target));
  }
  for (let i = 0; i < csCount; i++) {
    ids.push(add(RELATIONSHIP_TYPES.chartsheet, `chartsheets/sheet${i + 1}.xml`));
  }
  for (const target of dialogsheetTargets) {
    ids.push(add(RELATIONSHIP_TYPES.dialogsheet, target));
  }
  if (includeStyles) add(RELATIONSHIP_TYPES.styles, "styles.xml");
  if (includeTheme) add(RELATIONSHIP_TYPES.theme, themeTarget);
  return ids;
}

function hasMetadataContent(metadata: MetadataOptions): boolean {
  return (
    (metadata.types?.length ?? 0) > 0 ||
    (metadata.strings?.length ?? 0) > 0 ||
    (metadata.mdx?.length ?? 0) > 0 ||
    (metadata.futureMetadata?.length ?? 0) > 0 ||
    (metadata.cellMetadata?.length ?? 0) > 0 ||
    (metadata.valueMetadata?.length ?? 0) > 0
  );
}
