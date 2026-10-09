import type {
  DocumentOptions,
  RunPropertiesOptions,
  W14RunEffectsOptions,
} from "@office-open/docx";
import type { PresentationOptions } from "@office-open/pptx";
import type { WorkbookOptions } from "@office-open/xlsx";

export type OoxmlFamily = "document" | "presentation" | "workbook";

export interface OoxmlExecutableCase {
  id: string;
  schema: string;
  family: OoxmlFamily;
  element: string;
  attributes: readonly string[];
  options: unknown;
  partPath: string;
  projection: (options: never) => unknown;
  validationSchema?: string;
  validationDirectory?: string;
  expected: unknown;
}

interface DocumentExecutableCase extends Omit<
  OoxmlExecutableCase,
  "family" | "options" | "projection"
> {
  family: "document";
  options: DocumentOptions;
  projection: (options: DocumentOptions) => unknown;
}

interface PresentationExecutableCase extends Omit<
  OoxmlExecutableCase,
  "family" | "options" | "projection"
> {
  family: "presentation";
  options: PresentationOptions;
  projection: (options: PresentationOptions) => unknown;
}

interface WorkbookExecutableCase extends Omit<
  OoxmlExecutableCase,
  "family" | "options" | "projection"
> {
  family: "workbook";
  options: WorkbookOptions;
  projection: (options: WorkbookOptions) => unknown;
}

export type OoxmlExecutableCaseInput =
  | DocumentExecutableCase
  | PresentationExecutableCase
  | WorkbookExecutableCase;

export interface OoxmlExecutableRegistry {
  cases: readonly OoxmlExecutableCaseInput[];
}

const commentIdCase: OoxmlExecutableCaseInput = {
  id: "w16cid.commentsIds.required",
  schema: "w16cid",
  family: "document",
  element: "w16cid:commentId",
  attributes: ["w16cid:commentId/@paraId", "w16cid:commentId/@durableId"],
  options: {
    sections: [{ children: [] }],
    commentsIds: [{ paraId: "7FD6C115", durableId: "1EC7B1B1" }],
  },
  partPath: "word/commentsIds.xml",
  projection: (options) => options.commentsIds,
  expected: [{ paraId: "7FD6C115", durableId: "1EC7B1B1" }],
};

const extensibleCases: readonly OoxmlExecutableCaseInput[] = [
  {
    id: "w16cex.commentsExtensible.omitted-optional",
    schema: "w16cex",
    family: "document",
    element: "w16cex:commentExtensible",
    attributes: ["w16cex:commentExtensible/@durableId"],
    options: {
      sections: [{ children: [] }],
      commentsExtensible: [{ durableId: "1EC7B1B1" }],
    },
    partPath: "word/commentsExtensible.xml",
    projection: (options) => options.commentsExtensible,
    expected: [{ durableId: "1EC7B1B1" }],
  },
  {
    id: "w16cex.commentsExtensible.date-utc",
    schema: "w16cex",
    family: "document",
    element: "w16cex:commentExtensible",
    attributes: ["w16cex:commentExtensible/@durableId", "w16cex:commentExtensible/@dateUtc"],
    options: {
      sections: [{ children: [] }],
      commentsExtensible: [{ durableId: "1EC7B1B1", dateUtc: "2026-01-01T00:00:00Z" }],
    },
    partPath: "word/commentsExtensible.xml",
    projection: (options) => options.commentsExtensible,
    expected: [{ durableId: "1EC7B1B1", dateUtc: "2026-01-01T00:00:00Z" }],
  },
  {
    id: "w16cex.commentsExtensible.placeholder-false",
    schema: "w16cex",
    family: "document",
    element: "w16cex:commentExtensible",
    attributes: [
      "w16cex:commentExtensible/@durableId",
      "w16cex:commentExtensible/@intelligentPlaceholder",
    ],
    options: {
      sections: [{ children: [] }],
      commentsExtensible: [{ durableId: "1EC7B1B1", intelligentPlaceholder: false }],
    },
    partPath: "word/commentsExtensible.xml",
    projection: (options) => options.commentsExtensible,
    expected: [{ durableId: "1EC7B1B1", intelligentPlaceholder: false }],
  },
  {
    id: "w16cex.commentsExtensible.placeholder-true",
    schema: "w16cex",
    family: "document",
    element: "w16cex:commentExtensible",
    attributes: [
      "w16cex:commentExtensible/@durableId",
      "w16cex:commentExtensible/@intelligentPlaceholder",
    ],
    options: {
      sections: [{ children: [] }],
      commentsExtensible: [{ durableId: "1EC7B1B1", intelligentPlaceholder: true }],
    },
    partPath: "word/commentsExtensible.xml",
    projection: (options) => options.commentsExtensible,
    expected: [{ durableId: "1EC7B1B1", intelligentPlaceholder: true }],
  },
  {
    id: "w16cex.commentsExtensible.ext",
    schema: "w16cex",
    family: "document",
    element: "w16cex:commentExtensible",
    attributes: ["w16cex:commentExtensible/@durableId"],
    options: {
      sections: [{ children: [] }],
      commentsExtensible: [
        {
          durableId: "1EC7B1B1",
          ext: '<w16:ext w16:uri="example"><w16:value>example</w16:value></w16:ext>',
        },
      ],
    },
    partPath: "word/commentsExtensible.xml",
    projection: (options) => options.commentsExtensible,
    expected: [
      {
        durableId: "1EC7B1B1",
        ext: '<w16:ext w16:uri="example"><w16:value>example</w16:value></w16:ext>',
      },
    ],
  },
];

const lockedCanvasCase: OoxmlExecutableCaseInput = {
  id: "dml-lc.lockedCanvas.basic",
  schema: "dml-lc",
  family: "presentation",
  element: "lc:lockedCanvas",
  attributes: [],
  options: {
    slides: [
      {
        children: [
          {
            lockedCanvas: {
              id: 100,
              name: "Executable Locked Canvas",
              x: 100,
              y: 200,
              width: 300,
              height: 400,
            },
          },
        ],
      },
    ],
  },
  partPath: "ppt/slides/slide1.xml",
  validationSchema: "pml.xsd",
  validationDirectory: "ooxml-schemas/transitional",
  projection: (options) => {
    const child = options.slides?.[0]?.children?.[0];
    return child && "lockedCanvas" in child ? child.lockedCanvas : undefined;
  },
  expected: {
    id: 100,
    name: "Executable Locked Canvas",
    locking: { noGrp: true },
    x: 100,
    y: 200,
    width: 300,
    height: 400,
  },
};

const w15SdtExtensionsCase: OoxmlExecutableCaseInput = {
  id: "w15.sdt.color-and-data-binding",
  schema: "w15",
  family: "document",
  element: "w15:color",
  attributes: [
    "w15:color/@val",
    "w15:color/@themeColor",
    "w15:color/@themeTint",
    "w15:color/@themeShade",
    "w15:dataBinding/@xpath",
    "w15:dataBinding/@storeItemID",
    "w15:dataBinding/@prefixMappings",
  ],
  options: {
    sections: [
      {
        children: [
          {
            sdt: {
              properties: {
                w15Color: {
                  val: "FF0000",
                  themeColor: "accent1",
                  themeTint: "99",
                  themeShade: "BF",
                },
                w15DataBinding: {
                  xpath: "/invoice/customer",
                  storeItemID: "C9C24E26-8E6D-4C10-9B52-5A9D1C37D92E",
                  prefixMappings: "xmlns:invoice='http://example.com/invoice'",
                },
              },
              children: [],
            },
          },
        ],
      },
    ],
  },
  partPath: "word/document.xml",
  validationSchema: "wml.xsd",
  validationDirectory: "ooxml-schemas/transitional",
  projection: (options: DocumentOptions) => {
    const child = options.sections?.[0]?.children?.[0];
    return child && "sdt" in child ? child.sdt.properties : undefined;
  },
  expected: {
    w15Color: {
      val: "FF0000",
      themeColor: "accent1",
      themeTint: "99",
      themeShade: "BF",
    },
    w15DataBinding: {
      xpath: "/invoice/customer",
      storeItemID: "C9C24E26-8E6D-4C10-9B52-5A9D1C37D92E",
      prefixMappings: "xmlns:invoice='http://example.com/invoice'",
    },
    childOrder: ["w15DataBinding", "w15Color"],
  },
};

const w15SdtRepeatingSectionCase: OoxmlExecutableCaseInput = {
  id: "w15.sdt.repeating-section-properties",
  schema: "w15",
  family: "document",
  element: "w15:appearance",
  attributes: [
    "w15:appearance/@val",
    "w15:sectionTitle/@val",
    "w15:doNotAllowInsertDeleteSection/@val",
    "w15:webExtensionLinked/@val",
    "w15:webExtensionCreated/@val",
  ],
  options: {
    sections: [
      {
        children: [
          {
            sdt: {
              properties: {
                appearance: "tags",
                repeatingSection: {
                  sectionTitle: "Executable",
                  doNotAllowInsertDeleteSection: false,
                },
                webExtensionLinked: true,
                webExtensionCreated: false,
              },
              children: [],
            },
          },
        ],
      },
    ],
  },
  partPath: "word/document.xml",
  validationSchema: "wml.xsd",
  validationDirectory: "ooxml-schemas/transitional",
  projection: (options: DocumentOptions) => {
    const child = options.sections?.[0]?.children?.[0];
    return child && "sdt" in child ? child.sdt.properties : undefined;
  },
  expected: {
    appearance: "tags",
    repeatingSection: {
      sectionTitle: "Executable",
      doNotAllowInsertDeleteSection: false,
    },
    webExtensionLinked: true,
    webExtensionCreated: false,
    childOrder: ["appearance", "repeatingSection", "webExtensionLinked", "webExtensionCreated"],
  },
};

const w15SdtRepeatingSectionItemCase: OoxmlExecutableCaseInput = {
  id: "w15.sdt.repeating-section-item",
  schema: "w15",
  family: "document",
  element: "w15:repeatingSectionItem",
  attributes: [],
  options: {
    sections: [
      {
        children: [{ sdt: { properties: { repeatingSectionItem: true }, children: [] } }],
      },
    ],
  },
  partPath: "word/document.xml",
  validationSchema: "wml.xsd",
  validationDirectory: "ooxml-schemas/transitional",
  projection: (options: DocumentOptions) => {
    const child = options.sections?.[0]?.children?.[0];
    return child && "sdt" in child ? child.sdt.properties : undefined;
  },
  expected: { repeatingSectionItem: true, childOrder: ["repeatingSectionItem"] },
};

const w15CommentsExtendedCase: OoxmlExecutableCaseInput = {
  id: "w15.comments-extended",
  schema: "w15",
  family: "document",
  element: "w15:commentsEx",
  attributes: ["w15:commentEx/@paraId", "w15:commentEx/@paraIdParent", "w15:commentEx/@done"],
  options: {
    sections: [{ children: [] }],
    commentsExtended: [{ paraId: "7FD6C115", paraIdParent: "6CBA2F0C", done: true }],
  },
  partPath: "word/commentsExtended.xml",
  projection: (options: DocumentOptions) => options.commentsExtended,
  expected: [{ paraId: "7FD6C115", paraIdParent: "6CBA2F0C", done: true }],
};

const w15PeopleCase: OoxmlExecutableCaseInput = {
  id: "w15.people-presence",
  schema: "w15",
  family: "document",
  element: "w15:people",
  attributes: ["w15:person/@author", "w15:presenceInfo/@providerId", "w15:presenceInfo/@userId"],
  options: {
    sections: [{ children: [] }],
    people: [
      {
        author: "Executable Author",
        presenceInfo: { providerId: "None", userId: "author-1" },
      },
    ],
  },
  partPath: "word/people.xml",
  projection: (options: DocumentOptions) => options.people,
  expected: [
    {
      author: "Executable Author",
      presenceInfo: { providerId: "None", userId: "author-1" },
    },
  ],
};

const w15SettingsCase: OoxmlExecutableCaseInput = {
  id: "w15.settings-identity-and-charts",
  schema: "w15",
  family: "document",
  element: "w15:chartTrackingRefBased",
  attributes: ["w15:chartTrackingRefBased/@val", "w15:docId/@val"],
  options: {
    sections: [{ children: [] }],
    settings: {
      w15ChartTrackingRefBased: true,
      w15DocId: "{37FF56B8-11AD-4A0A-8C3F-F4BD7E1B24BE}",
    },
  },
  partPath: "word/settings.xml",
  validationSchema: "wml.xsd",
  validationDirectory: "ooxml-schemas/transitional",
  projection: (options: DocumentOptions) => ({
    w15ChartTrackingRefBased: options.settings?.w15ChartTrackingRefBased,
    w15DocId: options.settings?.w15DocId,
  }),
  expected: {
    w15ChartTrackingRefBased: true,
    w15DocId: "{37FF56B8-11AD-4A0A-8C3F-F4BD7E1B24BE}",
  },
};

const w15FootnoteColumnsCase: OoxmlExecutableCaseInput = {
  id: "w15.section-footnote-columns",
  schema: "w15",
  family: "document",
  element: "w15:footnoteColumns",
  attributes: ["w15:footnoteColumns/@val"],
  options: {
    sections: [{ children: [], properties: { footnoteColumns: 2 } }],
  },
  partPath: "word/document.xml",
  validationSchema: "wml.xsd",
  validationDirectory: "ooxml-schemas/transitional",
  projection: (options: DocumentOptions) => ({
    footnoteColumns: options.sections?.[0]?.properties?.footnoteColumns,
  }),
  expected: { footnoteColumns: 2 },
};

const w15CollapsedCase: OoxmlExecutableCaseInput = {
  id: "w15.paragraph-collapsed",
  schema: "w15",
  family: "document",
  element: "w15:collapsed",
  attributes: ["w15:collapsed/@val"],
  options: {
    sections: [
      { children: [{ paragraph: { children: [{ text: "Executable" }], collapsed: true } }] },
    ],
  },
  partPath: "word/document.xml",
  validationSchema: "wml.xsd",
  validationDirectory: "ooxml-schemas/transitional",
  projection: (options: DocumentOptions) => {
    const child = options.sections?.[0]?.children?.[0];
    return child && "paragraph" in child && typeof child.paragraph !== "string"
      ? { collapsed: child.paragraph.collapsed }
      : undefined;
  },
  expected: { collapsed: true },
};

const w14SettingsCase: OoxmlExecutableCaseInput = {
  id: "w14.settings.identity-and-image-options",
  schema: "w14",
  family: "document",
  element: "w14:docId",
  attributes: [
    "w14:docId/@val",
    "w14:conflictMode/@val",
    "w14:discardImageEditingData/@val",
    "w14:defaultImageDpi/@val",
  ],
  options: {
    sections: [{ children: [] }],
    settings: {
      w14DocId: "646534E6",
      w14ConflictMode: false,
      w14DiscardImageEditingData: false,
      w14DefaultImageDpi: 300,
    },
  },
  partPath: "word/settings.xml",
  projection: (options) => {
    const settings = options.settings;
    return settings
      ? {
          w14DocId: settings.w14DocId,
          w14ConflictMode: settings.w14ConflictMode,
          w14DiscardImageEditingData: settings.w14DiscardImageEditingData,
          w14DefaultImageDpi: settings.w14DefaultImageDpi,
        }
      : undefined;
  },
  expected: {
    w14DocId: "646534E6",
    w14ConflictMode: false,
    w14DiscardImageEditingData: false,
    w14DefaultImageDpi: 300,
  },
};

const w14TrackChangesCase: OoxmlExecutableCaseInput = {
  id: "w14.track-changes.conflict-markers-and-runs",
  schema: "w14",
  family: "document",
  element: "w14:conflictIns",
  attributes: [
    "w14:customXmlConflictInsRangeStart/@author",
    "w14:customXmlConflictInsRangeStart/@date",
    "w14:customXmlConflictInsRangeEnd/@id",
    "w14:customXmlConflictDelRangeStart/@author",
    "w14:customXmlConflictDelRangeStart/@date",
    "w14:customXmlConflictDelRangeEnd/@id",
  ],
  options: {
    sections: [
      {
        children: [
          {
            paragraph: {
              children: [
                {
                  customXmlConflictInsRangeStart: {
                    id: 9,
                    author: "Ada",
                    date: "2021-03-04T00:00:00Z",
                  },
                },
                { customXmlConflictInsRangeEnd: 9 },
                {
                  customXmlConflictDelRangeStart: {
                    id: 10,
                    author: "Ben",
                    date: "2021-03-05T00:00:00Z",
                  },
                },
                { customXmlConflictDelRangeEnd: 10 },
                {
                  conflictIns: {
                    id: 7,
                    author: "Ada",
                    date: "2021-03-04T00:00:00Z",
                    children: ["inserted"],
                  },
                },
                {
                  conflictDel: {
                    id: 8,
                    author: "Ben",
                    date: "2021-03-05T00:00:00Z",
                    children: ["deleted"],
                  },
                },
              ],
            },
          },
        ],
      },
    ],
  },
  partPath: "word/document.xml",
  projection: (options) => {
    const child = options.sections?.[0]?.children?.[0];
    const paragraph = child && "paragraph" in child ? child.paragraph : undefined;
    const names = [
      "customXmlConflictInsRangeStart",
      "customXmlConflictInsRangeEnd",
      "customXmlConflictDelRangeStart",
      "customXmlConflictDelRangeEnd",
      "conflictIns",
      "conflictDel",
    ];
    return paragraph?.children?.filter((entry) => names.some((name) => name in entry));
  },
  expected: [
    {
      customXmlConflictInsRangeStart: {
        id: 9,
        author: "Ada",
        date: "2021-03-04T00:00:00Z",
      },
    },
    { customXmlConflictInsRangeEnd: 9 },
    {
      customXmlConflictDelRangeStart: {
        id: 10,
        author: "Ben",
        date: "2021-03-05T00:00:00Z",
      },
    },
    { customXmlConflictDelRangeEnd: 10 },
    {
      conflictIns: {
        id: 7,
        author: "Ada",
        date: "2021-03-04T00:00:00Z",
        children: [{ text: "inserted" }],
      },
    },
    {
      conflictDel: {
        id: 8,
        author: "Ben",
        date: "2021-03-05T00:00:00Z",
        children: [{ text: "deleted", preserveSpace: false }],
      },
    },
  ],
};

const w14CheckboxCase: OoxmlExecutableCaseInput = {
  id: "w14.sdt.checkbox",
  schema: "w14",
  family: "document",
  element: "w14:checkbox",
  attributes: [
    "w14:checked/@val",
    "w14:checkedState/@font",
    "w14:checkedState/@val",
    "w14:uncheckedState/@font",
    "w14:uncheckedState/@val",
  ],
  options: {
    sections: [
      {
        children: [
          {
            sdt: {
              properties: {
                checkbox: {
                  checked: false,
                  checkedState: { val: "2612", font: "MS Gothic" },
                  uncheckedState: { val: "2610", font: "MS Gothic" },
                },
              },
              children: [],
            },
          },
        ],
      },
    ],
  },
  partPath: "word/document.xml",
  projection: (options) => {
    const child = options.sections?.[0]?.children?.[0];
    return child && "sdt" in child ? child.sdt.properties.checkbox : undefined;
  },
  expected: {
    checked: false,
    checkedState: { val: "2612", font: "MS Gothic" },
    uncheckedState: { val: "2610", font: "MS Gothic" },
  },
};

const w14EntityPickerCase: OoxmlExecutableCaseInput = {
  id: "w14.sdt.entity-picker",
  schema: "w14",
  family: "document",
  element: "w14:entityPicker",
  attributes: [],
  options: {
    sections: [
      {
        children: [
          {
            sdt: { properties: { entityPicker: true }, children: [] },
          },
        ],
      },
    ],
  },
  partPath: "word/document.xml",
  projection: (options) => {
    const child = options.sections?.[0]?.children?.[0];
    return child && "sdt" in child ? child.sdt.properties.entityPicker : undefined;
  },
  expected: true,
};

function paragraphRunProperties(options: DocumentOptions): RunPropertiesOptions | undefined {
  const child = options.sections?.[0]?.children?.[0];
  if (!child || !("paragraph" in child)) return undefined;
  const paragraphChild = child.paragraph.children?.[0];
  return paragraphChild && typeof paragraphChild === "object" && "text" in paragraphChild
    ? paragraphChild
    : undefined;
}

function w14EffectsProjection(
  options: DocumentOptions,
  keys: readonly (keyof W14RunEffectsOptions)[],
): Partial<W14RunEffectsOptions> {
  const properties = paragraphRunProperties(options);
  const result: Partial<W14RunEffectsOptions> = {};
  if (!properties) return result;
  for (const key of keys) {
    if (properties[key] !== undefined) result[key] = properties[key];
  }
  return result;
}

const w14EffectsCase: OoxmlExecutableCaseInput = {
  id: "w14.run-effects.complete",
  schema: "w14",
  family: "document",
  element: "w14:glow",
  attributes: [
    "w14:tint/@val",
    "w14:shade/@val",
    "w14:alpha/@val",
    "w14:hueMod/@val",
    "w14:sat/@val",
    "w14:satOff/@val",
    "w14:satMod/@val",
    "w14:lum/@val",
    "w14:lumOff/@val",
    "w14:lumMod/@val",
    "w14:srgbClr/@val",
    "w14:schemeClr/@val",
    "w14:glow/@rad",
    "w14:shadow/@blurRad",
    "w14:shadow/@dist",
    "w14:shadow/@dir",
    "w14:shadow/@sx",
    "w14:shadow/@sy",
    "w14:shadow/@kx",
    "w14:shadow/@ky",
    "w14:shadow/@algn",
    "w14:reflection/@blurRad",
    "w14:reflection/@stA",
    "w14:reflection/@stPos",
    "w14:reflection/@endA",
    "w14:reflection/@endPos",
    "w14:reflection/@dist",
    "w14:reflection/@dir",
    "w14:reflection/@fadeDir",
    "w14:reflection/@sx",
    "w14:reflection/@sy",
    "w14:reflection/@kx",
    "w14:reflection/@ky",
    "w14:reflection/@algn",
    "w14:textOutline/@w",
    "w14:textOutline/@cap",
    "w14:textOutline/@cmpd",
    "w14:textOutline/@algn",
    "w14:prstDash/@val",
    "w14:miter/@lim",
    "w14:rot/@lat",
    "w14:rot/@lon",
    "w14:rot/@rev",
    "w14:camera/@prst",
    "w14:lightRig/@rig",
    "w14:lightRig/@dir",
    "w14:bevelT/@w",
    "w14:bevelT/@h",
    "w14:bevelT/@prst",
    "w14:bevelB/@w",
    "w14:bevelB/@h",
    "w14:bevelB/@prst",
    "w14:props3d/@extrusionH",
    "w14:props3d/@contourW",
    "w14:props3d/@prstMaterial",
    "w14:gs/@pos",
    "w14:fillToRect/@l",
    "w14:fillToRect/@t",
    "w14:fillToRect/@r",
    "w14:fillToRect/@b",
    "w14:path/@path",
    "w14:styleSet/@id",
    "w14:styleSet/@val",
    "w14:ligatures/@val",
    "w14:numForm/@val",
    "w14:numSpacing/@val",
    "w14:cntxtAlts/@val",
  ],
  options: {
    sections: [
      {
        children: [
          {
            paragraph: {
              children: [
                {
                  text: "effects",
                  glow: {
                    radius: 63500,
                    color: {
                      srgb: {
                        val: "4472C4",
                        tint: 50,
                        shade: 60,
                        alpha: 70,
                        hueMod: 80,
                        sat: 90,
                        satOff: 10,
                        satMod: 20,
                        lum: 30,
                        lumOff: 40,
                        lumMod: 60,
                      },
                    },
                  },
                  shadowEffect: {
                    blurRadius: 50800,
                    distance: 38100,
                    direction: 45,
                    horizontalScale: 80,
                    verticalScale: 90,
                    horizontalSkew: 15,
                    verticalSkew: 20,
                    alignment: "bottomRight",
                    color: { scheme: { val: "accent1" } },
                  },
                  reflection: {
                    blurRadius: 6350,
                    startAlpha: 15,
                    startPosition: 25,
                    endAlpha: 2,
                    endPosition: 90,
                    distance: 12700,
                    direction: 45,
                    fadeDirection: 90,
                    horizontalScale: 80,
                    verticalScale: 90,
                    horizontalSkew: 10,
                    verticalSkew: 15,
                    alignment: "bottom",
                  },
                  textOutline: {
                    fill: { noFill: true },
                    dash: "dash",
                    join: { miterLimit: 150 },
                    width: 19050,
                    cap: "round",
                    compound: "single",
                    alignment: "center",
                  },
                  textFill: {
                    gradient: {
                      stops: [
                        { position: 10, color: { srgb: { val: "FFFFFF" } } },
                        { position: 90, color: { scheme: { val: "accent1" } } },
                      ],
                      path: {
                        shape: "circle",
                        fillToRect: { left: 5, top: 10, right: 15, bottom: 20 },
                      },
                    },
                  },
                  scene3d: {
                    camera: { preset: "isometricOffAxis1Left" },
                    lightRig: {
                      rig: "threePt",
                      direction: "topRight",
                      rotation: { lat: 10, lon: 20, rev: 30 },
                    },
                  },
                  props3d: {
                    topBevel: { w: 63500, h: 25400, prst: "circle" },
                    bottomBevel: { w: 31750, h: 12700, prst: "slope" },
                    extrusionColor: { srgb: { val: "808080" } },
                    contourColor: { srgb: { val: "404040" } },
                    extrusionHeight: 571500,
                    contourWidth: 34925,
                    material: "matte",
                  },
                  ligatures: "standardContextual",
                  numForm: "lining",
                  numSpacing: "proportional",
                  stylisticSets: { sets: [{ id: 1, enabled: true }] },
                  cntxtAlts: true,
                },
              ],
            },
          },
        ],
      },
    ],
  },
  partPath: "word/document.xml",
  projection: (options) =>
    w14EffectsProjection(options, [
      "glow",
      "shadowEffect",
      "reflection",
      "textOutline",
      "textFill",
      "scene3d",
      "props3d",
      "ligatures",
      "numForm",
      "numSpacing",
      "stylisticSets",
      "cntxtAlts",
    ]),
  expected: {
    glow: {
      radius: 63500,
      color: {
        srgb: {
          val: "4472C4",
          tint: 50,
          shade: 60,
          alpha: 70,
          hueMod: 80,
          sat: 90,
          satOff: 10,
          satMod: 20,
          lum: 30,
          lumOff: 40,
          lumMod: 60,
        },
      },
    },
    shadowEffect: {
      blurRadius: 50800,
      distance: 38100,
      direction: 45,
      horizontalScale: 80,
      verticalScale: 90,
      horizontalSkew: 15,
      verticalSkew: 20,
      alignment: "bottomRight",
      color: { scheme: { val: "accent1" } },
    },
    reflection: {
      blurRadius: 6350,
      startAlpha: 15,
      startPosition: 25,
      endAlpha: 2,
      endPosition: 90,
      distance: 12700,
      direction: 45,
      fadeDirection: 90,
      horizontalScale: 80,
      verticalScale: 90,
      horizontalSkew: 10,
      verticalSkew: 15,
      alignment: "bottom",
    },
    textOutline: {
      fill: { noFill: true },
      dash: "dash",
      join: { miterLimit: 150 },
      width: 19050,
      cap: "round",
      compound: "single",
      alignment: "center",
    },
    textFill: {
      gradient: {
        stops: [
          { position: 10, color: { srgb: { val: "FFFFFF" } } },
          { position: 90, color: { scheme: { val: "accent1" } } },
        ],
        path: { shape: "circle", fillToRect: { left: 5, top: 10, right: 15, bottom: 20 } },
      },
    },
    scene3d: {
      camera: { preset: "isometricOffAxis1Left" },
      lightRig: {
        rig: "threePt",
        direction: "topRight",
        rotation: { lat: 10, lon: 20, rev: 30 },
      },
    },
    props3d: {
      topBevel: { w: 63500, h: 25400, prst: "circle" },
      bottomBevel: { w: 31750, h: 12700, prst: "slope" },
      extrusionColor: { srgb: { val: "808080" } },
      contourColor: { srgb: { val: "404040" } },
      extrusionHeight: 571500,
      contourWidth: 34925,
      material: "matte",
    },
    ligatures: "standardContextual",
    numForm: "lining",
    numSpacing: "proportional",
    stylisticSets: { sets: [{ id: 1, enabled: true }] },
    cntxtAlts: true,
  },
};

const w14EffectJoinVariants: readonly OoxmlExecutableCaseInput[] = [
  {
    id: "w14.run-effects.round-join",
    schema: "w14",
    family: "document",
    element: "w14:round",
    attributes: [],
    options: runEffectsOptions({ textOutline: { join: "round" } }),
    partPath: "word/document.xml",
    projection: (options) => w14EffectsProjection(options, ["textOutline"]),
    expected: { textOutline: { join: "round" } },
  },
  {
    id: "w14.run-effects.bevel-join",
    schema: "w14",
    family: "document",
    element: "w14:bevel",
    attributes: [],
    options: runEffectsOptions({ textOutline: { join: "bevel" } }),
    partPath: "word/document.xml",
    projection: (options) => w14EffectsProjection(options, ["textOutline"]),
    expected: { textOutline: { join: "bevel" } },
  },
];

const w14LinearGradientCase: OoxmlExecutableCaseInput = {
  id: "w14.run-effects.linear-gradient",
  schema: "w14",
  family: "document",
  element: "w14:lin",
  attributes: ["w14:lin/@ang", "w14:lin/@scaled"],
  options: runEffectsOptions({
    textFill: {
      gradient: {
        stops: [
          { position: 10, color: { srgb: { val: "FFFFFF" } } },
          { position: 90, color: { scheme: { val: "accent1" } } },
        ],
        linear: { angle: 90, scaled: true },
      },
    },
  }),
  partPath: "word/document.xml",
  projection: (options) => w14EffectsProjection(options, ["textFill"]),
  expected: {
    textFill: {
      gradient: {
        stops: [
          { position: 10, color: { srgb: { val: "FFFFFF" } } },
          { position: 90, color: { scheme: { val: "accent1" } } },
        ],
        linear: { angle: 90, scaled: true },
      },
    },
  },
};

const w14SolidFillCase: OoxmlExecutableCaseInput = {
  id: "w14.run-effects.solid-fill",
  schema: "w14",
  family: "document",
  element: "w14:solidFill",
  attributes: [],
  options: runEffectsOptions({
    textFill: { solid: { srgb: { val: "FF0000" } } },
  }),
  partPath: "word/document.xml",
  projection: (options) => w14EffectsProjection(options, ["textFill"]),
  expected: { textFill: { solid: { srgb: { val: "FF0000" } } } },
};

function worksheetDrawing(
  key: "images" | "shapes" | "connectors" | "charts" | "groups" | "contentParts",
  value: unknown,
): WorkbookOptions {
  return { worksheets: [{ name: "Drawing", [key]: value }] as never };
}

function worksheetValue(options: WorkbookOptions, key: string): unknown {
  const worksheet = options.worksheets?.[0] as Record<string, unknown> | undefined;
  return worksheet?.[key];
}

function worksheetProjection(
  options: WorkbookOptions,
  key: string,
  pick: (entry: Record<string, unknown>) => Record<string, unknown> | number,
): unknown[] | undefined {
  const values = worksheetValue(options, key);
  return Array.isArray(values)
    ? values.map((value) => pick(value as Record<string, unknown>))
    : undefined;
}

const xdrTwoCellShapeCase: OoxmlExecutableCaseInput = {
  id: "xdr.shape.two-cell",
  schema: "dml-xdr",
  family: "workbook",
  element: "xdr:sp",
  attributes: [
    "xdr:sp/@macro",
    "xdr:sp/@textlink",
    "xdr:sp/@fLocksText",
    "xdr:sp/@fPublished",
    "xdr:cNvPr/@id",
    "xdr:cNvPr/@name",
    "xdr:cNvPr/@descr",
    "xdr:cNvPr/@hidden",
    "xdr:cNvPr/@title",
    "xdr:cNvSpPr/@txBox",
    "xdr:spPr/@bwMode",
    "xdr:twoCellAnchor/@editAs",
    "xdr:clientData/@fLocksWithSheet",
    "xdr:clientData/@fPrintsWithSheet",
  ],
  options: worksheetDrawing("shapes", [
    {
      col: 2,
      row: 3,
      colOffset: 100,
      rowOffset: 200,
      toCol: 5,
      toRow: 7,
      toColOffset: 300,
      toRowOffset: 400,
      anchorType: "twoCell",
      editAs: "oneCell",
      shapeId: 11,
      name: "Executable Shape",
      description: "Shape",
      title: "Shape",
      hidden: true,
      textBox: true,
      macro: "shape.macro",
      textlink: "A1",
      fLocksText: false,
      fPublished: false,
      blackWhiteMode: "gray",
      locksWithSheet: false,
      printsWithSheet: false,
      properties: { geometry: "rect", width: 914400, height: 914400 },
      textBody: { paragraphs: [{ text: "Shape text" }] },
      style: {
        lineReference: { index: 2, color: { value: "4472C4" } },
        fillReference: { index: 1, color: { value: "FFFFFF" } },
        effectReference: { index: 0 },
        fontReference: { collection: "minor" },
      },
    },
  ]),
  partPath: "xl/drawings/drawing1.xml",
  projection: (options) =>
    worksheetProjection(options, "shapes", (shape) => ({
      col: shape.col,
      row: shape.row,
      colOffset: shape.colOffset,
      rowOffset: shape.rowOffset,
      toCol: shape.toCol,
      toRow: shape.toRow,
      toColOffset: shape.toColOffset,
      toRowOffset: shape.toRowOffset,
      anchorType: shape.anchorType,
      editAs: shape.editAs,
      shapeId: shape.shapeId,
      name: shape.name,
      description: shape.description,
      title: shape.title,
      hidden: shape.hidden,
      textBox: shape.textBox,
      macro: shape.macro,
      textlink: shape.textlink,
      fLocksText: shape.fLocksText,
      fPublished: shape.fPublished,
      blackWhiteMode: shape.blackWhiteMode,
      locksWithSheet: shape.locksWithSheet,
      printsWithSheet: shape.printsWithSheet,
    })),
  expected: [
    {
      col: 2,
      row: 3,
      colOffset: 100,
      rowOffset: 200,
      toCol: 5,
      toRow: 7,
      toColOffset: 300,
      toRowOffset: 400,
      anchorType: "twoCell",
      editAs: "oneCell",
      shapeId: 11,
      name: "Executable Shape",
      description: "Shape",
      title: "Shape",
      hidden: true,
      textBox: true,
      macro: "shape.macro",
      textlink: "A1",
      fLocksText: false,
      fPublished: false,
      blackWhiteMode: "gray",
      locksWithSheet: false,
      printsWithSheet: false,
    },
  ],
};

const xdrOneCellPictureCase: OoxmlExecutableCaseInput = {
  id: "xdr.picture.one-cell",
  schema: "dml-xdr",
  family: "workbook",
  element: "xdr:pic",
  attributes: [
    "xdr:pic/@macro",
    "xdr:pic/@fPublished",
    "xdr:cNvPr/@id",
    "xdr:cNvPr/@name",
    "xdr:cNvPr/@descr",
    "xdr:cNvPr/@hidden",
    "xdr:cNvPr/@title",
    "xdr:cNvPicPr/@preferRelativeResize",
    "xdr:blipFill/@dpi",
    "xdr:blipFill/@rotWithShape",
    "xdr:spPr/@bwMode",
  ],
  options: worksheetDrawing("images", [
    {
      type: "png",
      data: new Uint8Array([1, 2, 3, 4]),
      col: 1,
      row: 1,
      colOffset: 10,
      rowOffset: 20,
      anchorType: "oneCell",
      extentCx: 123456,
      extentCy: 234567,
      shapeId: 21,
      name: "Executable Picture",
      description: "Picture",
      title: "Picture",
      hidden: true,
      dpi: 300,
      rotWithShape: false,
      blackWhiteMode: "gray",
      preferRelativeResize: false,
      properties: { rotation: 15, flipHorizontal: true, flipVertical: true },
    },
  ]),
  partPath: "xl/drawings/drawing1.xml",
  projection: (options) =>
    worksheetProjection(options, "images", (image) => ({
      col: image.col,
      row: image.row,
      colOffset: image.colOffset,
      rowOffset: image.rowOffset,
      anchorType: image.anchorType,
      extentCx: image.extentCx,
      extentCy: image.extentCy,
      shapeId: image.shapeId,
      name: image.name,
      description: image.description,
      title: image.title,
      hidden: image.hidden,
      dpi: image.dpi,
      rotWithShape: image.rotWithShape,
      blackWhiteMode: image.blackWhiteMode,
      preferRelativeResize: image.preferRelativeResize,
    })),
  expected: [
    {
      col: 1,
      row: 1,
      colOffset: 10,
      rowOffset: 20,
      anchorType: "oneCell",
      extentCx: 123456,
      extentCy: 234567,
      shapeId: 21,
      name: "Executable Picture",
      description: "Picture",
      title: "Picture",
      hidden: true,
      dpi: 300,
      rotWithShape: false,
      blackWhiteMode: "gray",
      preferRelativeResize: false,
    },
  ],
};

const xdrAbsoluteConnectorCase: OoxmlExecutableCaseInput = {
  id: "xdr.connector.absolute",
  schema: "dml-xdr",
  family: "workbook",
  element: "xdr:cxnSp",
  attributes: [
    "xdr:cxnSp/@macro",
    "xdr:cxnSp/@fPublished",
    "xdr:cNvPr/@id",
    "xdr:cNvPr/@name",
    "xdr:spPr/@bwMode",
    "xdr:pos/@x",
    "xdr:pos/@y",
    "xdr:ext/@cx",
    "xdr:ext/@cy",
  ],
  options: worksheetDrawing("connectors", [
    {
      col: 1,
      row: 1,
      anchorType: "absolute",
      absoluteX: 5000,
      absoluteY: 6000,
      extentCx: 100000,
      extentCy: 50000,
      shapeId: 31,
      name: "Executable Connector",
      macro: "connector.macro",
      fPublished: true,
      blackWhiteMode: "gray",
      properties: { geometry: "line", width: 100000, height: 50000 },
    },
  ]),
  partPath: "xl/drawings/drawing1.xml",
  projection: (options) =>
    worksheetProjection(options, "connectors", (connector) => ({
      col: connector.col,
      row: connector.row,
      anchorType: connector.anchorType,
      absoluteX: connector.absoluteX,
      absoluteY: connector.absoluteY,
      extentCx: connector.extentCx,
      extentCy: connector.extentCy,
      shapeId: connector.shapeId,
      name: connector.name,
      macro: connector.macro,
      fPublished: connector.fPublished,
      blackWhiteMode: connector.blackWhiteMode,
    })),
  expected: [
    {
      col: 1,
      row: 1,
      anchorType: "absolute",
      absoluteX: 5000,
      absoluteY: 6000,
      extentCx: 100000,
      extentCy: 50000,
      shapeId: 31,
      name: "Executable Connector",
      macro: "connector.macro",
      fPublished: true,
      blackWhiteMode: "gray",
    },
  ],
};

const xdrChartFrameCase: OoxmlExecutableCaseInput = {
  id: "xdr.chart.graphic-frame",
  schema: "dml-xdr",
  family: "workbook",
  element: "xdr:graphicFrame",
  attributes: [
    "xdr:graphicFrame/@macro",
    "xdr:graphicFrame/@fPublished",
    "xdr:cNvPr/@id",
    "xdr:cNvPr/@name",
    "xdr:xfrm/@rot",
    "xdr:xfrm/@flipH",
    "xdr:xfrm/@flipV",
  ],
  options: worksheetDrawing("charts", [
    {
      type: "column",
      title: "Executable Chart",
      categories: ["A", "B"],
      series: [{ name: "Series", values: [1, 2] }],
      rId: "chart",
      col: 1,
      row: 1,
      toCol: 5,
      toRow: 8,
      shapeId: 41,
      name: "Executable Chart",
      macro: "",
      fPublished: false,
      frameExtentCx: 100000,
      frameExtentCy: 80000,
      frameRotation: 15,
      frameFlipHorizontal: true,
      frameFlipVertical: false,
    },
  ]),
  partPath: "xl/drawings/drawing1.xml",
  projection: (options) =>
    worksheetProjection(options, "charts", (chart) => ({
      col: chart.col,
      row: chart.row,
      toCol: chart.toCol,
      toRow: chart.toRow,
      shapeId: chart.shapeId,
      name: chart.name,
      macro: chart.macro,
      fPublished: chart.fPublished,
      frameExtentCx: chart.frameExtentCx,
      frameExtentCy: chart.frameExtentCy,
      frameRotation: chart.frameRotation,
      frameFlipHorizontal: chart.frameFlipHorizontal,
      frameFlipVertical: chart.frameFlipVertical,
    })),
  expected: [
    {
      col: 1,
      row: 1,
      toCol: 5,
      toRow: 8,
      shapeId: 41,
      name: "Executable Chart",
      macro: "",
      fPublished: false,
      frameExtentCx: 100000,
      frameExtentCy: 80000,
      frameRotation: 15,
      frameFlipHorizontal: true,
      frameFlipVertical: false,
    },
  ],
};

const xdrGroupCase: OoxmlExecutableCaseInput = {
  id: "xdr.group.nested-objects",
  schema: "dml-xdr",
  family: "workbook",
  element: "xdr:grpSp",
  attributes: ["xdr:grpSpPr/@bwMode"],
  options: worksheetDrawing("groups", [
    {
      col: 1,
      row: 1,
      toCol: 6,
      toRow: 8,
      shapeId: 51,
      name: "Executable Group",
      blackWhiteMode: "gray",
      properties: {
        width: 300000,
        height: 200000,
        childOffsetX: 0,
        childOffsetY: 0,
        childExtentWidth: 300000,
        childExtentHeight: 200000,
        rotation: 10,
        flipHorizontal: true,
        flipVertical: true,
      },
      shapes: [
        {
          name: "Child Shape",
          shapeId: 52,
          properties: { geometry: "rect", width: 100000, height: 50000 },
        },
      ],
      images: [
        {
          rId: "",
          name: "Child Picture",
          shapeId: 53,
          properties: { width: 20000, height: 20000 },
        },
      ],
      connectors: [
        {
          name: "Child Connector",
          shapeId: 54,
          properties: { geometry: "line", width: 10000, height: 0 },
        },
      ],
    },
  ]),
  partPath: "xl/drawings/drawing1.xml",
  projection: (options) =>
    worksheetProjection(options, "groups", (group) => ({
      col: group.col,
      row: group.row,
      toCol: group.toCol,
      toRow: group.toRow,
      shapeId: group.shapeId,
      name: group.name,
      blackWhiteMode: group.blackWhiteMode,
      shapeCount: Array.isArray(group.shapes) ? group.shapes.length : 0,
      imageCount: Array.isArray(group.images) ? group.images.length : 0,
      connectorCount: Array.isArray(group.connectors) ? group.connectors.length : 0,
    })),
  expected: [
    {
      col: 1,
      row: 1,
      toCol: 6,
      toRow: 8,
      shapeId: 51,
      name: "Executable Group",
      blackWhiteMode: "gray",
      shapeCount: 1,
      imageCount: 1,
      connectorCount: 1,
    },
  ],
};

const xdrContentPartCase: OoxmlExecutableCaseInput = {
  id: "xdr.content-part.reference",
  schema: "dml-xdr",
  family: "workbook",
  element: "xdr:contentPart",
  attributes: ["xdr:contentPart/@r:id"],
  options: {
    worksheets: [
      {
        name: "Drawing",
        contentParts: [
          {
            col: 1,
            row: 1,
            toCol: 3,
            toRow: 4,
            shapeId: 61,
            relationshipType:
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
            relationshipTarget: "../embeddings/content1.bin",
            sourcePath: "xl/embeddings/content1.bin",
          },
        ],
      },
    ],
    rawParts: [
      {
        path: "xl/embeddings/content1.bin",
        data: new Uint8Array([1, 2, 3]),
        contentType: "application/vnd.openxmlformats-officedocument.oleObject",
      },
    ],
  },
  partPath: "xl/drawings/drawing1.xml",
  projection: (options) =>
    worksheetProjection(options, "contentParts", (part) => ({
      col: part.col,
      row: part.row,
      toCol: part.toCol,
      toRow: part.toRow,
      rId: part.rId,
      relationshipType: part.relationshipType,
      relationshipTarget: part.relationshipTarget,
      sourcePath: part.sourcePath,
    })),
  expected: [
    {
      col: 1,
      row: 1,
      toCol: 3,
      toRow: 4,
      rId: "rId1",
      relationshipType:
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
      relationshipTarget: "../embeddings/content1.bin",
      sourcePath: "xl/embeddings/content1.bin",
    },
  ],
};

function runEffectsOptions(effects: W14RunEffectsOptions): DocumentOptions {
  return {
    sections: [
      {
        children: [{ paragraph: { children: [{ text: "effect", ...effects }] } }],
      },
    ],
  };
}

function chartUserShapesWorkbook(
  userShapes: NonNullable<
    NonNullable<NonNullable<WorkbookOptions["worksheets"]>[number]["charts"]>[number]["userShapes"]
  >,
): WorkbookOptions {
  return {
    worksheets: [
      {
        name: "Data",
        rows: [{ cells: [{ value: "A" }, { value: 1 }] }],
        charts: [
          {
            type: "column",
            title: "Annotated",
            categories: ["A"],
            series: [{ name: "S", values: [1] }],
            col: 4,
            row: 1,
            userShapes,
          },
        ],
      },
    ],
  };
}

function chartUserShapesProjection(options: WorkbookOptions): unknown {
  const userShapes = options.worksheets?.[0]?.charts?.[0]?.userShapes;
  if (!userShapes) return undefined;
  return {
    relationshipId: userShapes.relationshipId,
    rootElement: userShapes.rootElement,
    anchors: userShapes.anchors.map((anchor) => {
      const { object, ...anchorFields } = anchor;
      const normalizedObject =
        object.type === "shape" ? { ...object, textBody: undefined } : object;
      return { ...anchorFields, object: normalizedObject };
    }),
  };
}

const chartDrawingShapeCase: OoxmlExecutableCaseInput = {
  id: "dml-cdr.relative-shape",
  schema: "dml-cdr",
  family: "workbook",
  element: "cdr:sp",
  attributes: [
    "cdr:cNvPr/@id",
    "cdr:cNvPr/@name",
    "cdr:cNvPr/@descr",
    "cdr:cNvPr/@title",
    "cdr:cNvPr/@hidden",
    "cdr:cNvSpPr/@txBox",
    "cdr:spPr/@bwMode",
    "cdr:sp/@macro",
    "cdr:sp/@textlink",
    "cdr:sp/@fLocksText",
    "cdr:sp/@fPublished",
  ],
  options: chartUserShapesWorkbook({
    rootElement: "chart",
    anchors: [
      {
        from: { x: 0.1, y: 0.2 },
        to: { x: 0.9, y: 0.8 },
        object: {
          type: "shape",
          id: 1,
          macro: "Shape.Macro",
          published: true,
          textBox: true,
          textLink: "Sheet1!$A$1",
          locksText: false,
          blackWhiteMode: "gray",
          style: {
            lineReference: { index: 2 },
            fillReference: { index: 1 },
            effectReference: { index: 0 },
            fontReference: { collection: "minor" },
          },
          nonVisualProperties: {
            name: "Executable Shape",
            description: "Chart shape",
            title: "Shape",
            hidden: true,
          },
          shapeProperties: {
            width: 914400,
            height: 457200,
            geometry: "rect",
          },
          textBody: { paragraphs: [{ children: [{ text: "Overlaid label" }] }] },
        },
      },
    ],
  }),
  partPath: "xl/charts/userShapes1.xml",
  validationSchema: "dml-chart.xsd",
  validationDirectory: "ooxml-schemas/transitional",
  projection: chartUserShapesProjection,
  expected: {
    relationshipId: "rId1",
    rootElement: "chart",
    anchors: [
      {
        from: { x: 0.1, y: 0.2 },
        to: { x: 0.9, y: 0.8 },
        object: {
          type: "shape",
          id: 1,
          macro: "Shape.Macro",
          published: true,
          textBox: true,
          textLink: "Sheet1!$A$1",
          locksText: false,
          style: {
            lineReference: { index: 2 },
            fillReference: { index: 1 },
            effectReference: { index: 0 },
            fontReference: { collection: "minor" },
          },
          nonVisualProperties: {
            name: "Executable Shape",
            description: "Chart shape",
            title: "Shape",
            hidden: true,
          },
          shapeProperties: {
            width: 914400,
            height: 457200,
            geometry: { preset: "rect" },
          },
          blackWhiteMode: "gray",
        },
      },
    ],
  },
};

const chartDrawingConnectorCase: OoxmlExecutableCaseInput = {
  id: "dml-cdr.absolute-connector",
  schema: "dml-cdr",
  family: "workbook",
  element: "cdr:cxnSp",
  attributes: ["cdr:cxnSp/@macro", "cdr:cxnSp/@fPublished", "cdr:ext/@cx", "cdr:ext/@cy"],
  options: chartUserShapesWorkbook({
    rootElement: "chart",
    anchors: [
      {
        from: { x: 0, y: 0 },
        extent: { width: 1828800, height: 914400 },
        object: {
          type: "connector",
          id: 2,
          nonVisualProperties: { name: "User Shape" },
          macro: "Connector.Macro",
          published: true,
          blackWhiteMode: "gray",
          shapeProperties: { width: 1828800, height: 914400 },
        },
      },
    ],
  }),
  partPath: "xl/charts/userShapes1.xml",
  validationSchema: "dml-chart.xsd",
  validationDirectory: "ooxml-schemas/transitional",
  projection: chartUserShapesProjection,
  expected: {
    relationshipId: "rId1",
    rootElement: "chart",
    anchors: [
      {
        from: { x: 0, y: 0 },
        extent: { width: 1828800, height: 914400 },
        object: {
          type: "connector",
          id: 2,
          nonVisualProperties: { name: "User Shape" },
          macro: "Connector.Macro",
          published: true,
          blackWhiteMode: "gray",
          shapeProperties: { width: 1828800, height: 914400 },
        },
      },
    ],
  },
};

const chartDrawingPictureCase: OoxmlExecutableCaseInput = {
  id: "dml-cdr.absolute-picture",
  schema: "dml-cdr",
  family: "workbook",
  element: "cdr:pic",
  attributes: [
    "cdr:pic/@macro",
    "cdr:pic/@fPublished",
    "cdr:cNvPicPr/@preferRelativeResize",
    "cdr:blipFill/@dpi",
    "cdr:blipFill/@rotWithShape",
  ],
  options: chartUserShapesWorkbook({
    rootElement: "chart",
    anchors: [
      {
        from: { x: 0.5, y: 0.5 },
        extent: { width: 914400, height: 914400 },
        object: {
          type: "picture",
          id: 3,
          nonVisualProperties: { name: "User Shape" },
          macro: "Picture.Macro",
          published: true,
          referenceId: "image1.png",
          preferRelativeResize: false,
          blipFill: { dpi: 300, rotWithShape: true },
          blackWhiteMode: "gray",
          shapeProperties: { width: 914400, height: 914400 },
        },
      },
    ],
  }),
  partPath: "xl/charts/userShapes1.xml",
  validationSchema: "dml-chart.xsd",
  validationDirectory: "ooxml-schemas/transitional",
  projection: chartUserShapesProjection,
  expected: {
    relationshipId: "rId1",
    rootElement: "chart",
    anchors: [
      {
        from: { x: 0.5, y: 0.5 },
        extent: { width: 914400, height: 914400 },
        object: {
          type: "picture",
          id: 3,
          nonVisualProperties: { name: "User Shape" },
          macro: "Picture.Macro",
          published: true,
          referenceId: "image1.png",
          preferRelativeResize: false,
          blipFill: { dpi: 300, rotWithShape: true },
          blackWhiteMode: "gray",
          shapeProperties: { width: 914400, height: 914400 },
        },
      },
    ],
  },
};

const chartDrawingGraphicFrameCase: OoxmlExecutableCaseInput = {
  id: "dml-cdr.graphic-frame",
  schema: "dml-cdr",
  family: "workbook",
  element: "cdr:graphicFrame",
  attributes: [
    "cdr:graphicFrame/@macro",
    "cdr:graphicFrame/@fPublished",
    "cdr:xfrm/@rot",
    "cdr:xfrm/@flipH",
    "cdr:xfrm/@flipV",
  ],
  options: chartUserShapesWorkbook({
    rootElement: "chart",
    anchors: [
      {
        from: { x: 0, y: 0 },
        extent: { width: 914400, height: 914400 },
        object: {
          type: "graphicFrame",
          id: 4,
          nonVisualProperties: { name: "User Shape" },
          macro: "Frame.Macro",
          published: true,
          graphicFrameLocks: { noChangeAspect: true },
          transform: {
            x: 0,
            y: 0,
            width: 914400,
            height: 914400,
            rotation: 45,
            flipHorizontal: true,
            flipVertical: true,
          },
          graphic:
            '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/></a:graphicData></a:graphic>',
        },
      },
    ],
  }),
  partPath: "xl/charts/userShapes1.xml",
  validationSchema: "dml-chart.xsd",
  validationDirectory: "ooxml-schemas/transitional",
  projection: chartUserShapesProjection,
  expected: {
    relationshipId: "rId1",
    rootElement: "chart",
    anchors: [
      {
        from: { x: 0, y: 0 },
        extent: { width: 914400, height: 914400 },
        object: {
          type: "graphicFrame",
          id: 4,
          nonVisualProperties: { name: "User Shape" },
          macro: "Frame.Macro",
          published: true,
          graphicFrameLocks: { noChangeAspect: true },
          transform: {
            x: 0,
            y: 0,
            width: 914400,
            height: 914400,
            rotation: 45,
            flipHorizontal: true,
            flipVertical: true,
          },
          graphic:
            '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/></a:graphicData></a:graphic>',
        },
      },
    ],
  },
};

const chartDrawingGroupCase: OoxmlExecutableCaseInput = {
  id: "dml-cdr.group",
  schema: "dml-cdr",
  family: "workbook",
  element: "cdr:grpSp",
  attributes: ["cdr:grpSpPr/@bwMode"],
  options: chartUserShapesWorkbook({
    rootElement: "chart",
    anchors: [
      {
        from: { x: 0, y: 0 },
        extent: { width: 1828800, height: 1828800 },
        object: {
          type: "group",
          id: 5,
          nonVisualProperties: { name: "User Shape" },
          blackWhiteMode: "gray",
          groupShapeProperties: {
            width: 1828800,
            height: 1828800,
            childOffsetX: 0,
            childOffsetY: 0,
            childExtentWidth: 1828800,
            childExtentHeight: 1828800,
          },
          children: [{ type: "shape", id: 6, shapeProperties: { width: 914400, height: 914400 } }],
        },
      },
    ],
  }),
  partPath: "xl/charts/userShapes1.xml",
  validationSchema: "dml-chart.xsd",
  validationDirectory: "ooxml-schemas/transitional",
  projection: chartUserShapesProjection,
  expected: {
    relationshipId: "rId1",
    rootElement: "chart",
    anchors: [
      {
        from: { x: 0, y: 0 },
        extent: { width: 1828800, height: 1828800 },
        object: {
          type: "group",
          id: 5,
          nonVisualProperties: { name: "User Shape" },
          blackWhiteMode: "gray",
          groupShapeProperties: {
            width: 1828800,
            height: 1828800,
            childOffsetX: 0,
            childOffsetY: 0,
            childExtentWidth: 1828800,
            childExtentHeight: 1828800,
          },
          children: [
            {
              type: "shape",
              id: 6,
              nonVisualProperties: { name: "User Shape" },
              shapeProperties: {
                width: 914400,
                height: 914400,
              },
            },
          ],
        },
      },
    ],
  },
};

const pictureCase: OoxmlExecutableCaseInput = {
  id: "dml-pic.picture.round-trip",
  schema: "dml-pic",
  family: "document",
  element: "pic:pic",
  attributes: [
    "pic:cNvPr/@id",
    "pic:cNvPr/@name",
    "pic:cNvPr/@descr",
    "pic:cNvPr/@title",
    "pic:cNvPr/@hidden",
    "pic:cNvPicPr/@preferRelativeResize",
    "pic:blipFill/@dpi",
    "pic:blipFill/@rotWithShape",
    "pic:spPr/@bwMode",
  ],
  options: {
    sections: [
      {
        children: [
          {
            paragraph: {
              children: [
                {
                  picture: {
                    type: "png",
                    data: new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10, 1, 2]),
                    dpi: 300,
                    rotWithShape: true,
                    blackWhiteMode: "gray",
                    geometry: "rect",
                    nonVisualProperties: {
                      id: 9,
                      name: "Executable Picture",
                      description: "Picture fixture",
                      title: "Picture",
                      hidden: true,
                      preferRelativeResize: false,
                    },
                    transformation: { width: 914400, height: 914400 },
                  },
                },
              ],
            },
          },
        ],
      },
    ],
  },
  partPath: "word/document.xml",
  validationSchema: "wml.xsd",
  validationDirectory: "ooxml-schemas/transitional",
  projection: (options: DocumentOptions) => {
    const child = options.sections?.[0]?.children?.[0];
    if (!child || !("paragraph" in child)) return undefined;
    const paragraphChild = child.paragraph.children?.[0];
    if (
      typeof paragraphChild !== "object" ||
      paragraphChild === null ||
      !("picture" in paragraphChild)
    ) {
      return undefined;
    }
    const picture = paragraphChild.picture;
    return {
      dpi: picture.dpi,
      rotWithShape: picture.rotWithShape,
      blackWhiteMode: picture.blackWhiteMode,
      geometry: picture.geometry,
      nonVisualProperties: picture.nonVisualProperties,
    };
  },
  expected: {
    dpi: 300,
    rotWithShape: true,
    blackWhiteMode: "gray",
    geometry: { preset: "rect" },
    nonVisualProperties: {
      id: 9,
      name: "Executable Picture",
      description: "Picture fixture",
      title: "Picture",
      hidden: true,
      preferRelativeResize: false,
      pictureLocks: { noChangeAspect: true },
    },
  },
};

export const OOXML_EXECUTABLE_REGISTRY: OoxmlExecutableRegistry = {
  cases: [
    commentIdCase,
    ...extensibleCases,
    lockedCanvasCase,
    w15SdtExtensionsCase,
    w15SdtRepeatingSectionCase,
    w15SdtRepeatingSectionItemCase,
    w15CommentsExtendedCase,
    w15PeopleCase,
    w15SettingsCase,
    w15FootnoteColumnsCase,
    w15CollapsedCase,
    w14SettingsCase,
    w14TrackChangesCase,
    w14CheckboxCase,
    w14EntityPickerCase,
    w14EffectsCase,
    ...w14EffectJoinVariants,
    w14LinearGradientCase,
    w14SolidFillCase,
    xdrTwoCellShapeCase,
    xdrOneCellPictureCase,
    xdrAbsoluteConnectorCase,
    xdrChartFrameCase,
    xdrGroupCase,
    xdrContentPartCase,
    chartDrawingShapeCase,
    chartDrawingConnectorCase,
    chartDrawingPictureCase,
    chartDrawingGraphicFrameCase,
    chartDrawingGroupCase,
    pictureCase,
  ],
};
