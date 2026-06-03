import { ContentTypes } from "@file/content-types";
import { CoreProperties, type CorePropertiesOptions } from "@file/core-properties";
import { Media } from "@file/media/media";
import { SharedStrings } from "@file/shared-strings";
import { Styles } from "@file/styles";
import type { DxfOptions } from "@file/styles";
import { DefaultTheme } from "@file/theme";
import {
  WorkbookXml,
  type SheetDefinition,
  type PivotCacheReference,
  type WorkbookProtectionOptions,
} from "@file/workbook";
import { Worksheet, type WorksheetOptions } from "@file/worksheet";
/**
 * File class (exported as Workbook) — the top-level container for XLSX documents.
 *
 * @module
 */
import { AppProperties, ChartCollection, Relationships } from "@office-open/core";

export interface WorkbookOptions extends CorePropertiesOptions {
  readonly worksheets?: readonly WorksheetOptions[];
  /** Pre-defined differential formats for conditional formatting */
  readonly dxfs?: readonly DxfOptions[];
  /** Workbook-level protection */
  readonly workbookProtection?: WorkbookProtectionOptions;
}

export class File {
  private readonly worksheetOptions: readonly WorksheetOptions[];
  private readonly corePropsOptions: CorePropertiesOptions;
  private readonly dxfOptions: readonly DxfOptions[];
  private readonly protectionOptions?: WorkbookProtectionOptions;

  // Lazy components
  private _coreProperties?: CoreProperties;
  private _appProperties?: AppProperties;
  private _contentTypes?: ContentTypes;
  private _styles?: Styles;
  private _theme?: DefaultTheme;
  private _workbookXml?: WorkbookXml;
  private _worksheets?: Worksheet[];
  private _sharedStrings?: SharedStrings;
  private _media?: Media;
  private _fileRels?: Relationships;
  private _workbookRels?: Relationships;
  private _charts?: ChartCollection;
  private _pivotCacheRefs: PivotCacheReference[] = [];

  public constructor(options: WorkbookOptions) {
    this.worksheetOptions = options.worksheets ?? [];
    this.corePropsOptions = options;
    this.dxfOptions = options.dxfs ?? [];
    this.protectionOptions = options.workbookProtection;
  }

  // ── Lazy getters ──

  public get coreProperties(): CoreProperties {
    return (this._coreProperties ??= new CoreProperties(this.corePropsOptions));
  }

  public get appProperties(): AppProperties {
    return (this._appProperties ??= new AppProperties());
  }

  public get contentTypes(): ContentTypes {
    if (!this._contentTypes) {
      this._contentTypes = new ContentTypes();
      for (let i = 0; i < this.worksheetOptions.length; i++) {
        this._contentTypes.addWorksheet(i + 1);
      }
      this._contentTypes.addStyles();
      this._contentTypes.addSharedStrings();
      this._contentTypes.addTheme();
    }
    return this._contentTypes;
  }

  public get styles(): Styles {
    return (this._styles ??= new Styles());
  }

  /**
   * Register a differential format and return its dxfId.
   * Call this before generating XML to use the ID in conditional formatting rules.
   */
  public registerDxf(opts: DxfOptions): number {
    return this.styles.registerDxf(opts);
  }

  public get theme(): DefaultTheme {
    return (this._theme ??= new DefaultTheme());
  }

  public get workbookXml(): WorkbookXml {
    if (!this._workbookXml) {
      const sheets: SheetDefinition[] = this.worksheetOptions.map((ws, i) => ({
        name: ws.name ?? `Sheet${i + 1}`,
        sheetId: i + 1,
        rId: `rId${i + 1}`,
      }));
      this._workbookXml = new WorkbookXml(sheets, this._pivotCacheRefs, this.protectionOptions);
    }
    return this._workbookXml;
  }

  public get sharedStrings(): SharedStrings {
    return (this._sharedStrings ??= new SharedStrings());
  }

  public get media(): Media {
    return (this._media ??= new Media());
  }

  public get charts(): ChartCollection {
    return (this._charts ??= new ChartCollection());
  }

  public get dxfEntries(): readonly DxfOptions[] {
    return this.dxfOptions;
  }

  public get pivotCacheRefs(): readonly PivotCacheReference[] {
    return this._pivotCacheRefs;
  }

  public registerPivotCache(cacheId: number, rId: string): void {
    this._pivotCacheRefs.push({ cacheId, rId });
  }

  public get worksheetConfigs(): readonly WorksheetOptions[] {
    return this.worksheetOptions;
  }

  public get worksheets(): readonly Worksheet[] {
    if (!this._worksheets) {
      this._worksheets = this.worksheetOptions.map((ws) => new Worksheet(ws));
    }
    return this._worksheets;
  }

  public get fileRelationships(): Relationships {
    if (!this._fileRels) {
      this._fileRels = new Relationships();
      this._fileRels.addRelationship(
        1,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
        "xl/workbook.xml",
      );
      this._fileRels.addRelationship(
        2,
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
        "docProps/core.xml",
      );
      this._fileRels.addRelationship(
        3,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
        "docProps/app.xml",
      );
    }
    return this._fileRels;
  }

  public get workbookRelationships(): Relationships {
    if (!this._workbookRels) {
      this._workbookRels = new Relationships();
      let rid = 1;
      for (let i = 0; i < this.worksheetOptions.length; i++) {
        this._workbookRels.addRelationship(
          rid++,
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
          `worksheets/sheet${i + 1}.xml`,
        );
      }
      this._workbookRels.addRelationship(
        rid++,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
        "styles.xml",
      );
      this._workbookRels.addRelationship(
        rid++,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
        "theme/theme1.xml",
      );
      this._workbookRels.addRelationship(
        rid++,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
        "sharedStrings.xml",
      );
    }
    return this._workbookRels;
  }
}
