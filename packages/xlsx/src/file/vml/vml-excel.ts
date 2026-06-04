/**
 * VML SpreadsheetDrawing elements (x: namespace).
 *
 * Factory functions for Excel-specific VML elements used in ClientData
 * for form controls, comments, and embedded objects.
 *
 * Reference: ISO/IEC 29500-4, vml-spreadsheetDrawing.xsd
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

// ── Boolean Elements (ST_TrueFalseBlank) ──

/**
 * Creates an x:MoveWithCells element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="MoveWithCells" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXMoveWithCells = (): XmlComponent =>
  new BuilderElement({ name: "x:MoveWithCells" });

/**
 * Creates an x:SizeWithCells element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="SizeWithCells" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXSizeWithCells = (): XmlComponent =>
  new BuilderElement({ name: "x:SizeWithCells" });

/**
 * Creates an x:Locked element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="Locked" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXLocked = (): XmlComponent => new BuilderElement({ name: "x:Locked" });

/**
 * Creates an x:DefaultSize element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="DefaultSize" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXDefaultSize = (): XmlComponent => new BuilderElement({ name: "x:DefaultSize" });

/**
 * Creates an x:PrintObject element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="PrintObject" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXPrintObject = (): XmlComponent => new BuilderElement({ name: "x:PrintObject" });

/**
 * Creates an x:Disabled element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="Disabled" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXDisabled = (): XmlComponent => new BuilderElement({ name: "x:Disabled" });

/**
 * Creates an x:AutoFill element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="AutoFill" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXAutoFill = (): XmlComponent => new BuilderElement({ name: "x:AutoFill" });

/**
 * Creates an x:AutoLine element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="AutoLine" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXAutoLine = (): XmlComponent => new BuilderElement({ name: "x:AutoLine" });

/**
 * Creates an x:AutoPict element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="AutoPict" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXAutoPict = (): XmlComponent => new BuilderElement({ name: "x:AutoPict" });

/**
 * Creates an x:AutoScale element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="AutoScale" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXAutoScale = (): XmlComponent => new BuilderElement({ name: "x:AutoScale" });

/**
 * Creates an x:LockText element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="LockText" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXLockText = (): XmlComponent => new BuilderElement({ name: "x:LockText" });

/**
 * Creates an x:JustLastX element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="JustLastX" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXJustLastX = (): XmlComponent => new BuilderElement({ name: "x:JustLastX" });

/**
 * Creates an x:SecretEdit element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="SecretEdit" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXSecretEdit = (): XmlComponent => new BuilderElement({ name: "x:SecretEdit" });

/**
 * Creates an x:Default element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="Default" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXDefault = (): XmlComponent => new BuilderElement({ name: "x:Default" });

/**
 * Creates an x:Help element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="Help" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXHelp = (): XmlComponent => new BuilderElement({ name: "x:Help" });

/**
 * Creates an x:Cancel element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="Cancel" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXCancel = (): XmlComponent => new BuilderElement({ name: "x:Cancel" });

/**
 * Creates an x:Dismiss element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="Dismiss" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXDismiss = (): XmlComponent => new BuilderElement({ name: "x:Dismiss" });

/**
 * Creates an x:Visible element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="Visible" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXVisible = (): XmlComponent => new BuilderElement({ name: "x:Visible" });

/**
 * Creates an x:RowHidden element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="RowHidden" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXRowHidden = (): XmlComponent => new BuilderElement({ name: "x:RowHidden" });

/**
 * Creates an x:ColHidden element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="ColHidden" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXColHidden = (): XmlComponent => new BuilderElement({ name: "x:ColHidden" });

/**
 * Creates an x:MultiLine element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="MultiLine" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXMultiLine = (): XmlComponent => new BuilderElement({ name: "x:MultiLine" });

/**
 * Creates an x:VScroll element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="VScroll" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXVScroll = (): XmlComponent => new BuilderElement({ name: "x:VScroll" });

/**
 * Creates an x:ValidIds element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="ValidIds" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXValidIds = (): XmlComponent => new BuilderElement({ name: "x:ValidIds" });

/**
 * Creates an x:NoThreeD element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="NoThreeD" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXNoThreeD = (): XmlComponent => new BuilderElement({ name: "x:NoThreeD" });

/**
 * Creates an x:NoThreeD2 element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="NoThreeD2" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXNoThreeD2 = (): XmlComponent => new BuilderElement({ name: "x:NoThreeD2" });

/**
 * Creates an x:FirstButton element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="FirstButton" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXFirstButton = (): XmlComponent => new BuilderElement({ name: "x:FirstButton" });

/**
 * Creates an x:Colored element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="Colored" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXColored = (): XmlComponent => new BuilderElement({ name: "x:Colored" });

/**
 * Creates an x:Horiz element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="Horiz" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXHoriz = (): XmlComponent => new BuilderElement({ name: "x:Horiz" });

/**
 * Creates an x:MapOCX element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="MapOCX" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXMapOCX = (): XmlComponent => new BuilderElement({ name: "x:MapOCX" });

/**
 * Creates an x:Camera element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="Camera" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXCamera = (): XmlComponent => new BuilderElement({ name: "x:Camera" });

/**
 * Creates an x:RecalcAlways element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="RecalcAlways" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXRecalcAlways = (): XmlComponent =>
  new BuilderElement({ name: "x:RecalcAlways" });

/**
 * Creates an x:DDE element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="DDE" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXDDE = (): XmlComponent => new BuilderElement({ name: "x:DDE" });

/**
 * Creates an x:UIObj element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="UIObj" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createXUIObj = (): XmlComponent => new BuilderElement({ name: "x:UIObj" });

// ── String Elements ──

/**
 * Creates an x:Anchor element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="Anchor" type="xsd:string"/>
 * ```
 */
export const createXAnchor = (): XmlComponent => new BuilderElement({ name: "x:Anchor" });

/**
 * Creates an x:FmlaMacro element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="FmlaMacro" type="xsd:string"/>
 * ```
 */
export const createXFmlaMacro = (): XmlComponent => new BuilderElement({ name: "x:FmlaMacro" });

/**
 * Creates an x:TextHAlign element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="TextHAlign" type="xsd:string"/>
 * ```
 */
export const createXTextHAlign = (): XmlComponent => new BuilderElement({ name: "x:TextHAlign" });

/**
 * Creates an x:TextVAlign element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="TextVAlign" type="xsd:string"/>
 * ```
 */
export const createXTextVAlign = (): XmlComponent => new BuilderElement({ name: "x:TextVAlign" });

/**
 * Creates an x:FmlaRange element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="FmlaRange" type="xsd:string"/>
 * ```
 */
export const createXFmlaRange = (): XmlComponent => new BuilderElement({ name: "x:FmlaRange" });

/**
 * Creates an x:SelType element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="SelType" type="xsd:string"/>
 * ```
 */
export const createXSelType = (): XmlComponent => new BuilderElement({ name: "x:SelType" });

/**
 * Creates an x:MultiSel element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="MultiSel" type="xsd:string"/>
 * ```
 */
export const createXMultiSel = (): XmlComponent => new BuilderElement({ name: "x:MultiSel" });

/**
 * Creates an x:LCT element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="LCT" type="xsd:string"/>
 * ```
 */
export const createXLCT = (): XmlComponent => new BuilderElement({ name: "x:LCT" });

/**
 * Creates an x:ListItem element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="ListItem" type="xsd:string"/>
 * ```
 */
export const createXListItem = (): XmlComponent => new BuilderElement({ name: "x:ListItem" });

/**
 * Creates an x:DropStyle element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="DropStyle" type="xsd:string"/>
 * ```
 */
export const createXDropStyle = (): XmlComponent => new BuilderElement({ name: "x:DropStyle" });

/**
 * Creates an x:FmlaLink element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="FmlaLink" type="xsd:string"/>
 * ```
 */
export const createXFmlaLink = (): XmlComponent => new BuilderElement({ name: "x:FmlaLink" });

/**
 * Creates an x:FmlaPict element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="FmlaPict" type="xsd:string"/>
 * ```
 */
export const createXFmlaPict = (): XmlComponent => new BuilderElement({ name: "x:FmlaPict" });

/**
 * Creates an x:FmlaGroup element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="FmlaGroup" type="xsd:string"/>
 * ```
 */
export const createXFmlaGroup = (): XmlComponent => new BuilderElement({ name: "x:FmlaGroup" });

/**
 * Creates an x:ScriptText element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="ScriptText" type="xsd:string"/>
 * ```
 */
export const createXScriptText = (): XmlComponent => new BuilderElement({ name: "x:ScriptText" });

/**
 * Creates an x:ScriptExtended element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="ScriptExtended" type="xsd:string"/>
 * ```
 */
export const createXScriptExtended = (): XmlComponent =>
  new BuilderElement({ name: "x:ScriptExtended" });

/**
 * Creates an x:FmlaTxbx element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="FmlaTxbx" type="xsd:string"/>
 * ```
 */
export const createXFmlaTxbx = (): XmlComponent => new BuilderElement({ name: "x:FmlaTxbx" });

// ── Integer Elements ──

/**
 * Creates an x:Accel element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="Accel" type="xsd:integer"/>
 * ```
 */
export const createXAccel = (): XmlComponent => new BuilderElement({ name: "x:Accel" });

/**
 * Creates an x:Accel2 element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="Accel2" type="xsd:integer"/>
 * ```
 */
export const createXAccel2 = (): XmlComponent => new BuilderElement({ name: "x:Accel2" });

/**
 * Creates an x:Row element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="Row" type="xsd:integer"/>
 * ```
 */
export const createXRow = (): XmlComponent => new BuilderElement({ name: "x:Row" });

/**
 * Creates an x:Column element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="Column" type="xsd:integer"/>
 * ```
 */
export const createXColumn = (): XmlComponent => new BuilderElement({ name: "x:Column" });

/**
 * Creates an x:VTEdit element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="VTEdit" type="xsd:integer"/>
 * ```
 */
export const createXVTEdit = (): XmlComponent => new BuilderElement({ name: "x:VTEdit" });

/**
 * Creates an x:WidthMin element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="WidthMin" type="xsd:integer"/>
 * ```
 */
export const createXWidthMin = (): XmlComponent => new BuilderElement({ name: "x:WidthMin" });

/**
 * Creates an x:Sel element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="Sel" type="xsd:integer"/>
 * ```
 */
export const createXSel = (): XmlComponent => new BuilderElement({ name: "x:Sel" });

/**
 * Creates an x:DropLines element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="DropLines" type="xsd:integer"/>
 * ```
 */
export const createXDropLines = (): XmlComponent => new BuilderElement({ name: "x:DropLines" });

/**
 * Creates an x:Checked element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="Checked" type="xsd:integer"/>
 * ```
 */
export const createXChecked = (): XmlComponent => new BuilderElement({ name: "x:Checked" });

/**
 * Creates an x:Val element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="Val" type="xsd:integer"/>
 * ```
 */
export const createXVal = (): XmlComponent => new BuilderElement({ name: "x:Val" });

/**
 * Creates an x:Min element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="Min" type="xsd:integer"/>
 * ```
 */
export const createXMin = (): XmlComponent => new BuilderElement({ name: "x:Min" });

/**
 * Creates an x:Max element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="Max" type="xsd:integer"/>
 * ```
 */
export const createXMax = (): XmlComponent => new BuilderElement({ name: "x:Max" });

/**
 * Creates an x:Inc element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="Inc" type="xsd:integer"/>
 * ```
 */
export const createXInc = (): XmlComponent => new BuilderElement({ name: "x:Inc" });

/**
 * Creates an x:Page element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="Page" type="xsd:integer"/>
 * ```
 */
export const createXPage = (): XmlComponent => new BuilderElement({ name: "x:Page" });

/**
 * Creates an x:Dx element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="Dx" type="xsd:integer"/>
 * ```
 */
export const createXDx = (): XmlComponent => new BuilderElement({ name: "x:Dx" });

// ── Non-Negative Integer Elements ──

/**
 * Creates an x:ScriptLanguage element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="ScriptLanguage" type="xsd:nonNegativeInteger"/>
 * ```
 */
export const createXScriptLanguage = (): XmlComponent =>
  new BuilderElement({ name: "x:ScriptLanguage" });

/**
 * Creates an x:ScriptLocation element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="ScriptLocation" type="xsd:nonNegativeInteger"/>
 * ```
 */
export const createXScriptLocation = (): XmlComponent =>
  new BuilderElement({ name: "x:ScriptLocation" });

// ── Special Type Elements ──

/**
 * Creates an x:CF element (Clipboard Format).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="CF" type="ST_CF"/>
 * ```
 */
export const createXCF = (): XmlComponent => new BuilderElement({ name: "x:CF" });
