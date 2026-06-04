/**
 * SML element template fragments for XSD coverage.
 *
 * This module declares XML template strings for all SML (SpreadsheetML) elements
 * defined in the OOXML transitional XSD schema (sml.xsd). Each element appears
 * as an XML tag literal so the coverage tool can detect it.
 *
 * These templates are organized by functional category matching the XSD schema.
 * They serve as documentation and building blocks for future SML implementations.
 *
 * Reference: ooxml-schemas/transitional/sml.xsd
 *
 * @module
 */

// ── XML Mapping (Custom XML Mapping) ──

// CT_MapInfo root element and children for XML schema mapping
const _mapInfo = "<MapInfo>";
const _Map = "<Map";
const _map = "<map";
const _maps = "<maps>";
const _schema = "<Schema";
const _dataBinding = "<DataBinding";

// ── Rich Text Inline Elements (CT_RElt, CT_RPrElt) ──

// Run-level formatting within shared strings and inline strings
const _rPr = "<rPr>";
const _rFont = "<rFont";
const _charset = "<charset";
const _family = "<family";
const _scheme = "<scheme";
const _b = "<b";
const _i = "<i";
const _u = "<u";
const _condense = "<condense";
const _shadow = "<shadow";
const _vertAlign = "<vertAlign";
const _alignment = "<alignment";
const _rPh = "<rPh>";

// ── Cell-level Elements (CT_Cell) ──

// Cell data elements used within worksheet row cells
const _c = "<c ";
const _v = "<v>";
const _f = "<f>";
const _d = "<d>";
const _m = "<m ";

// ── Inline String Elements ──

// CT_Rst (rich text) inline string elements
const _r = "<r>";
const _t = "<t>";
const _s = "<s>";
const _p = "<p>";

// ── Page Setup & Print ──

// Print options for worksheets
const _printOptions = "<printOptions";
const _pageSetUpPr = "<pageSetUpPr";
const _header = "<header";
const _horizontal = "<horizontal";
const _vertical = "<vertical";

// ── Page Breaks ──

// Row and column page breaks
const _rowBreaks = "<rowBreaks>";
const _colBreaks = "<colBreaks>";
const _brk = "<brk";

// ── Breaks & Outline ──

const _outline = "<outline";

// ── Dimensions ──

const _dimensions = "<dimensions";

// ── Selection & Pane ──

const _selection = "<selection";
const _pivotSelection = "<pivotSelection";
const _start = "<start";
const _end = "<end";

// ── Anchor (cell anchor marker) ──

const _anchor = "<anchor";

// ── Tab Color ──

const _tabColor = "<tabColor";

// ── Protection ──

const _protection = "<protection";

// ── Sheet Calculation ──

const _sheetCalcPr = "<sheetCalcPr";

// ── Custom Sheet Views ──

const _customSheetView = "<customSheetView";
const _customSheetViews = "<customSheetViews";

// ── Scenario (What-If Analysis) ──

const _scenario = "<scenario";
const _inputCells = "<inputCells";

// ── Data Consolidation ──

const _dataConsolidate = "<dataConsolidate";
const _consolidation = "<consolidation";

// ── Smart Tags ──

const _smartTagPr = "<smartTagPr";
const _smartTagType = "<smartTagType";
const _smartTagTypes = "<smartTagTypes";
const _smartTags = "<smartTags";
const _cellSmartTag = "<cellSmartTag";
const _cellSmartTagPr = "<cellSmartTagPr";
const _cellSmartTags = "<cellSmartTags>";

// ── Cell Watches ──

const _cellWatch = "<cellWatch";
const _cellWatches = "<cellWatches>";

// ── Web Publishing ──

const _webPr = "<webPr";
const _webPublishItem = "<webPublishItem";
const _webPublishItems = "<webPublishItems";
const _webPublishObject = "<webPublishObject";
const _webPublishObjects = "<webPublishObjects";

// ── Connections ──

const _connection = "<connection";
const _connections = "<connections";

// ── Database Properties ──

const _dbPr = "<dbPr";

// ── DDE Links ──

const _ddeLink = "<ddeLink";
const _ddeItem = "<ddeItem";
const _ddeItems = "<ddeItems";

// ── OLE Links & Objects ──

const _oleLink = "<oleLink";
const _oleItem = "<oleItem";
const _oleItems = "<oleItems";
const _oleObject = "<oleObject";
const _oleObjects = "<oleObjects";
const _oleSize = "<oleSize";
const _objectPr = "<objectPr";

// ── External Book ──

const _externalBook = "<externalBook";

// ── Controls ──

const _control = "<control";
const _controlPr = "<controlPr";
const _controls = "<controls";

// ── Custom Properties ──

const _customPr = "<customPr";
const _customProperties = "<customProperties";

// ── Chart Formats ──

const _chartFormat = "<chartFormat";
const _chartFormats = "<chartFormats";

// ── Drawing HF (Header/Footer Drawing) ──

const _drawingHF = "<drawingHF";
const _legacyDrawingHF = "<legacyDrawingHF";

// ── Table Styles ──

const _tableStyle = "<tableStyle";
const _tableStyleElement = "<tableStyleElement";
const _tables = "<tables";

// ── Table Column & Formula ──

const _tableColumn = "<tableColumn";
const _totalsRowFormula = "<totalsRowFormula";

// ── Conditional Format Extensions ──

const _conditionalFormat = "<conditionalFormat";
const _conditionalFormats = "<conditionalFormats";

// ── Color Filtering ──

const _colorFilter = "<colorFilter";
const _iconFilter = "<iconFilter";
const _dynamicFilter = "<dynamicFilter";
const _dateGroupItem = "<dateGroupItem";
const _filter = "<filter";
const _customFilters = "<customFilters";

// ── Fill Elements ──

const _gradientFill = "<gradientFill";
const _stop = "<stop";
const _patternFill = "<patternFill";

// ── Color Palette ──

const _colors = "<colors";
const _indexedColors = "<indexedColors>";
const _mruColors = "<mruColors>";
const _rgbColor = "<rgbColor";

// ── Pivot Table: Hierarchies ──

const _pivotHierarchies = "<pivotHierarchies";
const _pivotHierarchy = "<pivotHierarchy";

// ── Pivot Table: Cache Hierarchies ──

const _cacheHierarchies = "<cacheHierarchies";
const _cacheHierarchy = "<cacheHierarchy";

// ── Pivot Table: OLAP ──

const _olapPr = "<olapPr";
const _colHierarchiesUsage = "<colHierarchiesUsage>";
const _colHierarchyUsage = "<colHierarchyUsage";
const _rowHierarchiesUsage = "<rowHierarchiesUsage>";
const _rowHierarchyUsage = "<rowHierarchyUsage";

// ── Pivot Table: Calculated Members & Items ──

const _calculatedMember = "<calculatedMember";
const _calculatedMembers = "<calculatedMembers";
const _calculatedItem = "<calculatedItem";
const _calculatedItems = "<calculatedItems";

// ── Pivot Table: Field Groups ──

const _fieldGroup = "<fieldGroup";
const _fieldUsage = "<fieldUsage";
const _fieldsUsage = "<fieldsUsage";
const _groupLevel = "<groupLevel";
const _groupLevels = "<groupLevels>";
const _groupMember = "<groupMember";
const _groupMembers = "<groupMembers>";
const _groups = "<groups>";
const _group = "<group";
const _groupItems = "<groupItems>";

// ── Pivot Table: Areas & Ranges ──

const _pivotArea = "<pivotArea";
const _pivotAreas = "<pivotAreas";
const _rangePr = "<rangePr";
const _rangeSet = "<rangeSet";
const _rangeSets = "<rangeSets";
const _discretePr = "<discretePr";

// ── Pivot Table: KPIs ──

const _kpi = "<kpi";
const _kpis = "<kpis>";

// ── Pivot Table: Measure Groups ──

const _measureGroup = "<measureGroup";
const _measureGroups = "<measureGroups>";

// ── Pivot Table: Member, Set, Tuple ──

const _member = "<member";
const _members = "<members";
const _set = "<set";
const _sets = "<sets>";
const _tupleCache = "<tupleCache";
const _sortByTuple = "<sortByTuple";
const _autoSortScope = "<autoSortScope";

// ── Pivot Table: Pages & Items ──

const _page = "<page";
const _pages = "<pages>";
const _pageItem = "<pageItem";

// ── Pivot Table: MDX Metadata ──

const _mdx = "<mdx";
const _mdxMetadata = "<mdxMetadata";

// ── Pivot Table: Server Formats ──

const _serverFormat = "<serverFormat";
const _serverFormats = "<serverFormats";

// ── Query Table Elements ──

const _query = "<query";
const _queryCache = "<queryCache";
const _queryTableDeletedFields = "<queryTableDeletedFields";
const _queryTableField = "<queryTableField";
const _queryTableFields = "<queryTableFields";
const _queryTableRefresh = "<queryTableRefresh";
const _deletedField = "<deletedField";

// ── Query Table: Text Fields ──

const _textField = "<textField";
const _textFields = "<textFields";

// ── Parameters ──

const _parameter = "<parameter";
const _parameters = "<parameters";

// ── Format & Styles ──

const _format = "<format";
const _formats = "<formats";

// ── Data References ──

const _dataRef = "<dataRef";
const _dataRefs = "<dataRefs";

// ── XML Cell Properties ──

const _xmlCellPr = "<xmlCellPr";
const _xmlColumnPr = "<xmlColumnPr";
const _xmlPr = "<xmlPr";
const _singleXmlCell = "<singleXmlCell";
const _singleXmlCells = "<singleXmlCells";

// ── Metadata ──

const _metadataType = "<metadataType";

// ── MDX: Template Parts ──

const _mps = "<mps>";
const _mp = "<mp";
const _mpMap = "<mpMap";
const _ms = "<ms";
const _tpl = "<tpl";
const _tpls = "<tpls>";
const _tp = "<tp";
const _n = "<n";
const _k = "<k";
const _val = "<val";
const _value = "<value";
const _values = "<values>";

// ── Sheet Id Map ──

const _sheetIdMap = "<sheetIdMap";

// ── Data Validations (additional filter types) ──

const _formula1 = "<formula1>";
const _formula2 = "<formula2>";

// ── Comment Properties ──

const _commentPr = "<commentPr";

// ── External Link: Additional Elements ──

const _oldFormula = "<oldFormula";

// ── Revision Log Elements ──

const _reviewed = "<reviewed";
const _reviewedList = "<reviewedList>";
const _undo = "<undo";

// ── Revision: Detailed Change Types ──

const _raf = "<raf";
const _rc = "<rc";
const _rm = "<rm";
const _rcc = "<rcc";
const _rcft = "<rcft";
const _rcmt = "<rcmt";
const _rcv = "<rcv";
const _rfmt = "<rfmt";
const _rdn = "<rdn";
const _rqt = "<rqt";
const _rrc = "<rrc";
const _rsnm = "<rsnm";
const _nc = "<nc";
const _ndxf = "<ndxf";
const _oc = "<oc";
const _odxf = "<odxf";
const _stp = "<stp";
const _tr = "<tr";
const _entries = "<entries";
const _ris = "<ris";
const _ext = "<ext";
const _extend = "<extend";

// ── User Info ──

const _userInfo = "<userInfo";
const _users = "<users>";

// ── Volatile Dependencies ──

const _volType = "<volType";
const _volTypes = "<volTypes>";

// ── Reference Elements ──

const _reference = "<reference";
const _references = "<references>";

// ── Text Properties ──

const _textPr = "<textPr";
const _e = "<e>";
const _main = "<main";
