/**
 * SmartArt built-in definitions — layout, quick style, and color transforms.
 *
 * Provides XML generators for layoutDef, styleDef, and colorsDef stubs.
 * The "default" layout uses a full embedded XML; all others use minimal stubs
 * that carry only the uniqueId — PowerPoint/Word resolves these to built-in definitions.
 *
 * @module
 */

export {
    COLOR_CATEGORIES,
    LAYOUT_CATEGORIES,
    STYLE_CATEGORIES,
} from "@office-open/core/smartart/categories";

import {
    LAYOUT_CATEGORIES,
    STYLE_CATEGORIES,
    COLOR_CATEGORIES,
} from "@office-open/core/smartart/categories";

// ---------------------------------------------------------------------------
// Layout XML — full for "default", stubs for everything else
// ---------------------------------------------------------------------------

const DGM_NS = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** Full default list layout (urn:microsoft.com/office/officeart/2005/8/layout/default) */
const FULL_DEFAULT_LAYOUT_XML =
    '<dgm:layoutDef xmlns:dgm="' +
    DGM_NS +
    '" xmlns:a="' +
    A_NS +
    '" uniqueId="urn:microsoft.com/office/officeart/2005/8/layout/default">' +
    '<dgm:title val=""/><dgm:desc val=""/>' +
    '<dgm:catLst><dgm:cat type="list" pri="400"/></dgm:catLst>' +
    '<dgm:sampData><dgm:dataModel><dgm:ptLst><dgm:pt modelId="0" type="doc"/><dgm:pt modelId="1"><dgm:prSet phldr="1"/></dgm:pt><dgm:pt modelId="2"><dgm:prSet phldr="1"/></dgm:pt><dgm:pt modelId="3"><dgm:prSet phldr="1"/></dgm:pt><dgm:pt modelId="4"><dgm:prSet phldr="1"/></dgm:pt><dgm:pt modelId="5"><dgm:prSet phldr="1"/></dgm:pt></dgm:ptLst><dgm:cxnLst><dgm:cxn modelId="6" srcId="0" destId="1" srcOrd="0" destOrd="0"/><dgm:cxn modelId="7" srcId="0" destId="2" srcOrd="1" destOrd="0"/><dgm:cxn modelId="8" srcId="0" destId="3" srcOrd="2" destOrd="0"/><dgm:cxn modelId="9" srcId="0" destId="4" srcOrd="3" destOrd="0"/><dgm:cxn modelId="10" srcId="0" destId="5" srcOrd="4" destOrd="0"/></dgm:cxnLst><dgm:bg/><dgm:whole/></dgm:dataModel></dgm:sampData>' +
    '<dgm:styleData><dgm:dataModel><dgm:ptLst><dgm:pt modelId="0" type="doc"/><dgm:pt modelId="1"/><dgm:pt modelId="2"/></dgm:ptLst><dgm:cxnLst><dgm:cxn modelId="3" srcId="0" destId="1" srcOrd="0" destOrd="0"/><dgm:cxn modelId="4" srcId="0" destId="2" srcOrd="1" destOrd="0"/></dgm:cxnLst><dgm:bg/><dgm:whole/></dgm:dataModel></dgm:styleData>' +
    '<dgm:clrData><dgm:dataModel><dgm:ptLst><dgm:pt modelId="0" type="doc"/><dgm:pt modelId="1"/><dgm:pt modelId="2"/><dgm:pt modelId="3"/><dgm:pt modelId="4"/><dgm:pt modelId="5"/><dgm:pt modelId="6"/></dgm:ptLst><dgm:cxnLst><dgm:cxn modelId="7" srcId="0" destId="1" srcOrd="0" destOrd="0"/><dgm:cxn modelId="8" srcId="0" destId="2" srcOrd="1" destOrd="0"/><dgm:cxn modelId="9" srcId="0" destId="3" srcOrd="2" destOrd="0"/><dgm:cxn modelId="10" srcId="0" destId="4" srcOrd="3" destOrd="0"/><dgm:cxn modelId="11" srcId="0" destId="5" srcOrd="4" destOrd="0"/><dgm:cxn modelId="12" srcId="0" destId="6" srcOrd="5" destOrd="0"/></dgm:cxnLst><dgm:bg/><dgm:whole/></dgm:dataModel></dgm:clrData>' +
    '<dgm:layoutNode name="diagram"><dgm:varLst><dgm:dir/><dgm:resizeHandles val="exact"/></dgm:varLst>' +
    '<dgm:choose name="Name0"><dgm:if name="Name1" func="var" arg="dir" op="equ" val="norm"><dgm:alg type="snake"><dgm:param type="grDir" val="tL"/><dgm:param type="flowDir" val="row"/><dgm:param type="contDir" val="sameDir"/><dgm:param type="off" val="ctr"/></dgm:alg></dgm:if><dgm:else name="Name2"><dgm:alg type="snake"><dgm:param type="grDir" val="tR"/><dgm:param type="flowDir" val="row"/><dgm:param type="contDir" val="sameDir"/><dgm:param type="off" val="ctr"/></dgm:alg></dgm:else></dgm:choose>' +
    '<dgm:shape xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf/>' +
    '<dgm:constrLst><dgm:constr type="w" for="ch" forName="node" refType="w"/><dgm:constr type="h" for="ch" forName="node" refType="w" refFor="ch" refForName="node" fact="0.6"/><dgm:constr type="w" for="ch" forName="sibTrans" refType="w" refFor="ch" refForName="node" fact="0.1"/><dgm:constr type="sp" refType="w" refFor="ch" refForName="sibTrans"/><dgm:constr type="primFontSz" for="ch" forName="node" op="equ" val="65"/></dgm:constrLst><dgm:ruleLst/>' +
    '<dgm:forEach name="Name3" axis="ch" ptType="node">' +
    '<dgm:layoutNode name="node"><dgm:varLst><dgm:bulletEnabled val="1"/></dgm:varLst>' +
    '<dgm:alg type="tx"/><dgm:shape type="rect" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:blip="">' +
    '<dgm:adjLst/></dgm:shape><dgm:presOf axis="desOrSelf" ptType="node"/>' +
    '<dgm:constrLst><dgm:constr type="lMarg" refType="primFontSz" fact="0.3"/>' +
    '<dgm:constr type="rMarg" refType="primFontSz" fact="0.3"/>' +
    '<dgm:constr type="tMarg" refType="primFontSz" fact="0.3"/>' +
    '<dgm:constr type="bMarg" refType="primFontSz" fact="0.3"/></dgm:constrLst>' +
    '<dgm:ruleLst><dgm:rule type="primFontSz" val="5" fact="NaN" max="NaN"/></dgm:ruleLst></dgm:layoutNode>' +
    '<dgm:forEach name="Name4" axis="followSib" ptType="sibTrans" cnt="1">' +
    '<dgm:layoutNode name="sibTrans"><dgm:alg type="sp"/>' +
    '<dgm:shape xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:blip="">' +
    "<dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:constrLst/><dgm:ruleLst/>" +
    "</dgm:layoutNode></dgm:forEach></dgm:layoutNode></dgm:layoutDef>";

/**
 * Returns layout XML. Full XML for "default", minimal stub for others.
 * Stub has no layoutNode so PowerPoint falls back to built-in definitions
 * based on the uniqueId / loTypeId in the data model.
 */
export function getLayoutXml(layoutId: string): string {
    if (layoutId === "default") {
        return XML_DECL + FULL_DEFAULT_LAYOUT_XML;
    }
    const cat = LAYOUT_CATEGORIES[layoutId] ?? "list";
    return (
        XML_DECL +
        '<dgm:layoutDef xmlns:dgm="' +
        DGM_NS +
        '" xmlns:a="' +
        A_NS +
        '" uniqueId="urn:microsoft.com/office/officeart/2005/8/layout/' +
        layoutId +
        '">' +
        '<dgm:title val=""/><dgm:desc val=""/>' +
        '<dgm:catLst><dgm:cat type="' +
        cat +
        '" pri="400"/></dgm:catLst>' +
        '<dgm:sampData><dgm:dataModel><dgm:ptLst><dgm:pt modelId="0" type="doc"/><dgm:pt modelId="1"><dgm:prSet phldr="1"/></dgm:pt><dgm:pt modelId="2"><dgm:prSet phldr="1"/></dgm:pt></dgm:ptLst><dgm:cxnLst><dgm:cxn modelId="3" srcId="0" destId="1" srcOrd="0" destOrd="0"/><dgm:cxn modelId="4" srcId="0" destId="2" srcOrd="1" destOrd="0"/></dgm:cxnLst><dgm:bg/><dgm:whole/></dgm:dataModel></dgm:sampData>' +
        "</dgm:layoutDef>"
    );
}

/**
 * Returns style stub XML with the given uniqueId.
 */
export function getStyleXml(styleId: string): string {
    const cat = STYLE_CATEGORIES[styleId] ?? "simple";
    return (
        XML_DECL +
        '<dgm:styleDef xmlns:dgm="' +
        DGM_NS +
        '" xmlns:a="' +
        A_NS +
        '" uniqueId="urn:microsoft.com/office/officeart/2005/8/quickstyle/' +
        styleId +
        '">' +
        '<dgm:title val=""/><dgm:desc val=""/>' +
        '<dgm:catLst><dgm:cat type="' +
        cat +
        '" pri="10100"/></dgm:catLst>' +
        '<dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d>' +
        "</dgm:styleDef>"
    );
}

/**
 * Returns color stub XML with the given uniqueId.
 */
export function getColorXml(colorId: string): string {
    const cat = COLOR_CATEGORIES[colorId] ?? "accent1";
    return (
        XML_DECL +
        '<dgm:colorsDef xmlns:dgm="' +
        DGM_NS +
        '" xmlns:a="' +
        A_NS +
        '" uniqueId="urn:microsoft.com/office/officeart/2005/8/colors/' +
        colorId +
        '">' +
        '<dgm:title val=""/><dgm:desc val=""/>' +
        '<dgm:catLst><dgm:cat type="' +
        cat +
        '" pri="11200"/></dgm:catLst>' +
        "</dgm:colorsDef>"
    );
}

/** Minimal drawing cache for SmartArt (PowerPoint/WPS auto-regenerates this on open) */
export const DEFAULT_DRAWING_XML =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<dsp:drawing xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    "<dsp:spTree>" +
    '<dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr>' +
    "<dsp:grpSpPr/>" +
    "</dsp:spTree>" +
    "</dsp:drawing>";
