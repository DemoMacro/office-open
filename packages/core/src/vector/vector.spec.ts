import { parse as parseXml } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import { describe, it, expect } from "vite-plus/test";

import {
  stringifyVmlShapeDefaults,
  parseVmlShapeDefaults,
  stringifyVmlShapeLayout,
  parseVmlShapeLayout,
  type VmlShapeDefaultsOptions,
  type VmlShapeLayoutOptions,
} from "./office-shape-defaults";
import {
  stringifyVmlClientData,
  parseVmlClientData,
  type VmlClientDataOptions,
} from "./shape-elements/client-data";
import { stringifyVmlFill, parseVmlFill, type VmlFillOptions } from "./shape-elements/fill";
import {
  stringifyVmlFormulas,
  parseVmlFormulas,
  type VmlFormulasOptions,
} from "./shape-elements/formulas";
import {
  stringifyVmlHandles,
  parseVmlHandles,
  type VmlHandlesOptions,
} from "./shape-elements/handles";
import {
  stringifyVmlImageData,
  parseVmlImageData,
  type VmlImageDataOptions,
} from "./shape-elements/imagedata";
import {
  stringifyVmlSkew,
  parseVmlSkew,
  stringifyVmlExtrusion,
  parseVmlExtrusion,
  stringifyVmlCallout,
  parseVmlCallout,
  stringifyVmlLock,
  parseVmlLock,
  stringifyVmlClipPath,
  parseVmlClipPath,
  stringifyVmlSignatureLine,
  parseVmlSignatureLine,
  stringifyVmlInk,
  parseVmlInk,
  stringifyVmlDiagram,
  parseVmlDiagram,
  stringifyVmlOleObject,
  parseVmlOleObject,
  type VmlSkewOptions,
  type VmlExtrusionOptions,
  type VmlCalloutOptions,
  type VmlLockOptions,
  type VmlSignatureLineOptions,
  type VmlDiagramOptions,
  type VmlOleObjectOptions,
} from "./shape-elements/office-elements";
import { stringifyVmlPath, parseVmlPath, type VmlPathOptions } from "./shape-elements/path";
import {
  stringifyVmlTextData,
  parseVmlTextData,
  type VmlTextDataOptions,
} from "./shape-elements/presentation-elements";
import { stringifyVmlShadow, parseVmlShadow, type VmlShadowOptions } from "./shape-elements/shadow";
import { stringifyVmlStroke, parseVmlStroke, type VmlStrokeOptions } from "./shape-elements/stroke";
import {
  stringifyVmlTextbox,
  parseVmlTextbox,
  type VmlTextboxOptions,
} from "./shape-elements/textbox";
import {
  stringifyVmlTextPath,
  parseVmlTextPath,
  type VmlTextPathOptions,
} from "./shape-elements/textpath";
import {
  stringifyVmlWrap,
  parseVmlWrap,
  stringifyVmlBorder,
  parseVmlBorder,
  type VmlBorderOptions,
} from "./shape-elements/word-elements";
import {
  stringifyVmlShape,
  parseVmlShape,
  stringifyVmlShapetype,
  parseVmlShapetype,
  stringifyVmlGroup,
  parseVmlGroup,
  stringifyVmlBackground,
  parseVmlBackground,
  stringifyVmlArc,
  parseVmlArc,
  stringifyVmlCurve,
  parseVmlCurve,
  stringifyVmlImage,
  parseVmlImage,
  stringifyVmlLine,
  parseVmlLine,
  stringifyVmlOval,
  parseVmlOval,
  stringifyVmlPolyline,
  parseVmlPolyline,
  stringifyVmlRect,
  parseVmlRect,
  stringifyVmlRoundRect,
  parseVmlRoundRect,
  type VmlShapeOptions,
  type VmlGroupOptions,
  type VmlArcOptions,
  type VmlCurveOptions,
  type VmlLineOptions,
  type VmlPolylineOptions,
  type VmlRoundRectOptions,
} from "./shapes";
import { stringifyVmlStyle, parseVmlShapeStyle, parseVmlStyle, type VmlShapeStyle } from "./style";

/** Wrap standalone child-element XML in a carrier so it parses as a document. */
function roundTrip<T>(stringify: (opts: T) => string, parse: (el: XmlElement) => T, opts: T): T {
  const xml = `<v:carrier>${stringify(opts)}</v:carrier>`;
  const doc = parseXml(xml);
  const carrier = doc.elements?.[0];
  if (!carrier) throw new Error("parsed document has no root element");
  const el = carrier.elements?.find((child) => child.type === "element");
  if (!el) throw new Error("carrier has no child element");
  return parse(el);
}

describe("vml style", () => {
  it("round-trips a full style", () => {
    const style: VmlShapeStyle = {
      position: "absolute",
      left: "10pt",
      top: 20,
      width: "5cm",
      height: 50,
      zIndex: 3,
      visibility: "hidden",
    };
    expect(parseVmlShapeStyle(parseVmlStyle(stringifyVmlStyle(style)))).toEqual(style);
  });
});

describe("v:shape", () => {
  it("round-trips attrs and all v: children", () => {
    const opts: VmlShapeOptions = {
      id: "_x0000_s1026",
      type: "#_x0000_t202",
      adj: "1,500",
      filled: false,
      fillcolor: "#3366ff",
      stroked: true,
      strokecolor: "red",
      strokeweight: "1.5pt",
      style: { width: 100, height: 50 },
      pathElement: { v: "m0,0l100,0,0,100xe", textboxrect: "5,5,95,95" },
      formulas: { equations: ["val #0", "sum 0 0 @0"] },
      handles: { handles: [{ position: "topLeft,#0", xrange: "0 100" }] },
      fill: { type: "gradient", color: "#fff", color2: "#000", angle: 45, method: "linear sigma" },
      stroke: { weight: "2pt", color: "blue", startarrow: "classic", endarrow: "oval" },
      shadow: { on: true, color: "silver", offset: "2pt,2pt" },
      textbox: { txbxContent: "<w:p/>" },
      textpath: { on: true, string: "WordArt" },
      imagedata: { src: "{image1.png}", relationshipId: "rId5", cropleft: "-0.1" },
    };
    const result = roundTrip(stringifyVmlShape, parseVmlShape, opts);
    expect(result).toEqual(opts);
  });

  it("keeps the path attribute form as a string", () => {
    const result = roundTrip(stringifyVmlShape, parseVmlShape, { path: "m0,0l50,0,0,50xe" });
    expect(result.path).toBe("m0,0l50,0,0,50xe");
  });

  it("emits a self-closing tag with no children", () => {
    expect(stringifyVmlShape({ id: "s1" })).toBe('<v:shape id="s1"/>');
  });
});

describe("v:shapetype", () => {
  it("round-trips geometry vocabulary", () => {
    const result = roundTrip(stringifyVmlShapetype, parseVmlShapetype, {
      id: "_x0000_t202",
      coordsize: "21600,21600",
      adj: "1",
      pathElement: { v: "m0,0l100,0,0,100xe", textpathok: true },
      formulas: { equations: ["val #0"] },
    });
    expect(result.id).toBe("_x0000_t202");
    expect(result.coordsize).toBe("21600,21600");
    expect(result.pathElement).toEqual({ v: "m0,0l100,0,0,100xe", textpathok: true });
    expect(result.formulas).toEqual({ equations: ["val #0"] });
  });
});

describe("v:group", () => {
  it("round-trips nested children in order", () => {
    const opts: VmlGroupOptions = {
      id: "g1",
      coordsize: "200,200",
      editas: "orgchart",
      fillcolor: "green",
      children: [
        { rect: { id: "r1", style: { width: 100, height: 40 } } },
        { line: { id: "l1", from: "0,0", to: "100,40" } },
        { group: { id: "g2", children: [{ oval: { id: "o1" } }] } },
      ],
    };
    const result = roundTrip(stringifyVmlGroup, parseVmlGroup, opts);
    expect(result).toEqual(opts);
  });
});

describe("v:background", () => {
  it("round-trips fill", () => {
    const result = roundTrip(stringifyVmlBackground, parseVmlBackground, {
      id: "bg",
      fillcolor: "white",
      fill: { type: "tile", src: "{bg.png}", relationshipId: "rId1" },
    });
    expect(result).toEqual({
      id: "bg",
      fillcolor: "white",
      fill: { type: "tile", src: "{bg.png}", relationshipId: "rId1" },
    });
  });
});

describe("basic shapes", () => {
  it("round-trips v:arc angles", () => {
    const opts: VmlArcOptions = { startAngle: 90, endAngle: 270 };
    expect(roundTrip(stringifyVmlArc, parseVmlArc, opts)).toEqual(opts);
  });

  it("round-trips v:curve points", () => {
    const opts: VmlCurveOptions = { from: "0,0", control1: "10,5", control2: "20,15", to: "30,20" };
    expect(roundTrip(stringifyVmlCurve, parseVmlCurve, opts)).toEqual(opts);
  });

  it("round-trips v:image crop attributes", () => {
    expect(
      roundTrip(stringifyVmlImage, parseVmlImage, { src: "img.png", grayscale: true, gain: "1.2" }),
    ).toEqual({ src: "img.png", grayscale: true, gain: "1.2" });
  });

  it("round-trips v:line endpoints", () => {
    const opts: VmlLineOptions = { from: "10,10", to: "110,60", stroke: { weight: "1pt" } };
    expect(roundTrip(stringifyVmlLine, parseVmlLine, opts)).toEqual(opts);
  });

  it("round-trips v:oval", () => {
    expect(roundTrip(stringifyVmlOval, parseVmlOval, { style: { width: 80, height: 80 } })).toEqual(
      { style: { width: 80, height: 80 } },
    );
  });

  it("round-trips v:polyline points", () => {
    const opts: VmlPolylineOptions = { points: "0,0 50,50 100,0" };
    expect(roundTrip(stringifyVmlPolyline, parseVmlPolyline, opts)).toEqual(opts);
  });

  it("round-trips v:rect", () => {
    expect(roundTrip(stringifyVmlRect, parseVmlRect, { filled: false })).toEqual({ filled: false });
  });

  it("round-trips v:roundrect arcsize", () => {
    const opts: VmlRoundRectOptions = { arcsize: "0.2" };
    expect(roundTrip(stringifyVmlRoundRect, parseVmlRoundRect, opts)).toEqual(opts);
  });
});

describe("shape elements", () => {
  it("round-trips v:fill", () => {
    const opts: VmlFillOptions = {
      id: "f1",
      type: "gradientRadial",
      on: true,
      color: "#ff0000",
      color2: "#0000ff",
      opacity: "0.5",
      angle: 90,
      focus: "50%",
      method: "none",
      alignshape: false,
      relationshipId: "rId2",
    };
    expect(roundTrip(stringifyVmlFill, parseVmlFill, opts)).toEqual(opts);
  });

  it("round-trips v:stroke", () => {
    const opts: VmlStrokeOptions = {
      weight: "2.5pt",
      color: "#333",
      linestyle: "thickThin",
      endcap: "round",
      dashstyle: "dot",
      startarrow: "diamond",
      startarrowwidth: "wide",
      endarrowlength: "long",
      insetpen: true,
    };
    expect(roundTrip(stringifyVmlStroke, parseVmlStroke, opts)).toEqual(opts);
  });

  it("round-trips v:shadow", () => {
    const opts: VmlShadowOptions = {
      on: true,
      type: "perspective",
      color: "#111",
      offset: "3pt,3pt",
      matrix: "1,0.5,-0.2,1,0,0",
    };
    expect(roundTrip(stringifyVmlShadow, parseVmlShadow, opts)).toEqual(opts);
  });

  it("round-trips v:textbox with w:txbxContent", () => {
    const opts: VmlTextboxOptions = { txbxContent: "<w:p><w:r><w:t>Hi</w:t></w:r></w:p>" };
    expect(roundTrip(stringifyVmlTextbox, parseVmlTextbox, opts)).toEqual(opts);
  });

  it("round-trips v:textbox verbatim content", () => {
    const opts: VmlTextboxOptions = { content: '<div align="left">Cell text</div>' };
    expect(roundTrip(stringifyVmlTextbox, parseVmlTextbox, opts)).toEqual(opts);
  });

  it("round-trips v:imagedata", () => {
    const opts: VmlImageDataOptions = {
      src: "pic.png",
      relationshipId: "rId7",
      pictRelationshipId: "rId8",
      croptop: "0.1",
      cropbottom: "-0.1",
      gamma: "0.8",
      bilevel: true,
    };
    expect(roundTrip(stringifyVmlImageData, parseVmlImageData, opts)).toEqual(opts);
  });

  it("round-trips v:path", () => {
    const opts: VmlPathOptions = {
      v: "m0,0l100,0,0,100xe",
      limo: "10800,10800",
      textboxrect: "5,5,95,95",
      fillok: false,
      textpathok: true,
    };
    expect(roundTrip(stringifyVmlPath, parseVmlPath, opts)).toEqual(opts);
  });

  it("round-trips v:textpath", () => {
    const opts: VmlTextPathOptions = {
      on: true,
      fitshape: true,
      string: "Bent text",
      style: { fontFamily: '"Arial Black"', fontSize: "24pt", vTextAlign: "center" },
    };
    expect(roundTrip(stringifyVmlTextPath, parseVmlTextPath, opts)).toEqual(opts);
  });

  it("round-trips v:formulas", () => {
    const opts: VmlFormulasOptions = { equations: ["val #0", "prod @0 2 1", "sum 21600 0 @1"] };
    expect(roundTrip(stringifyVmlFormulas, parseVmlFormulas, opts)).toEqual(opts);
  });

  it("round-trips v:handles", () => {
    const opts: VmlHandlesOptions = {
      handles: [{ position: "topLeft,#0", invx: true, xrange: "0 21600" }],
    };
    expect(roundTrip(stringifyVmlHandles, parseVmlHandles, opts)).toEqual(opts);
  });
});

describe("office shape elements (o:)", () => {
  it("round-trips o:skew", () => {
    const opts: VmlSkewOptions = {
      ext: "view",
      on: true,
      offset: "2pt,0",
      matrix: "1,0,0.5,1,0,0",
    };
    expect(roundTrip(stringifyVmlSkew, parseVmlSkew, opts)).toEqual(opts);
  });

  it("round-trips o:extrusion with 3D lighting", () => {
    const opts: VmlExtrusionOptions = {
      on: true,
      type: "perspective",
      render: "wireFrame",
      plane: "ZX",
      skewangle: 30,
      backdepth: "30pt",
      metal: true,
      lightposition: "22000,0,12000",
      lightharsh2: false,
    };
    expect(roundTrip(stringifyVmlExtrusion, parseVmlExtrusion, opts)).toEqual(opts);
  });

  it("round-trips o:callout", () => {
    const opts: VmlCalloutOptions = {
      on: true,
      angle: "45",
      dropauto: true,
      accentbar: true,
      minusx: false,
    };
    expect(roundTrip(stringifyVmlCallout, parseVmlCallout, opts)).toEqual(opts);
  });

  it("round-trips o:lock", () => {
    const opts: VmlLockOptions = { ext: "edit", rotation: true, text: true, aspectratio: false };
    expect(roundTrip(stringifyVmlLock, parseVmlLock, opts)).toEqual(opts);
  });

  it("round-trips o:clippath qualified v:v attribute", () => {
    expect(roundTrip(stringifyVmlClipPath, parseVmlClipPath, { v: "m0,0l100,0,0,100xe" })).toEqual({
      v: "m0,0l100,0,0,100xe",
    });
  });

  it("round-trips o:signatureline", () => {
    const opts: VmlSignatureLineOptions = {
      issignatureline: true,
      id: "{00000000-1111-2222-3333-444444444444}",
      suggestedsigner: "A. Signer",
      suggestedsigneremail: "a@example.com",
      showsigndate: true,
    };
    expect(roundTrip(stringifyVmlSignatureLine, parseVmlSignatureLine, opts)).toEqual(opts);
  });

  it("round-trips o:ink", () => {
    expect(roundTrip(stringifyVmlInk, parseVmlInk, { i: "isf-data", annotation: true })).toEqual({
      i: "isf-data",
      annotation: true,
    });
  });

  it("round-trips o:diagram with relationtable", () => {
    const opts: VmlDiagramOptions = {
      dgmstyle: 2,
      autolayout: true,
      relations: [{ idsrc: "1", iddest: "2", idcntr: "3" }],
    };
    expect(roundTrip(stringifyVmlDiagram, parseVmlDiagram, opts)).toEqual(opts);
  });

  it("round-trips o:OLEObject with field children", () => {
    const opts: VmlOleObjectOptions = {
      Type: "Embed",
      ProgID: "Excel.Sheet.12",
      ShapeID: "_x0000_s1026",
      DrawAspect: "Icon",
      ObjectID: "_1234567890",
      relationshipId: "rId1",
      linkType: "bmp",
      lockedField: true,
      fieldCodes: "EMBED Excel.Sheet.12 \\s",
    };
    expect(roundTrip(stringifyVmlOleObject, parseVmlOleObject, opts)).toEqual(opts);
  });

  it("round-trips o:shapedefaults with children", () => {
    const opts: VmlShapeDefaultsOptions = {
      ext: "edit",
      spidmax: 1026,
      fillcolor: "#3366ff",
      strokecolor: "silver",
      allowincell: false,
      fillElement: { type: "tile", src: "bg.png" },
      textbox: { inset: "1mm,1mm,1mm,1mm" },
      colormru: { colors: "#ff0000,#00ff00" },
    };
    expect(roundTrip(stringifyVmlShapeDefaults, parseVmlShapeDefaults, opts)).toEqual(opts);
  });

  it("round-trips o:shapelayout with idmap and regrouptable", () => {
    const opts: VmlShapeLayoutOptions = {
      ext: "edit",
      idmap: { ext: "edit", data: "1" },
      regrouptable: { entries: [{ new: 1, old: 2 }] },
    };
    expect(roundTrip(stringifyVmlShapeLayout, parseVmlShapeLayout, opts)).toEqual(opts);
  });
});

describe("office attributes on shapes", () => {
  it("round-trips o: attributes folded into v:shape", () => {
    const result = roundTrip(stringifyVmlShape, parseVmlShape, {
      id: "s1",
      spt: 202,
      oned: false,
      userdrawn: true,
      connectortype: "straight",
      bwmode: "grayScale",
      hr: true,
      hralign: "center",
      bordertopcolor: "red",
      insetmode: "custom",
    });
    expect(result.spt).toBe(202);
    expect(result.oned).toBe(false);
    expect(result.userdrawn).toBe(true);
    expect(result.connectortype).toBe("straight");
    expect(result.bwmode).toBe("grayScale");
    expect(result.hr).toBe(true);
    expect(result.hralign).toBe("center");
    expect(result.bordertopcolor).toBe("red");
    expect(result.insetmode).toBe("custom");
  });

  it("round-trips o: children on v:shape", () => {
    const result = roundTrip(stringifyVmlShape, parseVmlShape, {
      id: "s1",
      skew: { on: true, offset: "5pt,5pt" },
      lock: { rotation: true },
      clippath: { v: "m0,0l100,0,0,100xe" },
    });
    expect(result.skew).toEqual({ on: true, offset: "5pt,5pt" });
    expect(result.lock).toEqual({ rotation: true });
    expect(result.clippath).toEqual({ v: "m0,0l100,0,0,100xe" });
  });

  it("round-trips v:stroke with o: sub-strokes", () => {
    const result = roundTrip(stringifyVmlStroke, parseVmlStroke, {
      weight: "1pt",
      leftStroke: { on: true, weight: "4pt", color: "blue" },
      topStroke: { weight: "2pt", dashstyle: "dash" },
    });
    expect(result.leftStroke).toEqual({ on: true, weight: "4pt", color: "blue" });
    expect(result.topStroke).toEqual({ weight: "2pt", dashstyle: "dash" });
  });

  it("round-trips v:background with o: bw attributes", () => {
    const result = roundTrip(stringifyVmlBackground, parseVmlBackground, {
      bwmode: "auto",
      targetscreensize: "800,600",
    });
    expect(result.bwmode).toBe("auto");
    expect(result.targetscreensize).toBe("800,600");
  });

  it("round-trips v:group with table properties and diagram", () => {
    const result = roundTrip(stringifyVmlGroup, parseVmlGroup, {
      tableproperties: "1",
      tablelimits: "10 20 30",
      diagram: { dgmstyle: 3, reverse: true },
    });
    expect(result.tableproperties).toBe("1");
    expect(result.tablelimits).toBe("10 20 30");
    expect(result.diagram).toEqual({ dgmstyle: 3, reverse: true });
  });

  it("round-trips v:shapetype with o:master and o:complex", () => {
    const result = roundTrip(stringifyVmlShapetype, parseVmlShapetype, {
      id: "_x0000_t202",
      master: "#_x0000_t1",
      complex: { ext: "view" },
    });
    expect(result.master).toBe("#_x0000_t1");
    expect(result.complex).toEqual({ ext: "view" });
  });
});

describe("word elements (w10:)", () => {
  it("round-trips w10:wrap", () => {
    const opts = {
      type: "tight" as const,
      side: "left" as const,
      anchorx: "text" as const,
      anchory: "line" as const,
    };
    expect(roundTrip(stringifyVmlWrap, parseVmlWrap, opts)).toEqual(opts);
  });

  it("round-trips w10: borders", () => {
    const opts: VmlBorderOptions = { type: "thinThickLarge", width: 8, shadow: true };
    expect(roundTrip((o) => stringifyVmlBorder("w10:bordertop", o), parseVmlBorder, opts)).toEqual(
      opts,
    );
  });

  it("round-trips w10: children on v:shape", () => {
    const result = roundTrip(stringifyVmlShape, parseVmlShape, {
      wrap: { type: "square", anchorx: "margin" },
      anchorlock: {},
      bordertop: { type: "single", width: 4 },
    });
    expect(result.wrap).toEqual({ type: "square", anchorx: "margin" });
    expect(result.anchorlock).toEqual({});
    expect(result.bordertop).toEqual({ type: "single", width: 4 });
  });
});

describe("client data (x:)", () => {
  it("round-trips x:ClientData with a note anchor", () => {
    const opts: VmlClientDataOptions = {
      objectType: "Note",
      Anchor: "\n_comment1\n1\n15\n90\n1\n12\n65\n85",
      MoveWithCells: true,
      SizeWithCells: true,
      Visible: false,
      Row: 3,
      Column: 5,
    };
    expect(roundTrip(stringifyVmlClientData, parseVmlClientData, opts)).toEqual(opts);
  });

  it("round-trips x:ClientData with form-control fields", () => {
    const opts: VmlClientDataOptions = {
      objectType: "Checkbox",
      FmlaLink: "Sheet1!$A$1",
      Checked: 1,
      FirstButton: true,
      NoThreeD: false,
      Accel: 65,
      ScriptLanguage: 0,
    };
    expect(roundTrip(stringifyVmlClientData, parseVmlClientData, opts)).toEqual(opts);
  });

  it("round-trips x:ClientData on v:shape", () => {
    const result = roundTrip(stringifyVmlShape, parseVmlShape, {
      id: "s1",
      type: "#_x0000_t202",
      clientData: { objectType: "Note", Visible: false },
    });
    expect(result.clientData).toEqual({ objectType: "Note", Visible: false });
  });
});

describe("presentation elements (pvml:)", () => {
  it("round-trips pvml:textdata", () => {
    const opts: VmlTextDataOptions = { id: "rId9" };
    expect(roundTrip(stringifyVmlTextData, parseVmlTextData, opts)).toEqual(opts);
  });

  it("round-trips pvml:iscomment on v:shape", () => {
    const result = roundTrip(stringifyVmlShape, parseVmlShape, {
      id: "s1",
      iscomment: {},
      textdata: { id: "rId2" },
    });
    expect(result.iscomment).toEqual({});
    expect(result.textdata).toEqual({ id: "rId2" });
  });
});
