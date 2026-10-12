import type { ModelRegistryEntry } from "./types";

function mathExecutableEvidence(localName: string): readonly string[] {
  if (localName === "mathPr") {
    return ["math.settings.round-trip", "math.settings.wrap-right"];
  }
  if (localName === "oMathPara" || localName === "oMathParaPr" || localName === "jc") {
    return ["math.inline.structures", "math.paragraph.round-trip"];
  }
  if (
    localName === "r" ||
    localName === "rPr" ||
    localName === "t" ||
    localName === "lit" ||
    localName === "nor" ||
    localName === "scr" ||
    localName === "sty" ||
    localName === "brk" ||
    localName === "aln"
  ) {
    return ["math.inline.structures", "math.run.normal"];
  }
  return ["math.inline.structures"];
}

const MATH_QNAMES = [
  "acc",
  "accPr",
  "aln",
  "alnScr",
  "argPr",
  "argSz",
  "bar",
  "barPr",
  "baseJc",
  "begChr",
  "borderBox",
  "borderBoxPr",
  "box",
  "boxPr",
  "brk",
  "brkBin",
  "brkBinSub",
  "cGp",
  "cGpRule",
  "chr",
  "count",
  "cSp",
  "ctrlPr",
  "d",
  "defJc",
  "deg",
  "degHide",
  "den",
  "diff",
  "dispDef",
  "dPr",
  "e",
  "endChr",
  "eqArr",
  "eqArrPr",
  "f",
  "fName",
  "fPr",
  "func",
  "funcPr",
  "groupChr",
  "groupChrPr",
  "grow",
  "hideBot",
  "hideLeft",
  "hideRight",
  "hideTop",
  "interSp",
  "intLim",
  "intraSp",
  "jc",
  "lim",
  "limLoc",
  "limLow",
  "limLowPr",
  "limUpp",
  "limUppPr",
  "lit",
  "lMargin",
  "m",
  "mathFont",
  "mathPr",
  "maxDist",
  "mc",
  "mcJc",
  "mcPr",
  "mcs",
  "mPr",
  "mr",
  "nary",
  "naryLim",
  "naryPr",
  "noBreak",
  "nor",
  "num",
  "objDist",
  "oMath",
  "oMathPara",
  "oMathParaPr",
  "opEmu",
  "phant",
  "phantPr",
  "plcHide",
  "pos",
  "postSp",
  "preSp",
  "r",
  "rad",
  "radPr",
  "rMargin",
  "rPr",
  "rSp",
  "rSpRule",
  "scr",
  "sepChr",
  "show",
  "shp",
  "smallFrac",
  "sPre",
  "sPrePr",
  "sSub",
  "sSubPr",
  "sSubSup",
  "sSubSupPr",
  "sSup",
  "sSupPr",
  "strikeBLTR",
  "strikeH",
  "strikeTLBR",
  "strikeV",
  "sty",
  "sub",
  "subHide",
  "sup",
  "supHide",
  "t",
  "transp",
  "type",
  "vertJc",
  "wrapIndent",
  "wrapRight",
  "zeroAsc",
  "zeroDesc",
  "zeroWid",
] as const;

export const MATH_REGISTRY: readonly ModelRegistryEntry[] = MATH_QNAMES.map((localName) => ({
  domain: "math",
  qname: `m:${localName}`,
  canonicalOptionPath: localName === "mathPr" ? "MathPropertiesOptions" : "MathInput",
  classification: "modeled",
  classificationReason: "Canonical typed parse and stringify both own this OMML surface.",
  owner: {
    stringify:
      localName === "mathPr"
        ? "packages/docx/src/parts/settings/descriptor.ts:settingsDesc"
        : "packages/docx/src/parts/paragraph/math/stringify.ts:stringifyMathInput",
    parse:
      localName === "mathPr"
        ? "packages/docx/src/parts/settings/descriptor.ts:settingsDesc"
        : "packages/docx/src/parts/paragraph/math/stringify.ts:parseMathChildren",
    compiler: localName === "mathPr" ? "DOCX settings compiler" : "DOCX paragraph compiler",
    publicApi: localName === "mathPr" ? "MathPropertiesOptions" : "MathInput",
  },
  evidence: {
    static:
      localName === "mathPr"
        ? ["packages/docx/src/parts/settings/descriptor.ts"]
        : ["packages/docx/src/parts/paragraph/math/stringify.ts"],
    executable: mathExecutableEvidence(localName),
    rawOnly: false,
  },
}));
