import { odfManifestMatrix, odtMissingEmbeddedObject } from "./odf";
import {
  docxPrinterSettings,
  docxRevisionIdentity,
  pptxIndefiniteTiming,
  xlsxPivotCalculatedItems,
} from "./ooxml";
import { pptNoTextShape } from "./ppt";
import { rtfBinaryToken, rtfListtable, rtfStylesheet } from "./rtf";
import { ProjectionError } from "./support";
import { xlsSupbookLinks } from "./xls";

export interface SyntheticDiagnostic {
  id: string;
  package: string;
  part: string;
  reason: string;
}

export interface SyntheticResult {
  diagnostics: SyntheticDiagnostic[];
  total: number;
  passed: number;
}

interface SyntheticCase {
  id: string;
  owner: string;
  part: string;
  verify: () => Promise<void>;
}

const CASES: readonly SyntheticCase[] = [
  {
    id: "docx-printer-settings",
    owner: "docx",
    part: "word/document.xml",
    verify: docxPrinterSettings,
  },
  {
    id: "docx-revision-identity",
    owner: "docx",
    part: "word/document.xml",
    verify: docxRevisionIdentity,
  },
  {
    id: "pptx-indefinite-timing",
    owner: "pptx",
    part: "ppt/slides/slide1.xml",
    verify: pptxIndefiniteTiming,
  },
  {
    id: "xlsx-pivot-calculated-items",
    owner: "xlsx",
    part: "xl/pivotCache/pivotCacheDefinition1.xml",
    verify: xlsxPivotCalculatedItems,
  },
  { id: "rtf-binary-token", owner: "rtf", part: "RTF objdata", verify: rtfBinaryToken },
  { id: "rtf-stylesheet", owner: "rtf", part: "RTF stylesheet", verify: rtfStylesheet },
  { id: "rtf-listtable", owner: "rtf", part: "RTF listtable", verify: rtfListtable },
  {
    id: "xls-supbook-dde-ole",
    owner: "xls",
    part: "Workbook globals SUPBOOK",
    verify: xlsSupbookLinks,
  },
  {
    id: "ppt-no-text-shape",
    owner: "ppt",
    part: "PowerPoint Document shape",
    verify: pptNoTextShape,
  },
  {
    id: "odt-missing-embedded-object",
    owner: "odt",
    part: "content.xml draw:object",
    verify: odtMissingEmbeddedObject,
  },
  {
    id: "odf-manifest-property-matrix",
    owner: "odf",
    part: "META-INF/manifest.xml",
    verify: odfManifestMatrix,
  },
];

export async function runSyntheticCorpus(): Promise<SyntheticResult> {
  const diagnostics: SyntheticDiagnostic[] = [];
  for (const testCase of CASES) {
    try {
      await testCase.verify();
      console.log(`  pass ${testCase.id}`);
    } catch (error) {
      const part = error instanceof ProjectionError ? error.part : testCase.part;
      const reason = error instanceof Error ? error.message : String(error);
      diagnostics.push({ id: testCase.id, package: testCase.owner, part, reason });
      console.log(`  FAIL ${testCase.id} [${testCase.owner}/${part}]: ${reason}`);
    }
  }
  return { diagnostics, total: CASES.length, passed: CASES.length - diagnostics.length };
}
