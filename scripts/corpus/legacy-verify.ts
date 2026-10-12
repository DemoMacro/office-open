import { parseDocument as parseLegacyDocument } from "../../packages/doc/dist/index.mjs";
import { generateDocument, parseDocument } from "../../packages/docx/dist/index.mjs";
import {
  generatePresentation as generateOdpPresentation,
  parsePresentation as parseOdpPresentation,
} from "../../packages/odp/dist/index.mjs";
import {
  generateWorkbook as generateOdsWorkbook,
  parseWorkbook as parseOdsWorkbook,
} from "../../packages/ods/dist/index.mjs";
import {
  generateDocument as generateOdtDocument,
  parseDocument as parseOdtDocument,
} from "../../packages/odt/dist/index.mjs";
import { parsePresentation as parseLegacyPresentation } from "../../packages/ppt/dist/index.mjs";
import {
  generatePresentation as generatePptxPresentation,
  parsePresentation,
} from "../../packages/pptx/dist/index.mjs";
import { parseDocument as parseRtfDocument } from "../../packages/rtf/dist/index.mjs";
import { parseWorkbook as parseLegacyWorkbook } from "../../packages/xls/dist/index.mjs";
import { generateWorkbook as generateXlsxWorkbook } from "../../packages/xlsx/dist/index.mjs";
import { parseWorkbook as parseXlsxDocument } from "../../packages/xlsx/dist/index.mjs";

export type LegacyFormat = "doc" | "xls" | "ppt" | "rtf" | "odt" | "ods" | "odp";

const CORPUS_PASSWORDS = ["tika", "VelvetSweatshop", "Password1234_", "password"];

export async function verify(format: LegacyFormat, data: Uint8Array): Promise<void> {
  if (format === "doc") {
    for (const password of [undefined, ...CORPUS_PASSWORDS]) {
      try {
        const options = parseLegacyDocument(data, password ? { password } : undefined);
        parseDocument(await generateDocument(options));
        return;
      } catch (error) {
        if (
          password === undefined &&
          !/encrypted|password/i.test(String((error as Error)?.message))
        ) {
          throw error;
        }
      }
    }
    throw new Error("Encrypted legacy DOC input could not be decrypted");
  }
  if (format === "xls") {
    for (const password of [undefined, ...CORPUS_PASSWORDS]) {
      try {
        parseXlsxDocument(
          await generateXlsxWorkbook(
            parseLegacyWorkbook(data, password ? { password } : undefined),
          ),
        );
        return;
      } catch (error) {
        if (
          password === undefined &&
          !/encrypted|password/i.test(String((error as Error)?.message))
        ) {
          throw error;
        }
      }
    }
    throw new Error("Encrypted legacy XLS input could not be decrypted");
  }
  if (format === "ppt") {
    for (const password of [undefined, ...CORPUS_PASSWORDS]) {
      try {
        const options = parseLegacyPresentation(data, password ? { password } : undefined);
        parsePresentation(await generatePptxPresentation(options));
        return;
      } catch (error) {
        if (
          password === undefined &&
          !/encrypted|password/i.test(String((error as Error)?.message))
        ) {
          throw error;
        }
      }
    }
    throw new Error("Encrypted legacy PPT input could not be decrypted");
  }
  if (format === "rtf") {
    parseDocument(await generateDocument(parseRtfDocument(Buffer.from(data).toString("latin1"))));
    return;
  }
  if (format === "odt") {
    parseOdtDocument(generateOdtDocument(parseOdtDocument(data)));
    return;
  }
  if (format === "ods") {
    parseOdsWorkbook(generateOdsWorkbook(parseOdsWorkbook(data)));
    return;
  }
  parseOdpPresentation(generateOdpPresentation(parseOdpPresentation(data)));
}
