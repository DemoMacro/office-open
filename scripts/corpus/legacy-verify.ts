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

const CORPUS_PASSWORDS = ["VelvetSweatshop", "Password1234_", "password"];

async function verifyWithPasswords(
  format: "DOC" | "XLS" | "PPT",
  verify: (password?: string) => Promise<void> | void,
): Promise<void> {
  for (const password of [undefined, ...CORPUS_PASSWORDS]) {
    try {
      await verify(password);
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
  throw new Error(`Encrypted legacy ${format} input could not be decrypted`);
}

export async function verify(format: LegacyFormat, data: Uint8Array): Promise<void> {
  if (format === "doc") {
    await verifyWithPasswords("DOC", async (password) => {
      const options = parseLegacyDocument(data, password ? { password } : undefined);
      return parseDocument(await generateDocument(options));
    });
    return;
  }
  if (format === "xls") {
    await verifyWithPasswords("XLS", async (password) => {
      parseXlsxDocument(
        await generateXlsxWorkbook(parseLegacyWorkbook(data, password ? { password } : undefined)),
      );
    });
    return;
  }
  if (format === "ppt") {
    await verifyWithPasswords("PPT", async (password) => {
      const options = parseLegacyPresentation(data, password ? { password } : undefined);
      return parsePresentation(await generatePptxPresentation(options));
    });
    return;
  }
  if (format === "rtf") {
    parseDocument(await generateDocument(parseRtfDocument(Buffer.from(data).toString("latin1"))));
    return;
  }
  if (format === "odt") {
    await parseOdtDocument(await generateOdtDocument(await parseOdtDocument(data)));
    return;
  }
  if (format === "ods") {
    await parseOdsWorkbook(await generateOdsWorkbook(await parseOdsWorkbook(data)));
    return;
  }
  await parseOdpPresentation(await generateOdpPresentation(await parseOdpPresentation(data)));
}
