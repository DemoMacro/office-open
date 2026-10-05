import { DocParseError } from "../errors";
import type { NumberPair } from "../streams/pieces";
import { readUint16, readUint32, requireRange } from "./binary";
import { FLAG_ENCRYPTED, FIB_BASE_LENGTH, FIB_SIGNATURE } from "./constants";
import type { Fib } from "./models";
export function parseFib(word: Uint8Array): Fib {
  const streamMessage = "Invalid Word document: FIB is truncated";
  requireRange(word, 0, FIB_BASE_LENGTH, streamMessage);
  if (readUint16(word, 0, streamMessage) !== FIB_SIGNATURE) {
    throw new DocParseError("Invalid Word document: FIB signature is not 0xA5EC", {
      part: "stream",
      path: "WordDocument",
      recordName: "FIB",
      offset: 0,
      length: 2,
      reason: "unsupported-required-structure",
    });
  }
  const flags = readUint16(word, 10, streamMessage);
  if ((flags & FLAG_ENCRYPTED) !== 0) {
    throw new DocParseError("Encrypted Word documents are not supported");
  }

  const csw = readUint16(word, 32, streamMessage);
  const cslwOffset = FIB_BASE_LENGTH + 2 + csw * 2;
  const cslw = readUint16(word, cslwOffset, streamMessage);
  const characterLengthsOffset = cslwOffset + 2;
  const blobCountOffset = characterLengthsOffset + cslw * 4;
  const cbRgFcLcb = readUint16(word, blobCountOffset, streamMessage);
  const blobOffset = blobCountOffset + 2;
  const readPair = (index: number, name: string): NumberPair => {
    if (index + 1 > cbRgFcLcb) {
      return { offset: 0, length: 0 };
    }
    const pairOffset = blobOffset + index * 8;
    const message = `Invalid Word document: ${name} is outside the WordDocument stream`;
    return {
      offset: readUint32(word, pairOffset, message),
      length: readUint32(word, pairOffset + 4, message),
    };
  };

  if (cslw < 11) {
    throw new DocParseError("Invalid Word document: FIB text lengths are truncated");
  }
  const readLength = (index: number): number =>
    readUint32(word, characterLengthsOffset + index * 4, streamMessage);
  const ccpText = readLength(3);
  const ccpFootnotes = readLength(4);
  const ccpHeaders = readLength(5);
  const ccpComments = readLength(7);
  const ccpEndnotes = readLength(8);
  const ccpTextboxes = readLength(9);
  const ccpHeaderTextboxes = readLength(10);
  return {
    ccpText,
    ccpFootnotes,
    ccpHeaders,
    ccpComments,
    ccpEndnotes,
    ccpTextboxes,
    ccpHeaderTextboxes,
    totalCharacters:
      ccpText +
      ccpFootnotes +
      ccpHeaders +
      ccpComments +
      ccpEndnotes +
      ccpTextboxes +
      ccpHeaderTextboxes,
    nFib: readUint16(word, 2, streamMessage),
    characterBinTable: readPair(12, "fcPlcfbteChpx"),
    paragraphBinTable: readPair(13, "fcPlcfbtePapx"),
    footnoteTable: readPair(3, "fcPlcffndTxt"),
    headerTable: readPair(11, "fcPlcfHdd"),
    endnoteTable: readPair(47, "fcPlcfendTxt"),
    clx: readPair(33, "fcClx"),
    textboxTable: readPair(56, "fcPlcftxbxTxt"),
    headerTextboxTable: readPair(58, "fcPlcfHdrTxbxTxt"),
    styleSheet: readPair(1, "fcStshf"),
    sectionTable: readPair(6, "fcPlcfSed"),
    bookmarkNames: readPair(21, "fcSttbfBkmk"),
    bookmarkStarts: readPair(22, "fcPlcfBkf"),
    bookmarkEnds: readPair(23, "fcPlcfBkl"),
    fields: readPair(16, "fcPlcffldMom"),
    comments: readPair(4, "fcPlcfandRef"),
    drawing: readPair(50, "fcDggInfo"),
    list: readPair(73, "fcPlfLst"),
    listOverrides: readPair(74, "fcPlcflfo"),
  };
}
