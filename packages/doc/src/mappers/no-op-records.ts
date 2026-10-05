import type { NumberPair } from "../streams/pieces";

export interface NoOpRecord {
  readonly recordType: number;
  readonly recordName: string;
  readonly reason: "absent-fib-table-pointer";
  readonly fixture: "empty-table-pointer";
}

export const NO_OP_FIB_TABLE_RECORDS: readonly NoOpRecord[] = [
  {
    recordType: 1,
    recordName: "fcStshf",
    reason: "absent-fib-table-pointer",
    fixture: "empty-table-pointer",
  },
  {
    recordType: 3,
    recordName: "fcPlcffndTxt",
    reason: "absent-fib-table-pointer",
    fixture: "empty-table-pointer",
  },
  {
    recordType: 4,
    recordName: "fcPlcfandRef",
    reason: "absent-fib-table-pointer",
    fixture: "empty-table-pointer",
  },
  {
    recordType: 6,
    recordName: "fcPlcfSed",
    reason: "absent-fib-table-pointer",
    fixture: "empty-table-pointer",
  },
  {
    recordType: 11,
    recordName: "fcPlcfHdd",
    reason: "absent-fib-table-pointer",
    fixture: "empty-table-pointer",
  },
  {
    recordType: 12,
    recordName: "fcPlcfbteChpx",
    reason: "absent-fib-table-pointer",
    fixture: "empty-table-pointer",
  },
  {
    recordType: 13,
    recordName: "fcPlcfbtePapx",
    reason: "absent-fib-table-pointer",
    fixture: "empty-table-pointer",
  },
  {
    recordType: 16,
    recordName: "fcPlcffldMom",
    reason: "absent-fib-table-pointer",
    fixture: "empty-table-pointer",
  },
  {
    recordType: 21,
    recordName: "fcSttbfBkmk",
    reason: "absent-fib-table-pointer",
    fixture: "empty-table-pointer",
  },
  {
    recordType: 22,
    recordName: "fcPlcfBkf",
    reason: "absent-fib-table-pointer",
    fixture: "empty-table-pointer",
  },
  {
    recordType: 23,
    recordName: "fcPlcfBkl",
    reason: "absent-fib-table-pointer",
    fixture: "empty-table-pointer",
  },
  {
    recordType: 33,
    recordName: "fcClx",
    reason: "absent-fib-table-pointer",
    fixture: "empty-table-pointer",
  },
  {
    recordType: 47,
    recordName: "fcPlcfendTxt",
    reason: "absent-fib-table-pointer",
    fixture: "empty-table-pointer",
  },
  {
    recordType: 50,
    recordName: "fcDggInfo",
    reason: "absent-fib-table-pointer",
    fixture: "empty-table-pointer",
  },
  {
    recordType: 56,
    recordName: "fcPlcftxbxTxt",
    reason: "absent-fib-table-pointer",
    fixture: "empty-table-pointer",
  },
  {
    recordType: 58,
    recordName: "fcPlcfHdrTxbxTxt",
    reason: "absent-fib-table-pointer",
    fixture: "empty-table-pointer",
  },
  {
    recordType: 73,
    recordName: "fcPlfLst",
    reason: "absent-fib-table-pointer",
    fixture: "empty-table-pointer",
  },
  {
    recordType: 74,
    recordName: "fcPlcflfo",
    reason: "absent-fib-table-pointer",
    fixture: "empty-table-pointer",
  },
];

export function noOpFibTableRecord(recordType: number, range: NumberPair): NoOpRecord | undefined {
  if (range.length !== 0) return undefined;
  return NO_OP_FIB_TABLE_RECORDS.find((record) => record.recordType === recordType);
}
