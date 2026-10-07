import { unzipSync, zipSync } from "@office-open/core";
import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { generateWorkbook } from "../generate";
import { parseWorkbookSync } from "../parse";
import {
  classificationLabelsDesc,
  type ClassificationLabelListOptions,
} from "./classification-labels";

const writeCtx = {} as unknown as WriteContext;
const readCtx = {} as unknown as ReadContext;

const options: ClassificationLabelListOptions = {
  labels: [
    {
      id: "{00000000-0000-0000-0000-000000000001}",
      enabled: true,
      method: "Standard",
      name: "Internal",
      siteId: "{00000000-0000-0000-0000-000000000002}",
      contentBits: 0,
      removed: false,
    },
  ],
  extensions: [
    {
      uri: "https://example.com/extension",
      innerXml: '<test:value xmlns:test="https://example.com/schema">kept</test:value>',
    },
  ],
};

describe("classificationLabelsDesc", () => {
  it("preserves an empty label list", () => {
    const xml = classificationLabelsDesc.stringify({ labels: [] }, writeCtx);
    const root = parseXml(xml!, { nativeTypeAttributes: true }).elements?.[0];
    expect(root).toBeDefined();
    expect(classificationLabelsDesc.parse(root!, readCtx)).toEqual({ labels: [] });
  });

  it("rejects unknown label-list children", () => {
    const root = parseXml(
      '<clbl:labelList xmlns:clbl="http://schemas.microsoft.com/office/2020/mipLabelMetadata"><clbl:unknown/></clbl:labelList>',
      { nativeTypeAttributes: true },
    ).elements?.[0];
    expect(() => classificationLabelsDesc.parse(root!, readCtx)).toThrowError(
      /unexpected child of labelList/,
    );
  });

  it("round-trips labels and extension payloads", () => {
    const xml = classificationLabelsDesc.stringify(options, writeCtx);
    const root = parseXml(xml!, { nativeTypeAttributes: true }).elements?.[0];
    expect(root).toBeDefined();
    expect(classificationLabelsDesc.parse(root!, readCtx)).toEqual(options);
  });

  it("emits the typed part, relationship, and content type", async () => {
    const output = (await generateWorkbook(
      { worksheets: [{ name: "Data" }], classificationLabels: options },
      { type: "uint8array" },
    )) as Uint8Array;
    const files = unzipSync(output);
    const xml = new TextDecoder().decode(files["docMetadata/LabelInfo.xml"]!);
    const rels = new TextDecoder().decode(files["_rels/.rels"]!);
    const types = new TextDecoder().decode(files["[Content_Types].xml"]!);

    expect(xml).toContain("clbl:labelList");
    expect(rels).toContain("relationships/classificationlabels");
    expect(types).toContain("application/vnd.ms-office.classificationlabels+xml");

    const parsed = parseWorkbookSync(zipSync(files));
    expect(parsed.classificationLabels).toEqual(options);
    expect(parsed.classificationLabelsPath).toBe("docMetadata/LabelInfo.xml");
    expect(parsed.rawParts).toBeUndefined();
  });
});
