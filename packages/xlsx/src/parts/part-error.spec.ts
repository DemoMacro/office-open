import { describe, expect, it } from "vite-plus/test";

import { ActiveXControlParseError } from "./active-x-control";
import { ClassificationLabelsParseError } from "./classification-labels";
import { ControlPropertiesParseError } from "./control-properties";
import { PersonsParseError } from "./persons";
import { RichDataParseError } from "./rich-data";

describe("XLSX part parse errors", () => {
  it("keeps Error.name separate from the rejected XML element name", () => {
    const errors = [
      new ActiveXControlParseError("xl/activeX.xml", "/ax:ocx", "ax:ocx", "invalid"),
      new ClassificationLabelsParseError("xl/metadata.xml", "/labels", "label", "invalid"),
      new ControlPropertiesParseError("xl/ctrlProps.xml", "/f", "f", "invalid"),
      new PersonsParseError("xl/persons.xml", "/persons", "person", "invalid"),
      new RichDataParseError("xl/richData.xml", "/rvb", "ext", "invalid"),
    ];

    expect(errors.map((error) => error.name)).toEqual([
      "ActiveXControlParseError",
      "ClassificationLabelsParseError",
      "ControlPropertiesParseError",
      "PersonsParseError",
      "RichDataParseError",
    ]);
    expect(errors.every((error) => error.part && error.path && error.reason)).toBe(true);
  });
});
