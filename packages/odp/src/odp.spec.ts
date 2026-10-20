import type { FormControlOptions, OfficeFormsOptions } from "@office-open/core";
import { generateOcf, ODF_NAMESPACES, readOcf } from "@office-open/odf";
import { describe, expect, it } from "vite-plus/test";

import { generatePresentation, OdpParseError, parsePresentation } from "./index";
import type { OdpPresentationOptions } from "./options";

describe("ODP codec", () => {
  it("round-trips a presentation through canonical PresentationOptions", async () => {
    const parsed = await parsePresentation(
      await generatePresentation({ slides: [{ notes: "ODP" }] }),
    );
    expect(parsed.slides?.[0]?.notes).toBe("ODP");
  });

  it("round-trips slide names and master bindings", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:presentation><draw:page draw:name="Overview" draw:master-page-name="Master"/></office:presentation></office:body></office:document-content>`;
    const source = await parsePresentation(
      generateOcf("application/vnd.oasis.opendocument.presentation", { "content.xml": content }),
    );
    expect(source.slides?.[0]).toMatchObject({ name: "Overview", master: "Master" });
    expect((await parsePresentation(await generatePresentation(source))).slides?.[0]).toMatchObject(
      {
        name: "Overview",
        master: "Master",
      },
    );
  });

  it("round-trips slide comments with authors and anchors", async () => {
    const generated = await generatePresentation({
      slides: [
        {
          comments: [
            {
              author: "Reviewer",
              initials: "R",
              date: "2026-01-31T12:34:56",
              text: "First\nSecond",
              x: 360000,
              y: 720000,
              width: 720000,
              height: 360000,
            },
          ],
        },
      ],
    });
    const parsed = await parsePresentation(generated);
    expect(parsed.slides?.[0]?.comments).toEqual([
      {
        author: "Reviewer",
        initials: "R",
        date: "2026-01-31T12:34:56",
        text: "First\nSecond",
        x: 360000,
        y: 720000,
        width: 720000,
        height: 360000,
      },
    ]);
    const regenerated = await parsePresentation(await generatePresentation(parsed));
    expect(regenerated.slides?.[0]?.comments).toEqual(parsed.slides?.[0]?.comments);
  });

  it("round-trips embedded media with playback parameters", async () => {
    const video = new Uint8Array([1, 2, 3]);
    const audio = new Uint8Array([4, 5, 6]);
    const generated = await generatePresentation({
      slides: [
        {
          children: [
            {
              video: {
                type: "mp4",
                data: video,
                x: 360000,
                y: 720000,
                width: 1440000,
                height: 900000,
                playback: { loop: false, mute: false, volume: 0 },
              },
            },
            {
              audio: {
                type: "wav",
                data: audio,
                x: 720000,
                y: 1440000,
                width: 720000,
                height: 720000,
                playback: { loop: true, mute: true },
              },
            },
          ],
        },
      ],
    });
    const parsed = await parsePresentation(generated);
    const children = parsed.slides?.[0]?.children ?? [];
    expect(children[0]).toMatchObject({
      video: {
        type: "mp4",
        data: video,
        x: 360000,
        y: 720000,
        width: 1440000,
        height: 900000,
        playback: { loop: false, mute: false, volume: 0 },
      },
    });
    expect(children[1]).toMatchObject({
      audio: {
        type: "wav",
        data: audio,
        playback: { loop: true, mute: true },
      },
    });
    expect(await parsePresentation(await generatePresentation(parsed))).toEqual(parsed);
  });

  it("round-trips slide text fields", async () => {
    const generated = await generatePresentation({
      slides: [
        {
          children: [
            {
              shape: {
                x: 0,
                y: 0,
                width: 1800000,
                height: 360000,
                textBody: {
                  paragraphs: [
                    {
                      children: [
                        { type: "slidenum", text: "1", fixed: false, pageAdjust: 1 },
                        { type: "slidecount", text: "2" },
                        { type: "slidename", text: "Overview" },
                        { type: "author", text: "Ada", fixed: true },
                        { type: "filename", text: "Report", display: "full" },
                        {
                          type: "datetime",
                          text: "2026",
                          fixed: true,
                          dateValue: "2026-01-01",
                          dataStyleName: "D1",
                        },
                      ],
                    },
                  ],
                },
              },
            },
          ],
        },
      ],
    });
    const parsed = await parsePresentation(generated);
    const child = parsed.slides?.[0]?.children?.[0];
    if (!child || !("shape" in child)) throw new Error("text-field shape was not parsed");
    const paragraph = child.shape.textBody?.paragraphs?.[0];
    if (!paragraph) throw new Error("text-field paragraph was not parsed");
    if (typeof paragraph === "string") throw new Error("text fields collapsed to plain text");
    expect(paragraph.children).toEqual([
      { type: "slidenum", text: "1", fixed: false, pageAdjust: 1 },
      { type: "slidecount", text: "2" },
      { type: "slidename", text: "Overview" },
      { type: "author", text: "Ada", fixed: true },
      { type: "filename", text: "Report", display: "full" },
      {
        type: "datetime",
        text: "2026",
        fixed: true,
        dateValue: "2026-01-01",
        dataStyleName: "D1",
      },
    ]);
    expect(await parsePresentation(await generatePresentation(parsed))).toEqual(parsed);
  });

  it("requires an embedded object body but tolerates a missing replacement", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:presentation><draw:page><draw:frame svg:x="1cm" svg:y="1cm" svg:width="1cm" svg:height="1cm"><draw:object-ole xlink:href="./Object 1" xlink:type="simple"/></draw:frame></draw:page></office:presentation></office:body></office:document-content>`;
    const source = generateOcf("application/vnd.oasis.opendocument.presentation", {
      "content.xml": content,
      "Object 1": new Uint8Array([1]),
    });
    const parsed = await parsePresentation(source);
    expect(parsed.packageMembers).toEqual([
      { path: "Object 1", mediaType: "application/binary", data: new Uint8Array([1]) },
    ]);
    expect((await parsePresentation(await generatePresentation(parsed))).packageMembers).toEqual(
      parsed.packageMembers,
    );

    let error: unknown;
    try {
      await parsePresentation(
        generateOcf("application/vnd.oasis.opendocument.presentation", {
          "content.xml": content,
        }),
      );
    } catch (cause) {
      error = cause;
    }
    expect(error).toBeInstanceOf(OdpParseError);
    expect(error).toMatchObject({ reason: "embedded-object-subdocument-missing" });
  });

  it("preserves a standalone embedded object with its replacement image", async () => {
    const source = generateOcf(
      "application/vnd.oasis.opendocument.presentation",
      {
        "content.xml": `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:presentation><draw:page><draw:frame draw:style-name="gr1" draw:name="Formula" svg:x="1cm" svg:y="2cm" svg:width="4cm" svg:height="3cm"><draw:object xlink:href="./Object 1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/><draw:image xlink:href="./ObjectReplacements/Object 1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame></draw:page></office:presentation></office:body></office:document-content>`,
        "Object 1/content.xml": "<math/>",
        "ObjectReplacements/Object 1": new Uint8Array([1, 2, 3]),
      },
      {
        "Object 1/": "application/vnd.oasis.opendocument.formula",
        "ObjectReplacements/Object 1": "application/x-openoffice-gdimetafile",
      },
    );
    const parsed = await parsePresentation(source);
    expect(parsed.embeddedObjects).toEqual([
      {
        path: "Object 1",
        mediaType: "application/vnd.oasis.opendocument.formula",
        members: [
          { path: "Object 1/content.xml", data: "<math/>" },
          {
            path: "ObjectReplacements/Object 1",
            mediaType: "application/x-openoffice-gdimetafile",
            data: new Uint8Array([1, 2, 3]),
          },
        ],
        frames: [{ shapeId: 1, styleName: "gr1", replacementPath: "ObjectReplacements/Object 1" }],
      },
    ]);
    const regenerated = await parsePresentation(await generatePresentation(parsed));
    expect(regenerated.embeddedObjects).toEqual(parsed.embeddedObjects);
    expect(regenerated.slides?.[0]?.children?.[0]).toMatchObject({
      shape: { id: 1, name: "Formula", x: 360000, y: 720000, width: 1440000, height: 1080000 },
    });
  });

  it("round-trips a linked picture with a parent-relative URL", async () => {
    const options: OdpPresentationOptions = {
      slides: [
        {
          children: [
            {
              picture: {
                type: "png",
                sourceUrl: "../images/logo.png",
                x: 360000,
                y: 360000,
                width: 360000,
                height: 360000,
              },
            },
          ],
        },
      ],
    };
    expect(
      (await parsePresentation(await generatePresentation(options))).slides?.[0]?.children?.[0],
    ).toMatchObject({ picture: { sourceUrl: "../images/logo.png" } });
  });

  it("round-trips line breaks and hyperlink runs", async () => {
    const options: OdpPresentationOptions = {
      slides: [
        {
          children: [
            {
              shape: {
                textBody: {
                  paragraphs: [
                    { children: ["A", { break: true }, "B"] },
                    { children: [{ text: "Link", hyperlink: { url: "https://example.com/" } }] },
                  ],
                },
              },
            },
          ],
        },
      ],
    };
    const child = (await parsePresentation(await generatePresentation(options))).slides?.[0]
      ?.children?.[0];
    if (!child || !("shape" in child)) throw new TypeError("Expected a shape child");
    const paragraphs = (child.shape?.textBody?.paragraphs ?? []).filter(
      (paragraph) => typeof paragraph !== "string",
    );
    expect(paragraphs[0]?.children).toEqual(["A", { break: true }, "B"]);
    expect(paragraphs[1]?.children).toEqual([
      { text: "Link", hyperlink: { url: "https://example.com/" } },
    ]);
  });

  it("restores percent and length semantics in style overlays", async () => {
    const options: OdpPresentationOptions = {
      slides: [{ notes: "Overlay" }],
      styleOverlays: [
        {
          name: "P1",
          family: "paragraph",
          properties: [
            {
              name: "style:paragraph-properties",
              attributes: { "fo:line-height": 150, "fo:margin-left": 360000 },
            },
          ],
        },
      ],
    };
    const generated = await generatePresentation(options);
    const first = (await readOcf(generated, "application/vnd.oasis.opendocument.presentation"))
      .files["content.xml"]!;
    expect(first).toContain('fo:line-height="150%"');
    expect(first).toContain('fo:margin-left="1cm"');

    const parsed = await parsePresentation(generated);
    const second = (
      await readOcf(
        await generatePresentation(parsed),
        "application/vnd.oasis.opendocument.presentation",
      )
    ).files["content.xml"]!;
    expect(second).toContain('fo:line-height="150%"');
    expect(second).toContain('fo:margin-left="1cm"');
  });

  it("wraps invalid packages in OdpParseError", async () => {
    await expect(parsePresentation(new Uint8Array([1, 2, 3]))).rejects.toThrow(OdpParseError);
  });

  it("round-trips slide form containers and controls", async () => {
    const controls: FormControlOptions[] = [
      { control: "text", id: "text", xmlId: "text", value: "", currentValue: "", maxLength: 0 },
      { control: "textArea", id: "textArea", xmlId: "textArea", paragraphs: ["Line"] },
      {
        control: "formattedText",
        id: "formattedText",
        xmlId: "formattedText",
        minValue: "0",
        maxValue: "9",
        validation: false,
      },
      {
        control: "number",
        id: "number",
        xmlId: "number",
        value: 0,
        currentValue: 1,
        minValue: 0,
        maxValue: 9,
      },
      {
        control: "date",
        id: "date",
        xmlId: "date",
        value: "2026-01-31",
        currentValue: "2026-02-01",
      },
      { control: "time", id: "time", xmlId: "time", value: "01:02:03" },
      { control: "comboBox", id: "comboBox", xmlId: "comboBox", items: [{ label: "" }] },
      {
        control: "listBox",
        id: "listBox",
        xmlId: "listBox",
        options: [{ label: "", currentSelected: false, selected: false, value: "" }],
      },
      {
        control: "checkbox",
        id: "checkbox",
        xmlId: "checkbox",
        currentState: "unknown",
        state: "unknown",
        triState: true,
      },
      { control: "password", id: "password", xmlId: "password", echoChar: "*" },
      { control: "file", id: "file", xmlId: "file", currentValue: "a.odp" },
      { control: "fixedText", id: "fixedText", xmlId: "fixedText", label: "", multiLine: true },
      {
        control: "button",
        id: "button",
        xmlId: "button",
        buttonType: "push",
        defaultButton: false,
        toggle: false,
        focusOnClick: false,
      },
      { control: "image", id: "image", xmlId: "image", buttonType: "url", href: "logo.png" },
      { control: "radio", id: "radio", xmlId: "radio", currentSelected: false, selected: false },
      { control: "frame", id: "frame", xmlId: "frame", label: "Frame" },
      { control: "imageFrame", id: "imageFrame", xmlId: "imageFrame", imageData: "logo.png" },
      { control: "hidden", id: "hidden", xmlId: "hidden", value: "" },
      {
        control: "grid",
        id: "grid",
        xmlId: "grid",
        columns: [{ name: "title", controls: [{ control: "text", id: "title", xmlId: "title" }] }],
      },
      {
        control: "valueRange",
        id: "valueRange",
        xmlId: "valueRange",
        minValue: 0,
        maxValue: 100,
        stepSize: 1,
        pageStepSize: 10,
      },
      { control: "genericControl", id: "genericControl", xmlId: "genericControl" },
    ];
    const forms: OfficeFormsOptions = {
      automaticFocus: true,
      designMode: true,
      containers: [
        {
          name: "Responses",
          controls,
        },
      ],
    };
    const parsed = await parsePresentation(await generatePresentation({ slides: [{ forms }] }));
    expect(parsed.slides?.[0]?.forms).toEqual(forms);
  });

  it("maps primitive ODF shapes to canonical pptx shapes", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:presentation><draw:page><draw:rect svg:x="1cm" svg:y="1cm" svg:width="2cm" svg:height="1cm" draw:name="Box"><text:p>Box</text:p></draw:rect><draw:path svg:x="4cm" svg:y="1cm" svg:width="2cm" svg:height="1cm" svg:d="M 0 0 L 2 0 L 2 1 Z"/><draw:g><draw:custom-shape svg:x="1cm" svg:y="3cm" svg:width="1cm" svg:height="1cm" draw:name="Grouped"/></draw:g></draw:page></office:presentation></office:body></office:document-content>`;
    const parsed = await parsePresentation(
      generateOcf("application/vnd.oasis.opendocument.presentation", { "content.xml": content }),
    );
    const children = parsed.slides?.[0]?.children ?? [];
    expect(children[0]).toMatchObject({
      shape: {
        name: "Box",
        x: 360000,
        y: 360000,
        width: 720000,
        height: 360000,
        properties: { geometry: "rect" },
      },
    });
    expect(children[1]).toMatchObject({
      shape: {
        properties: {
          customGeometry: {
            pathList: [
              {
                commands: [
                  { command: "moveTo", point: { x: "0", y: "0" } },
                  { command: "lineTo", point: { x: "720000", y: "0" } },
                  { command: "lineTo", point: { x: "720000", y: "360000" } },
                  { command: "close" },
                ],
              },
            ],
          },
        },
      },
    });
    expect(children[2]).toMatchObject({
      group: { children: [{ shape: { name: "Grouped" } }] },
    });
    const regenerated = await parsePresentation(await generatePresentation(parsed));
    expect(regenerated.slides?.[0]?.children).toEqual(parsed.slides?.[0]?.children);
  });

  it("round-trips ODF form-control shapes with typed references", async () => {
    const parsed = await parsePresentation(
      await generatePresentation({
        slides: [
          {
            children: [
              {
                shape: {
                  odfControlId: "control1",
                  name: "Control",
                  x: 360000,
                  y: 720000,
                  width: 720000,
                  height: 360000,
                  properties: { geometry: "rect" },
                },
              },
            ],
          },
        ],
      }),
    );
    expect(parsed.slides?.[0]?.children?.[0]).toMatchObject({
      shape: { odfControlId: "control1", properties: { geometry: "rect" } },
    });
    const regenerated = await parsePresentation(await generatePresentation(parsed));
    expect(regenerated.slides?.[0]?.children).toEqual(parsed.slides?.[0]?.children);
  });

  it("rejects unknown form controls with structured errors", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES} xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"><office:body><office:presentation><draw:page><office:forms><form:form><form:unknown/></form:form></office:forms></draw:page></office:presentation></office:body></office:document-content>`;
    let error: unknown;
    try {
      await parsePresentation(
        generateOcf("application/vnd.oasis.opendocument.presentation", {
          "content.xml": content,
        }),
      );
    } catch (cause) {
      error = cause;
    }
    expect(error).toBeInstanceOf(OdpParseError);
    expect(error).toMatchObject({ part: "content.xml", name: "form:unknown" });
  });

  it("round-trips slide charts with position, size, semantics, and order", async () => {
    const parsed = await parsePresentation(
      await generatePresentation({
        slides: [
          {
            children: [
              {
                chart: {
                  name: "First",
                  type: "column",
                  x: 360000,
                  y: 720000,
                  width: 2160000,
                  height: 1440000,
                  series: [{ name: "Sales", values: [1, 2] }],
                },
              },
              {
                chart: {
                  name: "Second",
                  type: "bar",
                  x: 1080000,
                  y: 1440000,
                  width: 1800000,
                  height: 1080000,
                  series: [{ name: "Margin", values: [2, 4] }],
                },
              },
            ],
          },
        ],
      }),
    );
    const children = parsed.slides?.[0]?.children ?? [];
    expect(children.map((child) => ("chart" in child ? child.chart.name : undefined))).toEqual([
      "First",
      "Second",
    ]);
    expect(children[0]).toMatchObject({
      chart: {
        name: "First",
        type: "column",
        x: 360000,
        y: 720000,
        width: 2160000,
        height: 1440000,
        series: [{ name: "Sales", values: [1, 2] }],
      },
    });
    expect(children[1]).toMatchObject({ chart: { type: "bar" } });
  });

  it("reports unknown slide elements with structured errors", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES} xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"><office:body><office:presentation><draw:page><draw:unknown/></draw:page></office:presentation></office:body></office:document-content>`;
    let error: unknown;
    try {
      await parsePresentation(
        generateOcf("application/vnd.oasis.opendocument.presentation", {
          "content.xml": content,
        }),
      );
    } catch (cause) {
      error = cause;
    }
    expect(error).toBeInstanceOf(OdpParseError);
    expect(error).toMatchObject({
      part: "content.xml",
      path: "/office:document-content/office:body/office:presentation/draw:page/draw:unknown",
      name: "draw:unknown",
      reason: "element has no canonical PresentationOptions mapping",
    });
  });
});
