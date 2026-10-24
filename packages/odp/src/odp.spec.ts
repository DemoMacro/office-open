import type { FormControlOptions, OfficeFormsOptions } from "@office-open/core";
import { generateOcf, ODF_NAMESPACES, readOcf } from "@office-open/odf";
import type { TransitionOptions } from "@office-open/pptx";
import { describe, expect, it } from "vite-plus/test";

import { generatePresentation, OdpParseError, parsePresentation } from "./index";
import type { OdpPresentationOptions } from "./options";
import type { OdpTransitionOptions } from "./transition";

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

  it("round-trips a typed slide transition", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:presentation><draw:page><anim:par><anim:transitionFilter smil:type="fade" smil:subtype="crossfade" smil:dur="1s"/></anim:par></draw:page></office:presentation></office:body></office:document-content>`;
    const parsed = await parsePresentation(
      generateOcf("application/vnd.oasis.opendocument.presentation", { "content.xml": content }),
    );
    expect(parsed.slides?.[0]?.transition).toEqual({ type: "fade", duration: 1000 });
    const generated = await generatePresentation(parsed);
    const regenerated = await parsePresentation(generated);
    expect(regenerated.slides?.[0]?.transition).toEqual(parsed.slides?.[0]?.transition);
  });

  it("round-trips ODF-only transition attributes", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:presentation><draw:page><anim:par><anim:transitionFilter smil:type="fade" smil:subtype="crossfade" smil:dur="500ms" smil:mode="out" smil:fadeColor="#102030"/></anim:par></draw:page></office:presentation></office:body></office:document-content>`;
    const parsed = await parsePresentation(
      generateOcf("application/vnd.oasis.opendocument.presentation", { "content.xml": content }),
    );
    expect(parsed.slides?.[0]?.transition).toEqual({
      type: "fade",
      duration: 500,
      mode: "out",
      fadeColor: "#102030",
    });
    const indefiniteContent = content.replace('smil:dur="500ms"', 'smil:dur="indefinite"');
    const indefinite = await parsePresentation(
      generateOcf("application/vnd.oasis.opendocument.presentation", {
        "content.xml": indefiniteContent,
      }),
    );
    expect(indefinite.slides?.[0]?.transition).toEqual({
      type: "fade",
      filterDuration: "indefinite",
      mode: "out",
      fadeColor: "#102030",
    });
    const regenerated = await parsePresentation(await generatePresentation(parsed));
    expect(regenerated.slides?.[0]?.transition).toEqual(parsed.slides?.[0]?.transition);
  });

  it("round-trips transition speed, advance timing, and sound through page styles", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:automatic-styles><style:style style:name="DP1" style:family="drawing-page"><style:drawing-page-properties presentation:transition-type="automatic" presentation:transition-speed="slow" presentation:duration="PT2S"><presentation:sound xlink:href="https://office-open.test/media/sound.wav" xlink:type="simple" presentation:play-full="true"/></style:drawing-page-properties></style:style></office:automatic-styles><office:body><office:presentation><draw:page draw:style-name="DP1"><anim:par><anim:transitionFilter smil:type="fade"/></anim:par></draw:page></office:presentation></office:body></office:document-content>`;
    const parsed = await parsePresentation(
      generateOcf("application/vnd.oasis.opendocument.presentation", { "content.xml": content }),
    );
    expect(parsed.slides?.[0]?.transition).toEqual({
      type: "fade",
      speed: "slow",
      advanceOnClick: false,
      advanceAfterTime: 2000,
      soundHref: "https://office-open.test/media/sound.wav",
      soundPlayFull: true,
    });
    const generated = await generatePresentation(parsed);
    const styles = (await readOcf(generated, "application/vnd.oasis.opendocument.presentation"))
      .files["content.xml"]!;
    expect(styles).toContain('presentation:transition-type="automatic"');
    expect(styles).toContain('presentation:transition-speed="slow"');
    expect(styles).toContain('presentation:duration="PT2S"');
    expect(styles).toContain('xlink:href="https://office-open.test/media/sound.wav"');
    const regenerated = await parsePresentation(generated);
    expect(regenerated.slides?.[0]?.transition).toEqual(parsed.slides?.[0]?.transition);
  });

  it("keeps transition filters when the same page also has animations", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:presentation><draw:page><anim:par><anim:transitionFilter smil:type="fade"/></anim:par><anim:par><anim:set smil:attributeName="visibility" smil:to="visible"/></anim:par></draw:page></office:presentation></office:body></office:document-content>`;
    const parsed = await parsePresentation(
      generateOcf("application/vnd.oasis.opendocument.presentation", { "content.xml": content }),
    );
    expect(parsed.slides?.[0]?.transition).toEqual({ type: "fade" });
    expect(parsed.slides?.[0]?.animations).toHaveLength(2);
    const regenerated = await parsePresentation(await generatePresentation(parsed));
    expect(regenerated.slides?.[0]?.transition).toEqual(parsed.slides?.[0]?.transition);
    expect(regenerated.slides?.[0]?.animations).toEqual(parsed.slides?.[0]?.animations);
  });

  it("reuses transition style names when source styles already occupy them", async () => {
    const generated = await generatePresentation({
      styleResources: [
        {
          name: "ODPTransition1",
          family: "drawing-page",
          properties: [{ name: "style:drawing-page-properties", attributes: {} }],
        },
      ],
      slides: [{ transition: { type: "fade", advanceOnClick: false } }],
    });
    const content = (await readOcf(generated, "application/vnd.oasis.opendocument.presentation"))
      .files["content.xml"]!;
    expect(content.match(/style:name="ODPTransition1"/g)).toHaveLength(1);
    expect(content).toContain('style:name="ODPTransition2"');
  });

  it("allocates unique transition style names across multiple generated slides", async () => {
    const generated = await generatePresentation({
      styleResources: [
        {
          name: "ODPTransition1",
          family: "drawing-page",
          properties: [{ name: "style:drawing-page-properties", attributes: {} }],
        },
      ],
      slides: [
        { transition: { type: "fade", advanceOnClick: false } },
        { transition: { type: "wipe", advanceOnClick: false } },
      ],
    });
    const content = (await readOcf(generated, "application/vnd.oasis.opendocument.presentation"))
      .files["content.xml"]!;
    expect(content.match(/style:name="ODPTransition2"/g)).toHaveLength(1);
    expect(content).toContain('style:name="ODPTransition3"');
  });

  it("materializes an embedded transition sound exactly once", async () => {
    const generated = await generatePresentation({
      slides: [
        {
          transition: {
            type: "fade",
            startSound: { type: "wav", data: new Uint8Array([1, 2, 3]) },
          },
        },
      ],
    });
    const archive = await readOcf(generated, "application/vnd.oasis.opendocument.presentation");
    expect(archive.binaries["Media/transition1.wav"]).toEqual(new Uint8Array([1, 2, 3]));
    const styles = archive.files["content.xml"]!;
    expect(styles).toContain('xlink:href="Media/transition1.wav"');
    const parsed = await parsePresentation(generated);
    expect(parsed.slides?.[0]?.transition).toMatchObject({
      type: "fade",
      soundHref: "Media/transition1.wav",
    });
    expect(parsed.slides?.[0]?.transition).not.toHaveProperty("startSound");
  });

  it("rejects an embedded transition sound missing from the package", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:automatic-styles><style:style style:name="DP1" style:family="drawing-page"><style:drawing-page-properties><presentation:sound xlink:href="Media/missing.wav"/></style:drawing-page-properties></style:style></office:automatic-styles><office:body><office:presentation><draw:page draw:style-name="DP1"/></office:presentation></office:body></office:document-content>`;
    await expect(
      parsePresentation(
        generateOcf("application/vnd.oasis.opendocument.presentation", { "content.xml": content }),
      ),
    ).rejects.toThrow(OdpParseError);
  });

  it("keeps generated transition sounds apart from package members", async () => {
    const member = new Uint8Array([9, 9, 9]);
    const generated = await generatePresentation({
      packageMembers: [{ path: "Media/transition1.wav", mediaType: "audio/wav", data: member }],
      slides: [
        {
          transition: {
            type: "fade",
            startSound: { type: "wav", data: new Uint8Array([1, 2, 3]) },
          },
        },
      ],
    });
    const archive = await readOcf(generated, "application/vnd.oasis.opendocument.presentation");
    expect(archive.binaries["Media/transition1.wav"]).toEqual(member);
    expect(archive.binaries["Media/transition2.wav"]).toEqual(new Uint8Array([1, 2, 3]));
    expect(archive.files["content.xml"]).toContain('xlink:href="Media/transition2.wav"');
  });

  it("allocates distinct transition sounds across slides", async () => {
    const generated = await generatePresentation({
      slides: [
        {
          transition: {
            type: "fade",
            startSound: { type: "wav", data: new Uint8Array([1, 2, 3]) },
          },
        },
        {
          transition: {
            type: "wipe",
            startSound: { type: "wav", data: new Uint8Array([4, 5, 6]) },
          },
        },
      ],
    });
    const archive = await readOcf(generated, "application/vnd.oasis.opendocument.presentation");
    expect(archive.binaries["Media/transition1.wav"]).toEqual(new Uint8Array([1, 2, 3]));
    expect(archive.binaries["Media/transition2.wav"]).toEqual(new Uint8Array([4, 5, 6]));
    expect(archive.files["content.xml"]).toContain('xlink:href="Media/transition1.wav"');
    expect(archive.files["content.xml"]).toContain('xlink:href="Media/transition2.wav"');
  });

  it("reuses identical transition sounds", async () => {
    const generated = await generatePresentation({
      slides: [
        {
          transition: {
            type: "fade",
            startSound: { type: "wav", data: new Uint8Array([1, 2, 3]) },
          },
        },
        {
          transition: {
            type: "wipe",
            startSound: { type: "wav", data: new Uint8Array([1, 2, 3]) },
          },
        },
      ],
    });
    const archive = await readOcf(generated, "application/vnd.oasis.opendocument.presentation");
    expect(Object.keys(archive.binaries).filter((path) => path.endsWith(".wav"))).toHaveLength(1);
    expect(
      archive.files["content.xml"]!.match(/xlink:href="Media\/transition1\.wav"/g),
    ).toHaveLength(2);
  });

  it("accepts XML package members as transition sounds", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:automatic-styles><style:style style:name="DP1" style:family="drawing-page"><style:drawing-page-properties><presentation:sound xlink:href="/Media/sound.xml"/></style:drawing-page-properties></style:style></office:automatic-styles><office:body><office:presentation><draw:page draw:style-name="DP1"/></office:presentation></office:body></office:document-content>`;
    const parsed = await parsePresentation(
      generateOcf("application/vnd.oasis.opendocument.presentation", {
        "content.xml": content,
        "Media/sound.xml": "<sound/>",
      }),
    );
    const transition = parsed.slides?.[0]?.transition;
    expect(typeof transition).toBe("object");
    expect((transition as OdpTransitionOptions).soundHref).toBe("/Media/sound.xml");
  });

  it("treats protocol-relative transition sound URLs as external", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:automatic-styles><style:style style:name="DP1" style:family="drawing-page"><style:drawing-page-properties><presentation:sound xlink:href="//media.example.test/sound.wav"/></style:drawing-page-properties></style:style></office:automatic-styles><office:body><office:presentation><draw:page draw:style-name="DP1"/></office:presentation></office:body></office:document-content>`;
    const parsed = await parsePresentation(
      generateOcf("application/vnd.oasis.opendocument.presentation", {
        "content.xml": content,
        "media.example.test/sound.wav": new Uint8Array([1]),
      }),
    );
    const transition = parsed.slides?.[0]?.transition;
    expect(typeof transition).toBe("object");
    expect((transition as OdpTransitionOptions).soundHref).toBe("//media.example.test/sound.wav");
  });

  it("rejects a missing package-relative transition sound", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:automatic-styles><style:style style:name="DP1" style:family="drawing-page"><style:drawing-page-properties><presentation:sound xlink:href="Sounds/missing.wav"/></style:drawing-page-properties></style:style></office:automatic-styles><office:body><office:presentation><draw:page draw:style-name="DP1"/></office:presentation></office:body></office:document-content>`;
    await expect(
      parsePresentation(
        generateOcf("application/vnd.oasis.opendocument.presentation", { "content.xml": content }),
      ),
    ).rejects.toMatchObject({ reason: "missing-transition-sound" });
  });

  it("merges patched transitions into the referenced drawing-page style", async () => {
    const generated = await generatePresentation({
      styleResources: [
        {
          name: "DP1",
          family: "drawing-page",
          properties: [
            {
              name: "style:drawing-page-properties",
              attributes: { "presentation:transition-type": "manual" },
            },
          ],
        },
      ],
      slides: [
        {
          drawingPageStyleName: "DP1",
          transition: {
            type: "fade",
            advanceOnClick: false,
            advanceAfterTime: 1000,
            speed: "fast",
            soundHref: "https://office-open.test/media/sound.wav",
            soundPlayFull: true,
            soundXmlId: "sound",
          },
        },
      ],
    });
    const styles = (await readOcf(generated, "application/vnd.oasis.opendocument.presentation"))
      .files["content.xml"]!;
    expect(styles.match(/style:name="DP1"/g)).toHaveLength(1);
    expect(styles).toContain('presentation:transition-type="automatic"');
    expect(styles).toContain('presentation:transition-speed="fast"');
    expect(styles).toContain('presentation:duration="PT1S"');
    expect(styles).toContain('xlink:href="https://office-open.test/media/sound.wav"');
    expect(styles).toContain('presentation:play-full="true"');
    expect(styles).toContain('xml:id="sound"');
    const parsed = await parsePresentation(generated);
    expect(parsed.slides?.[0]?.transition).toMatchObject({
      advanceOnClick: false,
      advanceAfterTime: 1000,
      speed: "fast",
      soundHref: "https://office-open.test/media/sound.wav",
      soundPlayFull: true,
      soundXmlId: "sound",
    });
  });

  it("rejects conflicting transitions projected into one shared style", async () => {
    await expect(
      generatePresentation({
        styleResources: [
          {
            name: "DP1",
            family: "drawing-page",
            properties: [{ name: "style:drawing-page-properties", attributes: {} }],
          },
        ],
        slides: [
          {
            drawingPageStyleName: "DP1",
            transition: { type: "fade", speed: "fast" },
          },
          {
            drawingPageStyleName: "DP1",
            transition: { type: "fade", speed: "slow" },
          },
        ],
      }),
    ).rejects.toThrow(/different transition styles/u);
  });

  it("rejects transitions ODF cannot represent without losing fields", async () => {
    await expect(
      generatePresentation({
        slides: [
          {
            transition: {
              type: "fade",
              soundHref: "https://office-open.test/media/sound.wav",
              startSound: { type: "wav", data: new Uint8Array([1]) },
            },
          },
        ],
      }),
    ).rejects.toThrow(/both embedded and linked sound/u);
    await expect(
      generatePresentation({
        slides: [{ transition: { type: "fade", stopPreviousSound: true } }],
      }),
    ).rejects.toThrow(/stopping the previous transition sound/u);
    await expect(
      generatePresentation({
        slides: [{ transition: { type: "fade", soundPlayFull: true } }],
      }),
    ).rejects.toThrow(/sound metadata requires embedded or linked sound/u);
  });

  it("round-trips OOXML cut, pull, strips, and wheel transitions", async () => {
    const transitions = [
      '<anim:transitionFilter smil:type="fade" smil:subtype="default"/>',
      '<anim:transitionFilter smil:type="slideWipe" smil:subtype="fromLeft" smil:direction="reverse"/>',
      '<anim:transitionFilter smil:type="waterfallWipe" smil:subtype="horizontalRight"/>',
      '<anim:transitionFilter smil:type="waterfallWipe" smil:subtype="horizontalRight" smil:direction="reverse"/>',
      '<anim:transitionFilter smil:type="pinWheelWipe" smil:subtype="threeBlade"/>',
    ];
    const expected: Partial<TransitionOptions>[] = [
      { type: "cut" },
      { type: "pull", direction: "right" },
      { type: "strips", direction: "leftDown" },
      { type: "strips", direction: "rightUp" },
      { type: "wheel", spokes: 3 },
    ];
    for (const [filter, transition] of transitions.map(
      (filter, index) => [filter, expected[index]] as const,
    )) {
      if (transition === undefined) throw new Error("missing expected transition");
      const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:presentation><draw:page><anim:par>${filter}</anim:par></draw:page></office:presentation></office:body></office:document-content>`;
      const parsed = await parsePresentation(
        generateOcf("application/vnd.oasis.opendocument.presentation", { "content.xml": content }),
      );
      expect(parsed.slides?.[0]?.transition).toMatchObject(transition);
      const regenerated = await parsePresentation(await generatePresentation(parsed));
      expect(regenerated.slides?.[0]?.transition).toEqual(parsed.slides?.[0]?.transition);
    }
  });

  it("preserves unmapped SMIL transition types in typed animations", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:presentation><draw:page><anim:par><anim:transitionFilter smil:type="barnDoorWipe"/></anim:par></draw:page></office:presentation></office:body></office:document-content>`;
    const parsed = await parsePresentation(
      generateOcf("application/vnd.oasis.opendocument.presentation", { "content.xml": content }),
    );
    expect(parsed.slides?.[0]?.animations).toEqual([
      {
        children: [
          {
            behavior: "transitionFilter",
            fadeColor: undefined,
            transitionSubType: undefined,
            transitionType: "barnDoorWipe",
          },
        ],
        timeNode: "par",
      },
    ]);
    const regenerated = await parsePresentation(await generatePresentation(parsed));
    expect(regenerated.slides?.[0]?.animations).toEqual(parsed.slides?.[0]?.animations);
  });

  it("preserves compatibility presentation additive tokens", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:presentation><draw:page><anim:par presentation:additive="base"/></draw:page></office:presentation></office:body></office:document-content>`;
    const parsed = await parsePresentation(
      generateOcf("application/vnd.oasis.opendocument.presentation", { "content.xml": content }),
    );
    expect(parsed.slides?.[0]?.animations?.[0]).toMatchObject({
      presentationAdditiveToken: "base",
      timeNode: "par",
    });
    const regenerated = await parsePresentation(await generatePresentation(parsed));
    expect(regenerated.slides?.[0]?.animations?.[0]).toMatchObject({
      presentationAdditiveToken: "base",
    });
  });

  it("treats inherited object names as unknown drawing presets", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:presentation><draw:page><draw:custom-shape svg:x="1cm" svg:y="1cm" svg:width="1cm" svg:height="1cm"><draw:enhanced-geometry draw:type="constructor"/></draw:custom-shape></draw:page></office:presentation></office:body></office:document-content>`;
    const parsed = await parsePresentation(
      generateOcf("application/vnd.oasis.opendocument.presentation", { "content.xml": content }),
    );
    expect(parsed.slides?.[0]?.children).toEqual([
      {
        shape: {
          height: 360000,
          name: undefined,
          odfGeometryType: "constructor",
          width: 360000,
          x: 360000,
          y: 360000,
        },
      },
    ]);
  });

  it("rejects inherited object names as style families", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:automatic-styles><style:style style:name="bad" style:family="constructor"/></office:automatic-styles><office:body><office:presentation><draw:page/></office:presentation></office:body></office:document-content>`;
    await expect(
      parsePresentation(
        generateOcf("application/vnd.oasis.opendocument.presentation", { "content.xml": content }),
      ),
    ).rejects.toMatchObject({ reason: "unknown style family" });
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

  it("round-trips a missing package image as a linked-only picture", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:presentation><draw:page><draw:frame svg:x="1cm" svg:y="1cm" svg:width="4cm" svg:height="3cm"><draw:image xlink:href="Pictures/missing.png" xlink:show="embed"/></draw:frame></draw:page></office:presentation></office:body></office:document-content>`;
    const parsed = await parsePresentation(
      generateOcf("application/vnd.oasis.opendocument.presentation", { "content.xml": content }),
    );
    expect(parsed.slides?.[0]?.children?.[0]).toMatchObject({
      picture: { type: "png", sourceUrl: "Pictures/missing.png" },
    });
    expect(
      (await parsePresentation(await generatePresentation(parsed))).slides?.[0]?.children?.[0],
    ).toEqual(parsed.slides?.[0]?.children?.[0]);
  });

  it("projects compatibility legacy drawings as positioned frames", async () => {
    const parsed = await parsePresentation(
      await generatePresentation({
        slides: [
          {
            children: [
              {
                rawXml:
                  '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2056" name="Legacy"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' +
                  '<p:xfrm><a:off x="5508625" y="4508500"/><a:ext cx="3419475" cy="1709738"/></p:xfrm>' +
                  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/compatibility">' +
                  '<com:legacyDrawing xmlns:com="http://schemas.openxmlformats.org/drawingml/2006/compatibility" spid="_x0000_s2056"/>' +
                  "</a:graphicData></a:graphic></p:graphicFrame>",
              },
            ],
          },
        ],
      }),
    );
    expect(parsed.slides?.[0]?.children?.[0]).toMatchObject({
      shape: { id: 2056, name: "Legacy", x: 5508625, y: 4508500, width: 3419475, height: 1709738 },
    });
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
      styleResources: [
        {
          name: "P1",
          family: "paragraph",
          properties: [
            {
              name: "style:paragraph-properties",
              attributes: {
                "fo:line-height": 150,
                "fo:margin-left": 360000,
                "style:contextual-spacing": true,
              },
              children: [
                {
                  name: "style:tab-stops",
                  stops: [{ attributes: { "style:position": 914400, "style:type": "right" } }],
                },
              ],
            },
          ],
          maps: [
            {
              name: "style:map",
              condition: "true",
              applyStyleName: "Strong",
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
    expect(first).toContain('style:contextual-spacing="true"');

    const parsed = await parsePresentation(generated);
    expect(parsed.styleResources).toEqual(options.styleResources);
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

  it("round-trips slide table names, mixed length units, and covered cells", async () => {
    const options: OdpPresentationOptions = {
      slides: [
        {
          children: [
            {
              table: {
                name: "Budget",
                columnWidths: [914400, 1828800],
                rows: [
                  {
                    cells: [{ text: "A", columnSpan: 2, rowSpan: 2 }, { text: "B" }],
                  },
                  { cells: [{ text: "C" }] },
                ],
              },
            },
          ],
        },
      ],
    };
    const parsed = await parsePresentation(await generatePresentation(options));
    expect(parsed.slides?.[0]?.children?.[0]).toMatchObject({
      table: {
        name: "Budget",
        columnWidths: [914400, 1828800],
        rows: [
          {
            cells: [
              { children: [{ text: "A" }], columnSpan: 2, rowSpan: 2 },
              { children: [{ text: "B" }] },
            ],
          },
          { cells: [{ children: [{ text: "C" }] }] },
        ],
      },
    });
  });

  it("rejects covered cells that do not match their preceding span", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:automatic-styles><style:style style:name="W" style:family="table-column"><style:table-column-properties style:column-width="25.4mm"/></style:style></office:automatic-styles><office:body><office:presentation><draw:page><table:table table:name="Budget"><table:table-column table:style-name="W"/><table:table-row><table:table-cell table:number-columns-spanned="2"><text:p>A</text:p></table:table-cell></table:table-row></table:table></draw:page></office:presentation></office:body></office:document-content>`;
    await expect(
      parsePresentation(
        generateOcf("application/vnd.oasis.opendocument.presentation", { "content.xml": content }),
      ),
    ).rejects.toThrow(OdpParseError);
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
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:presentation><draw:page><draw:rect svg:x="1cm" svg:y="1cm" svg:width="2cm" svg:height="1cm" draw:name="Box"><text:p>Box</text:p></draw:rect><draw:path svg:x="4cm" svg:y="1cm" svg:width="2cm" svg:height="1cm" svg:viewBox="0 0 2 1" svg:d="M 0 0 L 2 0 L 2 1 Z"/><draw:g><draw:custom-shape svg:x="1cm" svg:y="3cm" svg:width="1cm" svg:height="1cm" draw:name="Grouped"/></draw:g></draw:page></office:presentation></office:body></office:document-content>`;
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
                  controlId: "control1",
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
      shape: { controlId: "control1", properties: { geometry: "rect" } },
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
    expect(error).toMatchObject({ part: "content.xml", elementName: "form:unknown" });
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

  it("assigns an object path name to an unnamed slide chart", async () => {
    const parsed = await parsePresentation(
      await generatePresentation({
        slides: [{ children: [{ chart: { type: "pie", series: [{ values: [1] }] } }] }],
      }),
    );
    expect(parsed.slides?.[0]?.children?.[0]).toMatchObject({ chart: { name: "Object 1" } });
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
      elementName: "draw:unknown",
      reason: "element has no canonical PresentationOptions mapping",
    });
  });

  it("types and regenerates a legacy object-ole frame", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:presentation><draw:page><draw:frame svg:x="1cm" svg:y="1cm" svg:width="1cm" svg:height="1cm"><draw:object-ole draw:class-id="Legacy.Object" xlink:href="./Object 1" xlink:type="simple"/></draw:frame></draw:page></office:presentation></office:body></office:document-content>`;
    const parsed = await parsePresentation(
      generateOcf("application/vnd.oasis.opendocument.presentation", {
        "content.xml": content,
        "Object 1": new Uint8Array([1]),
      }),
    );
    expect(parsed.slides?.[0]?.odpNodes?.[0]?.node).toMatchObject({
      frame: {
        contents: [
          {
            kind: "object-ole",
            objectOle: {
              classId: "Legacy.Object",
              data: { href: "./Object 1" },
            },
          },
        ],
      },
    });
    expect(
      (await parsePresentation(await generatePresentation(parsed))).slides?.[0]?.odpNodes,
    ).toEqual(parsed.slides?.[0]?.odpNodes);
  });

  it("requires an embedded replacement image", async () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:presentation><draw:page><draw:frame svg:x="1cm" svg:y="1cm" svg:width="1cm" svg:height="1cm"><draw:object xlink:href="./Object 1" xlink:type="simple"/><draw:image xlink:href="./ObjectReplacements/Object 1" xlink:type="simple"/></draw:frame></draw:page></office:presentation></office:body></office:document-content>`;
    await expect(
      parsePresentation(
        generateOcf("application/vnd.oasis.opendocument.presentation", {
          "content.xml": content,
          "Object 1/content.xml": "<math/>",
        }),
      ),
    ).rejects.toMatchObject({ reason: "embedded-replacement-image-missing" });
  });

  it("round-trips picture accessibility metadata and unknown external image types", async () => {
    const parsed = await parsePresentation(
      await generatePresentation({
        slides: [
          {
            children: [
              {
                picture: {
                  type: "png",
                  data: new Uint8Array([1]),
                  x: 0,
                  y: 0,
                  width: 360000,
                  height: 360000,
                  title: "Chart image",
                  description: "Quarterly revenue",
                },
              },
            ],
          },
        ],
      }),
    );
    expect(parsed.slides?.[0]?.children?.[0]).toMatchObject({
      picture: { title: "Chart image", description: "Quarterly revenue" },
    });
    expect(await parsePresentation(await generatePresentation(parsed))).toEqual(parsed);
    const unsupported = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES}><office:body><office:presentation><draw:page><draw:frame svg:x="1cm" svg:y="1cm" svg:width="1cm" svg:height="1cm"><draw:image xlink:href="https://example.com/image.tiff"/></draw:frame></draw:page></office:presentation></office:body></office:document-content>`;
    const unknown = await parsePresentation(
      generateOcf("application/vnd.oasis.opendocument.presentation", {
        "content.xml": unsupported,
      }),
    );
    expect(unknown.slides?.[0]?.odpNodes?.[0]?.node).toMatchObject({
      frame: {
        contents: [{ kind: "image", image: { data: { href: "https://example.com/image.tiff" } } }],
      },
    });
    const regenerated = await parsePresentation(await generatePresentation(unknown));
    expect(regenerated.slides?.[0]?.odpNodes).toEqual(unknown.slides?.[0]?.odpNodes);
  });

  it("validates slide event listeners strictly", async () => {
    const frame = (listener: string) =>
      `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES} xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"><office:body><office:presentation><draw:page><draw:frame svg:x="1cm" svg:y="1cm" svg:width="1cm" svg:height="1cm"><draw:text-box/><office:event-listeners>${listener}</office:event-listeners></draw:frame></draw:page></office:presentation></office:body></office:document-content>`;
    const ambiguous = `<script:event-listener script:event-name="click" script:language="Basic" script:macro-name="Main" xlink:href="https://example.com"/>`;
    await expect(
      parsePresentation(
        generateOcf("application/vnd.oasis.opendocument.presentation", {
          "content.xml": frame(ambiguous),
        }),
      ),
    ).rejects.toMatchObject({ reason: "ambiguous-event-target" });
    const unknown = `<script:event-listener script:event-name="click" script:language="Basic" xlink:href="https://example.com" xlink:extra="1"/>`;
    await expect(
      parsePresentation(
        generateOcf("application/vnd.oasis.opendocument.presentation", {
          "content.xml": frame(unknown),
        }),
      ),
    ).rejects.toThrow(OdpParseError);
  });
});
