import { describe, expect, it } from "vite-plus/test";

import { Numbering } from "./numbering";

describe("Numbering", () => {
  describe("#constructor", () => {
    it("creates a default numbering with one abstract and one concrete instance", () => {
      const numbering = new Numbering({
        config: [],
      });

      const xml = numbering.toXml({ stack: [], file: null!, viewWrapper: null! });

      expect(xml).to.contain("<w:numbering");
      expect(xml).to.contain("<w:abstractNum");
      expect(xml).to.contain("<w:num ");

      // Should contain abstractNumId attribute
      expect(xml).to.contain('w:abstractNumId="');
    });

    describe("#createConcreteNumberingInstance", () => {
      it("should create a concrete numbering instance", () => {
        const numbering = new Numbering({
          config: [
            {
              levels: [
                {
                  level: 0,
                },
              ],
              reference: "test-reference",
            },
          ],
        });
        expect(numbering.concreteNumbering).to.have.length(1);

        numbering.createConcreteNumberingInstance("test-reference", 0);

        expect(numbering.concreteNumbering).to.have.length(2);
      });

      it("should not create a concrete numbering instance if reference is invalid", () => {
        const numbering = new Numbering({
          config: [
            {
              levels: [
                {
                  level: 0,
                },
              ],
              reference: "test-reference",
            },
          ],
        });
        expect(numbering.concreteNumbering).to.have.length(1);

        numbering.createConcreteNumberingInstance("invalid-reference", 0);

        expect(numbering.concreteNumbering).to.have.length(1);
      });

      it("should not create a concrete numbering instance if one already exists", () => {
        const numbering = new Numbering({
          config: [
            {
              levels: [
                {
                  level: 0,
                },
              ],
              reference: "test-reference",
            },
          ],
        });

        expect(numbering.concreteNumbering).to.have.length(1);

        numbering.createConcreteNumberingInstance("test-reference", 0);
        numbering.createConcreteNumberingInstance("test-reference", 0);

        expect(numbering.concreteNumbering).to.have.length(2);
      });
    });
    describe("#referenceConfigMap", () => {
      it("should store level configs into referenceConfigMap", () => {
        const numbering = new Numbering({
          config: [
            {
              levels: [
                {
                  level: 0,
                  start: 10,
                },
              ],
              reference: "test-reference",
            },
          ],
        });
        numbering.createConcreteNumberingInstance("test-reference", 0);
        const referenceConfig = numbering.referenceConfig[0];
        const zeroLevelConfig = referenceConfig[0];
        expect(zeroLevelConfig.start).to.be.equal(10);
      });
    });
  });
});
