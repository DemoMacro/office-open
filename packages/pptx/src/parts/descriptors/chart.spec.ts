import { describe, expect, it } from "vite-plus/test";

import type { ChartOptions } from "../chart-frame";
import { toChartSpaceOptions } from "./chart";

describe("chart frame payload", () => {
  it("keeps cNvPr extensions out of the chart-space extension list", () => {
    const options: ChartOptions = {
      type: "bar",
      name: "Frame",
      creationId: "{frame-id}",
      ext: '<a:ext uri="{frame-extension}"/>',
      chartExt: '<c:ext uri="{chart-extension}"/>',
      series: [],
    };

    expect(toChartSpaceOptions(options)).not.toHaveProperty("creationId");
    expect(toChartSpaceOptions(options).chartExt).toBe('<c:ext uri="{chart-extension}"/>');
  });
});
