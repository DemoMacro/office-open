import { EmptyElement, XmlComponent, chartAttr, wrapEl } from "../../xml-components";

export interface BubbleSeriesData {
  readonly name: string;
  readonly xValues: readonly number[];
  readonly yValues: readonly number[];
  readonly bubbleSize: readonly number[];
}

interface BubbleChartOptions {
  readonly series: readonly BubbleSeriesData[];
}

export class BubbleChart extends XmlComponent {
  public constructor(options: BubbleChartOptions) {
    super("c:bubbleChart");
    this.root.push(wrapEl("c:varyColors", chartAttr({ val: true })));

    for (let i = 0; i < options.series.length; i++) {
      this.root.push(new BubbleSeries(i, options.series[i]));
    }

    this.root.push(wrapEl("c:axId", chartAttr({ val: 10 })));
    this.root.push(wrapEl("c:axId", chartAttr({ val: 20 })));
  }
}

class BubbleSeries extends XmlComponent {
  public constructor(index: number, series: BubbleSeriesData) {
    super("c:ser");
    this.root.push(wrapEl("c:idx", chartAttr({ val: index })));
    this.root.push(wrapEl("c:order", chartAttr({ val: index })));
    this.root.push(new SeriesTx(series.name));
    this.root.push(new EmptyElement("c:spPr"));
    this.root.push(new SeriesXValues(series.xValues));
    this.root.push(new SeriesYValues(series.yValues));
    this.root.push(new SeriesBubbleSize(series.bubbleSize));
  }
}

class SeriesTx extends XmlComponent {
  public constructor(name: string) {
    super("c:tx");
    this.root.push(createStrRef(name));
  }
}

const createStrRef = (value: string): XmlComponent =>
  new (class extends XmlComponent {
    public constructor() {
      super("c:strRef");
      this.root.push(new EmptyElement("c:f"));
      this.root.push(new StrCache(value));
    }
  })();

class StrCache extends XmlComponent {
  public constructor(value: string) {
    super("c:strCache");
    this.root.push(wrapEl("c:ptCount", chartAttr({ val: 1 })));
    this.root.push(new StrPt(0, value));
  }
}

class StrPt extends XmlComponent {
  public constructor(index: number, value: string) {
    super("c:pt");
    this.root.push(chartAttr({ idx: index }));
    this.root.push(new StringValue("c:v", value));
  }
}

class StringValue extends XmlComponent {
  public constructor(name: string, val: string) {
    super(name);
    this.root.push(val);
  }
}

class SeriesXValues extends XmlComponent {
  public constructor(values: readonly number[]) {
    super("c:xVal");
    this.root.push(createNumRef(values));
  }
}

class SeriesYValues extends XmlComponent {
  public constructor(values: readonly number[]) {
    super("c:yVal");
    this.root.push(createNumRef(values));
  }
}

class SeriesBubbleSize extends XmlComponent {
  public constructor(values: readonly number[]) {
    super("c:bubbleSize");
    this.root.push(createNumRef(values));
  }
}

const createNumRef = (values: readonly number[]): XmlComponent =>
  new (class extends XmlComponent {
    public constructor() {
      super("c:numRef");
      this.root.push(new EmptyElement("c:f"));
      this.root.push(new NumCache(values));
    }
  })();

class NumCache extends XmlComponent {
  public constructor(values: readonly number[]) {
    super("c:numCache");
    this.root.push(new FormatCode("General"));
    this.root.push(wrapEl("c:ptCount", chartAttr({ val: values.length })));
    for (let i = 0; i < values.length; i++) {
      this.root.push(new NumPt(i, values[i]));
    }
  }
}

class NumPt extends XmlComponent {
  public constructor(index: number, value: number) {
    super("c:pt");
    this.root.push(chartAttr({ idx: index }));
    this.root.push(new StringValue("c:v", String(value)));
  }
}

class FormatCode extends XmlComponent {
  public constructor(code: string) {
    super("c:formatCode");
    this.root.push(code);
  }
}
