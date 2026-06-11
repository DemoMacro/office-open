// Example of how you would create a table and add data to it from a data source

import { writeFileSync } from "node:fs";

import {
  HeadingLevel,
  TextDirection,
  VerticalAlignTable,
  WidthType,
  generateDocument,
} from "@office-open/docx";

interface StockPrice {
  readonly date: Date;
  readonly ticker: string;
  readonly price: number;
}

const DATA: StockPrice[] = [
  {
    date: new Date("2007-08-28"),
    price: 18.12,
    ticker: "Apple",
  },
  {
    date: new Date("2007-08-29"),
    price: 19.15,
    ticker: "Apple",
  },
  {
    date: new Date("2007-08-30"),
    price: 19.46,
    ticker: "Apple",
  },
  {
    date: new Date("2007-08-31"),
    price: 19.78,
    ticker: "Apple",
  },
  {
    date: new Date("2007-09-04"),
    price: 20.59,
    ticker: "Apple",
  },
  {
    date: new Date("2007-09-05"),
    price: 19.54,
    ticker: "Apple",
  },
  {
    date: new Date("2007-09-06"),
    price: 19.29,
    ticker: "Apple",
  },
  {
    date: new Date("2007-09-07"),
    price: 18.82,
    ticker: "Apple",
  },
  {
    date: new Date("2007-09-10"),
    price: 19.53,
    ticker: "Apple",
  },
  {
    date: new Date("2007-09-11"),
    price: 19.36,
    ticker: "Apple",
  },
  {
    date: new Date("2007-09-12"),
    price: 19.55,
    ticker: "Apple",
  },
  {
    date: new Date("2007-09-13"),
    price: 19.6,
    ticker: "Apple",
  },
  {
    date: new Date("2007-09-14"),
    price: 19.83,
    ticker: "Apple",
  },
  {
    date: new Date("2007-09-17"),
    price: 19.77,
    ticker: "Apple",
  },
];

const generateRows = (prices: StockPrice[]) =>
  prices.map(({ date, ticker, price }) => ({
    cells: [
      {
        children: [{ paragraph: date.toString() }],
        textDirection: TextDirection.LEFT_TO_RIGHT_TOP_TO_BOTTOM,
        verticalAlign: VerticalAlignTable.CENTER,
      },
      {
        children: [{ paragraph: ticker }],
        textDirection: TextDirection.LEFT_TO_RIGHT_TOP_TO_BOTTOM,
        verticalAlign: VerticalAlignTable.CENTER,
      },
      {
        children: [{ paragraph: price.toString() }],
        textDirection: TextDirection.TOP_TO_BOTTOM_RIGHT_TO_LEFT,
        verticalAlign: VerticalAlignTable.CENTER,
      },
    ],
  }));

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          table: {
            rows: [
              {
                cells: [
                  {
                    children: [
                      {
                        paragraph: {
                          heading: HeadingLevel.HEADING_2,
                          children: [
                            {
                              text: "Date",
                              bold: true,
                              size: 40,
                            },
                          ],
                        },
                      },
                    ],
                    textDirection: TextDirection.LEFT_TO_RIGHT_TOP_TO_BOTTOM,
                    verticalAlign: VerticalAlignTable.CENTER,
                  },
                  {
                    children: [
                      {
                        paragraph: {
                          heading: HeadingLevel.HEADING_2,
                          children: [
                            {
                              text: "Ticker",
                              bold: true,
                              size: 40,
                            },
                          ],
                        },
                      },
                    ],
                    textDirection: TextDirection.LEFT_TO_RIGHT_TOP_TO_BOTTOM,
                    verticalAlign: VerticalAlignTable.CENTER,
                  },
                  {
                    children: [
                      {
                        paragraph: {
                          heading: HeadingLevel.HEADING_2,
                          children: [
                            {
                              text: "Price",
                              bold: true,
                              size: 40,
                            },
                          ],
                        },
                      },
                    ],
                    textDirection: TextDirection.TOP_TO_BOTTOM_RIGHT_TO_LEFT,
                    verticalAlign: VerticalAlignTable.CENTER,
                  },
                ],
              },
              ...generateRows(DATA),
            ],
            width: {
              size: 9070,
              type: WidthType.DXA,
            },
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
