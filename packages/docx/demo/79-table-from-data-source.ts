// Example of how you would create a table and add data to it from a data source

import * as fs from "fs";

import {
    Document,
    HeadingLevel,
    Packer,
    Paragraph,
    Table,
    TableCell,
    TableRow,
    TextDirection,
    TextRun,
    VerticalAlignTable,
    WidthType,
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

const generateRows = (prices: StockPrice[]): TableRow[] =>
    prices.map(
        ({ date, ticker, price }) =>
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph(date.toString())],
                        textDirection: TextDirection.LEFT_TO_RIGHT_TOP_TO_BOTTOM,
                        verticalAlign: VerticalAlignTable.CENTER,
                    }),
                    new TableCell({
                        children: [new Paragraph(ticker)],
                        textDirection: TextDirection.LEFT_TO_RIGHT_TOP_TO_BOTTOM,
                        verticalAlign: VerticalAlignTable.CENTER,
                    }),
                    new TableCell({
                        children: [new Paragraph(price.toString())],
                        textDirection: TextDirection.TOP_TO_BOTTOM_RIGHT_TO_LEFT,
                        verticalAlign: VerticalAlignTable.CENTER,
                    }),
                ],
            }),
    );

const doc = new Document({
    sections: [
        {
            children: [
                new Table({
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            heading: HeadingLevel.HEADING_2,
                                            children: [
                                                new TextRun({
                                                    text: "Date",
                                                    bold: true,
                                                    size: 40,
                                                }),
                                            ],
                                        }),
                                    ],
                                    textDirection: TextDirection.LEFT_TO_RIGHT_TOP_TO_BOTTOM,
                                    verticalAlign: VerticalAlignTable.CENTER,
                                }),
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            heading: HeadingLevel.HEADING_2,
                                            children: [
                                                new TextRun({
                                                    text: "Ticker",
                                                    bold: true,
                                                    size: 40,
                                                }),
                                            ],
                                        }),
                                    ],
                                    textDirection: TextDirection.LEFT_TO_RIGHT_TOP_TO_BOTTOM,
                                    verticalAlign: VerticalAlignTable.CENTER,
                                }),
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            heading: HeadingLevel.HEADING_2,
                                            children: [
                                                new TextRun({
                                                    text: "Price",
                                                    bold: true,
                                                    size: 40,
                                                }),
                                            ],
                                        }),
                                    ],
                                    textDirection: TextDirection.TOP_TO_BOTTOM_RIGHT_TO_LEFT,
                                    verticalAlign: VerticalAlignTable.CENTER,
                                }),
                            ],
                        }),
                        ...generateRows(DATA),
                    ],
                    width: {
                        size: 9070,
                        type: WidthType.DXA,
                    },
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
