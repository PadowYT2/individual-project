import { AlignmentType, BorderStyle, ExternalHyperlink, Paragraph, TableCell, TableRow, TextRun } from 'docx';

const FONT_SIZE = 28;

export const createTitle = (text: string): Paragraph =>
    new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 300 },
        children: [new TextRun({ text, size: 32, bold: true })],
    });

export const createSubtitle = (text: string): Paragraph =>
    new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 },
        children: [new TextRun({ text, size: FONT_SIZE, italics: true })],
    });

export const createHeading = (text: string): Paragraph =>
    new Paragraph({
        spacing: { before: 300, after: 150 },
        children: [new TextRun({ text, size: FONT_SIZE, bold: true })],
    });

export const createParagraph = (text: string): Paragraph =>
    new Paragraph({
        spacing: { after: 150 },
        indent: { firstLine: 400 },
        children: [new TextRun({ text, size: FONT_SIZE })],
    });

export const createBulletItem = (text: string, level = 0): Paragraph =>
    new Paragraph({
        bullet: { level },
        spacing: { line: 276, before: 80 },
        children: [new TextRun({ text, size: FONT_SIZE })],
    });

export const createNumberedItem = (text: string, index: number): Paragraph =>
    new Paragraph({
        spacing: { line: 276, before: 80 },
        indent: { left: 400 },
        children: [new TextRun({ text: `${index}. ${text}`, size: FONT_SIZE })],
    });

export const createHyperlink = (resource: { title: string; url: string }, index: number): Paragraph =>
    new Paragraph({
        spacing: { line: 276, before: 80 },
        indent: { left: 400 },
        children: [
            new TextRun({ text: `${index}. `, size: FONT_SIZE }),
            new ExternalHyperlink({
                children: [
                    new TextRun({
                        text: resource.title,
                        color: '0563c1',
                        underline: {},
                        size: FONT_SIZE,
                    }),
                ],
                link: resource.url,
            }),
        ],
    });

const createTableCell = (content: string, bold = false): TableCell =>
    new TableCell({
        children: [new Paragraph({ children: [new TextRun({ text: content, size: FONT_SIZE, bold })] })],
        borders: {
            top: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
            bottom: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
            left: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
            right: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
        },
        margins: { top: 80, bottom: 80, left: 100, right: 100 },
    });

export const createTableHeader = (cells: string[]): TableRow =>
    new TableRow({
        children: cells.map((cell) => createTableCell(cell, true)),
    });

export const createTableRowSimple = (cells: string[]): TableRow =>
    new TableRow({
        children: cells.map((cell) => createTableCell(cell)),
    });
