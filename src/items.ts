import { BorderStyle, ExternalHyperlink, Paragraph, TableCell, TableRow, TextRun } from 'docx';

const createTableCell = (content: string | Paragraph[]): TableCell =>
    new TableCell({
        children: Array.isArray(content)
            ? content
            : [new Paragraph({ children: [new TextRun({ text: content, size: 22 })] })],
        borders: {
            top: { style: BorderStyle.SINGLE, size: 1, color: '#000000' },
            bottom: { style: BorderStyle.SINGLE, size: 1, color: '#000000' },
            left: { style: BorderStyle.SINGLE, size: 1, color: '#000000' },
            right: { style: BorderStyle.SINGLE, size: 1, color: '#000000' },
        },
        margins: {
            top: 100,
            bottom: 100,
            left: 100,
            right: 100,
        },
    });

export const createHyperlink = (resource: { title: string; url: string }): Paragraph =>
    new Paragraph({
        children: [
            new ExternalHyperlink({
                children: [new TextRun({ text: resource.title, color: '#0563c1', underline: {}, size: 22 })],
                link: resource.url,
            }),
        ],
        bullet: { level: 0 },
        spacing: { line: 220, before: 50 },
    });

export const createRowData = (label: string, content: string | string[] | Paragraph[]): TableRow => {
    const array =
        Array.isArray(content) && content[0] instanceof Paragraph
            ? (content as Paragraph[])
            : Array.isArray(content)
              ? (content as string[]).map(
                    (item) =>
                        new Paragraph({
                            children: [new TextRun({ text: item, size: 22 })],
                            bullet: { level: 0 },
                            spacing: { line: 220, before: 50 },
                        }),
                )
              : [new Paragraph({ children: [new TextRun({ text: content, size: 22 })] })];

    return new TableRow({
        children: [createTableCell(label), createTableCell(array)],
    });
};
