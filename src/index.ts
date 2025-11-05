import Bun from 'bun';
import { AlignmentType, Document, Packer, Paragraph, Table, TableRow, TextRun, WidthType } from 'docx';
import data from '@/data.json' with { type: 'json' };
import { createHyperlink, createRowData } from '@/items';

const tableRows: TableRow[] = [
    createRowData('Тема', data.theme),
    createRowData('Актуальность', data.relevance),
    createRowData('Проблема', data.problem),
    createRowData('Гипотеза', data.hypothesis),
    createRowData('Объект исследования', data.researchObject),
    createRowData('Предмет исследования', data.researchSubject),
    createRowData('Замысел проекта', data.researchIdea),
    createRowData('Методы', data.methods),
    createRowData(
        'Ресурсы',
        [
            { label: 'Законодательные:', items: data.resources.legislative },
            { label: 'Теоретические:', items: data.resources.theoretical },
            { label: 'Статистические:', items: data.resources.statistical },
            { label: 'Технические:', items: data.resources.technical },
        ].flatMap(({ label, items }) => [
            new Paragraph({
                children: [new TextRun({ text: label, size: 22 })],
                spacing: { before: 100, after: 100 },
            }),
            ...items.map(createHyperlink),
        ]),
    ),
    createRowData('Цели', data.goals),
    createRowData('Задачи', data.tasks),
    createRowData('Этапы', data.stages),
    createRowData('Продукт', data.product),
    createRowData('Результаты', data.results),
    createRowData('Команда проекта', data.participants),
    createRowData('Руководитель проекта', data.supervisor),
];

const document = await Packer.toBuffer(
    new Document({
        sections: [
            {
                children: [
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 200 },
                        children: [new TextRun({ text: 'ПАСПОРТ ПРОЕКТА', size: 26, bold: true })],
                    }),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 300 },
                        children: [new TextRun({ text: data.theme, size: 24 })],
                    }),
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        columnWidths: [20, 80],
                        rows: tableRows,
                    }),
                ],
            },
        ],
    }),
);

await Bun.file('passport.docx').write(document);
