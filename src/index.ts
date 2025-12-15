import Bun from 'bun';
import { Document, Packer, Paragraph, Table, WidthType } from 'docx';
import data from '@/data.json' with { type: 'json' };
import {
    createHeading,
    createHyperlink,
    createNumberedItem,
    createParagraph,
    createSubtitle,
    createTableHeader,
    createTableRowSimple,
    createTitle,
} from '@/items';

const document = await Packer.toBuffer(
    new Document({
        sections: [
            {
                children: [
                    createTitle(`Реферат "${data.theme}"`),
                    createSubtitle(data.researchIdea),

                    createHeading('Актуальность проекта'),
                    createParagraph(data.relevance),

                    createHeading('Постановка проблемы'),
                    createParagraph(data.problem),

                    createHeading('Гипотеза'),
                    createParagraph(data.hypothesis),

                    createHeading('Объект исследования'),
                    createParagraph(data.researchObject),
                    createHeading('Предмет исследования'),
                    createParagraph(data.researchSubject),

                    createHeading('Цель проекта'),
                    createParagraph(data.goals),

                    createHeading('Задачи'),
                    ...data.tasks.map((task, i) => createNumberedItem(task, i + 1)),

                    createHeading('Методы исследования'),
                    createParagraph(data.methods),

                    createHeading('Этапы работы'),
                    ...data.stages.map((stage, i) => createNumberedItem(stage, i + 1)),

                    createHeading('Архитектура решения'),
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        rows: [
                            createTableHeader(['Продукт', 'Ожидаемые результаты']),
                            createTableRowSimple([data.product, data.results.join('; ')]),
                        ],
                    }),

                    new Paragraph({ spacing: { before: 200 } }),
                    createHeading('Команда проекта'),
                    createParagraph(data.participants),
                    createHeading('Руководители'),
                    createParagraph(data.supervisor),

                    new Paragraph({ spacing: { before: 200 } }),
                    createHeading('Список литературы'),
                    ...[
                        ...data.resources.legislative,
                        ...data.resources.theoretical,
                        ...data.resources.statistical,
                    ].map((resource, i) => createHyperlink(resource, i + 1)),
                ],
            },
        ],
    }),
);

await Bun.file('referat.docx').write(document);
