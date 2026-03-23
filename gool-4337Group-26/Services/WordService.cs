using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using gool_4337Group_26.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace gool_4337Group_26.Services
{
    public class WordService
    {
        public void Export(string path, List<Client> clients)
        {
            if (clients == null || clients.Count == 0)
                throw new ArgumentException("Список клиентов пуст");

            if (File.Exists(path))
                File.Delete(path);

            using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document();
                var body = new Body();

                // Заголовок
                body.Append(CreateParagraph("Отчет по клиентам", true, 28, JustificationValues.Center));
                body.Append(new Paragraph());

                var groups = clients.GroupBy(c =>
                {
                    if (c.Age <= 29) return "20-29";
                    if (c.Age <= 39) return "30-39";
                    return "40+";
                }).OrderBy(g => g.Key);

                foreach (var group in groups)
                {
                    // Заголовок группы
                    body.Append(CreateParagraph($"Возрастная группа: {group.Key}", true, 24));
                    body.Append(CreateParagraph($"Количество: {group.Count()}", false, 18, JustificationValues.Both, true));
                    body.Append(new Paragraph());

                    // Таблица
                    var table = new Table();
                    var tblProp = new TableProperties(
                        new TableBorders(
                            new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                            new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                            new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                            new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                            new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                            new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 }
                        )
                    );
                    table.Append(tblProp);

                    // Заголовок таблицы
                    var headerRow = new TableRow();
                    headerRow.Append(CreateTableCell("ФИО", true));
                    headerRow.Append(CreateTableCell("Email", true));
                    table.Append(headerRow);

                    foreach (var c in group.OrderBy(c => c.FullName))
                    {
                        var row = new TableRow();
                        row.Append(CreateTableCell(c.FullName));
                        row.Append(CreateTableCell(c.Email));
                        table.Append(row);
                    }

                    body.Append(table);
                    body.Append(new Paragraph());
                    body.Append(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
                }

                // Статистика на последней странице
                body.Append(CreateParagraph("Общая статистика", true, 24));

                var statsTable = new Table();
                statsTable.Append(new TableProperties(
                    new TableBorders(
                        new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                        new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                        new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                        new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                        new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                        new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 }
                    )
                ));

                var statHeader = new TableRow();
                statHeader.Append(CreateTableCell("Группа", true));
                statHeader.Append(CreateTableCell("Количество", true));
                statsTable.Append(statHeader);

                foreach (var group in groups)
                {
                    var row = new TableRow();
                    row.Append(CreateTableCell(group.Key));
                    row.Append(CreateTableCell(group.Count().ToString()));
                    statsTable.Append(row);
                }

                var totalRow = new TableRow();
                totalRow.Append(CreateTableCell("ВСЕГО", true));
                totalRow.Append(CreateTableCell(clients.Count.ToString(), true));
                statsTable.Append(totalRow);

                body.Append(statsTable);

                mainPart.Document.Append(body);
                mainPart.Document.Save();
            }
        }

        private Paragraph CreateParagraph(string text, bool bold, int fontSize, JustificationValues? justification = null, bool italic = false)
        {
            var para = new Paragraph();
            var run = new Run();
            var props = new RunProperties();

            if (bold) props.Append(new Bold());
            if (italic) props.Append(new Italic());
            props.Append(new FontSize { Val = fontSize.ToString() });

            run.Append(props);
            run.Append(new Text(text));
            para.Append(run);

            if (justification.HasValue)
            {
                para.Append(new ParagraphProperties(new Justification { Val = justification.Value }));
            }

            return para;
        }

        private TableCell CreateTableCell(string text, bool bold = false)
        {
            var cell = new TableCell();
            var para = new Paragraph();
            var run = new Run();

            if (bold) run.Append(new RunProperties(new Bold()));
            run.Append(new Text(text));
            para.Append(run);
            cell.Append(para);

            return cell;
        }
    }
}