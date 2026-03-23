using OfficeOpenXml;
using gool_4337Group_26.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace gool_4337Group_26.Services
{
    public class ExcelService
    {
        // Статический конструктор для установки лицензии один раз
        static ExcelService()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public List<Client> Import(string path)
        {
            if (!File.Exists(path))
                throw new FileNotFoundException($"Файл не найден: {path}");

            var list = new List<Client>();

            using (var package = new ExcelPackage(new FileInfo(path)))
            {
                var sheet = package.Workbook.Worksheets.FirstOrDefault();
                if (sheet == null || sheet.Dimension == null)
                    return list;

                // Начинаем со строки 2 (пропускаем заголовок)
                for (int i = 2; i <= sheet.Dimension.Rows; i++)
                {
                    // Проверяем что строка не пустая
                    if (string.IsNullOrWhiteSpace(sheet.Cells[i, 1].Text))
                        continue;

                    DateTime birthDate;
                    var dateText = sheet.Cells[i, 3].Text;

                    // Поддержка разных форматов даты
                    if (!DateTime.TryParse(dateText, out birthDate))
                    {
                        // Попробуем распарсить как число Excel (если дата в числовом формате)
                        double excelDate;
                        if (double.TryParse(dateText, out excelDate))
                            birthDate = DateTime.FromOADate(excelDate);
                        else
                            continue; // Пропускаем строку с невалидной датой
                    }

                    list.Add(new Client
                    {
                        FullName = sheet.Cells[i, 1].Text.Trim(),
                        Email = sheet.Cells[i, 2].Text.Trim(),
                        BirthDate = birthDate
                    });
                }
            }
            return list;
        }

        public void Export(string path, List<Client> clients)
        {
            if (clients == null || clients.Count == 0)
                throw new ArgumentException("Список клиентов пуст");

            using (var package = new ExcelPackage())
            {
                var groups = clients.GroupBy(c =>
                {
                    if (c.Age <= 29) return "20-29";
                    if (c.Age <= 39) return "30-39";
                    return "40+";
                }).OrderBy(g => g.Key);

                foreach (var group in groups)
                {
                    var sheet = package.Workbook.Worksheets.Add(group.Key);

                    // Заголовки
                    sheet.Cells[1, 1].Value = "ФИО";
                    sheet.Cells[1, 2].Value = "Email";

                    // Стили заголовков
                    using (var range = sheet.Cells[1, 1, 1, 2])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    }

                    int row = 2;
                    foreach (var c in group.OrderBy(c => c.FullName))
                    {
                        sheet.Cells[row, 1].Value = c.FullName;
                        sheet.Cells[row, 2].Value = c.Email;
                        row++;
                    }

                    // Автоширина колонок
                    sheet.Cells.AutoFitColumns();
                }

                // Добавляем сводный лист
                var summarySheet = package.Workbook.Worksheets.Add("Сводка");
                summarySheet.Cells[1, 1].Value = "Возрастная группа";
                summarySheet.Cells[1, 2].Value = "Количество";

                using (var range = summarySheet.Cells[1, 1, 1, 2])
                {
                    range.Style.Font.Bold = true;
                }

                int summaryRow = 2;
                foreach (var group in groups)
                {
                    summarySheet.Cells[summaryRow, 1].Value = group.Key;
                    summarySheet.Cells[summaryRow, 2].Value = group.Count();
                    summaryRow++;
                }
                summarySheet.Cells.AutoFitColumns();

                package.SaveAs(new FileInfo(path));
            }
        }
    }
}