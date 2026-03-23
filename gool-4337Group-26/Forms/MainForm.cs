using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using gool_4337Group_26.Data;
using gool_4337Group_26.Services;
using OfficeOpenXml;

namespace gool_4337Group_26.Forms
{
    public class MainForm : Form
    {
        AppDbContext db = new AppDbContext();

        public MainForm()
        {
            // Инициализация лицензий ПЕРЕД созданием UI
            InitializeLicenses();

            Text = "ЛР3-4 Вариант 7 - Управление клиентами";
            Width = 450;
            Height = 350;
            StartPosition = FormStartPosition.CenterScreen;

            db.Database.EnsureCreated();

            // Информационная метка
            var lblInfo = new Label()
            {
                Text = $"База данных: {db.Clients.Count()} клиентов",
                Top = 10,
                Left = 100,
                Width = 250,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            };

            var btn1 = new Button() { Text = "Импорт Excel", Top = 50, Width = 250, Left = 100, Height = 35 };
            var btn2 = new Button() { Text = "Экспорт Excel", Top = 95, Width = 250, Left = 100, Height = 35 };
            var btn3 = new Button() { Text = "Импорт JSON", Top = 140, Width = 250, Left = 100, Height = 35 };
            var btn4 = new Button() { Text = "Экспорт Word", Top = 185, Width = 250, Left = 100, Height = 35 };
            var btn5 = new Button() { Text = "Очистить базу данных", Top = 240, Width = 250, Left = 100, Height = 35, BackColor = System.Drawing.Color.LightCoral };

            btn1.Click += (s, e) => ImportExcel(lblInfo);
            btn2.Click += (s, e) => ExportExcel();
            btn3.Click += (s, e) => ImportJson(lblInfo);
            btn4.Click += (s, e) => ExportWord();
            btn5.Click += (s, e) => ClearDatabase(lblInfo);

            Controls.Add(lblInfo);
            Controls.Add(btn1);
            Controls.Add(btn2);
            Controls.Add(btn3);
            Controls.Add(btn4);
            Controls.Add(btn5);
        }

        private void InitializeLicenses()
        {
            // EPPlus - некоммерческая лицензи

            // Xceed Word - бесплатная лицензия (если не сработает, используем пробную)
            try
            {
                Xceed.Words.NET.Licenser.LicenseKey = "FREE-LIMITED-KEY";
            }
            catch
            {
                // Игнорируем ошибку лицензии, документ создастся с водяным знаком
            }
        }

        private void ImportExcel(Label lblInfo)
        {
            try
            {
                using (var dialog = new OpenFileDialog())
                {
                    dialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                    dialog.Title = "Выберите Excel файл для импорта";

                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        var service = new ExcelService();
                        var data = service.Import(dialog.FileName);

                        if (data.Count == 0)
                        {
                            MessageBox.Show("Файл пустой или не содержит данных", "Предупреждение",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        // Проверка на дубликаты по Email
                        var existingEmails = db.Clients.Select(c => c.Email).ToList();
                        var newClients = data.Where(c => !existingEmails.Contains(c.Email)).ToList();
                        var duplicates = data.Count - newClients.Count;

                        db.Clients.AddRange(newClients);
                        db.SaveChanges();

                        lblInfo.Text = $"База данных: {db.Clients.Count()} клиентов";

                        string msg = $"Импортировано: {newClients.Count} клиентов";
                        if (duplicates > 0)
                            msg += $"\nПропущено дубликатов: {duplicates}";

                        MessageBox.Show(msg, "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка импорта Excel:\n{ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportExcel()
        {
            try
            {
                var clients = db.Clients.ToList();
                if (clients.Count == 0)
                {
                    MessageBox.Show("База данных пуста. Нечего экспортировать.", "Предупреждение",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                using (var dialog = new SaveFileDialog())
                {
                    dialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                    dialog.Title = "Сохранить Excel файл";
                    dialog.FileName = "result.xlsx";

                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        var service = new ExcelService();
                        service.Export(dialog.FileName, clients);
                        MessageBox.Show($"Excel экспортирован:\n{dialog.FileName}", "Успех",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка экспорта Excel:\n{ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ImportJson(Label lblInfo)
        {
            try
            {
                using (var dialog = new OpenFileDialog())
                {
                    dialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*";
                    dialog.Title = "Выберите JSON файл для импорта";

                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        var service = new JsonService();
                        var data = service.Import(dialog.FileName);

                        if (data.Count == 0)
                        {
                            MessageBox.Show("Файл пустой или не содержит данных", "Предупреждение",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        // Проверка на дубликаты по Email
                        var existingEmails = db.Clients.Select(c => c.Email).ToList();
                        var newClients = data.Where(c => !existingEmails.Contains(c.Email)).ToList();
                        var duplicates = data.Count - newClients.Count;

                        db.Clients.AddRange(newClients);
                        db.SaveChanges();

                        lblInfo.Text = $"База данных: {db.Clients.Count()} клиентов";

                        string msg = $"Импортировано: {newClients.Count} клиентов";
                        if (duplicates > 0)
                            msg += $"\nПропущено дубликатов: {duplicates}";

                        MessageBox.Show(msg, "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка импорта JSON:\n{ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportWord()
        {
            try
            {
                var clients = db.Clients.ToList();
                if (clients.Count == 0)
                {
                    MessageBox.Show("База данных пуста. Нечего экспортировать.", "Предупреждение",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                using (var dialog = new SaveFileDialog())
                {
                    dialog.Filter = "Word documents (*.docx)|*.docx";
                    dialog.Title = "Сохранить Word документ";
                    dialog.FileName = "result.docx";

                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        var service = new WordService();
                        service.Export(dialog.FileName, clients);
                        MessageBox.Show($"Word экспортирован:\n{dialog.FileName}", "Успех",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка экспорта Word:\n{ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ClearDatabase(Label lblInfo)
        {
            var result = MessageBox.Show("Удалить всех клиентов из базы данных?", "Подтверждение",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                db.Clients.RemoveRange(db.Clients);
                db.SaveChanges();
                lblInfo.Text = $"База данных: {db.Clients.Count()} клиентов";
                MessageBox.Show("База данных очищена", "Готово",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        protected override void OnFormClosing(System.Windows.Forms.FormClosingEventArgs e)
        {
            db?.Dispose();
            base.OnFormClosing(e);
        }
    }
}