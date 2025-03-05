using System;
using System.Collections.Generic;
using System.Data.Entity.Validation;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace _4332Project.Students
{
    public partial class Zaripov_4332 : Window
    {
        public Zaripov_4332()
        {
            InitializeComponent();
        }

        private void B_Import_OnClick(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
                {
                    DefaultExt = ".xlsx",
                    Filter = "Xlsx files (*.xlsx)|*.xlsx",
                    Title = "Выберите файл базы данных"
                }
                ;
            if (ofd.ShowDialog() != true)
                return;

            string[,] list;
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook objWorkbook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)objWorkbook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = 51;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
            for (int i = 0; i < _rows; i++)
                list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
            objWorkbook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            using(OCSEntities userEntities = new OCSEntities())
            {
                for(int i = 1; i < _rows; i++)
                {
                    userEntities.Orders.Add(new Orders()
                    {
                        Id = list[i, 0],
                        Order_Id = list[i, 1],
                        Date = list[i, 2],
                        Client_Id = list[i, 4],
                        Services = list[i, 5],
                        Status = list[i, 6]
                    });
                }
                try
                {
                    userEntities.SaveChanges();
                }
                catch
                {
                MessageBox.Show("Ошибка в импорте данных","Ошибка",MessageBoxButton.OK, MessageBoxImage.Error);
                }
                MessageBox.Show("Данные импортированы успешно!", "Внимание!", MessageBoxButton.OK, MessageBoxImage.Information);

            }
        }

        private void B_Export_OnClick(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog()
            {
                DefaultExt = "*.xlsx",
                Filter = "файл Excel (*.xlsx)|*.xlsx",
                Title = "Сохранить файл базы данных"
            };

            if (sfd.ShowDialog() == true)
            {
                Excel.Application ObjWorkExcel = new Excel.Application();
                Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Add();

                try
                {
                    using (OCSEntities dbContext = new OCSEntities())
                    {
                        var orders = dbContext.Orders.ToList();
                        var groupedOrders = orders.GroupBy(c => c.Status);

                        foreach (var group in groupedOrders)
                        {
                            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets.Add();
                            ObjWorkSheet.Name = group.Key;
                            ObjWorkSheet.Cells[1, 1] = "Id";
                            ObjWorkSheet.Cells[1, 2] = "Код заказа";
                            ObjWorkSheet.Cells[1, 3] = "Дата создания";
                            ObjWorkSheet.Cells[1, 4] = "Код клиента";
                            ObjWorkSheet.Cells[1, 5] = "Услуги";
                            ObjWorkSheet.Cells[1, 6] = "Статус";

                            var sortedOrders = group.OrderBy(c => c.Id).ToList();
                            int row = 2;

                            foreach (var order in sortedOrders)
                            {
                                ObjWorkSheet.Cells[row, 1] = order.Id;
                                ObjWorkSheet.Cells[row, 2] = order.Order_Id;
                                ObjWorkSheet.Cells[row, 3] = order.Date;
                                ObjWorkSheet.Cells[row, 4] = order.Client_Id;
                                ObjWorkSheet.Cells[row, 5] = order.Services;
                                ObjWorkSheet.Cells[row, 6] = order.Status;
                                row++;
                            }
                        }
                    }

                    ObjWorkBook.SaveAs(sfd.FileName);
                    MessageBox.Show("Данные успешно экспортированы в Excel!", "Внимание!", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка: {ex.Message}", "Внимание!", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                finally
                {
                    ObjWorkBook.Close(false);
                    ObjWorkExcel.Quit();
                    GC.Collect();
                }
            }
        }

        private async void b_jsonImport_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Filter = "JSON файлы (*.json)|*.json",
                Title = "Выберите JSON-файл"
            };
            if (!(ofd.ShowDialog() == true))
                return;

            using (var fS = new FileStream(ofd.FileName, FileMode.Open, FileAccess.Read))
            {
                var db = new OCSEntities();
                var orders = await JsonSerializer.DeserializeAsync<List<Orders>>(fS);

                if (orders != null)
                {
                    foreach (Orders item in orders)
                    {
                        var order = new Orders
                        {
                            Id = Convert.ToString(item.Id),
                            Client_Id = item.Client_Id,
                            Date = item.Date,
                            Order_Id = item.Order_Id,
                            Services = item.Services,
                            Status = item.Status
                        };

                        db.Orders.Add(order);
                    }

                    try
                    {
                        await db.SaveChangesAsync();
                        MessageBox.Show("Данные импортированы успешно!",
                                        "Внимание!",
                                        MessageBoxButton.OK,
                                        MessageBoxImage.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Данные импортированы с ошибкой! Сообщение об ошибке: {ex.Message}",
                                        "Внимание!",
                                        MessageBoxButton.OK,
                                        MessageBoxImage.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Не удалось десериализовать данные из файла.",
                                    "Ошибка",
                                    MessageBoxButton.OK,
                                    MessageBoxImage.Error);
                }
            }
        }

        private void b_jsonExport_Click(object sender, RoutedEventArgs e)
        {
            var db = new OCSEntities();
            var orders = db.Orders.ToList();

            var newStatus = orders.Where(o => o.Status == "Новая").ToList();
            var inRentStatus = orders.Where(o => o.Status == "В прокате").ToList();
            var closedStatus = orders.Where(o => o.Status == "Закрыта").ToList();

            var app = new Word.Application();
            var document = app.Documents.Add();

            void CreateTable(string title, List<Orders> data)
            {
                var paragraph = document.Paragraphs.Add();
                var range = paragraph.Range;
                range.Text = title;
                range.InsertParagraphAfter();

                var tableParagraph = document.Paragraphs.Add();
                var tableRange = tableParagraph.Range;
                var table = document.Tables.Add(tableRange, data.Count + 1, 6);

                table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot;
                table.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                Word.Range cellRange;
                cellRange = table.Cell(1, 1).Range;
                cellRange.Text = "Id";
                cellRange = table.Cell(1, 2).Range;
                cellRange.Text = "Код Заказа";
                cellRange = table.Cell(1, 3).Range;
                cellRange.Text = "Дата создания";
                cellRange = table.Cell(1, 4).Range;
                cellRange.Text = "Код клиента";
                cellRange = table.Cell(1, 5).Range;
                cellRange.Text = "Услуги";
                cellRange = table.Cell(1, 6).Range;
                cellRange.Text = "Статус";
                table.Rows[1].Range.Bold = 1;
                table.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                int count = 1;
                foreach (var item in data)
                {
                    cellRange = table.Cell(count + 1, 1).Range;
                    cellRange.Text = item.Id;
                    cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    cellRange = table.Cell(count + 1, 2).Range;
                    cellRange.Text = item.Order_Id;
                    cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    cellRange = table.Cell(count + 1, 3).Range;
                    cellRange.Text = item.Date;
                    cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    cellRange = table.Cell(count + 1, 4).Range;
                    cellRange.Text = item.Client_Id;
                    cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    cellRange = table.Cell(count + 1, 5).Range;
                    cellRange.Text = item.Services;
                    cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    cellRange = table.Cell(count + 1, 6).Range;
                    cellRange.Text = item.Status;
                    cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    count++;
                }
            }
            CreateTable("Список новых заказов",newStatus);
            CreateTable("Список заказов в прокате", inRentStatus);
            CreateTable("Список закрытых заказов", closedStatus);
            app.Visible = true;
        }
    }


}