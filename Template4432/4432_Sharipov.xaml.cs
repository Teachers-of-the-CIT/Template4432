using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Microsoft.Office.Interop;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Text.Json;

namespace Template4432
{
    /// <summary>
    /// Логика взаимодействия для _4432_Sharipov.xaml
    /// </summary>
    public partial class _4432_Sharipov : Window
    {
        private class OrderDTO
        {
            public int Id { get; set; }
            public string CodeOrder { get; set; }
            public string CreateDate { get; set; }
            public string CodeClient { get; set; }
            public string Services { get; set; }
            public string ProkatTime { get; set; }
        }

        private Excel.Application _excel;
        public _4432_Sharipov()
        {
            InitializeComponent();
        }

        private void btnImport_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog()
            {
                DefaultExt = ".xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл для импорта"
            };
            if (dialog.ShowDialog() != true) {
                return;
            }

            string[,] list;

            _excel = new Excel.Application();
            var workbook = _excel.Workbooks.Open(dialog.FileName);

            var ws = workbook.Sheets[1] as Excel.Worksheet;

            var lastCell = ws.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);

            var columns = lastCell.Column;
            var rows = lastCell.Row;

            list = new string[rows, columns];

            for (var j = 0; j < columns; j++)
            {
                for (var i = 0; i < rows; i++)
                {
                    list[i, j] = ws.Cells[i + 1, j + 1].Text();
                }
            }

            workbook.Close(false, Type.Missing, Type.Missing);
            _excel.Quit();

            using (var db = new ISRPO2Entities())
            {
                for (var i = 1; i < rows; i++)
                {
                    if (list[i, 0] == String.Empty)
                    {
                        continue;
                    }

                    db.Order.Add(new Order()
                    {
                        Id = int.Parse(list[i, 0]),
                        OrderCode = list[i, 1],
                        CreationDate = DateTime.Parse(list[i, 2]),
                        Services = list[i, 5],
                        RentalTime = list[i, 8],
                        ClientCode = int.Parse(list[i, 4]),
                    });
                }
                db.SaveChanges();
            }

            MessageBox.Show("Данные успешно импортированы");
        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            List<IGrouping<string, Order>> ordersByRentalTime;
            _excel = new Excel.Application();

            using (var db = new ISRPO2Entities())
            {
                ordersByRentalTime = db.Order.GroupBy(order => order.RentalTime).ToList();
            }

            // _excel.SheetsInNewWorkbook = ordersByRentalTime.Count;
            var workbook = _excel.Workbooks.Add();

            foreach (var rentalTime in ordersByRentalTime)
            {
                var ws = workbook.Worksheets.Add() as Excel.Worksheet;
                ws.Name = rentalTime.Key.ToString();

                ws.Cells[1, 1] = "Id";
                ws.Cells[1, 2] = "Код заказа";
                ws.Cells[1, 3] = "Дата создания";
                ws.Cells[1, 4] = "Код клиента";
                ws.Cells[1, 5] = "Услуги";

                ws.Cells[1, 1].Font.Bold = true;
                ws.Cells[1, 2].Font.Bold = true;
                ws.Cells[1, 3].Font.Bold = true;
                ws.Cells[1, 4].Font.Bold = true;
                ws.Cells[1, 5].Font.Bold = true;

                var i = 2;
                foreach (var order in rentalTime.ToList())
                {
                    ws.Cells[i, 1] = order.Id.ToString();
                    ws.Cells[i, 2] = order.OrderCode;
                    ws.Cells[i, 3] = order.CreationDate.ToString();
                    ws.Cells[i, 4] = order.ClientCode.ToString();
                    ws.Cells[i, 5] = order.Services;
                    i++;
                }

                ws.Columns.AutoFit();

            }
            _excel.Visible = true;

        }

        private void btnImportJson_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Filter = "файл Json (Spisok.json)|*.json",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;

            List<OrderDTO> orders;
            using (var fs = new FileStream(ofd.FileName, FileMode.Open))
            {
                orders = JsonSerializer.Deserialize<List<OrderDTO>>(fs);
            }

            using (var db = new ISRPO2Entities())
            {
                foreach (var order in orders)
                {
                    db.Order.Add(new Order()
                    {
                        Id = order.Id,
                        OrderCode = order.CodeOrder,
                        CreationDate = DateTime.Parse(order.CreateDate),
                        Services = order.Services,
                        RentalTime = order.ProkatTime,
                        ClientCode = int.Parse(order.CodeClient),
                    });
                }
                db.SaveChanges();
            }
            MessageBox.Show("Данные успешно импортированы");
        }

        private void btnExportWord_Click(object sender, RoutedEventArgs e)
        {            
            List<IGrouping<string, Order>> ordersByRentalTime;

            using (var db = new ISRPO2Entities())
            {
                ordersByRentalTime = db.Order.GroupBy(order => order.RentalTime).ToList();
            }

            var word = new Word.Application();
            var document = word.Documents.Add();

            using (var db = new ISRPO2Entities())
            {
                ordersByRentalTime = db.Order.GroupBy(order => order.RentalTime).ToList();
            }

            foreach (var rentalTime in ordersByRentalTime)
            {
                var orders = rentalTime.ToList();
                var paragraph = document.Paragraphs.Add();
                paragraph.set_Style("Заголовок 1");

                var range = paragraph.Range;
                range.Text = rentalTime.Key.ToString();

                range.InsertParagraphAfter();

                #region Таблица
                var tableParagraph = document.Paragraphs.Add();
                var tableRange = tableParagraph.Range;

                var ordersTable = document.Tables.Add(tableRange, orders.Count(), 5);
                ordersTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                ordersTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

                ordersTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                #region Заголовок таблицы
                ordersTable.Rows[1].Range.Bold = 1;
                ordersTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                ordersTable.Cell(1, 1).Range.Text = "Id";
                ordersTable.Cell(1, 2).Range.Text = "Код заказа";
                ordersTable.Cell(1, 3).Range.Text = "Дата создания";
                ordersTable.Cell(1, 4).Range.Text = "Код клиента";
                ordersTable.Cell(1, 5).Range.Text = "Услуги";
                #endregion

                #region Заполнение таблицы
                var row = 2;
                foreach (var order in orders)
                {
                    ordersTable.Cell(row, 1).Range.Text = order.Id.ToString();
                    ordersTable.Cell(row, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    ordersTable.Cell(row, 2).Range.Text = order.OrderCode.ToString();
                    ordersTable.Cell(row, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    ordersTable.Cell(row, 3).Range.Text = order.CreationDate.Date.ToString();
                    ordersTable.Cell(row, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    ordersTable.Cell(row, 4).Range.Text = order.ClientCode.ToString();
                    ordersTable.Cell(row, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    ordersTable.Cell(row, 5).Range.Text = order.Services.ToString();
                    ordersTable.Cell(row, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    row++;
                }
                #endregion
                #endregion

                #region Дополнительная информация
                #region Дата первого заказа
                var firstOrderDate = document.Paragraphs.Add();
                firstOrderDate.Range.Text = $"Дата первого заказа - {orders.First().CreationDate.Date}";
                firstOrderDate.Range.InsertParagraphAfter();
                #endregion

                #region Дата последнего заказа
                var lastOrderDate = document.Paragraphs.Add();
                lastOrderDate.Range.Text = $"Дата первого заказа - {orders.Last().CreationDate.Date}";
                lastOrderDate.Range.InsertParagraphAfter();
                #endregion

                #endregion

                document.Words.Last.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);
            }

            word.Visible = true;
        }
    }
}
