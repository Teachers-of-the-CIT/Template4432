using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Template4432
{
    /// <summary>
    /// Логика взаимодействия для _4432_Naumkina.xaml
    /// </summary>
    public partial class _4432_Naumkina : System.Windows.Window
    {
        public _4432_Naumkina()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls; *.xlsx",
                Title = "Выберите файл"
            };
            if (!(ofd.ShowDialog() == true))
            {
                return;
            }
            string[,] list;
            Microsoft.Office.Interop.Excel.Application Excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = Excel.Workbooks.Open(ofd.FileName);
            Worksheet worksheet = workbook.Sheets[1];
            var lastCell = worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);
            int lastCol = lastCell.Column;
            int lastRow = 51;
            list = new string[lastRow, lastCol];
            for (int j = 0; j < lastCol; j++)
            {
                for (int i = 0; i < lastRow; i++)
                {
                    list[i, j] = worksheet.Cells[i + 1, j + 1].Text;
                }
            }
            workbook.Close(false, Type.Missing, Type.Missing);
            Excel.Quit();
            GC.Collect();
            using (Entities db = new Entities())
            {
                for (int i = 1; i < lastRow; i++)
                {
                    OrderSet o = new OrderSet() { Code = list[i, 1], Date = list[i, 2], Time = list[i, 3], ClientCode = list[i, 4], Services = list[i, 5], Status = list[i, 6], RentTime = list[i, 8] };
                    if (list[i, 7] == "")
                    {
                        o.ClosingDate = null;
                    }
                    else
                    {
                        o.ClosingDate = list[i, 7];
                    }
                    db.OrderSet.Add(o);
                }
                db.SaveChanges();
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            List<OrderSet> orders = new List<OrderSet>();
            using (Entities db = new Entities())
            {
                orders = db.OrderSet.ToList();
            }
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.SheetsInNewWorkbook = orders.Count();
            Workbook workbook = app.Workbooks.Add(Type.Missing);
            Worksheet[] worksheets = new Worksheet[3];
            for (int i = 0; i < 3; i++)
            {
                worksheets[i] = app.Worksheets.Item[i + 1];
                worksheets[i].Cells[1][1] = "ID";
                worksheets[i].Cells[2][1] = "Код заказа";
                worksheets[i].Cells[3][1] = "Дата создания";
                worksheets[i].Cells[4][1] = "Код клиента";
                worksheets[i].Cells[5][1] = "Услуги";
            }
            worksheets[0].Name = "Новая";
            worksheets[1].Name = "В прокате";
            worksheets[2].Name = "Закрыта";
            int _new = 2;
            int rent = 2;
            int close = 2;
            foreach (OrderSet order in orders)
            {
                int k = 0;
                int m = 1;
                if (order.Status == "Новая")
                {
                    k = 0;
                    m = _new;
                    _new++;
                }
                else if (order.Status == "В прокате")
                {
                    k = 1;
                    m = rent;
                    rent++;
                }
                else
                {
                    k = 2;
                    m = close;
                    close++;
                }
                worksheets[k].Cells[1][m] = order.ID;
                worksheets[k].Cells[2][m] = order.Code;
                worksheets[k].Cells[3][m] = order.Date;
                worksheets[k].Cells[4][m] = order.ClientCode;
                worksheets[k].Cells[5][m] = order.Services;
            }
            for (int i = 0; i < 3; i++)
            {
                int k = 1;
                if (i == 0)
                {
                    k = _new - 1;
                }
                else if (i == 1)
                {
                    k = rent - 1;
                }
                else
                {
                    k = close - 1;
                }
                Microsoft.Office.Interop.Excel.Range borders = worksheets[i].Range[worksheets[i].Cells[1][1], worksheets[i].Cells[5][k]];
                borders.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = borders.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = borders.Borders[XlBordersIndex.xlEdgeTop].LineStyle = borders.Borders[XlBordersIndex.xlEdgeRight].LineStyle = borders.Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = borders.Borders[XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                worksheets[i].Columns.AutoFit();
            }
            app.Visible = true;
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            string f = File.ReadAllText(@"D:\Документы\Шарага\ИСРПО\Лабораторная работа 3\Данные для импорта.json");
            f = f.Trim('[');
            f = f.Trim(']');
            string[] ordersString = f.Split('}');
            OrderSet[] orders = new OrderSet[ordersString.Length - 1];
            using (Entities db = new Entities())
            {
                for (int i = 0; i < ordersString.Length - 1; i++)
                {
                    ordersString[i] = ordersString[i].Trim(',');
                    ordersString[i] += "}";
                    orders[i] = JsonSerializer.Deserialize<OrderSet>(ordersString[i]);
                    db.OrderSet.Add(orders[i]);
                }
                db.SaveChanges();
            }
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            var app = new Microsoft.Office.Interop.Word.Application();
            Document document = app.Documents.Add();
            Microsoft.Office.Interop.Word.Paragraph[] paragraphs = new Microsoft.Office.Interop.Word.Paragraph[3];
            string[] p = { "Новая", "В прокате", "Закрыта" };
            int[] count = { 0, 0, 0 };
            using (Entities db = new Entities())
            {
                foreach (OrderSet order in db.OrderSet)
                {
                    if (order.Status == "Новая")
                    {
                        count[0]++;
                    }
                    else if (order.Status == "В прокате")
                    {
                        count[1]++;
                    }
                    else
                    {
                        count[2]++;
                    }
                }

            }
            for (int i = 0; i < 3; i++)
            {
                paragraphs[i] = document.Paragraphs.Add();
                Microsoft.Office.Interop.Word.Range range = paragraphs[i].Range;
                range.Text = p[i];
                paragraphs[i].set_Style("Заголовок 1");
                range.InsertParagraphAfter();
                Microsoft.Office.Interop.Word.Paragraph tableParagraph = document.Paragraphs.Add();
                Microsoft.Office.Interop.Word.Range tableRange = tableParagraph.Range;
                Microsoft.Office.Interop.Word.Table table = document.Tables.Add(tableRange, count[i] + 1, 5);
                table.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
                table.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                Microsoft.Office.Interop.Word.Range cellRange;
                cellRange = table.Cell(1, 1).Range;
                cellRange.Text = "Id";
                cellRange = table.Cell(1, 2).Range;
                cellRange.Text = "Код заказа";
                cellRange = table.Cell(1, 3).Range;
                cellRange.Text = "Дата создания";
                cellRange = table.Cell(1, 4).Range;
                cellRange.Text = "Код клиента";
                cellRange = table.Cell(1, 5).Range;
                cellRange.Text = "Услуги";
                table.Rows[1].Range.Bold = 1;
                table.Rows[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                using (Entities db = new Entities())
                {
                    int j = 2;
                    foreach (OrderSet order in db.OrderSet)
                    {
                        if (order.Status == p[i])
                        {
                            cellRange = table.Cell(j, 1).Range;
                            cellRange.Text = order.ID.ToString();
                            cellRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            cellRange = table.Cell(j, 2).Range;
                            cellRange.Text = order.Code.ToString();
                            cellRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            cellRange = table.Cell(j, 3).Range;
                            cellRange.Text = order.Date.ToString();
                            cellRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            cellRange = table.Cell(j, 4).Range;
                            cellRange.Text = order.ClientCode.ToString();
                            cellRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            cellRange = table.Cell(j, 5).Range;
                            cellRange.Text = order.Services.ToString();
                            cellRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            j++;
                        }
                    }
                }
                Microsoft.Office.Interop.Word.Paragraph countOrdersParagraph = document.Paragraphs.Add();
                Microsoft.Office.Interop.Word.Range countOrdersRange = countOrdersParagraph.Range;
                countOrdersRange.Text = "Количество заказов: " + count[i];
                countOrdersRange.Font.Color = WdColor.wdColorDarkRed;
                countOrdersRange.InsertParagraphAfter();
                if (i != 2)
                {
                    document.Words.Last.InsertBreak(WdBreakType.wdPageBreak);
                }
            }
            app.Visible = true;
        }
    }
}
