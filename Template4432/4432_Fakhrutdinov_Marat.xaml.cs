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
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Xml.Linq;
using System.Text.Json;
using System.Drawing;

namespace Template4432
{
    /// <summary>
    /// Логика взаимодействия для _4432_Fakhrutdinov_Marat.xaml
    /// </summary>
    public partial class _4432_Fakhrutdinov_Marat : System.Windows.Window
    {
        public FakhrutdinovDBEntities db = new FakhrutdinovDBEntities();
        public _4432_Fakhrutdinov_Marat()
        {
            InitializeComponent();
        }

        private void close_window(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        List<ForOrderList> OrderListSmall = new List<ForOrderList>();
        public class ForOrderList
        {
            public int Id;
            public DateTime dateTime1;
            public string Order_code1;
            public string Client_code1;
            public string Services1;
        }

        private void export_to_excel(object sender, RoutedEventArgs e)
        {
            var app = new Excel.Application();
            var book = app.Workbooks.Add(Type.Missing);
            app.SheetsInNewWorkbook = 1;
            var sheet1 = app.Worksheets.Item[1];
            sheet1.Name = "Отсортированные заказы по дате";
            sheet1.Cells[1][1] = "Id";
            sheet1.Cells[2][1] = "Код заказа";
            sheet1.Cells[3][1] = "Код клиента";
            sheet1.Cells[4][1] = "Услуги";
            string FullDate = null;

            foreach (Order order in db.Order)
            {
                FullDate = order.Order_date + " " + order.Order_time;
                DateTime dateTime = Convert.ToDateTime(FullDate);
                ForOrderList forOrderList = new ForOrderList();
                forOrderList.Id = order.Id;
                forOrderList.dateTime1 = dateTime;
                forOrderList.Order_code1 = order.Order_code;
                forOrderList.Client_code1 = order.Client_code;
                forOrderList.Services1 = order.Services;
                OrderListSmall.Add(forOrderList);
            }
            OrderListSmall = OrderListSmall.ToList().OrderBy(p => p.dateTime1).ToList();

            int startRowIndex = 2;
            foreach (ForOrderList forOrderList1 in OrderListSmall)
            {
                sheet1.Cells[1][startRowIndex] = forOrderList1.Id;
                sheet1.Cells[2][startRowIndex] = forOrderList1.Order_code1;
                sheet1.Cells[3][startRowIndex] = forOrderList1.Client_code1;
                sheet1.Cells[4][startRowIndex] = forOrderList1.Services1;
                startRowIndex++;
            }
            app.Visible = true;
        }

        private void Import_Data(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls; *xlsx";
            ofd.Filter = "файл Excel .xlsx|*.xlsx";
            ofd.Title = "Выбор файла";
            bool? resultdialog = ofd.ShowDialog();
            if (resultdialog == true)
            {
                Excel.Application app = new Excel.Application();
                Excel.Workbook book = app.Workbooks.Open(ofd.FileName);
                Excel.Worksheet sheet = app.Worksheets.Item[1];
                var lastCell = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                int lastColumn = (int)lastCell.Column;
                int lastRow = (int)lastCell.Row;
                for (int i = 0; i < lastRow - 1; i++)
                {
                    if (sheet.Cells[2][i + 2].Text != "")
                    {
                        Order order = new Order();
                        order.Order_code = sheet.Cells[2][i + 2].Text;
                        order.Order_date = sheet.Cells[3][i + 2].Text;
                        order.Order_time = sheet.Cells[4][i + 2].Text;
                        order.Client_code = sheet.Cells[5][i + 2].Text;
                        order.Services = sheet.Cells[6][i + 2].Text;
                        order.Status = sheet.Cells[7][i + 2].Text;
                        order.Closing_date = sheet.Cells[8][i + 2].Text;
                        order.Rental_time = sheet.Cells[9][i + 2].Text;
                        db.Order.Add(order);
                    }
                    else
                    {
                        break;
                    }
                }
                try
                {
                    db.SaveChanges();
                    MessageBox.Show("Импорт данных завершён");
                }
                catch
                {
                    MessageBox.Show("Ошибка с переносом в бд! Попробуйте повторить перенос снова!");
                }
                book.Close();
                app.Quit();               
            }
        }

        public class OrderJSON
        {
            public int Id { get; set; }
            public string CodeOrder { get; set; }
            public string CreateDate { get; set; }
            public string CreateTime { get; set; }
            public string CodeClient { get; set; }
            public string Services { get; set; }
            public string Status { get; set; }
            public string ClosedDate { get; set; }
            public string ProkatTime { get; set; }
        }

        private void Import_from_JSON(object sender, RoutedEventArgs e)
        {          
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == true)
            {               
                FileStream fs = new FileStream(ofd.FileName, FileMode.Open);
                List<OrderJSON> jsonOrders = JsonSerializer.Deserialize<List<OrderJSON>>(fs);
                foreach (OrderJSON orders in jsonOrders)
                {
                    Order order = new Order();
                    order.Order_code = orders.CodeOrder;
                    order.Order_date = orders.CreateDate;
                    order.Order_time = orders.CreateTime;
                    order.Client_code = orders.CodeClient;
                    order.Services = orders.Services;
                    order.Status = orders.Status;
                    order.Closing_date = orders.ClosedDate;
                    order.Rental_time = orders.ProkatTime;
                    db.Order.Add(order);
                }
                try
                {
                    db.SaveChanges();
                    MessageBox.Show("Импорт данных завершён");
                }
                catch
                {
                    MessageBox.Show("Ошибка с переносом в бд! Попробуйте повторить перенос снова!");
                }               
            }
        }

        private void export_to_word(object sender, RoutedEventArgs e)
        {
            var app = new Word.Application();
            Word.Document doc = app.Documents.Add();
            Word.Paragraph paragraph = doc.Paragraphs.Add();
            Word.Range range = paragraph.Range;
            paragraph.set_Style("Заголовок 1");
            range.Text = "ОТСОРТИРОВАННЫЕ ЗАКАЗЫ ПО ДАТЕ СОЗДАНИЯ";
            range.Font.Size = 20;
            range.Font.Name = "Times New Roman";
            range.Font.Color = Word.WdColor.wdColorBlack;
            range.Font.Bold = 2;
            paragraph.Alignment = (WdParagraphAlignment)StringAlignment.Center;
            range.InsertParagraphAfter();
            string FullDate = null;

            foreach (Order order in db.Order)
            {
                FullDate = order.Order_date + " " + order.Order_time;
                DateTime dateTime = Convert.ToDateTime(FullDate);
                ForOrderList forOrderList = new ForOrderList();
                forOrderList.Id = order.Id;
                forOrderList.dateTime1 = dateTime;
                forOrderList.Order_code1 = order.Order_code;
                forOrderList.Client_code1 = order.Client_code;
                forOrderList.Services1 = order.Services;
                OrderListSmall.Add(forOrderList);
            }
            OrderListSmall = OrderListSmall.ToList().OrderBy(p => p.dateTime1).ToList();

            Word.Paragraph table = doc.Paragraphs.Add();
            Word.Range tablerange = table.Range;
            Word.Table employeetable = doc.Tables.Add(tablerange, OrderListSmall.Count + 1, 4);
            employeetable.Borders.InsideLineStyle = employeetable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            employeetable.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            Word.Range cellRange;
            cellRange = employeetable.Cell(1, 1).Range;
            cellRange.Text = "Id";
            cellRange = employeetable.Cell(1, 2).Range;
            cellRange.Text = "Код заказа";
            cellRange = employeetable.Cell(1, 3).Range;
            cellRange.Text = "Код клиента";
            cellRange = employeetable.Cell(1, 4).Range;
            cellRange.Text = "Услуги";
            employeetable.Rows[1].Range.Bold = 1;
            employeetable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //Заполнение
            int i = 1;
            foreach (ForOrderList forOrderList1 in OrderListSmall)
            {
                cellRange = employeetable.Cell(i + 1, 1).Range;
                cellRange.Text = forOrderList1.Id.ToString();
                cellRange.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphCenter;
                cellRange = employeetable.Cell(i + 1, 2).Range;
                cellRange.Text = forOrderList1.Order_code1;
                cellRange.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphCenter;
                cellRange = employeetable.Cell(i + 1, 3).Range;
                cellRange.Text = forOrderList1.Client_code1;
                cellRange.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphCenter;
                cellRange = employeetable.Cell(i + 1, 4).Range;
                cellRange.Text = forOrderList1.Services1;
                cellRange.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphCenter;
                i++;
            }
            app.Visible = true;
        }
    }
}
