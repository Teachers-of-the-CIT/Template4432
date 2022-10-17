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
                try
                {
                    db.SaveChanges();
                }
                catch
                {
                    MessageBox.Show("Ошибка с переносом в бд! Попробуйте повторить перенос снова!");
                }
                book.Close();
                app.Quit();
                MessageBox.Show("Импорт данных завершён");
            }
        } 
    }
}
