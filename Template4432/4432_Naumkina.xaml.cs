using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
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
            using (Model1Container1 db = new Model1Container1())
            {
                for (int i = 1; i < lastRow; i++)
                {
                    Order o = new Order() { Code = list[i, 1], Date = DateTime.ParseExact(list[i, 2], "d.M.yy", CultureInfo.InvariantCulture), Time = DateTime.ParseExact(list[i, 3], "H:mm", CultureInfo.InvariantCulture), ClientCode = int.Parse(list[i, 4]), Services = list[i, 5], Status = list[i, 6], RentTime = list[i, 8] };
                    if (list[i, 7] == "")
                    {
                        o.ClosingDate = null;
                    }
                    else
                    {
                        o.ClosingDate = DateTime.ParseExact(list[i, 7], "d.M.yy", CultureInfo.InvariantCulture);
                    }
                    db.OrderSet.Add(o);
                }
                db.SaveChanges();
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            List<Order> orders;
            using (Model1Container1 db = new Model1Container1())
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
            foreach (Order order in orders)
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
                Range borders = worksheets[i].Range[worksheets[i].Cells[1][1], worksheets[i].Cells[5][k]];
                borders.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = borders.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = borders.Borders[XlBordersIndex.xlEdgeTop].LineStyle = borders.Borders[XlBordersIndex.xlEdgeRight].LineStyle = borders.Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = borders.Borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
                worksheets[i].Columns.AutoFit();
            }
            app.Visible = true;
        }
    }
}
