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

namespace Template4432
{
    /// <summary>
    /// Логика взаимодействия для _4432_Sharipov.xaml
    /// </summary>
    public partial class _4432_Sharipov : Window
    {
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
    }
}
