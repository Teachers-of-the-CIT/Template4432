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
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
namespace Template4432
{
    /// <summary>
    /// Логика взаимодействия для _4432_LatypovaDina.xaml
    /// </summary>
    public partial class _4432_LatypovaDina : System.Windows.Window
    {
        public _4432_LatypovaDina()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (2.xlsx)|*.xlsx",
                Title = "Выберите файл БД"
            };
            if (!(ofd.ShowDialog() == true))
                return;

            string[,] list;
            Excel.Application objWorkExcel = new Excel.Application();
            Excel.Workbook objWorkBook = objWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet objWorkSheet = (Excel.Worksheet)objWorkBook.Sheets[1];
            var lastCell = objWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
            {
                for (int i = 0; i < _rows; i++)
                {
                    list[i, j] = objWorkSheet.Cells[i + 1, j + 1].Text;
                    objWorkBook.Close(false, Type.Missing, Type.Missing);
                    objWorkExcel.Quit();
                    GC.Collect();
                }

            }
            using (ISRPOEntities iSRPOEntities = new ISRPOEntities())
            {
                iSRPOEntities.SaveChanges();
            }
        }
    }
}
