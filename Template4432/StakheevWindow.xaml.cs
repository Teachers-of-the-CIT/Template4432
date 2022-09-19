using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Messaging;
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

namespace Template4432
{
    /// <summary>
    /// Логика взаимодействия для StakheevWindow.xaml
    /// </summary>
    public partial class StakheevWindow : Window
    {
        public static StakheevModelContainer db = new StakheevModelContainer();
        public StakheevWindow()
        {
            InitializeComponent();
        }

        private void StakheevWindowClose(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void ImportExcelStakheev(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "файл Excel .xlsx|*.xlsx";
            ofd.Title = "Выберите файл";
          bool? resultdialog=  ofd.ShowDialog();
            if (resultdialog == true)
            {
                Excel.Application app = new Excel.Application();
                Excel.Workbook book = app.Workbooks.Open(ofd.FileName);
                Excel.Worksheet sheet = app.Worksheets.Item[1];
                var lastCell = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//последнюю ячейку
                int lastColumn = (int)lastCell.Column;
                int lastRow = (int)lastCell.Row;
                for (int i = 0; i < lastRow - 1; i++)
                {
                    Employee employee = new Employee();
                    string str = sheet.Cells[1][i + 2].Text;
                    str = str.Remove(0, 3);
                    employee.Id = int.Parse(str);
                    employee.Post = sheet.Cells[2][i + 2].Text;
                    employee.FIO = sheet.Cells[3][i + 2].Text;
                    employee.Login = sheet.Cells[4][i + 2].Text;
                    employee.Password = sheet.Cells[5][i + 2].Text;
                    employee.LastAuth = sheet.Cells[6][i + 2].Text;
                    employee.AuthType = sheet.Cells[7][i + 2].Text;
                    db.EmployeeSet.Add(employee);
                }
                try
                {
                    db.SaveChanges();
                }
                catch
                {
                    MessageBox.Show("Такой ключ уже есть в базе данных");
                }
                book.Close();
                app.Quit();
                MessageBox.Show("Данные были импортированы");
                /* string result = "";
                 foreach (Employee item in users)
                 {
                     result += "Пользователь: " + item.Id + " \n";
                     result += "Должность: " + item.Post + " \n";
                     result += "ФИО: " + item.FIO + " \n";
                     result += "Логин: " + item.Login + " \n";
                     result += "Пароль: " + item.Password + " \n";
                     result += "Последний вход: " + item.LastAuth + " \n";
                     result += "Тип входа: " + item.AuthType + " \n";
                     result += "\n";
                 }
                 MessageBox.Show(result);*/
            }
        }

        private void ExportExcelStakheev(object sender, RoutedEventArgs e)
        {
            ///Export
            Excel.Application app = new Excel.Application();
            Excel.Workbook book = app.Workbooks.Add(Type.Missing);
            app.SheetsInNewWorkbook = 3;
            Excel.Worksheet sheet = app.Worksheets[1];
            //Группировка
            List<Employee> employeelistgrouped = db.EmployeeSet.ToList().OrderBy(p => p.Post).ToList();
            int counter = 0;
            string lastcategory = "";
            int sheets = 0;
            foreach (Employee item in employeelistgrouped)
            {
                if (item.Post != lastcategory)
                {
                    sheet.Columns.AutoFit();
                    counter = 0;
                    sheets++;
                    sheet = app.Worksheets[sheets];
                    sheet.Name = item.Post;
                    sheet.Cells[1][1] = "Код сотрудника";
                    sheet.Cells[2][1] = "ФИО";
                    sheet.Cells[3][1] = "Логин";
                    //Стиль
                    Excel.Range range = sheet.Range[sheet.Cells[1][1], sheet.Cells[7][1]];
                    range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    range.Font.Bold = true;
                    range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDot;
                    range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlDot;
                    range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlDot;
                    range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlDot;
                    range.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlDot;
                    range.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlDot;
                    counter++;
                    sheet.Cells[1][counter + 1] = item.Id;
                    sheet.Cells[2][counter + 1] = item.FIO;
                    sheet.Cells[3][counter + 1] = item.Login;
                    //Стиль
                    Excel.Range range2 = sheet.Range[sheet.Cells[1][counter + 1], sheet.Cells[7][counter + 1]];
                    range2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    range2.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDot;
                    range2.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlDot;
                    range2.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlDot;
                    range2.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlDot;
                    range2.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlDot;
                    range2.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlDot;
                    lastcategory = item.Post;
                }
                else
                {
                    sheet.Cells[1][counter + 1] = item.Id;
                    sheet.Cells[2][counter + 1] = item.FIO;
                    sheet.Cells[3][counter + 1] = item.Login;
                    //Стиль
                    Excel.Range range = sheet.Range[sheet.Cells[1][counter + 1], sheet.Cells[7][counter + 1]];
                    range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDot;
                    range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlDot;
                    range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlDot;
                    range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlDot;
                    range.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlDot;
                    range.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlDot;
                }
                counter++;
            }
            sheet.Columns.AutoFit();
            app.Visible = true;
        }
    }
}
