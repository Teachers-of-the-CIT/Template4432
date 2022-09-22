using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
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
using System.Text.Json;
using Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.Xml.Linq;

namespace Template4432
{
    /// <summary>
    /// Логика взаимодействия для StakheevWindow.xaml
    /// </summary>
    public partial class StakheevWindow : System.Windows.Window
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
        public class Person
        {
            public int Id { get; set; }
            public string CodeStaff { get; set; }
            public string Position { get; set; }
            public string FullName { get; set; }
            public string Log { get; set; }
            public string Password { get; set; }
            public string LastEnter { get; set; }
            public string TypeEnter { get; set; }
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
            app.SheetsInNewWorkbook = 3;
            Excel.Workbook book = app.Workbooks.Add(Type.Missing);
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
                    sheet = app.Worksheets.Item[sheets];
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

        private void ImportFromJsonStakheev(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == true)
            {
                FileStream fs = new FileStream(ofd.FileName, FileMode.Open);
                List<Person> personlist = JsonSerializer.Deserialize<List<Person>>(fs);
              foreach(Person person in personlist)
                {
                    Employee emp = new Employee();
                    emp.Id = int.Parse(person.CodeStaff.Remove(0, 3));
                    emp.Post = person.Position;
                    emp.Password = person.Password;
                    emp.LastAuth = person.LastEnter;
                    emp.AuthType = person.TypeEnter;
                    emp.FIO = person.FullName;
                    emp.Login = person.Log;
                    db.EmployeeSet.Add(emp);
                }
                try
                {
                    db.SaveChanges();
                }
                catch
                {
                    MessageBox.Show("Пользователя с таким ключём уже есть в базе данных");
                }
            }
        }

        private void ExportToWordStakheev(object sender, RoutedEventArgs e)
        {
            ///Export
            List<Employee> employees = db.EmployeeSet.ToList().OrderBy(p => p.AuthType).ToList();
            var app = new Word.Application();
            Word.Document doc = app.Documents.Add();
                    Word.Paragraph paragraph =
                    doc.Paragraphs.Add();
                    Word.Range range = paragraph.Range;
            paragraph.set_Style("Заголовок 1");
            range.InsertParagraphAfter();

            var employeeTypeSuccess = employees.Where(p => p.AuthType == "Успешно").ToList();
            var employeeTypeFail = employees.Where(p => p.AuthType == "Неуспешно").ToList();
            Word.Paragraph table = doc.Paragraphs.Add();
            Word.Range tablerange = table.Range;
            Word.Table employeetable = doc.Tables.Add(tablerange, employeeTypeSuccess.Count + 1, 3);
            employeetable.Borders.InsideLineStyle = employeetable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            employeetable.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            Word.Range cellRange;
            cellRange = employeetable.Cell(1, 1).Range;
            cellRange.Text = "Код";
            cellRange = employeetable.Cell(1, 2).Range;
            cellRange.Text = "Должность";
            cellRange = employeetable.Cell(1, 3).Range;
            cellRange.Text = "Логин";
            employeetable.Rows[1].Range.Bold = 1;
            employeetable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //Заполнение
            int i = 1;
            foreach (var item in employeeTypeSuccess)
            {
                cellRange = employeetable.Cell(i + 1, 1).Range;
                cellRange.Text = item.Id.ToString();
                cellRange.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphCenter;
                cellRange = employeetable.Cell(i + 1, 2).Range;
                cellRange.Text = item.Post;
                cellRange.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphCenter;
                cellRange = employeetable.Cell(i + 1, 3).Range;
                cellRange.Text = item.Login;
                cellRange.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphCenter;
                i++;
            }

            Word.Paragraph par = doc.Paragraphs.Add();
            Word.Range ran = par.Range;
            ran.Text = "Всего работников в успешным входом: " + employeeTypeSuccess.Count;
            ran.Font.Color = Word.WdColor.wdColorDarkRed;
            ran.InsertParagraphAfter();
            ran.Font.Size = 26;
            doc.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);

            Word.Paragraph table2 = doc.Paragraphs.Add();
            Word.Range tablerange2 = table2.Range;
            Word.Table employeetable2 = doc.Tables.Add(tablerange2, employeeTypeFail.Count + 1, 3);
            employeetable2.Borders.InsideLineStyle = employeetable2.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            employeetable2.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            Word.Range cellRange2;
            cellRange2 = employeetable2.Cell(1, 1).Range;
            cellRange2.Text = "Код";
            cellRange2 = employeetable2.Cell(1, 2).Range;
            cellRange2.Text = "Должность";
            cellRange2 = employeetable2.Cell(1, 3).Range;
            cellRange2.Text = "Логин";
            employeetable2.Rows[1].Range.Bold = 1;
            employeetable2.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //Заполнение
            i = 1;
            foreach (var item in employeeTypeFail)
            {
                cellRange2 = employeetable2.Cell(i + 1, 1).Range;
                cellRange2.Text = item.Id.ToString();
                cellRange2.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphCenter;
                cellRange2 = employeetable2.Cell(i + 1, 2).Range;
                cellRange2.Text = item.Post;
                cellRange2.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphCenter;
                cellRange2 = employeetable2.Cell(i + 1, 3).Range;
                cellRange2.Text = item.Login;
                cellRange2.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphCenter;
                i++;
            }

            Word.Paragraph par2 = doc.Paragraphs.Add();
            Word.Range ran2 = par2.Range;
            ran2.Text = "Всего работников с неуспешным входом: " + employeeTypeFail.Count;
            ran2.Font.Color = Word.WdColor.wdColorDarkRed;
            ran2.InsertParagraphAfter();
            ran2.Font.Size = 26;
            doc.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
            app.Visible = true;
        }
    }
}
