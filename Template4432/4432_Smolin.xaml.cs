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
using Microsoft.Win32;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.Json;
using Microsoft.Office.Interop.Word;

namespace Template4432
{
    
    /// <summary>
    /// Логика взаимодействия для _4432_Smolin.xaml
    /// </summary>
    public partial class _4432_Smolin : System.Windows.Window
    {
        LR2IsrpoEntities db = new LR2IsrpoEntities();

        public _4432_Smolin()
        {
            InitializeComponent();

        }

        private void SmolinImportBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog()
                {
                    DefaultExt = "*.xls;*.xlsx",
                    Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                    Title = "Выберите файл"
                };
                if (!(ofd.ShowDialog() == true))
                    return;

                string[,] list;
                Excel.Application ObjWorkExcel = new Excel.Application();
                Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
                Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
                var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                int _columns = (int)lastCell.Column;
                int _rows = (int)lastCell.Row;
                list = new string[_rows, _columns];
                for (int j = 0; j < _columns; j++)
                    for (int i = 0; i < _rows; i++)
                        list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                ObjWorkBook.Close(false, Type.Missing, Type.Missing);
                ObjWorkExcel.Quit();
                GC.Collect();
                for (int i = 1; i < _rows; i++)
                {
                    db.Worker.Add(new Worker()
                    {
                        WorkerID = int.Parse(list[i, 0]),
                        Post = list[i, 1],
                        Full_name = list[i, 2],
                        Login = list[i, 3],
                        Password = list[i, 4],
                        Last_authorization = list[i, 5],
                        Type_authorization = list[i, 6],
                    });
                }
                db.SaveChanges();
                MessageBox.Show("Данные успешно ипортированы.");
            }
            catch
            {
                MessageBox.Show("Произошла ошибка");

            }
        }

        private void SmolinExportBtn_Click(object sender, RoutedEventArgs e)
        {
            List<Worker> allWorker;
            
                allWorker = db.Worker.ToList().OrderBy(work => work.Type_authorization).ToList();
            
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = 2;
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            var EmployeeEntryTypes = allWorker.OrderBy(o => o.Type_authorization).GroupBy(g => g.Type_authorization).ToDictionary(d => d.Key, d => d.Select(g => new { g.WorkerID, g.Post, g.Login }).ToArray());

            for (int i = 0; i < 2; i++)
            {
                int startRowIndex = 1;
                var worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = $"Тип входа {i + 1}";
                worksheet.Cells[1][startRowIndex] = "Код сотрудника";
                worksheet.Cells[2][startRowIndex] = "Должность";
                worksheet.Cells[3][startRowIndex] = "Логин";
                startRowIndex++;

                
                    List<Worker> workeer;
                    workeer = (from em in db.Worker select em).ToList<Worker>();
                    foreach (var work in workeer)
                    {
                        if (i == 0)
                        {
                            if (work.Type_authorization == "Успешно")
                            {
                                worksheet.Cells[1][startRowIndex] = work.WorkerID;
                                worksheet.Cells[2][startRowIndex] = work.Post;
                                worksheet.Cells[3][startRowIndex] = work.Login;
                                startRowIndex++;
                            }
                        }
                        else if (i == 1)
                        {
                            if (work.Type_authorization == "Неуспешно")
                            {
                                worksheet.Cells[1][startRowIndex] = work.WorkerID;
                                worksheet.Cells[2][startRowIndex] = work.Post;
                                worksheet.Cells[3][startRowIndex] = work.Login;
                                startRowIndex++;
                            }
                        }
                    }
                
                worksheet.Columns.AutoFit();
            }
            app.Visible = true;
        }

        public class Worker1
        {
            public string CodeStaff { get; set; }
            public string Position { get; set; }
            public string FullName { get; set; }
            public string Log { get; set; }
            public string Password { get; set; }
            public string LastEnter { get; set; }
            public string TypeEnter { get; set; }
        }

        private void SmolinImport2_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == true)
            {
                FileStream fs = new FileStream(ofd.FileName, FileMode.Open);
                List<Worker1> personlist = JsonSerializer.Deserialize<List<Worker1>>(fs);
                foreach (Worker1 person in personlist)
                {
                    Worker work = new Worker();
                    work.WorkerID = int.Parse(person.CodeStaff.Remove(0, 3));
                    work.Post = person.Position;
                    work.Password = person.Password;
                    work.Last_authorization = person.LastEnter;
                    work.Type_authorization = person.TypeEnter;
                    work.Full_name = person.FullName;
                    work.Login = person.Log;
                    db.Worker.Add(work);
                  
                    
                }
                try
                {
                    db.SaveChanges();
                    MessageBox.Show("Данные успешно импортированны");
                }
                catch
                {

                    MessageBox.Show("Произошла ошибка");
                }
              
            }
        }

        private void SmolinExport2_Click(object sender, RoutedEventArgs e)
        {
            List<Worker> employees = db.Worker.ToList().OrderBy(p => p.Type_authorization).ToList();
            var app = new Word.Application();
            Word.Document doc = app.Documents.Add();
            Word.Paragraph paragraph =
            doc.Paragraphs.Add();
            Word.Range range = paragraph.Range;
            paragraph.set_Style("Заголовок 1");
            range.InsertParagraphAfter();

            var employeeTypeSuccess = employees.Where(p => p.Type_authorization == "Успешно").ToList();
            var employeeTypeFail = employees.Where(p => p.Type_authorization == "Неуспешно").ToList();
            int Seller1Count= employeeTypeSuccess.Where(p => p.Post == "Продавец").ToList().Count;
            int Seller2Count = employeeTypeFail.Where(p => p.Post == "Продавец").ToList().Count;
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
                cellRange.Text = item.WorkerID.ToString();
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
            ran.Text = "Всего продавцов с успешным входом: " + Seller1Count;
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
                cellRange2.Text = item.WorkerID.ToString();
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
            ran2.Text = "Всего продавцов с неуспешным входом: " + Seller2Count;
            ran2.InsertParagraphAfter();
            ran2.Font.Size = 26;
            doc.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
            app.Visible = true;

        }

       
    }
}
