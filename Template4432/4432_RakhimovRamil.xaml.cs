using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.IO;
using Template4432.Models;

namespace Template4432
{
    /// <summary>
    /// Логика взаимодействия для _4432_RakhimovRamil.xaml
    /// </summary>
    public partial class _4432_RakhimovRamil : Window
    {
        public _4432_RakhimovRamil()
        {
            InitializeComponent();
        }

        private void Button_Import_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
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

            using(var db = new ISRPODBEntities())
            {
                for (int i = 1; i < _rows; i++)
                {
                    int nullCollumns = 0;
                    for (int j = 0; j < _columns; j++)
                    {
                        if (String.IsNullOrEmpty(list[i, j]))
                            nullCollumns++;
                    }
                    if (nullCollumns == _columns)
                        continue;
                    db.Client.Add(new Client()
                    {
                        FIO = list[i, 0],
                        UserId = Convert.ToInt32(list[i, 1]),
                        BirthDate = list[i, 2],
                        Index = Convert.ToInt32(list[i, 3]),
                        City = list[i, 4],
                        Street = list[i, 5],           
                        House = Convert.ToInt32(list[i, 6]),
                        Apartment = Convert.ToInt32(list[i, 7]),
                        Email = list[i, 8]
                    }); 
                }
                try
                {
                    db.SaveChanges();
                }
                catch{ }
            }
        }

        private void Button_Export_Click(object sender, RoutedEventArgs e)
        {
            List<Client> allClients;
            List<string> allStreets;
            using (var isrpoEntities = new ISRPODBEntities())
            {
                allClients = (from c in isrpoEntities.Client
                              orderby c.FIO
                              select c).ToList();
                allStreets = (from s in allClients
                              group s by s.Street into g
                              select g.Key).ToList();
            }
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = allStreets.Count;
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            for (int i = 0; i < allStreets.Count; i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = allStreets[i];
                worksheet.Cells[1][startRowIndex] = "Код клиента";
                worksheet.Cells[2][startRowIndex] = "ФИО";
                worksheet.Cells[3][startRowIndex] = "E-mail";
                Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[3][1]];
                headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                headerRange.Font.Bold = true;
                startRowIndex++;
                foreach(Client client in allClients)
                {
                    if (client.Street != allStreets[i])
                        continue;
                    worksheet.Cells[1][startRowIndex] = client.UserId;
                    worksheet.Cells[2][startRowIndex] = client.FIO;
                    worksheet.Cells[3][startRowIndex] = client.Email;
                    startRowIndex++;
                }
                Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[3][startRowIndex - 1]];             
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                worksheet.Columns.AutoFit();
            }
            app.Visible = true;
        }

        private async void Button_ImportJSON_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Filter = "Файл Json (*.json)|*.json|Text files (*.txt)|*.txt",
                Title = "Выберите файл"
            };
            if (!(ofd.ShowDialog() == true))
                return;

            var sclientList = new List<JSONClient>();
            var clientList = new List<Client>();
            using(FileStream fs = new FileStream(ofd.FileName, FileMode.OpenOrCreate))
            {
                sclientList = await JsonSerializer.DeserializeAsync<List<JSONClient>>(fs);
            }
            foreach (var sc in sclientList)
            {
                clientList.Add(Client.MakeClient(sc));
            }
            using (var db = new ISRPODBEntities())
            {
                foreach (var client in clientList)
                {
                    db.Client.Add(client);
                }
                db.SaveChanges();
            }
        }

        private void Button_ExportWord_Click(object sender, RoutedEventArgs e)
        {
            List<Client> allClients;
            List<string> allStreets;
            using (var isrpoEntities = new ISRPODBEntities())
            {
                allClients = (from c in isrpoEntities.Client
                              orderby c.FIO
                              select c).ToList();
                allStreets = (from s in allClients
                              group s by s.Street into g
                              select g.Key).ToList();
            }
            var app = new Word.Application();
            Word.Document document = app.Documents.Add();

            foreach (var street in allStreets)
            {
                Word.Paragraph paragraph = document.Paragraphs.Add();
                Word.Range range = paragraph.Range;
                range.Text = street;
                paragraph.set_Style("Заголовок 1");
                range.InsertParagraphAfter();
                Word.Paragraph tableParagraph = document.Paragraphs.Add();
                Word.Range tableRange = tableParagraph.Range;
                Word.Table clientsTable = document.Tables.Add(tableRange, allClients.Where(c => c.Street==street).Count() + 1, 3);
                clientsTable.Borders.InsideLineStyle = clientsTable.Borders.OutsideLineStyle =  Word.WdLineStyle.wdLineStyleSingle;
                clientsTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                Word.Range cellRange;
                cellRange = clientsTable.Cell(1, 1).Range;
                cellRange.Text = "Код клиента";
                cellRange = clientsTable.Cell(1, 2).Range;
                cellRange.Text = "ФИО";
                cellRange = clientsTable.Cell(1, 3).Range;
                cellRange.Text = "Email";
                clientsTable.Rows[1].Range.Bold = 1;
                clientsTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                int i = 1;
                foreach (var currentStudent in allClients.Where(c => c.Street == street))
                {
                    cellRange = clientsTable.Cell(i + 1, 1).Range;
                    cellRange.Text = currentStudent.UserId.ToString();
                    cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    cellRange = clientsTable.Cell(i + 1, 2).Range;
                    cellRange.Text = currentStudent.FIO;
                    cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    cellRange = clientsTable.Cell(i + 1, 3).Range;
                    cellRange.Text = currentStudent.Email;
                    cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    i++;
                }
                document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
            }
            app.Visible = true;
            document.SaveAs2(@"D:\outputFileWord.docx");
        }

        private void Button_ClearDb_Click(object sender, RoutedEventArgs e)
        {
            using (var db = new ISRPODBEntities())
            {
                foreach (var client in db.Client)
                {
                    db.Entry(client).State = System.Data.Entity.EntityState.Deleted;
                }
                db.SaveChanges();
            }
        }
    }
}
