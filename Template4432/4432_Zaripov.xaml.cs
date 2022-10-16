using Microsoft.Win32;
using System.Text.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
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
using Newtonsoft.Json.Serialization;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;


namespace Template4432
{
    /// <summary>
    /// Логика взаимодействия для _4432_Zaripov.xaml
    /// </summary>
    public partial class _4432_Zaripov : Window
    {

        public _4432_Zaripov()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private async void ButtonImpJSON_Click(object sender, RoutedEventArgs e)
        {
            var text = "";
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Filter = "файл JSON(Spisok.json) | *.json",
                Title = "Выберите файл базы данных"
            };

            if (!(ofd.ShowDialog() == true))
                return;

            var path = ofd.FileName;

            using (StreamReader reader = new StreamReader(path))
            {
                text = await reader.ReadToEndAsync();
            }         
            using (FileStream fs = new FileStream(path, FileMode.Open))
            {
                List<tableIsrpo3> users = await JsonSerializer.DeserializeAsync<List<tableIsrpo3>>(fs);
                using (JSONEntities us = new JSONEntities())
                {
                    foreach(var person in users)
                    {
                        us.tableIsrpo3.Add(new tableIsrpo3()
                        {
                            Id = person.Id,
                            Position = person.Position,
                            FullName = person.FullName,
                            Log = person.Log,
                            Password = person.Password
                        });
                        us.SaveChanges();
                    }                                     
                }
                MessageBox.Show("Успешно!");

            }
        }

        private void ButtonImp_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel(Spisok.xlsx) | *.xlsx",
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
            for(int j = 0; j < _columns; j++)
            {
                for(int i = 0; i < _rows; i++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;                  
                }               
            }
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();
            using (LR2ISRPOEntities usersEntities = new LR2ISRPOEntities())
            {
                var countItems = usersEntities.tableIsrpo2.ToList().Count();
                for (int i = 1; i < _rows; i++)
                {
                    usersEntities.tableIsrpo2.Add(new tableIsrpo2()
                    {
                        id = ++countItems,
                        role_person = list[i,0],
                        fullName = list[i,1],
                        loginPerson = list[i,2],
                        passwordPerson = list[i,3]                     
                    });
                }
                usersEntities.SaveChanges();
            }
            MessageBox.Show("Успешно!");
        }

        public string HashPassword(string password)
        {
            byte[] salt;
            byte[] buffer2;
            if (password == null)
            {
                throw new ArgumentNullException("password");
            }
            using (Rfc2898DeriveBytes bytes = new Rfc2898DeriveBytes(password, 0x10, 0x3e8))
            {
                salt = bytes.Salt;
                buffer2 = bytes.GetBytes(0x20);
            }
            byte[] dst = new byte[0x31];
            Buffer.BlockCopy(salt, 0, dst, 1, 0x10);
            Buffer.BlockCopy(buffer2, 0, dst, 0x11, 0x20);
            return Convert.ToBase64String(dst);
        }

        private void ButtonIksWord_Click(object sender, RoutedEventArgs e)
        {
            List<tableIsrpo3> allItems;
            List<string> allItemsGroupRole;
            using(JSONEntities ent = new JSONEntities())
            {
                allItems = ent.tableIsrpo3.ToList().OrderBy(f=>f.Position).ToList();
                allItemsGroupRole = allItems.Select(f => f.Position).Distinct().ToList();
                var grouping = allItems.GroupBy(f => f.Position).ToList();
                var app = new Word.Application();
                Word.Document document = app.Documents.Add();
                foreach(var group in allItemsGroupRole)
                {
                    var groupObjCount = grouping.Find(f => f.Key == group).Select(g=>g.FullName).Count();
                    Word.Paragraph paragraph = document.Paragraphs.Add();
                    Word.Range range = paragraph.Range;
                    range.Text = group;
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();
                   
                    Word.Paragraph tableParagraph = document.Paragraphs.Add();
                    Word.Range tableRange = tableParagraph.Range;
                    Word.Table studentsTable = document.Tables.Add(tableRange, groupObjCount + 1, 3);
                    
                    studentsTable.Borders.InsideLineStyle = studentsTable.Borders.OutsideLineStyle =
                    Word.WdLineStyle.wdLineStyleSingle;
                    
                    Word.Range cellRange;
                    studentsTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                   
                    cellRange = studentsTable.Cell(1, 1).Range;
                    cellRange.Text = "Логин";
                    cellRange = studentsTable.Cell(1, 2).Range;
                    cellRange.Text = "Пароль";
                    cellRange = studentsTable.Cell(1, 3).Range;
                    cellRange.Text = "Хеш пароля";
                    studentsTable.Rows[1].Range.Bold = 1;
                    studentsTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    int i = 1;
                    
                    foreach(var cur in allItems)
                    {
                        if (cur.Position == group)
                        {
                            cellRange = studentsTable.Cell(i + 1, 1).Range;
                            cellRange.Text = cur.Log.ToString();
                            cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            cellRange = studentsTable.Cell(i + 1, 2).Range;
                            cellRange.Text = cur.Password;
                            cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            cellRange = studentsTable.Cell(i + 1, 3).Range;
                            cellRange.Text = HashPassword(cur.Log);                            
                            cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            i++;
                        }
                        else
                        {
                            continue;
                        }
                    }                    
                }
                app.Visible = true;
            }
        }

        private void ButtonIks_Click(object sender, RoutedEventArgs e)
        {
            List<tableIsrpo2> allItems;
            List<string> allItemsGroupRole;

            using (LR2ISRPOEntities usersEntities = new LR2ISRPOEntities())
            {
                allItems = usersEntities.tableIsrpo2.ToList().OrderBy(f => f.role_person).ToList();
                allItemsGroupRole = allItems.Select(f=>f.role_person).Distinct().ToList();
                var app = new Excel.Application();
                app.SheetsInNewWorkbook = allItemsGroupRole.Count();
                Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
                for(int i = 0; i < allItemsGroupRole.Count(); i++)
                {
                    int startRowIndex = 1;
                    Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                    worksheet.Name = Convert.ToString(allItemsGroupRole[i]);
                    worksheet.Cells[1][startRowIndex] = "Login";
                    worksheet.Cells[2][startRowIndex] = "Password";
                    worksheet.Cells[3][startRowIndex] = "Hash password";
                    startRowIndex++;
                    foreach(var us in allItems)
                    {
                        if (us.role_person == allItemsGroupRole[i])
                        {
                            Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[2][1]];
                            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            headerRange.Font.Italic = true;
                            worksheet.Cells[1][startRowIndex] = us.loginPerson;
                            worksheet.Cells[2][startRowIndex] = us.passwordPerson;
                            worksheet.Cells[3][startRowIndex] = HashPassword(us.passwordPerson);
                            startRowIndex++;
                        }
                        else
                        {
                            continue;
                        }
                    }
                    Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[3][startRowIndex - 1]];
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                            rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                                    rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                                        rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle =
                                            Excel.XlLineStyle.xlContinuous;
                    worksheet.Columns.AutoFit();                                
                }
                app.Visible = true;            
            }
        }

        

       
    }
}
