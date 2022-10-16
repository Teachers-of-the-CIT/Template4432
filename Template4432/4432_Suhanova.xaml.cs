using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
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
using Word = Microsoft.Office.Interop.Word;

namespace Template4432
{
    /// <summary>
    /// Логика взаимодействия для _4432_Suhanova.xaml
    /// </summary>
    public partial class _4432_Suhanova : Window
    {
        public _4432_Suhanova()
        {
            InitializeComponent();
        }

        private void BnImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true)) return;

            string[,] list;
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
            {
                for (int i = 0; i < _rows; i++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                }
            }
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            using (isrpo_lr2Entities db = new isrpo_lr2Entities())
            {
                for (int i = 1; i < _rows; i++)
                {
                    db.data.Add(new data()
                    {
                        id = int.Parse(list[i, 0]),
                        name_service = list[i, 1],
                        kind_service = list[i, 2],
                        id_service = list[i, 3],
                        cost = int.Parse(list[i, 4])
                    });
                }
                db.SaveChanges();
            }
            MessageBox.Show("Готово!");
        }

        private void BnExport_Click(object sender, RoutedEventArgs e)
        {
            var listData = new List<data>();

            using (isrpo_lr2Entities db = new isrpo_lr2Entities())
            {
                listData = db.data.ToList().OrderBy(x => x.kind_service).OrderBy(x => x.cost).ToList();
            }
            var allKindServise = listData.GroupBy(x => x.kind_service).ToList();

            var app = new Excel.Application();
            app.SheetsInNewWorkbook = allKindServise.Count();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            for (int i = 0; i < allKindServise.Count(); i++) 
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = Convert.ToString(allKindServise[i].Key);
                worksheet.Cells[1][startRowIndex] = "id";
                worksheet.Cells[2][startRowIndex] = "Название услуги";
                worksheet.Cells[3][startRowIndex] = "Стоимость";
                startRowIndex++;
                foreach (var item in listData)
                {
                    if (item.kind_service == allKindServise[i].Key)
                    {
                        Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][1],
                        worksheet.Cells[2][1]];
                        headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        headerRange.Font.Italic = true;
                        worksheet.Cells[1][startRowIndex] = item.id;
                        worksheet.Cells[2][startRowIndex] = item.name_service;
                        worksheet.Cells[3][startRowIndex] = item.cost;
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
                rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                worksheet.Columns.AutoFit();
            }

            app.Visible = true;
        }

        private async void BnImportJson_Click(object sender, RoutedEventArgs e)
        {
            var text = "";
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Filter = "файл JSON(Spisok.json) | *.json",
                Title = "Выберите файл базы данных"
            };

            if (!(ofd.ShowDialog() == true)) return;

            var path = ofd.FileName;

            using (StreamReader reader = new StreamReader(path))
            {
                text = await reader.ReadToEndAsync();
            }

            using (FileStream fs = new FileStream(path, FileMode.Open))
            {
                List<data> table = await JsonSerializer.DeserializeAsync<List<data>>(fs);
                using (isrpo_lr2Entities db = new isrpo_lr2Entities())
                {
                    foreach(var item in table)
                    {
                        db.data.Add(new data() 
                        { 
                            id = item.id,
                            name_service = item.name_service,
                            kind_service = item.kind_service,
                            id_service = item.id_service,
                            cost = item.cost
                        });
                        db.SaveChanges();
                    }
                }
                MessageBox.Show("Готово!");
            }
        }

        private void BnExportWord_Click(object sender, RoutedEventArgs e)
        {
            List<data> allItems;
            List<string> allItemsGroup;
            using (isrpo_lr2Entities db = new isrpo_lr2Entities())
            {
                allItems = db.data.ToList().OrderBy(f => f.kind_service).OrderBy(f => f.cost).ToList();
                allItemsGroup = allItems.Select(f => f.kind_service).Distinct().ToList();
                var grouping = allItems.GroupBy(f => f.kind_service).ToList();
                var app = new Word.Application();
                Word.Document document = app.Documents.Add();
                foreach (var group in allItemsGroup)
                {
                    var groupObjCount = grouping.Find(f => f.Key == group).Select(g => g.name_service).Count();
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
                    cellRange.Text = "id";
                    cellRange = studentsTable.Cell(1, 2).Range;
                    cellRange.Text = "Название услуги";
                    cellRange = studentsTable.Cell(1, 3).Range;
                    cellRange.Text = "Стоимость";
                    studentsTable.Rows[1].Range.Bold = 1;
                    studentsTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    int i = 1;

                    foreach (var cur in allItems)
                    {
                        if (cur.kind_service == group)
                        {
                            cellRange = studentsTable.Cell(i + 1, 1).Range;
                            cellRange.Text = cur.id.ToString();
                            cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            cellRange = studentsTable.Cell(i + 1, 2).Range;
                            cellRange.Text = cur.name_service;
                            cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            cellRange = studentsTable.Cell(i + 1, 3).Range;
                            cellRange.Text = cur.cost.ToString();
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
    }
}
