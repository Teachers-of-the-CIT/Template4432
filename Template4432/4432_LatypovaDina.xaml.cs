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
using System.Text.Json;
using System.Text.Json.Serialization;
using System.IO;
using JsonSerializer = System.Text.Json.JsonSerializer;
using Word = Microsoft.Office.Interop.Word;

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
                DefaultExt = "*.json",
                Filter = "Файл Json (*.json)|*.json|Text files (*.txt)|*.txt",
                Title = "Выберите файл"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            FileStream fs = File.OpenRead(ofd.FileName);
            List<Zakaz> zakazi;
            zakazi = JsonSerializer.Deserialize<List<Zakaz>>(fs);

            using (ISRPOEntities3 ie3 = new ISRPOEntities3())
            {
                foreach (var a in zakazi)
                {
                    ie3.Zakaz.Add(a);
                }
                ie3.SaveChanges();
            }

        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            List<Zakaz> zakazes;
            using (ISRPOEntities3 ie3 = new ISRPOEntities3())
            {
                zakazes = ie3.Zakaz.ToList().OrderBy(g => g.Data_zakaza).ToList();
                var group = zakazes
                        .GroupBy(s => s.Data_zakaza);

                var app = new Word.Application();
                Word.Document document = app.Documents.Add();

                int i = 0;
                foreach (var data in group)
                {
                    Word.Paragraph paragraph = document.Paragraphs.Add();
                    Word.Range range = paragraph.Range;
                    range.Text = $"Дата заказа {i + 1}";
                    paragraph.set_Style("История заказа 1");
                    range.InsertParagraphAfter();
                    Word.Paragraph tableParagraph = document.Paragraphs.Add();
                    Word.Range tableRange = tableParagraph.Range;
                    Word.Table newTable = document.Tables.Add(tableRange, data.Count() + 1, 4);
                    newTable.Borders.InsideLineStyle = newTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    newTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                    Word.Range cellRange;
                    cellRange = newTable.Cell(1, 1).Range;
                    cellRange.Text = "Id";
                    cellRange = newTable.Cell(1, 2).Range;
                    cellRange.Text = "Код заказа";
                    cellRange = newTable.Cell(1, 3).Range;
                    cellRange.Text = "Код клиента";
                    cellRange = newTable.Cell(1, 4).Range;
                    cellRange.Text = "Услуги";
                    newTable.Rows[1].Range.Bold = 1;
                    newTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    int k = 1;
                    foreach (var vivod in data)
                    {
                        cellRange = newTable.Cell(k + 1, 1).Range;
                        cellRange.Text = Convert.ToString(vivod.ID);
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = newTable.Cell(k + 1, 2).Range;
                        cellRange.Text = vivod.Kod_zakaza;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = newTable.Cell(k + 1, 3).Range;
                        cellRange.Text = vivod.Kod_klienta;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = newTable.Cell(k + 1, 4).Range;
                        cellRange.Text = vivod.Uslugi;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        k++;
                    }

                    Word.Paragraph newParagraph = document.Paragraphs.Add();
                    Word.Range countEmployeeRange = newParagraph.Range;
                    countEmployeeRange.Font.Color = Word.WdColor.wdColorBlue;
                    countEmployeeRange.InsertParagraphAfter();

                    document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
                    app.Visible = true;
                    document.SaveAs2(@"C:\Users\Латыповлар\LatypovaLR3.docx");
                    document.SaveAs2(@"C:\Users\Латыповлар\LatypovaLR3.pdf", Word.WdExportFormat.wdExportFormatPDF);
                    i++;
                }
            }
        }
    }
    
}
