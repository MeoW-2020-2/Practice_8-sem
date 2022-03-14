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
using Word = Microsoft.Office.Interop.Word;

namespace ReportExport
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Entities _context = new Entities();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void ExelExportBtn_Click(object sender, RoutedEventArgs e)
        {
            var allCustomers = _context.Customers.ToList().OrderBy(p => p.Name).ToList();

            var exсelApp = new Excel.Application();
            exсelApp.SheetsInNewWorkbook = allCustomers.Count();

            Excel.Workbook workbook = exсelApp.Workbooks.Add(Type.Missing);

            for(int i = 0; i < allCustomers.Count; i++)
            {
                int startRowIndex = 1;

                Excel.Worksheet worksheet = exсelApp.Worksheets.Item[i + 1];
                worksheet.Name = allCustomers[i].Name;

                worksheet.Cells[1][startRowIndex] = "Дата продажи";
                worksheet.Cells[2][startRowIndex] = "Наименование товара";
                worksheet.Cells[3][startRowIndex] = "Стоимость, руб";
                worksheet.Cells[4][startRowIndex] = "Количество, шт";
                worksheet.Cells[5][startRowIndex] = "Сумма, руб";

                startRowIndex++;

                var customersCategories = allCustomers[i].Sales.OrderBy(p => p.Date).GroupBy(p => p.Product.ProductType).OrderBy(p => p.Key.ID);

                foreach (var productCategory in customersCategories)
                {
                    Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][startRowIndex], worksheet.Cells[5][startRowIndex]];
                    headerRange.Merge();
                    headerRange.Value = productCategory.Key.Name;
                    headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    headerRange.Font.Italic = true;

                    startRowIndex++;

                    foreach(var sale in productCategory)
                    {
                        worksheet.Cells[1][startRowIndex] = sale.Date;
                        worksheet.Cells[2][startRowIndex] = sale.Product.Name;
                        worksheet.Cells[3][startRowIndex] = sale.Product.Price;
                        worksheet.Cells[4][startRowIndex] = sale.Quantity;

                        worksheet.Cells[5][startRowIndex].Formula = $"=C{startRowIndex}*D{startRowIndex}";

                        worksheet.Cells[3][startRowIndex].NumberFormat =
                            worksheet.Cells[5][startRowIndex].NumberFormat = "# ###,00";

                        startRowIndex++;
                    }

                    Excel.Range sumRange = worksheet.Range[worksheet.Cells[1][startRowIndex], worksheet.Cells[4][startRowIndex]];
                    sumRange.Merge();
                    sumRange.Value = "ИТОГО:";
                    sumRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

                    worksheet.Cells[5][startRowIndex].Formula = $"=SUM(E{startRowIndex - productCategory.Count()}: E{startRowIndex - 1})";
                    sumRange.Font.Bold = worksheet.Cells[5][startRowIndex].Font.Bold = true;
                    worksheet.Cells[5][startRowIndex].NumberFormat = "# ###,00";

                    startRowIndex++;
                }

                Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[5][startRowIndex - 1]];
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;

                worksheet.Columns.AutoFit();
            }

            exсelApp.Visible = true;
        }

        private void WordExportBtn_Click(object sender, RoutedEventArgs e)
        {
            var allCustomers = _context.Customers.ToList().OrderBy(p => p.Name).ToList();
            var allCategories = _context.ProductTypes.ToList();

            var wordApp = new Word.Application();

            Word.Document document = wordApp.Documents.Add();

            foreach(var customer in allCustomers)
            {
                Word.Paragraph customerParagraph = document.Paragraphs.Add();
                Word.Range customerRange = customerParagraph.Range;
                customerRange.Text = customer.Name;
                customerParagraph.set_Style("Заголовок");
                customerRange.InsertParagraphAfter();

                Word.Paragraph tableParagraph = document.Paragraphs.Add();
                Word.Range tableRange = tableParagraph.Range;
                Word.Table salesTable = document.Tables.Add(tableRange, allCategories.Count() + 1, 3);
                salesTable.Borders.InsideLineStyle = salesTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                salesTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                Word.Range cellRange;

                cellRange = salesTable.Cell(1, 1).Range;
                cellRange.Text = "Иконка";
                cellRange = salesTable.Cell(1, 2).Range;
                cellRange.Text = "Категория товара";
                cellRange = salesTable.Cell(1, 3).Range;
                cellRange.Text = "Сумма платежа";

                salesTable.Rows[1].Range.Bold = 1;
                salesTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                for(int i = 0; i < allCategories.Count(); i++)
                {
                    var currentCategory = allCategories[i];

                    salesTable.AllowAutoFit = true;
                    Word.Column firstCol = salesTable.Columns[1];
                    Word.Column lastCol = salesTable.Columns[3];
                    firstCol.AutoFit();
                    lastCol.AutoFit();
                    salesTable.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);

                    cellRange = salesTable.Cell(i + 2, 1).Range;
                    Word.InlineShape imageShape = cellRange.InlineShapes.AddPicture(AppDomain.CurrentDomain.BaseDirectory + "..\\..\\" + currentCategory.Icon);
                    imageShape.Width = imageShape.Height = 40;
                    cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    cellRange = salesTable.Cell(i + 2, 2).Range;
                    cellRange.Text = currentCategory.Name;

                    cellRange = salesTable.Cell(i + 2, 3).Range;
                    cellRange.Text = customer.Sales.ToList().Where(p => p.Product.ProductType == currentCategory).Sum(p => p.Quantity * p.Product.Price).ToString("N2") + " руб.";
                }

                Sale maxSale = customer.Sales.OrderByDescending(p => p.Product.Price * p.Quantity).FirstOrDefault();
                if(maxSale != null)
                {
                    Word.Paragraph maxSaleParagraph = document.Paragraphs.Add();
                    Word.Range maxSaleRange = maxSaleParagraph.Range;
                    maxSaleRange.Text = $"Наибольший платеж - {maxSale.Product.Name} за {(maxSale.Product.Price * maxSale.Quantity).ToString("N2")} руб. от {maxSale.Date.ToString("dd.MM.yyyy")}";
                    maxSaleParagraph.set_Style("Выделенная цитата");
                    maxSaleRange.Font.Color = Word.WdColor.wdColorDarkRed;
                    maxSaleRange.InsertParagraphAfter();
                }

                Sale minSale = customer.Sales.OrderBy(p => p.Product.Price * p.Quantity).FirstOrDefault();
                if (minSale != null)
                {
                    Word.Paragraph minSaleParagraph = document.Paragraphs.Add();
                    Word.Range minSaleRange = minSaleParagraph.Range;
                    minSaleRange.Text = $"Наименьший платеж - {minSale.Product.Name} за {(minSale.Product.Price * minSale.Quantity).ToString("N2")} руб. от {maxSale.Date.ToString("dd.MM.yyyy")}";
                    minSaleParagraph.set_Style("Выделенная цитата");
                    minSaleRange.Font.Color = Word.WdColor.wdColorDarkGreen;
                    minSaleRange.InsertParagraphAfter();
                }

                if(customer != allCustomers.LastOrDefault())
                {
                    document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
                }
            }
            wordApp.Visible = true;
        }
    }
}
