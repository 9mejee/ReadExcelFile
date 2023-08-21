using Microsoft.Office.Interop.Excel;
namespace ReadFileFromExcel
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            Application excelApp = new Application();
            Workbook excelBook = excelApp.Workbooks.Open(@"G:\document-table.xlsx");
            _Worksheet excelSheet = excelBook.Sheets[1];
            
            int rowCount = excelSheet.UsedRange.Rows.Count;
            int colCount = excelSheet.UsedRange.Columns.Count;

            for (int i = 1; i <= rowCount; i++)
            {
                var v1 = excelSheet.UsedRange.Cells[i, 1];
                string id = v1.Value2.ToString();
            }
        }
    }
}
