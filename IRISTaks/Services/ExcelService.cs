using excel = Microsoft.Office.Interop.Excel;

namespace IRISTaks.Services
{
    public class ExcelService
    {
        public static string ConvertLetterToColumn(string request, string filePath)
        {
            excel.Application application = new excel.Application();
            excel.Workbook workbook = application.Workbooks.Open(filePath, false, ReadOnly: true);
            excel._Worksheet worksheet = workbook.Worksheets.get_Item("Sheet1");
            try
            {
                var range = (excel.Range)worksheet.UsedRange;
                var column = range.Columns[1];
                var columnHeader = (System.Array)column.Cells.Value;
                return "";
            }
            catch (System.Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                workbook.Close();
                application.Quit();
            }

        }
    }
}