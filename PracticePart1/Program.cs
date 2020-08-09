using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace PracticePart1
{
    public class Program
    {
        public static void Main(string[] args)
        {
            //当前运行时路径
            var directoryInfo = new DirectoryInfo(Directory.GetCurrentDirectory());
            var fileName = $@"PracticePart1-{DateTime.Now:yyyyMMddHHmmss}.xlsx";

            //文件路径，保存在运行时路径下
            var filepath = Path.Combine(directoryInfo.ToString(), fileName);
            Console.WriteLine($"FilePath: {filepath}");

            //创建SpreadsheetDocument对象，xlsx类型，通过路径
            var spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);

            //通过Stream对象
            //MemoryStream ms = new MemoryStream();
            //SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook);

            //调用AddWorkbookPart, 创建WorkbookPart对象， 创建Workbook对象（相当于XML根元素）关联到WorkbookPart
            var workbookPart = spreadsheetDocument.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            //通过上面的WorkbookPart，创建WorksheetPart对象，创建Worksheet对象（相当于XML根元素）关联到 WorksheetPart
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // 创建Sheets 到 Workbook
            var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            // 创建添加Sheet对象， Id关联 Worksheet， 从而命名工作表的名称
            var sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "myFirstSheet"
            };

            //追加到 Sheets
            sheets.Append(sheet);

            //保存到磁盘
            workbookPart.Workbook.Save();

            // Close the document.
            spreadsheetDocument.Close();
        }
    }
}