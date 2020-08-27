using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace PracticePart4
{
    public class Program
    {
        public static void Main(string[] args)
        {
            //构建一个MemoryStream
            var ms = new MemoryStream();

            //Excel Workbook (*.xlsx).
            var document = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook);
            //创建WorkbookPart（工作簿）
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            //创建WorksheetPart（工作簿中的工作表）
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

            //Workbook 下创建Sheets节点, 建立一个子节点Sheet，关联工作表WorksheetPart
            var sheets = workbookPart.Workbook.AppendChild<Sheets>(
                new Sheets(new Sheet()
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "MegerSheet"
                }));

            //构建Worksheet根节点，同时追加子节点SheetData
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            


            //保存
            document.WorkbookPart.Workbook.Save();
            document.Close();

            //保存到文件
            SaveToFile(ms);
            Console.WriteLine("End.");
        }



        /// <summary>
        /// 保存到文件
        /// </summary>
        public static void SaveToFile(MemoryStream ms)
        {
            //当前运行时路径
            var directoryInfo = new DirectoryInfo(Directory.GetCurrentDirectory());
            var fileName = $@"PracticePart4-{DateTime.Now:yyyyMMddHHmmss}.xlsx";

            //文件路径，保存在运行时路径下
            var filepath = Path.Combine(directoryInfo.ToString(), fileName);

            var bytes = ms.ToArray();
            var fileStream = new FileStream(filepath, FileMode.Create, FileAccess.Write, FileShare.Read);
            fileStream.Write(bytes, 0, bytes.Length);
            fileStream.Flush();

            Console.WriteLine($"Save Path: {filepath}");
        }
    }
}
