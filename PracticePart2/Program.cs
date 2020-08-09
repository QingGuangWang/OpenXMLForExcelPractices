using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace PracticePart2
{
    public class Program
    {
        public static void Main(string[] args)
        {
            //构建一个MemoryStream
            var ms = new MemoryStream();

            //创建Workbook, 指定为Excel Workbook (*.xlsx).
            var document = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook);

            //创建WorkbookPart（工作簿）
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            //构建SharedStringTablePart
            var shareStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
            shareStringPart.SharedStringTable = new SharedStringTable(); //创建根元素

            //创建WorksheetPart（工作簿中的工作表）
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            
            //Workbook 下创建Sheets节点, 建立一个子节点Sheet，关联工作表WorksheetPart
            var sheets = document.WorkbookPart.Workbook.AppendChild<Sheets>(
                new Sheets(new Sheet()
                {
                    Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "myFirstSheet"
                }));

            //初始化Worksheet
            InitWorksheet(worksheetPart);

            //创建表头 （序号，学生姓名，学生年龄，学生班级，辅导老师）
            CreateTableHeader(worksheetPart, shareStringPart);

            //创建内容数据
            CreateTableBody(worksheetPart);

            workbookPart.Workbook.Save();
            document.Close();

            //保存到文件
            SaveToFile(ms);
            Console.WriteLine("End.");
        }

        /// <summary>
        /// 初始化工作表
        /// </summary>
        /// <param name="worksheetPart"></param>
        public static void InitWorksheet(WorksheetPart worksheetPart)
        {
            //构建Worksheet根节点，同时追加子节点SheetData
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            //获取Worksheet对象
            var worksheet = worksheetPart.Worksheet;

            //SheetFormatProperties, 设置默认行高度，宽度， 值类型是Double类型。
            var sheetFormatProperties = new SheetFormatProperties()
            {
                DefaultColumnWidth = 15d,
                DefaultRowHeight = 15d
            };

            //插入SheetFormatProperties，插入到SheetData的前面。通过InsertBefore方法，而不是Append
            //顺序不能错误，否则会导致office打开提示错误，所以一般最好提前在一个列表或者数组，放好顺序再一次性加入
            worksheet.InsertBefore(sheetFormatProperties, worksheet.GetFirstChild<SheetData>());

            //初始化列宽 第一列 5 个单位， 第二列~第四列 30个单位
            var columns = new Columns();
            //列，从1开始算起。
            var column1 = new Column
            {
                Min = 1, Max = 1, Width = 5d, CustomWidth = true
            };
            var column2 = new Column
            {
                Min = 2, Max =3, Width = 30d, CustomWidth = true
            };

            columns.Append(column1, column2);

            //插入Column1对象， 它的位置是在SheetFormatProperties 的后面，但是在SheetData的前面。
            //worksheet.Append(columns); //直接追加在后面，office打开提示错误
            worksheet.InsertAfter(columns, worksheet.GetFirstChild<SheetFormatProperties>());

            //最好是前面弄好对象，这里一次性插入，或者初始化时先创建对象。用的时候直接拿出来
            //worksheet.Append(new OpenXmlElement[]
            //{
            //    new SheetFormatProperties(),
            //    new Columns(),
            //    new SheetData()
            //});
        }

        /// <summary>
        /// 创建表头。 （序号，学生姓名，学生年龄，学生班级，辅导老师）
        /// </summary>
        /// <param name="worksheetPart">WorksheetPart 对象</param>
        /// <param name="shareStringPart">SharedStringTablePart 对象</param>
        public static void CreateTableHeader(WorksheetPart worksheetPart, SharedStringTablePart shareStringPart)
        {
            //获取Worksheet对象
            var worksheet = worksheetPart.Worksheet;

            //获取表格的数据对象，SheetData
            var sheetData = worksheet.GetFirstChild<SheetData>();

            //插入第一行数据，作为表头数据 创建 Row 对象，表示一行
            var row = new Row
            {
                //设置行号，从1开始，不是从0
                RowIndex = 1
            };
           
            //Row下面，追加Cell对象
            row.AppendChild(CreateTableHeaderCell("序号", shareStringPart));
            row.AppendChild(CreateTableHeaderCell("学生姓名", shareStringPart));
            row.AppendChild(CreateTableHeaderCell("学生年龄", shareStringPart));
            row.AppendChild(CreateTableHeaderCell("学生班级", shareStringPart));
            row.AppendChild(CreateTableHeaderCell("辅导老师", shareStringPart));

            sheetData.AppendChild(row);
        }

        /// <summary>
        /// 创建表头的单元格
        /// </summary>
        public static Cell CreateTableHeaderCell(string headerStr, SharedStringTablePart shareStringPart)
        {
            //共享字符串表
            var sharedStringTable = shareStringPart.SharedStringTable;

            //把字符串追加到共享
            sharedStringTable.AppendChild(new SharedStringItem(new Text(headerStr)));
            var index = sharedStringTable.ChildElements.Count - 1; //获取索引

            var cell = new Cell
            {
                //设置值，这里的值是引用 共享字符串里面的对应的索引，就是上面添加的SharedStringItem的子元素的位置。
                CellValue = new CellValue(index.ToString()),
                //设置值类型是共享字符串
                DataType = new EnumValue<CellValues>(CellValues.SharedString)
            };

            return cell;
        }

        public static void CreateTableBody(WorksheetPart worksheetPart)
        {
            //获取Worksheet对象
            var worksheet = worksheetPart.Worksheet;

            //获取表格的数据对象，SheetData
            var sheetData = worksheet.GetFirstChild<SheetData>();

            //插入第一行数据，作为表头数据 创建 Row 对象，表示一行
            var row1 = new Row
            {
                RowIndex = 2
            };

            row1.Append(new OpenXmlElement[]
            {
                new Cell()
                {
                    CellValue = new CellValue("1"),
                    DataType = new EnumValue<CellValues>(CellValues.String) 
                },
                new Cell()
                {
                    CellValue = new CellValue("王同学"),
                    DataType = new EnumValue<CellValues>(CellValues.String) 
                },
                new Cell()
                {
                    CellValue = new CellValue("18岁"),
                    DataType = new EnumValue<CellValues>(CellValues.String) 
                },
                new Cell()
                {
                    CellValue = new CellValue("一班"),
                    DataType = new EnumValue<CellValues>(CellValues.String) 
                },
                new Cell()
                {
                    CellValue = new CellValue("林老师"),
                    DataType = new EnumValue<CellValues>(CellValues.String) 
                }
            });

            sheetData.AppendChild(row1);

            var row2 = new Row
            {
                RowIndex = 3
            };

            row2.Append(new OpenXmlElement[]
            {
                new Cell()
                {
                    CellValue = new CellValue("2"),
                    DataType = new EnumValue<CellValues>(CellValues.String) 
                },
                new Cell()
                {
                    CellValue = new CellValue("李同学"),
                    DataType = new EnumValue<CellValues>(CellValues.String) 
                },
                new Cell()
                {
                    CellValue = new CellValue("19岁"),
                    DataType = new EnumValue<CellValues>(CellValues.String) 
                },
                new Cell()
                {
                    CellValue = new CellValue("二班"),
                    DataType = new EnumValue<CellValues>(CellValues.String) 
                },
                new Cell()
                {
                    CellValue = new CellValue("林老师"),
                    DataType = new EnumValue<CellValues>(CellValues.String) 
                }
            });

            sheetData.AppendChild(row2);
        }

        /// <summary>
        /// 保存到文件
        /// </summary>
        public static void SaveToFile(MemoryStream ms)
        {
            //当前运行时路径
            var directoryInfo = new DirectoryInfo(Directory.GetCurrentDirectory());
            var fileName = $@"PracticePart1-{DateTime.Now:yyyyMMddHHmmss}.xlsx";

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
