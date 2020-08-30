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

            //初始化样式
            InitWorkbookStyles(workbookPart);

            //创建WorksheetPart（工作簿中的工作表）
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

            //Workbook 下创建Sheets节点, 建立一个子节点Sheet，关联工作表WorksheetPart
            var sheets = workbookPart.Workbook.AppendChild<Sheets>(
                new Sheets(new Sheet()
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "MergeSheet"
                }));
            
            //构建Worksheet根节点，同时追加子节点SheetData
            worksheetPart.Worksheet = new Worksheet();
            var sheetData = new SheetData();
            var mergeCells = new MergeCells();
            worksheetPart.Worksheet.Append(sheetData, mergeCells);

            //合并A1 - A3
            mergeCells.AppendChild(new MergeCell()
            {
                Reference = new StringValue("A1:A3")
            });
            //合并B1 - B5
            mergeCells.AppendChild(new MergeCell()
            {
                Reference = new StringValue("B1:B5")
            });

            sheetData.AppendChild(new Row(new []
            {
                new Cell()
                {
                    CellValue = new CellValue("A1-A3 合并"),
                    StyleIndex = 1,
                    DataType = new EnumValue<CellValues>(CellValues.String)
                },
                new Cell()
                {
                    CellValue = new CellValue("B1-B5 合并"),
                    StyleIndex = 1,
                    DataType = new EnumValue<CellValues>(CellValues.String)
                },
                new Cell()
                {
                    CellValue = new CellValue("C1"),
                    StyleIndex = 1,
                    DataType = new EnumValue<CellValues>(CellValues.String)
                }
            }));

            //保存
            document.WorkbookPart.Workbook.Save();
            document.Close();

            //保存到文件
            SaveToFile(ms);
            Console.WriteLine("End.");
        }

        /// <summary>
        /// 初始化Excel的样式
        /// </summary>
        public static void InitWorkbookStyles(WorkbookPart workbookPart)
        {
            //WorkbookStylesPart， excel的样式相关
            var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();

            // 创建Stylesheet对象（相当于XML根元素）
            // 一个styleSheet节点下，一般包含：numFmts，fonts，fills，borders，cellStyleXfs，cellXfs，cellStyles，extLst 等节点
            stylesPart.Stylesheet = new Stylesheet
            {
                NumberingFormats = new NumberingFormats(), //numFmts节点 （数字格式化相关）
                Fonts = new Fonts(), //fonts节点 （字体）
                Fills = new Fills(), //fills节点 （填充）
                Borders = new Borders(), //borders节点 （边框）
                CellStyleFormats = new CellStyleFormats(), //cellStyleXfs 节点
                CellFormats = new CellFormats(), //cellXfs 节点
                CellStyles = new CellStyles() //cellStyles 节点
            };


            //（1）初始化填充 Fill
            InitWorkbookStyleFill(stylesPart);
            //（2）初始化字体 Font
            InitWorkbookStyleFont(stylesPart);
            //（3）初始化边框 Border
            InitWorkbookStyleBorder(stylesPart);

            //构建cellStyleXfs及其子元素xf
            stylesPart.Stylesheet.CellStyleFormats = new CellStyleFormats();

            //无填充，字体雅黑，不加粗, 无边框
            var defaultCellFormat = new CellFormat
            {
                FillId = 0,
                FontId = 0,
                BorderId = 0,
                ApplyFill = true,
                ApplyBorder = true,
                ApplyFont = true
            };

            //无填充，字体雅黑，加粗, 有边框
            var cellFormat = new CellFormat(
                new Alignment()
                {
                    Horizontal = HorizontalAlignmentValues.Center
                })
            {
                FillId = 0,
                FontId = 1,
                BorderId = 1,
                ApplyFill = true,
                ApplyBorder = true,
                ApplyFont = true
            };
            
            stylesPart.Stylesheet.CellStyleFormats.Append(defaultCellFormat, cellFormat);
            stylesPart.Stylesheet.CellStyleFormats.Count = 2;

            //构建cellStyle FormatId 关联 cellStyleXfs 的序号
            stylesPart.Stylesheet.CellStyles = new CellStyles();
            stylesPart.Stylesheet.CellStyles.AppendChild(new CellStyle { Name = "常规", FormatId = 0 }); //常规
            stylesPart.Stylesheet.CellStyles.AppendChild(new CellStyle { Name = "自定义", FormatId = 1 }); //自定义
            stylesPart.Stylesheet.CellStyles.Count = 2;

            stylesPart.Stylesheet.CellFormats = new CellFormats();
            stylesPart.Stylesheet.CellFormats.AppendChild(
                new CellFormat
                {
                    FillId = 0, FontId = 0, BorderId = 0, FormatId = 0
                });

            stylesPart.Stylesheet.CellFormats.AppendChild(
                new CellFormat(
                    new Alignment()
                    {
                        Horizontal = HorizontalAlignmentValues.Center,
                        Vertical = VerticalAlignmentValues.Center
                    } )
                {
                    FillId = 0, FontId = 1, BorderId = 1, FormatId = 1
                });
            
            stylesPart.Stylesheet.CellFormats.Count = 2;
        }

        /// <summary>
        /// 初始化填充
        /// </summary>
        public static void InitWorkbookStyleFill(WorkbookStylesPart stylesPart)
        {
            // 一个默认选项none，即不填充， 一个patternType=PatternValues.Gray125的属性值的填充
            //网络资料（未验证）：今天查了资料发现fill属性中默认的前两个为fillid=0的项必须为不填充，fillid=1的项为gray125，自定义项从id=2，也就是第三项开始设置。不这样设定无法正常打开excel，office会提示你必须修复，而修复的结果是将以上两个提到的属性强制添加并覆盖你的属性。
            stylesPart.Stylesheet.Fills.AppendChild(
                new Fill(
                    new PatternFill()
                    {
                        PatternType = PatternValues.None
                    }));

            stylesPart.Stylesheet.Fills.AppendChild(
                new Fill(
                    new PatternFill()
                    {
                        PatternType = PatternValues.Gray125
                    }));
        }

        /// <summary>
        /// 初始化边框样式
        /// </summary>
        public static void InitWorkbookStyleBorder(WorkbookStylesPart stylesPart)
        {
            //创建一个无边框
            var defaultBorder = new Border(
                new LeftBorder(),
                new RightBorder(),
                new TopBorder(),
                new BottomBorder(),
                new DiagonalBorder());

            //创建一个黑色边框
            var blackBorder = new Border();
            blackBorder.Append(
                new LeftBorder(new Color()
                {
                    Rgb = new HexBinaryValue() { Value = "FF000000" }
                })
                {
                    Style = BorderStyleValues.Thin
                },
                new RightBorder(new Color()
                {
                    Rgb = new HexBinaryValue() { Value = "FF000000" }
                })
                {
                    Style = BorderStyleValues.Thin
                },
                new TopBorder(new Color()
                {
                    Rgb = new HexBinaryValue() { Value = "FF000000" }
                })
                {
                    Style = BorderStyleValues.Thin
                },
                new BottomBorder(new Color()
                {
                    Rgb = new HexBinaryValue() { Value = "FF000000" }
                })
                {
                    Style = BorderStyleValues.Thin
                },
                new DiagonalBorder());

            stylesPart.Stylesheet.Borders.AppendChild(defaultBorder);
            stylesPart.Stylesheet.Borders.AppendChild(blackBorder);
            stylesPart.Stylesheet.Borders.Count = 2; 
        }

        /// <summary>
        /// 初始化字体样式
        /// </summary>
        public static void InitWorkbookStyleFont(WorkbookStylesPart stylesPart)
        {
            //字体 Font 黑色微软雅黑，不加粗
            var blackYaHeiFont = new Font();
            blackYaHeiFont.Append(new OpenXmlElement[]{
                new FontSize() { Val = 11 },
                new FontName() { Val = "微软雅黑" },
                new Color()
                {
                    Rgb = new HexBinaryValue()
                    {
                        Value = System.Drawing.ColorTranslator.ToHtml(
                                System.Drawing.Color.FromArgb(
                                    System.Drawing.Color.Black.A,  System.Drawing.Color.Black.R,
                                    System.Drawing.Color.Black.G, System.Drawing.Color.Black.B))
                            .Replace("#", "")
                    }
                }
            });

            //字体 Font 黑色微软雅黑，加粗
            var blackYaHeiBoldFont = new Font();
            blackYaHeiBoldFont.Append(new OpenXmlElement[]{
                new Bold(),  //加粗
                new FontSize() { Val = 11 },
                new FontName() { Val = "微软雅黑" },
                new Color()
                {
                    Rgb = new HexBinaryValue()
                    {
                        Value = System.Drawing.ColorTranslator.ToHtml(
                            System.Drawing.Color.FromArgb(
                                System.Drawing.Color.Black.A,  System.Drawing.Color.Black.R,
                                System.Drawing.Color.Black.G, System.Drawing.Color.Black.B))
                            .Replace("#", "")
                    }
                }
            }); 

            stylesPart.Stylesheet.Fonts.Append(blackYaHeiFont, blackYaHeiBoldFont);
            stylesPart.Stylesheet.Fonts.Count = 2; //这个Count 并不会自己进行统计，因为只是xml文件中的一个属性值，需要自己进行赋值操作。 
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
