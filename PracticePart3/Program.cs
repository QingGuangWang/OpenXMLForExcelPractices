using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace PracticePart3
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

            //创建Workbook 和初始化其下的 样式，工作表 等部件
            InitSharedStringTablePart(workbookPart);
            
            //创建样式
            InitWorkbookStyles(document.WorkbookPart);

            //创建工作簿内容
            CreateWorkSheet(workbookPart);
            
            //保存
            document.WorkbookPart.Workbook.Save();
            document.Close();

            //保存到文件
            SaveToFile(ms);
            Console.WriteLine("End."); 
        }

        /// <summary>
        /// 初始化 SharedStringTablePart 的样式
        /// </summary>
        public static void InitSharedStringTablePart(WorkbookPart workbookPart)
        {
            //构建SharedStringTablePart
            var shareStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
            shareStringPart.SharedStringTable = new SharedStringTable();
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

            //以下关于 cellStyleXfs，cellXfs，cellStyles， 这三者
            //cellStyleXfs，单元格自定义样式详细定义，里面拥有FontId, fillId, borderId等属性。
            //cellXfs，是单元格的样式，里面也拥有FontId, fillId, borderId等属性, 可以说，他就是CellStyleXfs的克隆，同CellStyleXfs相比，它多了一个xfid属性，表示它对应CellStyleXfs的第几项索引，从0开始。
            //cellStyle, 单元格自定义样式总纲，比如excel文件我们打开，选项框可以选择那些 常规，数字，短日期等定义好的样式。"name"属性表示的是CellStyleXfs中样式的名称，"xfId"属性表示的是CellStyleXfs中"xf"子节点的索引，从0开始。

            //构建cellStyleXfs及其子元素xf
            stylesPart.Stylesheet.CellStyleFormats = new CellStyleFormats();

            //无填充，字体雅黑，不加粗, 无边框
            var defaultCellFormat = new CellFormat
            {
                FillId = 0, FontId = 0, BorderId = 0, ApplyFill = true, ApplyBorder = true, ApplyFont = true
            };

            //无填充，字体雅黑，加粗, 有边框
            var tableHeaderCellFormat = new CellFormat(
                new Alignment()
                {
                    Horizontal = HorizontalAlignmentValues.Center
                })
            {
                FillId = 0, FontId = 1, BorderId = 1, ApplyFill = true, ApplyBorder = true, ApplyFont = true
            };

            //无填充，不加粗，字体雅黑, 有边框
            var tableBodyCellFormat = new CellFormat(
                new Alignment()
                {
                    Horizontal = HorizontalAlignmentValues.Center
                })
            {
                FillId = 0, FontId = 0, BorderId = 1
            };

            stylesPart.Stylesheet.CellStyleFormats.Append (defaultCellFormat, tableHeaderCellFormat, tableBodyCellFormat); 
            stylesPart.Stylesheet.CellStyleFormats.Count = 3;

            //构建cellStyle FormatId 关联 cellStyleXfs 的序号
            stylesPart.Stylesheet.CellStyles = new CellStyles();
            stylesPart.Stylesheet.CellStyles.AppendChild(new CellStyle { Name = "常规", FormatId = 0 }); //常规
            stylesPart.Stylesheet.CellStyles.AppendChild(new CellStyle { Name = "表头", FormatId = 1 }); //自定义
            stylesPart.Stylesheet.CellStyles.AppendChild(new CellStyle { Name = "表数据", FormatId = 2 }); //自定义
            stylesPart.Stylesheet.CellStyles.Count = 3;
            
            stylesPart.Stylesheet.CellFormats = new CellFormats();
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat { FillId = 0, FontId = 0, BorderId = 0, FormatId = 0 });
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat(new Alignment()
            {
                Horizontal = HorizontalAlignmentValues.Center
            }) { FillId = 0, FontId = 1, BorderId = 1, FormatId = 1 });
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat(new Alignment()
            {
                Horizontal = HorizontalAlignmentValues.Center
            }) { FillId = 0, FontId = 0, BorderId = 1, FormatId = 2 });
            stylesPart.Stylesheet.CellFormats.Count = 3;
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
            
            //字体 Font 蓝色微软雅黑，加粗
            var blueYaHeiBoldFont = new Font();
            blueYaHeiBoldFont.Append(new OpenXmlElement[]{
                new Bold(),  //加粗
                new FontSize() { Val = 11 },
                new FontName() { Val = "微软雅黑" },
                new Color()
                {
                    Rgb = new HexBinaryValue()
                    {
                        Value = System.Drawing.ColorTranslator.ToHtml(
                            System.Drawing.Color.FromArgb( 
                                System.Drawing.Color.DodgerBlue.A, System.Drawing.Color.DodgerBlue.R, 
                                System.Drawing.Color.DodgerBlue.G, System.Drawing.Color.DodgerBlue.B))
                            .Replace("#", "")
                    }
                }
            });

            //字体 Font 红色微软雅黑，不加粗
            var redYaHeiFont = new Font();
            redYaHeiFont.Append(new OpenXmlElement[]{
                new FontSize() { Val = 11 },
                new FontName() { Val = "微软雅黑" },
                new Color()
                {
                    Rgb = new HexBinaryValue()
                    {
                        Value = System.Drawing.ColorTranslator.ToHtml(
                                System.Drawing.Color.FromArgb(
                                    System.Drawing.Color.Red.A, System.Drawing.Color.Red.R,
                                    System.Drawing.Color.Red.G, System.Drawing.Color.Red.B))
                            .Replace("#", "")
                    }
                }
            });

            //字体 Font 黄色微软雅黑，不加粗
            var yellowYaHeiFont = new Font();
            yellowYaHeiFont.Append(new OpenXmlElement[]{
                new FontSize() { Val = 11 },
                new FontName() { Val = "微软雅黑" },
                new Color()
                {
                    Rgb = new HexBinaryValue()
                    {
                        Value = System.Drawing.ColorTranslator.ToHtml(
                                System.Drawing.Color.FromArgb(
                                    System.Drawing.Color.Yellow.A, System.Drawing.Color.Yellow.R,
                                    System.Drawing.Color.Yellow.G, System.Drawing.Color.Yellow.B))
                            .Replace("#", "")
                    }
                }
            });

            //字体 Font 黄色微软雅黑，不加粗
            var greenYaHeiFont = new Font();
            greenYaHeiFont.Append(new OpenXmlElement[]{
                new FontSize() { Val = 11 },
                new FontName() { Val = "微软雅黑" },
                new Color()
                {
                    Rgb = new HexBinaryValue()
                    {
                        Value = System.Drawing.ColorTranslator.ToHtml(
                                System.Drawing.Color.FromArgb(
                                    System.Drawing.Color.Green.A, System.Drawing.Color.Green.R,
                                    System.Drawing.Color.Green.G, System.Drawing.Color.Green.B))
                            .Replace("#", "")
                    }
                }
            });

            stylesPart.Stylesheet.Fonts.Append(blackYaHeiFont, blackYaHeiBoldFont, blueYaHeiBoldFont,
                redYaHeiFont, yellowYaHeiFont, greenYaHeiFont);
            stylesPart.Stylesheet.Fonts.Count = 6; //这个Count 并不会自己进行统计，因为只是xml文件中的一个属性值，需要自己进行赋值操作。 
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
            stylesPart.Stylesheet.Borders.Count = 2 ; //这个Count 并不会自己进行统计，因为只是xml文件中的一个属性值，需要自己进行赋值操作。
        }

        /// <summary>
        /// 创建和初始化工作表
        /// </summary>
        /// <param name="workbookPart"></param>
        public static void CreateWorkSheet(WorkbookPart workbookPart)
        {
            //创建WorksheetPart（工作簿中的工作表）
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

            //Workbook 下创建Sheets节点, 建立一个子节点Sheet，关联工作表WorksheetPart
            var sheets = workbookPart.Workbook.AppendChild<Sheets>(
                new Sheets(new Sheet()
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "AttendanceSheet"
                }));

            //初始化工作簿的整体
            InitWorksheet(worksheetPart);

            //创建表头
            CreateTableHeader(worksheetPart, workbookPart.SharedStringTablePart);

            //创建表数据
            CreateTableBody(worksheetPart);
        }

        /// <summary>
        /// 初始化工作表的整体
        /// </summary>
        /// <param name="worksheetPart"></param>
        public static void InitWorksheet(WorksheetPart worksheetPart)
        {
            //构建Worksheet根节点，同时追加子节点SheetData
            worksheetPart.Worksheet = new Worksheet();

            //获取Worksheet对象
            var worksheet = worksheetPart.Worksheet;

            //SheetFormatProperties, 设置默认行高度，宽度， 值类型是Double类型。
            var sheetFormatProperties = new SheetFormatProperties()
            {
                DefaultColumnWidth = 15,
                DefaultRowHeight = 13.5
            };

            //初始化列宽 第一列 5 个单位， 第二列~第四列 30个单位
            var columns = new Columns();
            //列，从1开始算起。
            var column1 = new Column
            {
                Min = 1,
                Max = 1,
                Width = 5d,
                CustomWidth = true
            };
            var column2 = new Column
            {
                Min = 2,
                Max = 3,
                Width = 10d,
                CustomWidth = true
            };

            columns.Append(column1, column2); 

            //插入到工作表
            worksheet.Append(new OpenXmlElement[]
            {
                sheetFormatProperties,
                columns,
                new SheetData()
            });
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
            row.AppendChild(CreateTableHeaderCell("姓名", shareStringPart));
            row.AppendChild(CreateTableHeaderCell("部门", shareStringPart));
            row.AppendChild(CreateTableHeaderCell("出勤天数", shareStringPart));
            row.AppendChild(CreateTableHeaderCell("缺勤天数", shareStringPart));

            //最后一列，是富文本 出勤率(单位: %), 富文本则不是简单的Text类，而是需要用到 Run类，RunProperties类

            var richTextStringItem = new SharedStringItem();
            richTextStringItem.AppendChild(new Run(new Text("出勤率(")));

            richTextStringItem.AppendChild(new Run(new OpenXmlElement[]
            {
                new RunProperties(new OpenXmlElement []
                {
                    new Bold(),  //加粗
                    new FontSize() { Val = 11 },
                    new RunFont() { Val = "微软雅黑" },
                    new Color()
                    {
                        Rgb = new HexBinaryValue()
                        {
                            Value = System.Drawing.ColorTranslator.ToHtml(
                                    System.Drawing.Color.FromArgb(
                                        System.Drawing.Color.DodgerBlue.A, System.Drawing.Color.DodgerBlue.R,
                                        System.Drawing.Color.DodgerBlue.G, System.Drawing.Color.DodgerBlue.B))
                                .Replace("#", "")
                        }
                    }
                }),
                new Text("单位: %")
            }));

            richTextStringItem.AppendChild(new Run(new Text(")")));
            shareStringPart.SharedStringTable.AppendChild(richTextStringItem);

            var cell = new Cell
            {
                //设置值，这里的值是引用 共享字符串里面的对应的索引，就是上面添加的SharedStringItem的子元素的位置。
                CellValue = new CellValue((shareStringPart.SharedStringTable.ChildElements.Count - 1).ToString()),
                StyleIndex = 1, // 表头的样式
                //设置值类型是共享字符串
                DataType = new EnumValue<CellValues>(CellValues.SharedString)
            };

            row.AppendChild(cell);

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
                StyleIndex = 1, // 表头的样式
                //设置值类型是共享字符串
                DataType = new EnumValue<CellValues>(CellValues.SharedString)
            };

            return cell;
        }

        /// <summary>
        /// 创建表头。 （序号，学生姓名，学生年龄，学生班级，辅导老师）
        /// </summary>
        /// <param name="worksheetPart">WorksheetPart 对象</param>
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
                    StyleIndex = 2,
                    DataType = new EnumValue<CellValues>(CellValues.String)
                },
                new Cell()
                {
                    CellValue = new CellValue("张同事"),
                    StyleIndex = 2,
                    DataType = new EnumValue<CellValues>(CellValues.String)
                },
                new Cell()
                {
                    CellValue = new CellValue("技术部"),
                    StyleIndex = 2,
                    DataType = new EnumValue<CellValues>(CellValues.String)
                },
                new Cell()
                {
                    CellValue = new CellValue("19"),
                    StyleIndex = 2,
                    DataType = new EnumValue<CellValues>(CellValues.String)
                },
                new Cell()
                {
                    InlineString = new InlineString(
                        new Run(new OpenXmlElement []
                        {
                            new RunProperties(new OpenXmlElement []
                            {
                                new FontSize() { Val = 11 },
                                new RunFont() { Val = "微软雅黑" },
                                new Color(){
                                    Rgb = new HexBinaryValue()
                                    {
                                        Value = System.Drawing.ColorTranslator.ToHtml(
                                                System.Drawing.Color.FromArgb(
                                                    System.Drawing.Color.Red.A, System.Drawing.Color.Red.R,
                                                    System.Drawing.Color.Red.G, System.Drawing.Color.Red.B))
                                            .Replace("#", "")
                                    }
                                },
                            }),
                            new Text("1")
                        })),
                    StyleIndex = 2,
                    DataType = new EnumValue<CellValues>(CellValues.InlineString)
                },
                new Cell()
                {
                    InlineString = new InlineString(
                        new Run(new OpenXmlElement []
                        {
                            new RunProperties(new OpenXmlElement []
                            {
                                new FontSize() { Val = 11 },
                                new RunFont() { Val = "微软雅黑" },
                                new Color(){
                                    Rgb = new HexBinaryValue()
                                    {
                                        Value = System.Drawing.ColorTranslator.ToHtml(
                                                System.Drawing.Color.FromArgb(
                                                    System.Drawing.Color.Orange.A, System.Drawing.Color.Orange.R,
                                                    System.Drawing.Color.Orange.G, System.Drawing.Color.Orange.B))
                                            .Replace("#", "")
                                    }
                                },
                            }),
                            new Text("95")
                        }),
                        new Run(new RunProperties(new OpenXmlElement []
                            {
                                new FontSize() { Val = 11 },
                                new RunFont() { Val = "微软雅黑" },
                            }),
                            new Text("%"))),
                    StyleIndex = 2,
                    DataType = new EnumValue<CellValues>(CellValues.InlineString)
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
                    StyleIndex = 2,
                    DataType = new EnumValue<CellValues>(CellValues.String)
                },
                new Cell()
                {
                    CellValue = new CellValue("李同事"),
                    StyleIndex = 2,
                    DataType = new EnumValue<CellValues>(CellValues.String)
                },
                new Cell()
                {
                    CellValue = new CellValue("技术部"),
                    StyleIndex = 2,
                    DataType = new EnumValue<CellValues>(CellValues.String)
                },
                new Cell()
                {
                    CellValue = new CellValue("18"),
                    StyleIndex = 2,
                    DataType = new EnumValue<CellValues>(CellValues.String)
                },
                new Cell()
                {
                    InlineString = new InlineString(
                        new Run(new OpenXmlElement []
                        {
                            new RunProperties(new OpenXmlElement []
                            {
                                new FontSize() { Val = 11 },
                                new RunFont() { Val = "微软雅黑" },
                                new Color(){
                                    Rgb = new HexBinaryValue()
                                    {
                                        Value = System.Drawing.ColorTranslator.ToHtml(
                                                System.Drawing.Color.FromArgb(
                                                    System.Drawing.Color.Red.A, System.Drawing.Color.Red.R,
                                                    System.Drawing.Color.Red.G, System.Drawing.Color.Red.B))
                                            .Replace("#", "")
                                    }
                                },
                            }),
                            new Text("2")
                        })),
                    StyleIndex = 2,
                    DataType = new EnumValue<CellValues>(CellValues.InlineString)
                },
                new Cell()
                {
                    InlineString = new InlineString(
                        new Run(new OpenXmlElement []
                        {
                            new RunProperties(new OpenXmlElement []
                            {
                                new FontSize() { Val = 11 },
                                new RunFont() { Val = "微软雅黑" },
                                new Color(){
                                    Rgb = new HexBinaryValue()
                                    {
                                        Value = System.Drawing.ColorTranslator.ToHtml(
                                                System.Drawing.Color.FromArgb(
                                                    System.Drawing.Color.Orange.A, System.Drawing.Color.Orange.R,
                                                    System.Drawing.Color.Orange.G, System.Drawing.Color.Orange.B))
                                            .Replace("#", "")
                                    }
                                },
                            }),
                            new Text("90")
                        }),
                        new Run(new RunProperties(new OpenXmlElement []
                            {
                                new FontSize() { Val = 11 },
                                new RunFont() { Val = "微软雅黑" },
                            }),
                            new Text("%"))),
                    StyleIndex = 2,
                    DataType = new EnumValue<CellValues>(CellValues.InlineString)
                }
            });

            sheetData.AppendChild(row2);

            var row3 = new Row
            {
                RowIndex = 4
            };

            row3.Append(new OpenXmlElement[]
            {
                new Cell()
                {
                    CellValue = new CellValue("3"),
                    StyleIndex = 2,
                    DataType = new EnumValue<CellValues>(CellValues.String)
                },
                new Cell()
                {
                    CellValue = new CellValue("王同事"),
                    StyleIndex = 2,
                    DataType = new EnumValue<CellValues>(CellValues.String)
                },
                new Cell()
                {
                    CellValue = new CellValue("技术部"),
                    StyleIndex = 2,
                    DataType = new EnumValue<CellValues>(CellValues.String)
                },
                new Cell()
                {
                    CellValue = new CellValue("20"),
                    StyleIndex = 2,
                    DataType = new EnumValue<CellValues>(CellValues.String)
                },
                new Cell()
                {
                    CellValue = new CellValue("0"),
                    StyleIndex = 2,
                    DataType = new EnumValue<CellValues>(CellValues.String)
                },
                new Cell()
                {
                    InlineString = new InlineString(
                        new Run(new OpenXmlElement []
                        {
                            new RunProperties(new OpenXmlElement []
                            {
                                new FontSize() { Val = 11 },
                                new RunFont() { Val = "微软雅黑" },
                                new Color(){
                                    Rgb = new HexBinaryValue()
                                    {
                                        Value = System.Drawing.ColorTranslator.ToHtml(
                                                System.Drawing.Color.FromArgb(
                                                    System.Drawing.Color.Green.A, System.Drawing.Color.Green.R,
                                                    System.Drawing.Color.Green.G, System.Drawing.Color.Green.B))
                                            .Replace("#", "")
                                    }
                                },
                            }),
                            new Text("100")
                        }),
                        new Run(new RunProperties(new OpenXmlElement []
                            {
                                new FontSize() { Val = 11 },
                                new RunFont() { Val = "微软雅黑" },
                            }),
                            new Text("%"))),
                    StyleIndex = 2,
                    DataType = new EnumValue<CellValues>(CellValues.InlineString)
                }
            });

            sheetData.AppendChild(row3);

            var row4 = new Row
            {
                RowIndex = 5
            };

            row4.Append(new OpenXmlElement[]
            {
                new Cell()
                {
                    CellValue = new CellValue("4"),
                    StyleIndex = 2,
                    DataType = new EnumValue<CellValues>(CellValues.String)
                },
                new Cell()
                {
                    CellValue = new CellValue("刘同事"),
                    StyleIndex = 2,
                    DataType = new EnumValue<CellValues>(CellValues.String)
                },
                new Cell()
                {
                    CellValue = new CellValue("人力资源部"),
                    StyleIndex = 2,
                    DataType = new EnumValue<CellValues>(CellValues.String)
                },
                new Cell()
                {
                    CellValue = new CellValue("20"),
                    StyleIndex = 2,
                    DataType = new EnumValue<CellValues>(CellValues.String)
                },
                new Cell()
                {
                    CellValue = new CellValue("0"),
                    StyleIndex = 2,
                    DataType = new EnumValue<CellValues>(CellValues.String)
                },
                new Cell()
                {
                    InlineString = new InlineString(
                        new Run(new OpenXmlElement []
                        {
                            new RunProperties(new OpenXmlElement []
                            {
                                new FontSize() { Val = 11 },
                                new RunFont() { Val = "微软雅黑" },
                                new Color(){
                                    Rgb = new HexBinaryValue()
                                    {
                                        Value = System.Drawing.ColorTranslator.ToHtml(
                                                System.Drawing.Color.FromArgb(
                                                    System.Drawing.Color.Green.A, System.Drawing.Color.Green.R,
                                                    System.Drawing.Color.Green.G, System.Drawing.Color.Green.B))
                                            .Replace("#", "")
                                    }
                                },
                            }),
                            new Text("100")
                        }),
                        new Run(new RunProperties(new OpenXmlElement []
                            {
                                new FontSize() { Val = 11 },
                                new RunFont() { Val = "微软雅黑" },
                            }),
                            new Text("%"))),
                    StyleIndex = 2,
                    DataType = new EnumValue<CellValues>(CellValues.InlineString)
                }
            });

            sheetData.AppendChild(row4);
        }

        /// <summary>
        /// 保存到文件
        /// </summary>
        public static void SaveToFile(MemoryStream ms)
        {
            //当前运行时路径
            var directoryInfo = new DirectoryInfo(Directory.GetCurrentDirectory());
            var fileName = $@"PracticePart3-{DateTime.Now:yyyyMMddHHmmss}.xlsx";

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
