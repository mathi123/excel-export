using System;
using System.Globalization;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelExport
{
    public class ExcelExporter
    {
        public static MemoryStream Generate(SheetConfiguration configuration, ExcelStyle style = null)
        {
            return Generate(new[] {configuration}, style);
        }

        public static MemoryStream Generate(SheetConfiguration[] configurations, ExcelStyle style = null)
        {
            if (style == null) style = new ExcelStyle();
            var stream = new MemoryStream();
            var spreadSheet = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            
            // Add a WorkbookPart to the document.
            var workbookpart = spreadSheet.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorkbookStylesPart to the WorkbookPart
            var stylesPart = workbookpart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = style.Stylesheet;
            stylesPart.Stylesheet.Save();

            // Add Sheets to the Workbook.
            var sheets = spreadSheet.WorkbookPart.Workbook.
                AppendChild(new Sheets());

            foreach (var sheetConfiguration in configurations)
            {
                // Add a WorksheetPart to the WorkbookPart.
                var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();

                // building sheet columns must happen before creating sheet data
                BuildSheetColumns(worksheetPart.Worksheet, sheetConfiguration);

                worksheetPart.Worksheet.Append(new SheetData());

                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                //row start
                sheetData.AppendChild(GetTitleRow(sheetConfiguration, style));

                foreach (var record in sheetConfiguration.Data)
                {
                    var row = new Row();

                    foreach (var column in sheetConfiguration.Columns)
                    {
                        var val = GetRawValue(record, column.PropertyPath);

                        var cell = BuildCell(val, column);

                        if (cell != null)
                        {
                            cell.StyleIndex = style.CellFormatDefaultId;
                            row.AppendChild(cell);
                        }
                    }
                    sheetData.AppendChild<Row>(row);
                }

                // Append a new worksheet and associate it with the workbook.
                var sheet = new Sheet()
                {
                    Id = spreadSheet.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = sheetConfiguration.Name
                };
                sheets.Append(sheet);
            }

            workbookpart.Workbook.Save();
            spreadSheet.Close();

            return stream;
        }

        private static void BuildSheetColumns(Worksheet worksheet, SheetConfiguration configuration)
        {
            Columns columns = new Columns();
            uint i = 1;

            foreach (var column in configuration.Columns)
            {
                var excelColumn = new Column()
                {
                    Min = i,
                    Max = i,
                    CustomWidth = true
                };
                i++;

                if (column.Width.HasValue)
                {
                    var val = column.Width.Value == 0 ? 0 : column.Width.Value / 6.5;
                    excelColumn.Width = new DoubleValue(val);
                }
                else
                {
                    excelColumn.Width = new DoubleValue(10.0);
                }

                columns.Append(excelColumn);
            }

            worksheet.Append(columns);
        }

        private static Row GetTitleRow(SheetConfiguration configuration, ExcelStyle style)
        {
            var row = new Row();
            foreach (var column in configuration.Columns)
            {
                var cell = new Cell()
                {
                    DataType = CellValues.InlineString,
                    InlineString = new InlineString()
                    {
                        Text = new Text(column.Header)
                    },
                    StyleIndex = style.CellFormatTitleRowId
                };
                row.AppendChild(cell);
            }
            return row;
        }

        private static Cell BuildCell(object rawValue, ColumnBase column)
        {
            if (column is TextColumn textColumn)
            {
                return BuildInlineTextCell(rawValue, textColumn);
            }
            if (column is NumberColumn numberColumn)
            {
                return BuildNumberCell(rawValue, numberColumn);
            }
            if (column is DateColumn dateColumn)
            {
                return BuildDateCell(rawValue, dateColumn);
            }
            if (column is BooleanColumn booleanColumn)
            {
                return BuildBooleanCell(rawValue, booleanColumn);
            }

            return BuildErrorCell();
        }

        private static Cell BuildDateCell(object rawValue, DateColumn column)
        {
            if (!(rawValue is DateTime))
            {
                return BuildErrorCell();
            }

            var date = (DateTime) rawValue;

            var cell = new Cell()
            {
                DataType = CellValues.InlineString,
                InlineString = new InlineString()
                {
                    Text = new Text(date.ToString(column.Format))
                }
            };

            return cell;
        }

        private static Cell BuildNumberCell(object rawValue, NumberColumn column)
        {
            if (!(rawValue is short) && !(rawValue is int) && !(rawValue is long) && !(rawValue is double) &&
                !(rawValue is decimal))
            {
                return BuildErrorCell();
            }

            var asDecimal = Convert.ToDecimal(rawValue);

            if (column.ShouldRound)
            {
                asDecimal = Math.Round(asDecimal, column.Round);
            }

            // Todo: format currency sign

            var cell = new Cell()
            {
                DataType = CellValues.Number,
                CellValue = new CellValue(asDecimal.ToString(CultureInfo.InvariantCulture))
            };

            return cell;
        }

        private static Cell BuildBooleanCell(object rawValue, BooleanColumn column)
        {
            if (!(rawValue is bool))
            {
                return BuildErrorCell();
            }

            bool asBool = (bool) rawValue;

            var cell = new Cell()
            {
                DataType = CellValues.Boolean,
                CellValue = new CellValue(asBool ? "1" : "0")
            };

            return cell;
        }

        private static Cell BuildInlineTextCell(object rawValue, TextColumn column)
        {
            var stringValue = string.Empty;

            if (rawValue != null)
            {
                stringValue = rawValue.ToString();
            }

            if (column.Prefix != null)
            {
                stringValue = column.Prefix + stringValue;
            }

            if (column.Suffix!= null)
            {
                stringValue = stringValue + column.Suffix;
            }

            var cell = new Cell()
            {
                DataType = CellValues.InlineString,
                InlineString = new InlineString()
                {
                    Text = new Text(stringValue)
                }
            };

            return cell;
        }

        private static Cell BuildErrorCell()
        {
            var cell = new Cell()
            {
                DataType = CellValues.Error,
                CellValue = new CellValue("-")
            };

            return cell;
        }

        private static object GetRawValue(object obj, string pathName)
        {
            var fieldNames = pathName.Split('.');

            foreach (string fieldName in fieldNames)
            {
                var property = obj.GetType().GetProperty(fieldName);

                if (property != null)
                {
                    obj = property.GetValue(obj, null);
                }
                else
                {
                    obj = null;
                    break;
                }
            }
            return obj;
        }
    }
}