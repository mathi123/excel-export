using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelExport.Tests
{
    [TestClass]
    public class ExcelExporterTests
    {
        public const string TestFilePath = "test.xlsx";
        public const string TestMultiSheetFilePath = "testmultisheet.xlsx";

        [TestMethod]
        public void WriteTestFile()
        {
            var configuration = new SheetConfiguration()
            {
                Name = "Sheet A",
                Columns = new List<ColumnBase>()
                {
                    new TextColumn("Title", nameof(Record.Name)),
                    new BooleanColumn("Active", nameof(Record.IsActive)),
                    new DateColumn("Edited on...", nameof(Record.LastEditTime)),
                    new NumberColumn("Length", nameof(Record.Size))
                },
                Data = new List<object>()
                {
                    new Record()
                    {
                        LastEditTime = DateTime.Now,
                        IsActive = true,
                        Name = "Test a",
                        Size = 123.4
                    },
                    new Record()
                    {
                        LastEditTime = DateTime.MaxValue,
                        IsActive = false,
                        Name = "Test b",
                        Size = 4
                    }
                }
            };

            using (var stream = ExcelExporter.Generate(configuration))
            {
                File.WriteAllBytes(TestFilePath, stream.ToArray());
            }
        }
        [TestMethod]
        public void WriteTestMultiSheetFile()
        {
            var sheetA = new SheetConfiguration()
            {
                Name = "Sheet A",
                Columns = new List<ColumnBase>()
                {
                    new TextColumn("Title", nameof(Record.Name)),
                    new BooleanColumn("Active", nameof(Record.IsActive)),
                    new DateColumn("Edited on...", nameof(Record.LastEditTime)),
                    new NumberColumn("Length", nameof(Record.Size))
                },
                Data = new List<object>()
                {
                    new Record()
                    {
                        LastEditTime = DateTime.Now,
                        IsActive = true,
                        Name = "Test a",
                        Size = 123.4
                    },
                    new Record()
                    {
                        LastEditTime = DateTime.MaxValue,
                        IsActive = false,
                        Name = "Test b",
                        Size = 4
                    }
                }
            };
            var sheetB = new SheetConfiguration()
            {
                Name = "Sheet B",
                Columns = new List<ColumnBase>()
                {
                    new TextColumn("Title", nameof(Record.Name)),
                    new BooleanColumn("Active", nameof(Record.IsActive)),
                    new DateColumn("Edited on...", nameof(Record.LastEditTime)),
                    new NumberColumn("Length", nameof(Record.Size))
                },
                Data = new List<object>()
                {
                    new Record()
                    {
                        LastEditTime = DateTime.Now,
                        IsActive = true,
                        Name = "Test d",
                        Size = 123.4
                    },
                    new Record()
                    {
                        LastEditTime = DateTime.MaxValue,
                        IsActive = false,
                        Name = "Test e",
                        Size = 4
                    }
                }
            };

            using (var stream = ExcelExporter.Generate(new [] { sheetA, sheetB }))
            {
                File.WriteAllBytes(TestMultiSheetFilePath, stream.ToArray());
            }
        }
    }
}
