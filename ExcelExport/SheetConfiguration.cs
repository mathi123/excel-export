using System.Collections.Generic;

namespace ExcelExport
{
    public class SheetConfiguration
    {
        public string Name { get; set; } = "data";
        public IList<object> Data { get; set; }
        public List<ColumnBase> Columns { get; set; } = new List<ColumnBase>();
    }
}