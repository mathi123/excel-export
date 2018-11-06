using System;

namespace ExcelExport.Tests
{
    public class Record
    {
        public string Name { get; set; }
        public DateTime? LastEditTime{ get; set; }
        public bool IsActive { get; set; }
        public double Size { get; set; }
    }
}
