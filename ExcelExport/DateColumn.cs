namespace ExcelExport
{
    public class DateColumn : ColumnBase
    {
        public string Format { get; set; }

        public DateColumn()
        {
            Format = "yyyy/MM/dd";
        }

        public DateColumn(string header, string propertyPath):base(header, propertyPath)
        {
            Format = "yyyy/MM/dd";
        }
    }
}
