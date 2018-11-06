namespace ExcelExport
{
    public class TextColumn : ColumnBase
    {
        public string Prefix { get; set; }
        public string Suffix { get; set; }

        public TextColumn()
        {
            
        }

        public TextColumn(string header, string propertyPath) 
            : base(header, propertyPath)
        {
        }
    }
}
