namespace ExcelExport
{
    public abstract class ColumnBase
    {
        public string Header { get; set; }
        public string PropertyPath { get; set; }
        public double? Width { get; set; }

        protected ColumnBase()
        {
            
        }

        protected ColumnBase(string header, string propertyPath)
        {
            Header = header;
            PropertyPath = propertyPath;
        }
    }
}