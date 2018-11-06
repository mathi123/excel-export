namespace ExcelExport
{
    public class NumberColumn : ColumnBase
    {
        public NumberColumn()
        {
            
        }

        public NumberColumn(string header, string propertyPath)
            : base(header, propertyPath)
        {
            
        }
        public bool ShouldRound { get; set; }
        public int Round { get; set; }
        public bool IsCurrency { get; set; }
    }
}
