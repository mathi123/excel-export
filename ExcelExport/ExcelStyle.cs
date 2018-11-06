using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelExport
{
    public class ExcelStyle
    {
        public uint FontDefaultId;
        public uint FontBoldId;
        public uint FontItalicId;

        public uint FillDefaultId;
        public uint BorderDefaultId;

        public uint CellFormatDefaultId;
        public uint CellFormatTitleRowId;

        public Stylesheet Stylesheet { get; private set; }

        public ExcelStyle()
        {
            Stylesheet = BuildStyleSheet();
        }

        private Stylesheet BuildStyleSheet()
        {
            return new Stylesheet(BuildFonts(), BuildFills(), BuildBorders(), BuildCellFormats());
        }

        private CellFormats BuildCellFormats()
        {
            var cellFormats = new CellFormats();
            var cellFormatDefault = new CellFormat()
            {
                FontId = FontDefaultId,
                FillId = FillDefaultId,
                BorderId = BorderDefaultId
            };

            cellFormats.AppendChild(cellFormatDefault);

            var cellFormatTitle= new CellFormat(new Alignment()
            {
                Horizontal = HorizontalAlignmentValues.Left,
                Vertical = VerticalAlignmentValues.Center
            })
            {
                FontId = FontBoldId,
                FillId = FillDefaultId,
                BorderId = BorderDefaultId,
                ApplyAlignment = true
            };

            cellFormats.AppendChild(cellFormatTitle);
            CellFormatTitleRowId = 1;

            var cellFormatBody = new CellFormat(new Alignment()
            {
                Horizontal = HorizontalAlignmentValues.Left,
                Vertical = VerticalAlignmentValues.Center
            })
            {
                FontId = FontDefaultId,
                FillId = FillDefaultId,
                BorderId = BorderDefaultId,
                ApplyAlignment = true
            };

            cellFormats.AppendChild(cellFormatBody);
            CellFormatDefaultId = 2;

            return cellFormats;
        }

        private Fonts BuildFonts()
        {
            var fonts = new Fonts();
            fonts.AppendChild(BuildFont(11, "000000", "Calibri"));
            FontDefaultId = 0;

            fonts.AppendChild(BuildFont(11, "000000", "Calibri", true));
            FontBoldId = 1;

            fonts.AppendChild(BuildFont(11, "000000", "Calibri", false, true));
            FontItalicId = 2;

            return fonts;
        }

        private Font BuildFont(int fontSize, string hexBinaryColor, string fontName, bool bold = false, bool italic = false)
        {
            var font = new Font();

            if (bold)
            {
                font.AppendChild(new Bold());
            }
            if (italic)
            {
                font.AppendChild(new Italic());
            }

            font.AppendChild(new FontSize() { Val = fontSize });
            font.AppendChild(new Color() { Rgb = new HexBinaryValue() { Value = hexBinaryColor } });
            font.AppendChild(new FontName() { Val = fontName });
            return font;
        }

        private Fills BuildFills()
        {
            var fills = new Fills();
            var fillDefault = new Fill(new PatternFill() {PatternType = PatternValues.None});

            fills.AppendChild(fillDefault);
            FillDefaultId = 0;

            return fills;
        }

        private Borders BuildBorders()
        {
            BorderDefaultId = 0;

            return new Borders(
                new Border(
                    new LeftBorder(),
                    new RightBorder(),
                    new TopBorder(),
                    new BottomBorder(),
                    new DiagonalBorder())
            );
        }
    }
}
