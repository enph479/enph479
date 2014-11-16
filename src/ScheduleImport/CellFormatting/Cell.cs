using System.Drawing;

namespace ElectricalToolSuite.ScheduleImport.CellFormatting
{
    internal class Cell
    {
        public int RowIndex { get; set; }
        public int ColumnIndex { get; set; }

        public int NumberOfColumns { get; set; }
        public int NumberOfRows { get; set; }

        public string Text { get; set; }

        public string FontName { get; set; }
        public double FontSize { get; set; }
        public bool FontBold { get; set; }
        public bool FontItalic { get; set; }
        public bool FontUnderline { get; set; }
        public Color FontColor { get; set; }

        public Color BackgroundColor { get; set; }

        public HorizontalAlignment HorizontalAlignment { get; set; }
        public VerticalAlignment VerticalAlignment { get; set; }

        public BorderLineStyle BottomBorderLineStyle { get; set; }
        public BorderLineStyle TopBorderLineStyle { get; set; }
        public BorderLineStyle LeftBorderLineStyle { get; set; }
        public BorderLineStyle RightBorderLineStyle { get; set; }

        public Orientation Orientation { get; set; }
    }
}
