using Spire.Xls;

namespace MBXel_Core.Core.Units
{
    /// <summary>
    /// Represents a range of cells style
    /// </summary>
    public class RangeStyle
    {
        public HorizontalAlignType HorizontalAlign  { get; set; } = HorizontalAlignType.Left;
        public VerticalAlignType   VerticalAlign    { get; set; } = VerticalAlignType.Bottom;

        public string              FontColor        { get; set; } = "#000000";
        public string              BackColor        { get; set; } = "#FFFFFF";

        public double              FontSize         { get; set; } = 12;
        public bool                IsFontBold       { get; set; } = false;
        public bool                IsFontItalic     { get; set; } = false;
        public FontUnderlineType   FontUnderline    { get; set; } = FontUnderlineType.None;

        public LineStyleType       BordersLineStyle { get; set; } = LineStyleType.None;
        public string              BordersColor     { get; set; } = "#000000";

        public double              RowHeight     { get; set; } = 15;
    }
}