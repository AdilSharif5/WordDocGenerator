namespace WordDocGenerator
{
    public class DataModel
    {
        public Row[] rows { get; set; }
        public int totalCols { get; set; }
    }

    public class Row
    {
        public Col[] cols { get; set; }
    }

    public class Col
    {
        public Font font { get; set; }
        public string bgColor { get; set; }
        public string color { get; set; }
        public string colSpan { get; set; }
        public string rowSpan { get; set; }
        public string textAlign { get; set; }
        public Cellcontent[] cellContent { get; set; }
    }

    public class Font
    {
    }

    public class Cellcontent
    {
        public string? component { get; set; }
        public Tablejson? tableJson { get; set; }
        public string cellType { get; set; }
        public string label { get; set; }
        public string? display { get; set; }
        public string[]? valuesList { get; set; }
        public string? align { get; set; }
        public string? imageName { get; set; }
        public string? imageSrc { get; set; }
        public string? ImageWidth { get; set; }
        public string? ImageHeight { get; set; }
    }

    public class Tablejson
    {
        public SingleRow[] rows { get; set; }
        public int totalCols { get; set; }
    }

    public class SingleRow
    {
        public SingleCol[] cols { get; set; }
    }

    public class SingleCol
    {
        public Font1 font { get; set; }
        public string bgColor { get; set; }
        public string color { get; set; }
        public Cellcontent[] cellContent { get; set; }
        public string colSpan { get; set; }
        public string rowSpan { get; set; }
        public string textAlign { get; set; }
    }

    public class Font1
    {
    }

    public class Cellcontent1
    {
        public string cellType { get; set; }
        public string label { get; set; }
    }

}