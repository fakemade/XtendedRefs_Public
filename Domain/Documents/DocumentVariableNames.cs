namespace XtendedRefs.ComAddIn.Domain.Documents;

internal static class DocumentVariableNames
{
    public const string FigurePrefix = "FigurePrefix";
    public const string FigureStyle = "FigureStyle";
    public const string FormulaPrefix = "FormulaPrefix";
    public const string LiteraturePrefix = "LitPrefix";
    public const string LiteratureStyle = "LitStyle";
    public const string LiteratureSymbol = "LitSymbol";
    public const string TablePrefix = "TablePrefix";
    public const string TableStyle = "TableStyle";
    public const string TableCaptionNextLine = "TableCaptionNextLine";
    public const string MediaMap = "XR_MediaMap";
    public const string IsOpened = "IsOpened";

    public static string GetFigureLevelStyle(int level) => $"FigureLevelStyle{level}";
}
