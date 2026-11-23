using System.Xml.Serialization;
using Word = Microsoft.Office.Interop.Word;


namespace XtendedRefs.ComAddIn.Domain.Models.NewModels;

public abstract class DocumentAnchor
{
    public string BookmarkName { get; set; } = string.Empty;

    [XmlIgnore]
    public int Start { get; set; }
    
    [XmlIgnore]
    public int End { get; set; }
    
    [XmlIgnore]
    public int Shift { get; set; }

    [XmlIgnore]
    public abstract string DisplayText {get; set; }
    
    public abstract string GetNumber(Word.Range range);
}