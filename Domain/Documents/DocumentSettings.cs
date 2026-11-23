using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls.Ribbon;
using System.Xml.Serialization;
using XtendedRefs.ComAddIn.Domain.Abstractions;

namespace XtendedRefs.ComAddIn.Domain.Documents;

public static class Brackets
{
    private const string SquareBracket = "[]";
    private const string CircleBracket = "()";
    private const string CurlyBracket = "{}"; 

    public static string GetInput(string text)
    {
        switch (text)
        {
            case "[...]":
                return SquareBracket;
            case "{...}":
                return CircleBracket;
            case "(...)":
                return CurlyBracket;
            default:
                return string.Empty;
        }
    }

    public static string GetSquareText()
    {
        return "[...]";
    }
    public static string GetCircleText()
    {
        return "(...)";
    }

    public static string GetCurlyText()
    {
        return "{...}";
    }
}

public enum SymbolType
{
    Literature,
    Formula,
    Table,
    Picture
}

[XmlRoot(ElementName = "LevelStyleItem", Namespace = "XtendedRefs.ComAddIn.Settings")]
public class LevelStyleItem
{
    
    public LevelStyleItem() { }
    
    /// <summary>
    /// Название стиля
    /// </summary>
    public string StyleName { get; set; }
    
    /// <summary>
    /// Уровень стиля
    /// </summary>
    public int StyleLevel { get; set; }
}

[XmlRoot(ElementName = "DocumentSettings", Namespace = "XtendedRefs.ComAddIn.Settings")]
public class DocumentSettings
{
    public DocumentSettings() { }

    public void CheckSettings()
    {
        if (IsOpened == false)
        {
            IsOpened = true;
            IsDirty = false;
            IsScanned = false;
            PictureSignTemplate = BasicTemplates.PictureSignTemplate;
            PictureInsertTemplate = BasicTemplates.PictureInsertTemplate;
            FormulaSignTemplate = BasicTemplates.FormulaSignTemplate;
            FormulaInsertTemplate = BasicTemplates.FormulaInsertTemplate;
            LiteratureSignTemplate = BasicTemplates.LiteratureSignTemplate;
            LiteratureInsertTemplate = BasicTemplates.LiteratureInsertTemplate;
            TableSignTemplate = BasicTemplates.TableSignTemplate;
            TableInsertTemplate = BasicTemplates.TableInsertTemplate;
            LevelStyles = new List<LevelStyleItem>();
        }
        else
        {
            if(string.IsNullOrWhiteSpace(PictureSignTemplate))
                PictureSignTemplate =  BasicTemplates.PictureSignTemplate;
            if(string.IsNullOrWhiteSpace(PictureInsertTemplate))
                PictureInsertTemplate =  BasicTemplates.PictureInsertTemplate;
            if(string.IsNullOrWhiteSpace(LiteratureSignTemplate))
                LiteratureSignTemplate =  BasicTemplates.LiteratureSignTemplate;
            if(string.IsNullOrWhiteSpace(LiteratureInsertTemplate))
                LiteratureInsertTemplate =  BasicTemplates.LiteratureInsertTemplate;
            if(string.IsNullOrWhiteSpace(TableSignTemplate))
                TableSignTemplate =  BasicTemplates.TableSignTemplate;
            if(string.IsNullOrWhiteSpace(TableInsertTemplate))
                TableInsertTemplate =  BasicTemplates.TableInsertTemplate;
            if(string.IsNullOrWhiteSpace(FormulaSignTemplate))
                FormulaSignTemplate =  BasicTemplates.FormulaSignTemplate;
            if(string.IsNullOrWhiteSpace(FormulaInsertTemplate))
                FormulaInsertTemplate =  BasicTemplates.FormulaInsertTemplate;
            
        }
    }
    
    [XmlArray(ElementName = "LevelStyles",Namespace = "XtendedRefs.ComAddIn.Settings")]
    [XmlArrayItem(ElementName = "LevelStyleItem",Namespace = "XtendedRefs.ComAddIn.Settings")]
    public List<LevelStyleItem> LevelStyles { get; set; }
    
    public string PictureSignTemplate { get; set; } = string.Empty;
    public string PictureInsertTemplate { get; set; } = string.Empty;

    public string PictureStyle { get; set; } = string.Empty;

    public string FormulaSignTemplate { get; set; } = string.Empty;
    public string FormulaInsertTemplate { get; set; } = string.Empty;

    public string LiteratureSignTemplate { get; set; } = string.Empty;
    public string LiteratureInsertTemplate { get; set; } = string.Empty;
    public string LiteratureStyle { get; set; } = string.Empty;
    public string LiteratureSymbol { get; set; } = string.Empty;
    
    public string LiteratureBracketInsertTemplate { get; set; } = string.Empty;

    public string TableSignTemplate { get; set; } = string.Empty;
    public string TableInsertTemplate { get; set; } = string.Empty;

    public string TableStyle { get; set; } = string.Empty;

    public bool TableCaptionNextLine { get; set; }
    
    public bool IsOpened { get; set; }

    /// <summary>
    /// Параметр "ленивого" обновления документа
    /// </summary>
    public bool IsDirty { get; set; }
    
    public bool IsScanned { get; set; }


    public string NormalizeLiteratureSeparator()
    {
        string separator = LiteratureSymbol;
        if (string.Equals(separator, "SPACE", StringComparison.Ordinal))
        {
            return " ";
        }

        if (string.Equals(separator, "TAB", StringComparison.Ordinal))
        {
            return "\t";
        }

        return separator ?? string.Empty;
    }

    public int GetSymbolCount(SymbolType type)
    {
        int count = 0;
        char target = '$';
        switch (type)
        {
            case SymbolType.Literature:
            {
                count = LiteratureSignTemplate.Count( c => c == target );
                break;
            }
            case SymbolType.Formula:
            {
                count = FormulaSignTemplate.Count( c => c == target );
                break;
            }
            case SymbolType.Table:
            {
                count = TableSignTemplate.Count( c => c == target );
                break;
            }
            case SymbolType.Picture:
            {
                count = PictureSignTemplate.Count( c => c == target );
                break;
            }
        }

        return count;
    }
}
