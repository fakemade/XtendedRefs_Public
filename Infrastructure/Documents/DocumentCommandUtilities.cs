using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows;
using XtendedRefs.ComAddIn.Domain.Documents;
using Word = Microsoft.Office.Interop.Word;

namespace XtendedRefs.ComAddIn.Infrastructure.Documents;

public class DocumentCommandUtilities
{
    public Word.Document Document { get; set; }
    public DocumentCommandUtilities(Word.Document document)
    {
        Document = document;
    }
    
    public  void CheckNullProperties(object obj)
    {
        if (obj == null)
        {
            MessageBox.Show("Object is null");
            return;
        }

        foreach (PropertyInfo prop in obj.GetType().GetProperties())
        {
            object value = prop.GetValue(obj);

            if (value == null)
            {
                MessageBox.Show($"{prop.Name} is NULL");
            }
        }
    }
    
    public bool RangesIntersect(Word.Range a, Word.Range b)
    {
        return a.Start <= b.End && b.Start <= a.End;
    }
    
    public string Get(string name)
    {
        try
        {
            return Document.Variables[name].Value ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    public void Set(string name, string value)
    {
        try
        {
            Document.Variables[name].Value = value ?? string.Empty;
        }
        catch
        {
            Document.Variables.Add(name, value ?? string.Empty);
        }
    }
    

    public  int CountDollars(string template)
    {
        int count = 0;
        foreach (char ch in template ?? string.Empty)
        {
            if (ch == '$')
            {
                count++;
            }
        }

        return count;
    }

    public string GetFigureLevelStyle(int level)
    {
        return Get(DocumentVariableNames.GetFigureLevelStyle(level));
    }

    public string GetFigureTemplate()
    {
        string result = Get(DocumentVariableNames.FigurePrefix);
        return string.IsNullOrEmpty(result) ? "Рисунок $" : result;
    }

    public string GetFigureStyle()
    {
        return Get(DocumentVariableNames.FigureStyle);
    }

    public string GetFormulaTemplate()
    {
        string result = Get(DocumentVariableNames.FormulaPrefix);
        return string.IsNullOrEmpty(result) ? "($)" : result;
    }

    public string GetLiteratureTemplate()
    {
        string result = Get(DocumentVariableNames.LiteraturePrefix);
        return string.IsNullOrEmpty(result) ? "$. " : result;
    }

    public string GetLiteratureStyle()
    {
        return Get(DocumentVariableNames.LiteratureStyle);
    }
    
    public string GetTableTemplate()
    {
        string result = Get(DocumentVariableNames.TablePrefix);
        return string.IsNullOrEmpty(result) ? "Таблица $" : result;
    }

    public string GetTableStyle()
    {
        return Get( DocumentVariableNames.TableStyle);
    }

    public bool GetTableCaptionNextLine()
    {
        return string.Equals(
            Get( DocumentVariableNames.TableCaptionNextLine),
            "True",
            StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Returns true when the document contains at least one paragraph using the configured
    /// level-1 heading style (or built-in Heading 1 if no custom style is configured).
    /// </summary>
    public bool DocumentHasChapters()
    {
        string levelStyleName = GetFigureLevelStyle(1);
        if (!string.IsNullOrWhiteSpace(levelStyleName))
        {
            foreach (Word.Paragraph paragraph in Document.Paragraphs)
            {
                try
                {
                    Word.Style style = paragraph.get_Style() as Word.Style;
                    if (style != null &&
                        string.Equals(style.NameLocal, levelStyleName, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
                catch
                {
                }
            }

            return false;
        }

        Word.Range range = Document.Range();
        Word.Find find = range.Find;
        find.ClearFormatting();
        find.set_Style(Document.Styles[Word.WdBuiltinStyle.wdStyleHeading1]);
        find.Text = string.Empty;
        find.Forward = true;
        find.Wrap = Word.WdFindWrap.wdFindStop;
        return find.Execute();
    }

    public void SaveVariable(string name, string value)
    {
        Set( name, value);
    }

    public string GetVariable(string name)
    {
        return Get( name) is null
            ? string.Empty
            : Get( name);
    }

    public void SaveFigureHash(string bookmarkName, long hashValue)
    {
        string folder = Path.Combine(Path.GetTempPath(), "WordFigures");
        Directory.CreateDirectory(folder);
        File.WriteAllText(Path.Combine(folder, bookmarkName + ".hash"), hashValue.ToString());
    }

    public long LoadFigureHash(string bookmarkName)
    {
        string hashFile = Path.Combine(Path.GetTempPath(), "WordFigures", bookmarkName + ".hash");
        if (!File.Exists(hashFile))
        {
            return -1;
        }

        return long.TryParse(File.ReadAllText(hashFile).Trim(), out long value)
            ? value
            : -1;
    }

    public List<string> GetParagraphStyleNames()
    {
        List<string> styles = new List<string>();

        foreach (Word.Style style in Document.Styles)
        {
            try
            {
                if (style.Type == Word.WdStyleType.wdStyleTypeParagraph)
                {
                    styles.Add(style.NameLocal);
                }
            }
            catch
            {
            }
        }

        return styles;
    }
}