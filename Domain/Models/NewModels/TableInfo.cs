using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Serialization;
using XtendedRefs.ComAddIn.Application.Contracts;
using XtendedRefs.ComAddIn.Domain.Abstractions;
using XtendedRefs.ComAddIn.Domain.Documents;
using XtendedRefs.ComAddIn.Domain.State;
using Word = Microsoft.Office.Interop.Word;

namespace XtendedRefs.ComAddIn.Domain.Models.NewModels;

[XmlRoot(ElementName = "TableList", Namespace = "XtendedRefs.ComAddIn.Table")]
public class TableList : IRestorable
{
    public TableList() { }

    [XmlElement("Table", Namespace = "XtendedRefs.ComAddIn.Table")]
    public List<TableInfo> Tables { get; set; } = new List<TableInfo>();

    public void Restore(Word.Document document)
    {
        
    }
}


// Техническая информация о таблице: номер SEQ, закладка на весь абзац подписи.
// Из закладки после обновления полей вытягиваем номер и всё нужное.

[XmlRoot(ElementName = "Table", Namespace = "XtendedRefs.ComAddIn.Table")]
public class TableInfo : DocumentAnchor
{
    
    public TableInfo() { }

    public TableInfo (List<string> numberParts, (int Start, int End) range)
    {
        NumberParts.AddRange(numberParts);
        Start = range.Start;
        End = range.End;
    }

    //------------ номер из диапазона (SEQ + HYPERLINK по главам) ------------

    [XmlArray("Part")]
    public List<string> NumberParts { get; set; } = new List<string>();

    public List<string> Number(Word.Range range)
    {
        List<string> nums = range.Fields.Cast<Word.Field>()
            .Select<Word.Field, string>(f =>
            {
                if (f.Type == Word.WdFieldType.wdFieldSequence)
                    return f.Result.Text;

                if (f.Type == Word.WdFieldType.wdFieldHyperlink)
                {
                    if (f.Result.Hyperlinks.Count > 0)
                    {
                        return f.Result.Hyperlinks[1].TextToDisplay;
                    }
                }

                return "Пропущено";
            })
            .ToList();

        return nums;
    }

    public override string GetNumber(Word.Range range)
    {
        return string.Join(".", Number(range));
    }

    //------------ текст для вставки по шаблону ------------

    [XmlIgnore]
    public string InsertText
    {
        get
        {
            RuntimeStore rs = Singleton<RuntimeStore>.Instance;
            DocumentSettings documentSettings = rs.DocumentSettings;

            int index = 0;
            return Regex.Replace(documentSettings.TableInsertTemplate, @"\$", m =>
            {
                if (index < NumberParts.Count)
                    return NumberParts[index++];
                return m.Value;
            });
        }
    }

    //------------ отображаемое имя ------------

    [XmlIgnore]
    public override string DisplayText
    {
        get => NumberParts?.Count > 0 ? "Таблица " + string.Join(".", NumberParts) : "Нет данных";
        set => throw new NotImplementedException();
    }

    //------------ восстановление при открытии документа ------------

    public void Restore()
    {
        // Таблицы — физические объекты Word, превью не требуется.
        // При необходимости можно добавить загрузку из XML.
    }
    
    public override string ToString()
    {
        var sb = new StringBuilder();

        // ── поля унаследованного класса DocumentAnchor (доступны через base) ──
        if (!string.IsNullOrEmpty(BookmarkName))
            sb.AppendLine($"[Base] BookmarkName = {BookmarkName}");

        if (Start != 0)
            sb.AppendLine($"[Base] Start       = {Start}");

        if (End != 0)
            sb.AppendLine($"[Base] End         = {End}");

        if (Shift != 0)
            sb.AppendLine($"[Base] Shift       = {Shift}");

        // ── поля собственного класса TableInfo ──
        if (NumberParts.Count > 0)
            sb.AppendLine($"Own     NumberParts   = [{string.Join(", ", NumberParts)}]");

        // var numberStr = GetNumber(null!); // null — только для отображения полей, не для Word-операций
        // if (!string.IsNullOrEmpty(numberStr))
        //     sb.AppendLine($"Own     GetNumber()   = {numberStr}");

        return sb.ToString();
    }

}
