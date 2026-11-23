using System;
using System.Drawing;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Serialization;
using XtendedRefs.ComAddIn.Application.Contracts;
using XtendedRefs.ComAddIn.Domain.Abstractions;
using XtendedRefs.ComAddIn.Domain.Documents;
using XtendedRefs.ComAddIn.Domain.State;
using Word = Microsoft.Office.Interop.Word;

namespace XtendedRefs.ComAddIn.Domain.Models.NewModels;

[XmlRoot(ElementName = "PictureList", Namespace = "XtendedRefs.ComAddIn.Picture")]
public class PictureList : IRestorable
{
    public PictureList() { }

    [XmlElement("Picture", Namespace = "XtendedRefs.ComAddIn.Picture")]
    public List<PictureInfo> Pictures { get; set; } = new List<PictureInfo>();

    public void Restore(Word.Document document)
    {
        
    }
}



[XmlRoot(ElementName = "Picture", Namespace = "XtendedRefs.ComAddIn.Picture")]
public class PictureInfo : DocumentAnchor
{

    // Текст подписи (после номера) — аналог LiteratureText для формул с текстом
    public (int Start, int End) textRange = (0, 0);
    private string _pictureText;


    [XmlElement("PictureText")]
    public string PictureText
    {
        get => _pictureText;
        set => _pictureText = value;
    }

    // Диапазон абзаца подписи в исходном документе (для синхронизации)


    //------------ constructors ------------
    public PictureInfo() { }

    public PictureInfo( List<string> numberParts, (int Start, int End) range)
    {
        NumberParts.AddRange(numberParts);
        Start = range.Start;
        End = range.End;
    }


    //------------ номер из диапазона (поля SEQ + HYPERLINK по главам) ------------

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
            return Regex.Replace(documentSettings.PictureInsertTemplate, @"\$", m =>
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
        get => NumberParts?.Count > 0 ? "Рисунок " + string.Join(".", NumberParts) : "Нет данных";
        set => throw new NotImplementedException();
    }


    


    //------------ восстановление при открытии документа ------------

    public void Restore()
    {
        // Загружаем превью из сериализованной строки (если не null)
        // PictureList.Restore() вызывает это для каждого элемента
        /*if (!string.IsNullOrEmpty(_preview))
            _picture = LoadPreviewFromBase64(_preview);*/
    }
}
