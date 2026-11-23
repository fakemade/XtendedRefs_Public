using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Serialization;
using XtendedRefs.ComAddIn.Application.Contracts;
using XtendedRefs.ComAddIn.Domain.Abstractions;
using XtendedRefs.ComAddIn.Domain.Documents;
using XtendedRefs.ComAddIn.Domain.State;
using Word = Microsoft.Office.Interop.Word;

namespace XtendedRefs.ComAddIn.Domain.Models.NewModels;


[XmlRoot(ElementName = "LitreatureList", Namespace = "XtendedRefs.ComAddIn.Literature")]
public class LiteratureList : IRestorable
{
    public LiteratureList() { }

    [XmlElement("Literature", Namespace = "XtendedRefs.ComAddIn.Literature")]
    public List<LiteratureInfo> Literature { get; set; } = new List<LiteratureInfo>();
    
    public void Restore(Word.Document document)
    {
        foreach (var literatureInfo in Literature)
        {
            literatureInfo.LiteratureTextChanged += literatureInfo.OnLitTextChanged;
        }

    }
}

//Technical info
// Номер литературы вставляем как SEQ, ставим на всю литературу закладку. Из нее потом вытягиваем и номер после обновления полей и все, что нужно

[XmlRoot(ElementName = "Literature",  Namespace = "XtendedRefs.ComAddIn.Literature")]
public class LiteratureInfo : DocumentAnchor
{
    private (int Start, int End) _textRange = (0,0);
    public (int Start, int End) TextRange
    {
        get => _textRange;
        set
        {
            _textRange.Start = value.Start + Shift;
            _textRange.End = value.End;
        }
    } 

    public LiteratureInfo(){}


    public LiteratureInfo(string literatureText, List<string> numberParts, (int Start, int End) range, int shift)
    {
        NumberParts.AddRange(numberParts);
        TextRange = range;
        Start = range.Start;
        End = range.End;
        LiteratureText = literatureText;
        Shift = shift;
        LiteratureTextChanged += OnLitTextChanged;
    }
    

    private string _literatureText;
    public event EventHandler<string> LiteratureTextChanged;
    
    [XmlElement("LiteratureText")]
    public string LiteratureText
    {
        get => _literatureText;
        set
        {
            _literatureText = value;
            LiteratureTextChanged?.Invoke(this, _literatureText);
            // 3. Вызываем событие при изменении значения
        }
    }

    

    //--------------
    /// <summary>
    /// Метод, который вызывается при изменении текста в окне литературы, для изменения его в самом документе
    /// </summary>
    /// <param name="sender">Отправитель события</param>
    /// <param name="text">Текст, на который нужно исправить</param>
    public void OnLitTextChanged(object sender, string text)
    {
        Word.Document doc = Singleton<RuntimeStore>.Instance.Document;
        LiteratureInfo obj = sender as LiteratureInfo;
        Word.Range range = doc.Range(obj.TextRange.Start, obj.TextRange.End);
        range.Text = text;
        obj.TextRange = (range.Start, range.End);
        obj.Start = range.Start;
        obj.End = range.End;
    }
    
    //--------------
    [XmlArray("Part")]
    public List<string> NumberParts { get; set; } = new List<string>();
    
    public List<string> Number(Word.Range range)
    {
        List<string> nums = range.Fields.Cast<Word.Field>()
            .Select(f => 
            {
                // Пишем полноценную логику с ветвлением прямо внутри LINQ
                if (f.Type == Word.WdFieldType.wdFieldSequence)
                {
                    return f.Result.Text;
                }

                if (f.Type == Word.WdFieldType.wdFieldHyperlink)
                {
                    if (f.Result.Hyperlinks.Count > 0)
                    {
                        Word.Hyperlink link = f.Result.Hyperlinks[1];
                        string visibleText = link.TextToDisplay;

                        // Теперь вы можете передать объект link в свой метод обработки
                        return visibleText;
                    }
                }

                return "Пропущено"; // Возвращаем строковый статус для остальных
            })
            .ToList(); // Запускает итерацию по коллекции
        return nums;
    }

    
    
    [XmlIgnore]
    public string InsertText
    {
        get
        {
            RuntimeStore rs = Singleton<RuntimeStore>.Instance;
            DocumentSettings documentSettings = rs.DocumentSettings;
            
            int index = 0;
            string result = Regex.Replace(documentSettings.LiteratureInsertTemplate, @"\$", m => 
            {
                if (index < NumberParts.Count)
                {
                    return NumberParts[index++];
                }
                return m.Value;
            });

            return result;
        }
    }
    
    [XmlIgnore] 
    public override string DisplayText
    {
        get
        {
            RuntimeStore rs = Singleton<RuntimeStore>.Instance;
            DocumentSettings documentSettings = rs.DocumentSettings;
            string separator = documentSettings.NormalizeLiteratureSeparator();
            return string.Join(" ", string.Join(".", NumberParts), LiteratureText);
        }
        set => throw new NotImplementedException();
    }


    public override string GetNumber(Word.Range range)
    {
        return string.Join(".", Number(range));
    }
}