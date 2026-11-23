using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using XtendedRefs.ComAddIn.Application.Composition;
using XtendedRefs.ComAddIn.Application.Utilities;
using XtendedRefs.ComAddIn.Domain.Abstractions;
using XtendedRefs.ComAddIn.Domain.Documents;
using XtendedRefs.ComAddIn.Domain.Models.NewModels;
using XtendedRefs.ComAddIn.Domain.State;
using XtendedRefs.ComAddIn.Infrastructure.Events;
using Word = Microsoft.Office.Interop.Word;

namespace XtendedRefs.ComAddIn.Application.Commands.NewCommands;

public class LiteratureCommands
{
    public void AddLiterature()
    {
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;
        Word.Document doc = rs.Document;
        DocumentSettings documentSettings = rs.DocumentSettings;

        bool previousScreenUpdating = doc.ActiveWindow.Application.ScreenUpdating;
        doc.ActiveWindow.Application.ScreenUpdating = false;

        try
        {
            //Собираем шаблон, стиль и разделитель
            string template = documentSettings.LiteratureSignTemplate;
            string styleName = documentSettings.LiteratureStyle;
            string separator = documentSettings.NormalizeLiteratureSeparator();

            //Берем место курсора и параграф
            Word.Selection selection = doc.ActiveWindow.Selection;
            Word.Paragraph paragraph = selection.Paragraphs[1];
            Word.Range targetRange = paragraph.Range.Duplicate;
            targetRange.End -= 1;
            string literatureText = targetRange.Text;
            int initEnd = paragraph.Range.End;
            //Если назначен стиль, применяем
            if (!string.IsNullOrEmpty(styleName))
            {
                try
                {
                    paragraph.set_Style(doc.Styles[styleName]);
                }
                catch
                {
                    // Style names can be document-specific; the literature item can still be inserted.
                }
            }

            Word.Range findRange = paragraph.Range.Duplicate;
            findRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            //Вставили шаблон
            findRange.Text = template + separator;

            //Ищем $ для вставки поля.
            //TODO: Необходимость вставки для нескольких местозаполнителей литературы
            //------------

            List<ChapterInfo> chapters = DocumentUtilities.FindNearChaptersByTemplate(doc, SymbolType.Literature);
            Word.Find find = findRange.Find;
            find.Text = "$";
            find.Wrap = Word.WdFindWrap.wdFindStop;
            Word.Field litField = null;
            string forNumParts = string.Empty;
            if (chapters.Count == 0 | chapters.Count == 1)
            {
                find.Execute();
                
                Word.Range hit = findRange.Duplicate;
                Word.Range dollar = hit.Duplicate;
                dollar.End = dollar.Start + 1;
                dollar.Text = "";
                litField = targetRange.Fields.Add(
                    dollar,
                    Word.WdFieldType.wdFieldSequence,
                    @"Literature_Seq"
                );
                forNumParts = litField.Result.Text;
            }
            else
            {
                List<string> nums = new List<string>();
                for (int i = 0; i < chapters.Count; i++)
                {
                    find.Execute();
                    Word.Range hit = findRange.Duplicate;
                    Word.Range dollar = hit.Duplicate;
                    dollar.End = dollar.Start + 1;
                    dollar.Text = "";

                    if (i == chapters.Count - 1)
                    {
                        litField = targetRange.Fields.Add(
                            dollar,
                            Word.WdFieldType.wdFieldSequence,
                            $"Literature_{string.Join("", nums)}_Seq"
                        );
                        string number = litField.Result.Text;
                        forNumParts += number;
                        
                    }
                    else
                    {
                        string number = chapters[i].NumberParts[chapters[i].NumberParts.Count - 1];
                        targetRange.Hyperlinks.Add(
                            Anchor: dollar,
                            SubAddress: chapters[i].BookmarkName,
                            TextToDisplay: number,
                            ScreenTip: chapters[i].Title
                        );
                        forNumParts += number + ".";
                        nums.Add(number);
                    }

                    findRange.Start = dollar.Start + 1;
                    find = findRange.Find;
                    find.Text = "$";
                }
            }
            //---------------
            string bmName = string.Join("", "_lt_", Guid.NewGuid().ToString("N"));
            List<string> pts = DocumentUtilities.ExtractNumbers(forNumParts);
            int shift = targetRange.End - initEnd;
            LiteratureInfo lit = new LiteratureInfo(literatureText, pts, (targetRange.Start , targetRange.End), shift);
            lit.BookmarkName = bmName;
            try
            {
                rs.LiteratureList.Literature.Add(lit);
                Word.Bookmark bm = doc.Bookmarks.Add(bmName, targetRange);
                doc.Fields.Update();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        finally
        {
            doc.Application.ScreenUpdating = previousScreenUpdating;
            Update();
            EventBus<LiteratureInfo>.Raise(doc);
        }
    }

    #region MyRegion

    //test
    public static string RevealHiddenChars(string input)
    {
        if (string.IsNullOrEmpty(input)) return string.Empty;

        var sb = new StringBuilder();

        foreach (char c in input)
        {
            switch (c)
            {
                case ' ':
                    sb.Append('·'); // Обычный пробел
                    break;
                case '\u00A0':
                    sb.Append('°'); // Неразрывный пробел (часто встречается в Word)
                    break;
                case '\t':
                    sb.Append(@"\t"); // Табуляция
                    break;
                case '\r':
                    sb.Append(@"\r"); // Возврат каретки (конец абзаца в Word)
                    break;
                case '\n':
                    sb.Append(@"\n"); // Перевод строки
                    break;
                case '\v':
                    sb.Append(@"\v"); // Вертикальная табуляция (мягкий перенос Shift+Enter в Word)
                    break;
                case '\f':
                    sb.Append(@"\f"); // Разрыв страницы или раздела в Word
                    break;
                case '\a':
                    sb.Append(@"\a"); // Маркер конца ячейки или строки таблицы в Word
                    break;
                default:
                    // Если символ управляющий (невидимый), выводим его код \uXXXX
                    if (char.IsControl(c))
                    {
                        sb.Append($@"\u{(int)c:X4}");
                    }
                    else
                    {
                        sb.Append(c);
                    }
                    break;
            }
        }

        return sb.ToString();
    }
        

    #endregion
    
    
    private void Update()
    {
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;
        Word.Document doc = rs.Document;
        doc.Fields.Update();
        Word.Bookmarks bms = doc.Bookmarks;
        foreach (LiteratureInfo lit in rs.LiteratureList.Literature)
        {
            Word.Bookmark bm = bms[lit.BookmarkName];
            if (bm.Range.Paragraphs.Count > 1)
            {
                Word.Paragraph lastPara = bm.Range.Paragraphs[bm.Range.Paragraphs.Count];
                bm.Start = lastPara.Range.Start;
                bm.End = lastPara.Range.End - 1;
            }
            if (bm.Range.Fields.Count == 1)
            {
                bool res = int.TryParse(bm.Range.Fields[1].Result.Text, out int num);
                List<string> nums = new List<string>() { num.ToString() };
                lit.NumberParts.Clear();
                lit.NumberParts.AddRange(nums);
            }
            else
            {
                List<string> nums = lit.Number(bm.Range);
                lit.NumberParts.Clear();
                lit.NumberParts.AddRange(nums);
            }
        }
        doc.Fields.Update();
    }

    public void InsertLiteratureReference()
    {
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;
        Word.Document doc = rs.Document;

        ServiceLocator srv = Singleton<ServiceLocator>.Instance;

        if (rs.LiteratureList.Literature.Count == 0)
        {
            CommonCommands.ShowWarning("Источники литературы не найдены!", "Предупреждение");
            return;
        }

        LiteratureInfo literature = srv.LiteraturePickerService.PickBookmark(doc);
        if (literature == null)
        {
            return;
        }

        Word.Range targetRange = doc.Application.Selection.Range.Duplicate;
        targetRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);


        targetRange.Hyperlinks.Add(
            Anchor: targetRange,
            SubAddress: literature.BookmarkName,
            TextToDisplay: literature.InsertText,
            ScreenTip: literature.LiteratureText
        );
    }

    public void DeleteLiteratureTextReference()
    {
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;
        DocumentSettings documentSettings = rs.DocumentSettings;
        Word.Document doc = rs.Document;

        //Параграф, в котором удаляем ссылку
        Word.Paragraph paragraph = doc.Application.Selection.Paragraphs[1];


        Word.Range paraRange = paragraph.Range.Duplicate;

        if (paraRange.Fields.Count == 0)
        {
            CommonCommands.ShowError("Здесь нет литературы", "Ошибка!");
            return;
        }

        string bmName = paraRange.Bookmarks
            .Cast<Word.Bookmark>()
            .Where(b => b.Name.Contains("_lt_"))
            .Select(b => b.Name)
            .FirstOrDefault();

        if (string.IsNullOrWhiteSpace(bmName))
        {
            CommonCommands.ShowError("Здесь нет литературы", "Ошибка!");
            return;
        }

        DocumentUtilities.ReplaceAllLinks(doc, LinkType.Literature, bmName);
        Word.Bookmark bm = doc.Bookmarks[bmName];
        bm.Delete();
        //Удаляем текст вместе с шаблоном
        try
        {
            LiteratureInfo lit = rs.LiteratureList.Literature.Find(x =>x.BookmarkName == bmName);
            int length = lit.TextRange.Start - lit.TextRange.End;
            int end = lit.End - length-1;
            doc.Range(lit.Start, lit.End).Delete();
            rs.LiteratureList.Literature.RemoveAll(l => l.BookmarkName == bmName);
            
        }
        catch (Exception e)
        {
            MessageBox.Show(e.ToString());
        }
        Update();
        MessageBox.Show("Литература удалена, можно стереть лишнее");
        EventBus<LiteratureInfo>.Raise(doc);
    }
}