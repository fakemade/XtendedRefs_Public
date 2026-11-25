using System;
using System.Collections.Generic;
using System.Linq;
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

public sealed class PictureCommands
{
    public void CreatePicture()
    {
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;
        Word.Document doc = rs.Document;
        DocumentSettings documentSettings = rs.DocumentSettings;

        bool previousScreenUpdating = doc.ActiveWindow.Application.ScreenUpdating;
        bool prevPagination = doc.ActiveWindow.Application.Options.Pagination;
        Word.WdViewType oldView = doc.ActiveWindow.View.Type;
        doc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;
        doc.ActiveWindow.Application.ScreenUpdating = false;
        doc.ActiveWindow.Application.Options.Pagination = false;
        
        //Rescan chapters
        DocumentUtilities.RescanDocument2(doc, rs.DocumentSettings.LevelStyles);
        
        try
        {
            // Собираем шаблон, стиль и разделитель
            string template = documentSettings.PictureSignTemplate;
            string styleName = documentSettings.PictureStyle;

            // Берём место курсора и параграф
            Word.Selection selection = doc.ActiveWindow.Selection;
            Word.Paragraph paragraph = selection.Paragraphs[1];
            Word.Range tRange = paragraph.Range.Duplicate;
            tRange.End -= 1;

            // Если назначен стиль, применяем
            if (!string.IsNullOrEmpty(styleName))
            {
                try
                {
                    paragraph.set_Style(doc.Styles[styleName]);
                }
                catch
                {
                    // Style names can be document-specific; the picture item can still be inserted.
                }
            }

            //Делаем рисунок по центру и не отрываем от следующего
            paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            Word.Paragraph prevPara = paragraph.Previous();
            if (prevPara != null)
            {
                // Выравнивание по центру
                prevPara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                // Свойство «Не отрывать от следующего» (-1 в Word Interop означает true)
                prevPara.KeepWithNext = -1;
            }

            Word.Range targetRange = paragraph.Range.Duplicate;
            Word.Range findRange = paragraph.Range.Duplicate;
            findRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            // Вставили шаблон
            findRange.Text = template;

            // Ищем $ для вставки поля.
            List<ChapterInfo> chapters = DocumentUtilities.FindNearChaptersByTemplate(doc, SymbolType.Picture);
            int sCount = rs.DocumentSettings.GetSymbolCount(SymbolType.Picture);

            Word.Find find = findRange.Find;
            find.Text = "$";
            find.Wrap = Word.WdFindWrap.wdFindStop;
            Word.Field picField = null;
            string forNumParts = string.Empty;

            if (sCount == 1)
            {
                find.Execute();

                Word.Range hit = findRange.Duplicate;
                Word.Range dollar = hit.Duplicate;
                dollar.End = dollar.Start + 1;
                dollar.Text = "";
                picField = targetRange.Fields.Add(
                    dollar,
                    Word.WdFieldType.wdFieldSequence,
                    @"Picture_seq \* MERGEFORMAT"
                );
                forNumParts = picField.Result.Text;
            }
            else
            {
                List<string> nums = new List<string>();
                for (int i = 0; i <= chapters.Count; i++)
                {
                    find.Execute();
                    Word.Range hit = findRange.Duplicate;
                    Word.Range dollar = hit.Duplicate;
                    dollar.End = dollar.Start + 1;
                    dollar.Text = "";

                    if (i == chapters.Count)
                    {
                        picField = targetRange.Fields.Add(
                            dollar,
                            Word.WdFieldType.wdFieldSequence,
                            $"Picture_{string.Join("", nums)}_seq \\* MERGEFORMAT"
                        );
                        forNumParts += picField.Result.Text;
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

            // Записываем закладку на весь абзац подписи (скрываем в обычном окне Word)
            string bmName = string.Join("", "_pc_", Guid.NewGuid().ToString("N"));
            List<string> pts = DocumentUtilities.ExtractNumbers(forNumParts);

            PictureInfo pic = new PictureInfo(pts, (targetRange.Start, targetRange.End));
            pic.BookmarkName = bmName;
            try
            {
                rs.PictureList.Pictures.Add(pic);
                targetRange.End -= 1;
                doc.Bookmarks.Add(bmName, targetRange);
                doc.Content.Fields.Update();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        finally
        {
            doc.Application.ScreenUpdating = previousScreenUpdating;
            doc.Application.Options.Pagination = prevPagination;
            doc.ActiveWindow.View.Type = oldView;
            Update();
            EventBus<PictureInfo>.Raise(doc);
        }
    }

    // Обновление номерных частей (SEQ / HYPERLINK) из закладок в памяти — аналог LiteratureCommands.Update()
    private void Update()
    {
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;
        Word.Document doc = rs.Document;
        
        
        
        bool previousScreenUpdating = doc.ActiveWindow.Application.ScreenUpdating;
        bool prevPagination = doc.ActiveWindow.Application.Options.Pagination;
        Word.WdViewType oldView = doc.ActiveWindow.View.Type;
        doc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;
        doc.ActiveWindow.Application.ScreenUpdating = false;
        doc.ActiveWindow.Application.Options.Pagination = false;

        try
        {
            // Обновляем только поля SEQ и REF в основном тексте (включая таблицы основного текста)
            doc.Content.Fields.Update();
            foreach (PictureInfo pic in rs.PictureList.Pictures)
            {
                if (!doc.Bookmarks.Exists(pic.BookmarkName)) continue;
                Word.Bookmark bm = doc.Bookmarks[pic.BookmarkName];
                // Если закладка растянулась на несколько абзацев (внутри появился символ \r)
                if (bm.Range.Paragraphs.Count > 1)
                {
                    // Берем только ПОСЛЕДНИЙ абзац внутри этой закладки (где и живет сам источник)
                    Word.Paragraph lastPara = bm.Range.Paragraphs[bm.Range.Paragraphs.Count];
    
                    // Переназначаем границы закладки строго на этот абзац (без символа \r в конце)
                    bm.Start = lastPara.Range.Start;
                    bm.End = lastPara.Range.End - 1;
                }
                //string empty;
                if (bm.Range.Fields.Count == 1)
                {
                    bool res = int.TryParse(bm.Range.Fields[1].Result.Text, out int num);
                    List<string> nums = new List<string>() { num.ToString() };
                    pic.NumberParts.Clear();
                    pic.NumberParts.AddRange(nums);
                    //empty = pic.DisplayText;
                }
                else
                {
                    List<string> nums = pic.Number(bm.Range);
                    pic.NumberParts.Clear();
                    pic.NumberParts.AddRange(nums);
                    //empty = pic.DisplayText;
                }
            }
        }
        finally
        {
            doc.Application.ScreenUpdating = previousScreenUpdating;
            doc.Application.Options.Pagination = prevPagination;
            doc.ActiveWindow.View.Type = oldView;
        }
        doc.Fields.Update();
    }

    public void DeletePicture()
    {
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;
        Word.Document doc = rs.Document;

        // Параграф, в котором удаляем подпись рисунка
        Word.Selection selection = doc.Application.Selection;
        Word.Paragraph paragraph = selection.Paragraphs[1];

        Word.Range paraRange = paragraph.Range.Duplicate;
        bool previousScreenUpdating = doc.ActiveWindow.Application.ScreenUpdating;
        doc.ActiveWindow.Application.ScreenUpdating = false;
        // Ищем закладку "_pc_..." в текущем параграфе
        string bmName = paraRange.Bookmarks
            .Cast<Word.Bookmark>()
            .Where(b => b.Name.Contains("_pc_"))
            .Select(b => b.Name)
            .FirstOrDefault();
        
        
        if (string.IsNullOrWhiteSpace(bmName))
        {
            CommonCommands.ShowError("Здесь нет рисунка", "Ошибка!");
            return;
        }

        // Заменяем все гиперссылки, ссылающиеся на удалённый рисунок, на текст-заглушку
        DocumentUtilities.ReplaceAllLinks(doc, LinkType.Picture, bmName);
        Word.Bookmark bm = doc.Bookmarks[bmName];

        try
        {
            doc.Range(bm.Range.Start, bm.Range.End).Delete();
            rs.PictureList.Pictures.RemoveAll(l => l.BookmarkName == bmName);
        }
        catch (Exception e)
        {
            MessageBox.Show(e.ToString());
        }
        
        doc.ActiveWindow.Application.ScreenUpdating = previousScreenUpdating;
        Update();
        EventBus<PictureInfo>.Raise(doc);
    }

    public void InsertPictureReference()
    {
        
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;
        Word.Document doc = rs.Document;
        ServiceLocator srv = Singleton<ServiceLocator>.Instance;
        
        if (rs.PictureList.Pictures.Count == 0)
        {
            CommonCommands.ShowWarning("Рисунки не найдены!", "Предупреждение");
            return;
        }

        // Подбираем картинку из списка (сервис PicturePickerService будет зарегистрирован в ServiceLocator)
        PictureInfo picture = (PictureInfo)srv.PicturePickerService.PickBookmark(doc);
        if (picture == null)
        {
            return;
        }


        Word.Range targetRange = doc.Application.Selection.Range.Duplicate;
        targetRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);

        // Вставляем гиперссылку на закладку рисунка с номером как отображаемым текстом
        targetRange.Hyperlinks.Add(
            Anchor: targetRange,
            SubAddress: picture.BookmarkName,
            TextToDisplay: picture.InsertText,
            ScreenTip: string.Join("",picture.NumberParts)
        );
        
    }
}