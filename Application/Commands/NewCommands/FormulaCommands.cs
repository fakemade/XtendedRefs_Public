using System;
using System.Collections.Generic;
using XtendedRefs.ComAddIn.Application.Composition;
using XtendedRefs.ComAddIn.Application.Utilities;
using XtendedRefs.ComAddIn.Domain.Abstractions;
using XtendedRefs.ComAddIn.Domain.Documents;
using XtendedRefs.ComAddIn.Domain.Models.NewModels;
using XtendedRefs.ComAddIn.Domain.State;
using XtendedRefs.ComAddIn.Infrastructure.Events;
using Word = Microsoft.Office.Interop.Word;

namespace XtendedRefs.ComAddIn.Application.Commands.NewCommands;

public class FormulaCommands
{
    public void AddFormula()
    {
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;
        Word.Document doc = rs.Document;
        DocumentSettings documentSettings = rs.DocumentSettings;
        
        //Rescan chapters
        DocumentUtilities.RescanDocument2(doc, rs.DocumentSettings.LevelStyles);
        
        Word.Selection selection = doc.Application.Selection;
        Word.Paragraph currentParagraph = selection.Paragraphs[1];
        //Проверка на выделение объекта MathType
        if (selection.InlineShapes.Count == 0)
        {
            CommonCommands.ShowWarning("Выберите объект MathType!", "Предупреждение");
            return;
        }

        if (currentParagraph.Range.Tables.Count >= 1)
        {
            CommonCommands.ShowWarning("Выделенная формула уже форматирована или находится в другой таблице!",
                "Предупреждение");
            return;
        }


        bool previousScreenUpdating = doc.ActiveWindow.Application.ScreenUpdating;
        bool prevPagination = doc.ActiveWindow.Application.Options.Pagination;
        Word.WdViewType oldView = doc.ActiveWindow.View.Type;

        doc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;
        doc.ActiveWindow.Application.ScreenUpdating = false;
        doc.ActiveWindow.Application.Options.Pagination = false;
        try
        {
            //Собираем шаблон
            string template = documentSettings.FormulaSignTemplate;

            Word.InlineShape selectedObject = selection.InlineShapes[1];
            Word.Range equationRange = selectedObject.Range;
            equationRange.Cut();

            //Проверяем, есть ли сверху или снизу таблица, в следующем или предыдущем абзаце
            Word.Paragraph prevParagraph = currentParagraph.Previous();
            Word.Paragraph nextParagraph = currentParagraph.Next();
            Word.Table table;
            int index = 1;
            if (prevParagraph.Range.Tables.Count >= 1)
            {
                table = prevParagraph?.Range.Tables[1];
                table.Rows.Add();
                index = table.Rows.Count;
                currentParagraph.Range.Delete();
            }
            else if (nextParagraph?.Range.Tables.Count >= 1)
            {
                table = nextParagraph.Range.Tables[1];
                Word.Row firstRow = table.Rows[1];
                table.Rows.Add(firstRow);
                currentParagraph.Range.Delete();
            }
            else
            {
                table = doc.Tables.Add(selection.Range, 1, 3);
            }


            table.Borders.Enable = 0;

            float totalWidth = doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin;
            float col1Width = doc.Application.CentimetersToPoints(2);
            float col3Width = doc.Application.CentimetersToPoints(2);

            table.Cell(index, 1).Width = col1Width;
            table.Cell(index, 2).Width = totalWidth - col1Width - col3Width;
            table.Cell(index, 3).Width = col3Width;
            table.Cell(index, 2).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            table.Cell(index, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            table.Cell(index, 3).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            table.Cell(index, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            table.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
            Word.Range targetEquationRange = table.Cell(index, 2).Range;
            targetEquationRange.End -= 1;
            targetEquationRange.Paste();
            // equationRange.Delete();


            Word.Range targetRange = table.Cell(index, 3).Range;
            Word.Range findRange = targetRange.Duplicate;
            targetRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            findRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            //Вставили шаблон
            findRange.Text = template;

            //Получаем количество символов местозаполнителя для формулы.
            List<ChapterInfo> chapters = DocumentUtilities.FindNearChaptersByTemplate(doc, SymbolType.Formula);
            int sCount = rs.DocumentSettings.GetSymbolCount(SymbolType.Formula);


            Word.Find find = findRange.Find;
            find.Text = "$";
            find.Wrap = Word.WdFindWrap.wdFindStop;
            Word.Field fmField = null;

            string forNumParts = string.Empty;

            if (sCount == 1)
            {
                find.Execute();

                Word.Range hit = findRange.Duplicate;
                Word.Range dollar = hit.Duplicate;
                dollar.End = dollar.Start + 1;
                dollar.Text = "";
                fmField = targetRange.Fields.Add(
                    dollar,
                    Word.WdFieldType.wdFieldSequence,
                    @"Formula_seq"
                );
                forNumParts = fmField.Result.Text;
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
                        fmField = targetRange.Fields.Add(
                            dollar,
                            Word.WdFieldType.wdFieldSequence,
                            $"Formula_{string.Join("", nums)}_seq"
                        );
                        forNumParts += fmField.Result.Text;
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
            string bmName = string.Join("", "_fm_", Guid.NewGuid().ToString("N"));
            List<string> pts = DocumentUtilities.ExtractNumbers(forNumParts);
            FormulaInfo fm = new FormulaInfo(pts);
            fm.BookmarkName = bmName;
            Word.Range bmRange = table.Cell(index, 3).Range;
            bmRange.End -= 1;
            try
            {
                rs.FormulaList.Formulas.Add(fm);
                targetRange.End -= 1;
                doc.Bookmarks.Add(bmName, bmRange);
                // doc.Fields.Update();
                doc.Content.Fields.Update();
            }
            catch (Exception ex)
            {
                CommonCommands.ShowError(ex.ToString(), "Неизвестная ошибка!");
            }
        }
        finally
        {
            doc.Application.ScreenUpdating = previousScreenUpdating;
            doc.Application.Options.Pagination = prevPagination;
            doc.ActiveWindow.View.Type = oldView;
            Update();
            EventBus<FormulaInfo>.Raise(doc);
        }
    }

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
            doc.Content.Fields.Update();
            // Обновляем только поля SEQ и REF в основном тексте (включая таблицы основного текста)
            foreach (FormulaInfo fm in rs.FormulaList.Formulas)
            {
                Word.Bookmark bm = doc.Bookmarks[fm.BookmarkName];
                if (bm.Range.Fields.Count == 1)
                {
                    bool res = int.TryParse(bm.Range.Fields[1].Result.Text, out int num);
                    List<string> nums = new List<string>() { num.ToString() };
                    fm.NumberParts.Clear();
                    fm.NumberParts.AddRange(nums);
                }
                else
                {
                    List<string> nums = fm.Number(bm.Range);
                    fm.NumberParts.Clear();
                    fm.NumberParts.AddRange(nums);
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

    public void DeleteFormula()
    {
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;
        Word.Document doc = rs.Document;
        Word.Selection selection = doc.Application.Selection;
       
        //Проверка, что курсор в таблице
        if (selection.Tables.Count == 0)
        {
            CommonCommands.ShowError("Курсор не находится в таблице!", "Ошибка!");
            return;
        }

        //Проверка, что 3 столбца и в 3 есть закладка с "_fm_" в названии закладки
        if (selection.Tables[1].Columns.Count != 3)
        {
            CommonCommands.ShowError("Формат таблицы не тот!", "Ошибка!");
            return;
        }

        int rowIndex = (int)selection.Information[Word.WdInformation.wdStartOfRangeRowNumber];
        if (selection.Tables[1].Cell(rowIndex, 3).Range.Bookmarks.Count < 1)
        {
            CommonCommands.ShowError("В 3 столбце нет закладки (соответственно номера формулы)!", "Ошибка!");
            return;
        }

        Word.Bookmarks cellBookmarks = selection.Tables[1].Cell(rowIndex, 3).Range.Bookmarks;
        int cnt = 0;
        string bmName = string.Empty;
        foreach (Word.Bookmark bm in cellBookmarks)
        {
            // Сразу пропускаем системные скрытые закладки (_Toc, _Ref и т.д.)
            string name = bm.Name;
            if (name.Contains("_fm_"))
            {
                bmName = bm.Name;
                cnt++;
            }
        }

        if (cnt == 0 || cnt > 1)
        {
            CommonCommands.ShowError("Это закладка не формулы или!", "Ошибка!");
            return;
        }
        //Удаление закладки, удаление из общего хранилища и удаление самого текста и строки в таблице

        Word.Table table = selection.Tables[1];
        bool previousScreenUpdating = doc.ActiveWindow.Application.ScreenUpdating;
        bool prevPagination = doc.ActiveWindow.Application.Options.Pagination;
        Word.WdViewType oldView = doc.ActiveWindow.View.Type;
        doc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;
        doc.ActiveWindow.Application.ScreenUpdating = false;
        doc.ActiveWindow.Application.Options.Pagination = false;
        try
        {
            
            rs.FormulaList.Formulas.RemoveAll(x => x.BookmarkName == bmName);
            doc.Bookmarks[bmName].Delete();
            table.Rows[rowIndex].Delete();
            DocumentUtilities.ReplaceAllLinks(doc, LinkType.Formula, bmName);
        }
        catch (Exception ex)
        {
            CommonCommands.ShowError(ex.ToString(), "Непредвиденная ошибка!");
        }
        finally
        {
            Update();
            doc.Application.ScreenUpdating = previousScreenUpdating;
            doc.Application.Options.Pagination = prevPagination;
            doc.ActiveWindow.View.Type = oldView;
            EventBus<FormulaInfo>.Raise(doc);
        }
    }

    public void InsertFormula()
    {
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;
        Word.Document doc = rs.Document;
        ServiceLocator srv = Singleton<ServiceLocator>.Instance;
        if (rs.FormulaList.Formulas.Count == 0)
        {
            CommonCommands.ShowWarning("Формулы не найдены!", "Предупреждение");
            return;
        }

        FormulaInfo frm = (FormulaInfo)srv.FormulaPickerService.PickBookmark(doc);
        if (frm == null)
        {
            return;
        }

        Word.Range targetRange = doc.Application.Selection.Range.Duplicate;
        targetRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);


        targetRange.Hyperlinks.Add(
            Anchor: targetRange,
            SubAddress: frm.BookmarkName,
            TextToDisplay: frm.InsertText,
            ScreenTip: frm.DisplayText
        );
    }
}