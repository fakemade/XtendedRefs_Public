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

public class TableCommands
{
    // ---------- CreateTable (AddTable) ----------
    public void CreateTable()
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

        try
        {
            // Собираем шаблон, стиль и разделитель
            string template = documentSettings.TableSignTemplate;
            string styleName = documentSettings.TableStyle;

            // Берём место курсора и параграф
            Word.Selection selection = doc.ActiveWindow.Selection;
            Word.Paragraph paragraph = selection.Paragraphs[1];
            paragraph.KeepWithNext = -1;
            
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

            

            Word.Range targetRange = paragraph.Range.Duplicate;
            Word.Range findRange = paragraph.Range.Duplicate;
            findRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            // Вставили шаблон
            findRange.Text = template;

            // Ищем $ для вставки поля.
            List<ChapterInfo> chapters = DocumentUtilities.FindNearChaptersByTemplate(doc, SymbolType.Table);
            
            CommonCommands.ShowInfo(chapters.Count.ToString());
            
            Word.Find find = findRange.Find;
            find.Text = "$";
            find.Wrap = Word.WdFindWrap.wdFindStop;
            Word.Field tbField = null;
            string forNumParts = string.Empty;

            if (chapters.Count == 0 || chapters.Count == 1)
            {
                find.Execute();

                Word.Range hit = findRange.Duplicate;
                Word.Range dollar = hit.Duplicate;
                dollar.End = dollar.Start + 1;
                dollar.Text = "";
                tbField = targetRange.Fields.Add(
                    dollar,
                    Word.WdFieldType.wdFieldSequence,
                    @"Table_seq"
                );
                forNumParts = tbField.Result.Text;
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
                        tbField = targetRange.Fields.Add(
                            dollar,
                            Word.WdFieldType.wdFieldSequence,
                            $"Table_{string.Join("", nums)}_seq"
                        );
                        forNumParts += tbField.Result.Text;
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
            string bmName = string.Join("", "_tb_", Guid.NewGuid().ToString("N"));
            List<string> pts = DocumentUtilities.ExtractNumbers(forNumParts);

            TableInfo tbl = new TableInfo(pts, (targetRange.Start, targetRange.End));
            tbl.BookmarkName = bmName;
            try
            {
                rs.TableList.Tables.Add(tbl);
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
            EventBus<TableInfo>.Raise(doc);
        }
    }

    // ---------- Обновление номерных частей (SEQ / HYPERLINK) из закладок в памяти ----------
    private void Update()
    {
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;
        Word.Document doc = rs.Document;
        doc.Content.Fields.Update();
        
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
            foreach (TableInfo tbl in rs.TableList.Tables)
            {
                if (!doc.Bookmarks.Exists(tbl.BookmarkName)) continue;
                Word.Bookmark bm = doc.Bookmarks[tbl.BookmarkName];
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
                    tbl.NumberParts.Clear();
                    tbl.NumberParts.AddRange(nums);
                    //empty = pic.DisplayText;
                }
                else
                {
                    List<string> nums = tbl.Number(bm.Range);
                    tbl.NumberParts.Clear();
                    tbl.NumberParts.AddRange(nums);
                    //empty = pic.DisplayText;
                }
            }
        }
        catch (Exception ex)
        {
            CommonCommands.ShowError(ex.ToString());
        }
        finally
        {
            doc.Application.ScreenUpdating = previousScreenUpdating;
            doc.Application.Options.Pagination = prevPagination;
            doc.ActiveWindow.View.Type = oldView;
        }
        doc.Fields.Update();
    }

    // ---------- DeleteTableCaption (удаление зарегистрированной таблицы) ----------
    public void DeleteTableCaption()
    {
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;
        Word.Document doc = rs.Document;

        // Параграф, в котором удаляем ссылку на таблицу
        Word.Paragraph paragraph = doc.Application.Selection.Paragraphs[1];
        Word.Range paraRange = paragraph.Range.Duplicate;

        if (paraRange.Fields.Count == 0)
        {
            CommonCommands.ShowError("Здесь нет таблицы", "Ошибка!");
            return;
        }

        // Ищем закладку "_tb_..." в текущем параграфе
        string bmName = paraRange.Bookmarks
            .Cast<Word.Bookmark>()
            .Where(b => b.Name.Contains("_tb_"))
            .Select(b => b.Name)
            .FirstOrDefault();

        if (string.IsNullOrWhiteSpace(bmName))
        {
            CommonCommands.ShowError("Здесь нет таблицы", "Ошибка!");
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
            DocumentUtilities.ReplaceAllLinks(doc, LinkType.Table, bmName);
            Word.Bookmark bm = doc.Bookmarks[bmName];
            doc.Range(bm.Range.Start, bm.Range.End).Delete();
            rs.TableList.Tables.RemoveAll(l => l.BookmarkName == bmName);
        }
        catch (Exception e)
        {
            MessageBox.Show(e.ToString());
        }
        finally
        {
            Update();
            doc.Application.ScreenUpdating = previousScreenUpdating;
            doc.Application.Options.Pagination = prevPagination;
            doc.ActiveWindow.View.Type = oldView;
            EventBus<TableInfo>.Raise(doc);
        }
        Update();
        EventBus<TableInfo>.Raise(doc);
    }

    // ---------- InsertTableReference (вставка гиперссылки на таблицу) ----------
    public void InsertTableReference()
    {
        
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;
        Word.Document doc = rs.Document;
        ServiceLocator srv = Singleton<ServiceLocator>.Instance;

        if (rs.TableList.Tables.Count == 0)
        {
            CommonCommands.ShowWarning("Таблицы не найдены!", "Предупреждение");
            return;
        }

        // Подбираем картинку из списка (сервис PicturePickerService будет зарегистрирован в ServiceLocator)
        TableInfo table = (TableInfo)srv.TablePickerService.PickBookmark(doc);
        if (table == null)
        {
            return;
        }


        Word.Range targetRange = doc.Application.Selection.Range.Duplicate;
        targetRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);

        // Вставляем гиперссылку на закладку рисунка с номером как отображаемым текстом
        targetRange.Hyperlinks.Add(
            Anchor: targetRange,
            SubAddress: table.BookmarkName,
            TextToDisplay: table.InsertText,
            ScreenTip: string.Join("",table.NumberParts)
        );
    }

    // ---------- Вспомогательные методы ----------

    /// <summary>
    /// Извлекает текст подписи из первой ячейки выбранной таблицы.
    /// Обычно это параграф с текстом вида «Таблица X» или просто описание.
    /// </summary>
    private static string GetCaptionFromTable(Word.Table table)
    {
        // Берём первый параграф из первой ячейки (или второго, если первая — номер)
        Word.Paragraph para = table.Rows[1].Cells[1].Range.Paragraphs[1];
        return para.Range.Text.Trim().Replace("\r", "").Trim();
    }

    /// <summary>
    /// Получает все таблицы из документа для хранения в TableInfo.Tables.
    /// Используется при создании элемента: сохраняем ссылку на физический объект Word.Table.
    /// </summary>
    private static List<Word.Table> GetTablesFromDocument(Word.Document doc)
    {
        return doc.Tables.Cast<Word.Table>().ToList();
    }
}
