using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using XtendedRefs.ComAddIn.Domain.Abstractions;
using XtendedRefs.ComAddIn.Domain.Documents;
using XtendedRefs.ComAddIn.Domain.Models.NewModels;
using XtendedRefs.ComAddIn.Domain.State;
using Word = Microsoft.Office.Interop.Word;

namespace XtendedRefs.ComAddIn.Application.Utilities;

public enum LinkType
{
    Literature,
    Formula,
    Picture,
    Table
}

public static class DocumentUtilities
{
    public static List<ChapterInfo> ScanChapters(Word.Document doc, List<LevelStyleItem> styles)
    {
        List<ChapterInfo> ch = new List<ChapterInfo>();

        for (int i = 0; i < styles.Count; i++)
        {
            Word.Range rng = doc.Content;
            Word.Find finding = rng.Find;

            finding.ClearFormatting();
            finding.Text = "";
            finding.set_Style(doc.Styles[styles[i].StyleName]);
            finding.Forward = true;
            finding.Wrap = Word.WdFindWrap.wdFindStop;

            while (finding.Execute())
            {
                Word.Range found = rng.Duplicate;
                Word.Paragraph paragraph = found.Paragraphs[1];
                Word.Range paraRange = paragraph.Range;

                ChapterInfo chInfo = new ChapterInfo(paraRange.Text, styles[i].StyleName, i);

                // Добавляем закладку на начало главы
                string bmName = "_ch_" + Guid.NewGuid().ToString("N");
                // Word.Range startRange = paraRange.Duplicate;
                // startRange.End = paraRange.Start;
                Word.Bookmark bookmark = doc.Bookmarks.Add(bmName, paraRange);
                chInfo.BookmarkName = bookmark.Name;

                //Добавим части номера
                chInfo.NumberParts = ExtractNumbers(paraRange.ListFormat.ListString);

                ch.Add(chInfo);

                int nextStart = paraRange.End;

                rng.SetRange(paraRange.End, doc.Content.End);
                finding = rng.Find;
                finding.ClearFormatting();
                finding.Text = "";
                finding.set_Style(doc.Styles[styles[i].StyleName]);
                finding.Forward = true;
                finding.Wrap = Word.WdFindWrap.wdFindStop;
                if (nextStart >= doc.Content.End)
                    break;
            }
        }

        return ch;
    }


    public static int CompareParts(List<string> a, List<string> b)
    {
        int len = Math.Min(a.Count, b.Count);

        for (int i = 0; i < len; i++)
        {
            int cmp = int.Parse(a[i]).CompareTo(int.Parse(b[i]));
            if (cmp != 0)
                return cmp;
        }

        return a.Count.CompareTo(b.Count);
    }

    public static void RescanDocument(Word.Document doc, List<LevelStyleItem> styles)
    {
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;

        if (rs.ChaptersList.Chapters.Count == 0)
        {
            MessageBox.Show(@"No chapters found");
            return;
        }


        var existingRanges = new HashSet<(int start, int end)>(
            rs.ChaptersList.Chapters
                .Where(c => c.BookmarkName != null && doc.Bookmarks.Exists(c.BookmarkName))
                .Select(c =>
                {
                    Word.Range r = doc.Bookmarks[c.BookmarkName].Range;
                    return (r.Start, r.End);
                })
        );

        //TODO:Добавление в случае удаления какой-либо из глав удаление ее и из списка
        for (int i = 0; i < styles.Count; i++)
        {
            Word.Range rng = doc.Content;
            Word.Find find = rng.Find;

            find.ClearFormatting();
            find.Text = "";
            find.set_Style(doc.Styles[styles[i].StyleName]);
            find.Forward = true;
            find.Wrap = Word.WdFindWrap.wdFindStop;

            while (find.Execute())
            {
                Word.Range foundRange = rng.Duplicate;
                Word.Paragraph paragraph = foundRange.Paragraphs.First;
                Word.Range paraRange = paragraph.Range;

                var key = (paraRange.Start, paraRange.End);

                if (!existingRanges.Contains(key))
                {
                    string bmName = "_ch_" + Guid.NewGuid().ToString("N");

                    doc.Bookmarks.Add(bmName, paraRange);

                    ChapterInfo chInfo = new ChapterInfo(
                        paraRange.Text,
                        styles[i].StyleName,
                        i
                    );

                    chInfo.BookmarkName = bmName;

                    chInfo.NumberParts = ExtractNumbers(
                        paraRange.ListFormat.ListString
                    );

                    rs.ChaptersList.Chapters.Add(chInfo);

                    existingRanges.Add(key);
                }

                int nextStart = paraRange.End;
                rng.SetRange(nextStart, doc.Content.End);

                find = rng.Find;
                find.ClearFormatting();
                find.Text = "";
                find.set_Style(doc.Styles[styles[i].StyleName]);
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindStop;

                if (nextStart >= doc.Content.End)
                    break;
            }
        }

        foreach (var chapter in rs.ChaptersList.Chapters)
        {
            if (!doc.Bookmarks.Exists(chapter.BookmarkName))
                continue;

            var range = doc.Bookmarks[chapter.BookmarkName].Range;

            var current = ExtractNumbers(range.ListFormat.ListString);

            if (!chapter.NumberParts.SequenceEqual(current))
            {
                chapter.NumberParts = current;
            }
        }

        rs.ChaptersList.Chapters.Sort((a, b) =>
            a.Level != b.Level
                ? a.Level.CompareTo(b.Level)
                : CompareParts(a.NumberParts, b.NumberParts)
        );
    }


    public static List<string> ExtractNumbers(string input)
    {
        Match match = Regex.Match(input, @"\d+(?:\.\d+)*");

        if (!match.Success)
            return new List<string>();

        return match.Value.Split('.').ToList();
    }

    public static void RemoveBookmarksByPrefix(Word.Document doc, string prefix)
    {
        var bookmarksToDelete = doc.Bookmarks
            .Cast<Word.Bookmark>()
            .Where(b => b.Name.StartsWith(prefix))
            .Select(b => b.Name)
            .ToList();

        foreach (var name in bookmarksToDelete)
        {
            doc.Bookmarks[name].Delete();
        }
    }


    /// <summary>
    /// Метод, возвращающий словарь, где каждой главе соответствует список ссылок в тексте. Необходимо для обновления. Главы, не имеющие ссылок
    /// не попадают в результирующий словарь
    /// </summary>
    /// <param name="doc">Документ, в котором необходимо найти ссылки (по факту текущий открытый)</param>
    /// <returns>
    /// Возвращает словарь глава-список ссылок.
    /// <see cref="Dictionary{TKey, TValue}"/>:
    /// <list type="bullet">
    /// <item>Key: <see cref="ChapterInfo"/></item>
    /// <item>Value: <see cref="Word.Hyperlink"/> list</item>
    /// </list>
    /// </returns>
    private static Dictionary<ChapterInfo, List<Word.Hyperlink>> GetChaptersLinks(Word.Document doc)
    {
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;

        List<Word.Hyperlink> hyperlinks = doc.Hyperlinks
            .Cast<Word.Hyperlink>()
            .Where(h => !string.IsNullOrEmpty(h.SubAddress))
            .ToList();

        Dictionary<string, List<Word.Hyperlink>> groupedLinks = hyperlinks
            .GroupBy(h => h.SubAddress)
            .ToDictionary(g => g.Key, g => g.ToList());

        Dictionary<ChapterInfo, List<Word.Hyperlink>> chaptersWithLinks = rs.ChaptersList.Chapters
            .Where(c => groupedLinks.ContainsKey(c.BookmarkName))
            .ToDictionary(
                c => c,
                c => groupedLinks[c.BookmarkName]
            );
        return chaptersWithLinks;
    }


    private static bool UpdateHyperlinks<T>(Word.Document doc, Dictionary<T, List<Word.Hyperlink>> links)
        where T : DocumentAnchor
    {
        Dictionary<T, Word.Range> ranges = links.ToDictionary<KeyValuePair<T, List<Word.Hyperlink>>, T, Word.Range>(
            kvp => kvp.Key,
            kvp => doc.Bookmarks[kvp.Key.BookmarkName].Range.Duplicate
        );


        foreach (KeyValuePair<T, List<Word.Hyperlink>> kvp in links)
        {
            if (!ranges.TryGetValue(kvp.Key, out Word.Range range))
                continue;

            string newNum = kvp.Key.GetNumber(range);
            string newTip = kvp.Key.DisplayText;
            foreach (var link in kvp.Value)
            {
                link.TextToDisplay = newNum;
                link.ScreenTip = newTip;
            }
        }

        return true;
    }

    public static bool UpdateChapterHyperLinks(Word.Document doc)
    {
        Dictionary<ChapterInfo, List<Word.Hyperlink>> links = GetChaptersLinks(doc);
        return UpdateHyperlinks(doc, links);
    }

    public static bool UpdateStyles(Word.Document doc)
    {
        Word.Style normal = doc.Styles[Word.WdBuiltinStyle.wdStyleHyperlink];
        Word.Style visited = doc.Styles[Word.WdBuiltinStyle.wdStyleHyperlinkFollowed];

        normal.Font.ColorIndex = Word.WdColorIndex.wdAuto;
        normal.Font.Underline = Word.WdUnderline.wdUnderlineNone;

        visited.Font.ColorIndex = Word.WdColorIndex.wdAuto;
        visited.Font.Underline = Word.WdUnderline.wdUnderlineNone;
        return true;
    }

    public static bool UpdateLiteratureFields(Word.Document doc)
    {
        doc.Application.ScreenUpdating = false;          // Отключаем прорисовку интерфейса
        doc.Application.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone; // Отключаем всплывающие окна
        doc.ActiveWindow.View.ShowFieldCodes = true; // Включаем коды полей, чтобы Find их видел
        try
        {
            Word.Range range = doc.Content;
            Word.Find find = range.Find;

            find.ClearFormatting();
            find.Text = $"^d Literature_Seq"; 
            find.Forward = true;
            find.Wrap = Word.WdFindWrap.wdFindStop;
            while (find.Execute())
            {
                if (range.Fields.Count > 0)
                {
                    range.Fields[1].Update(); 
                }
                
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(find);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
        }
        catch (Exception ex)
        {
            MessageBox.Show($@"Ошибка при обновлении: {ex.Message}", @"Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            doc.ActiveWindow.View.ShowFieldCodes = false;
            doc.Application.DisplayAlerts = Word.WdAlertLevel.wdAlertsAll;
            doc.Application.ScreenUpdating = true;
        }
        
        //----------------
        
        
        return true;
    }

    private static Dictionary<LiteratureInfo, List<Word.Hyperlink>> GetLiteratureLinks(Word.Document doc)
    {
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;

        List<Word.Hyperlink> hyperlinks = doc.Hyperlinks
            .Cast<Word.Hyperlink>()
            .Where(h => !string.IsNullOrEmpty(h.SubAddress))
            .ToList();

        Dictionary<string, List<Word.Hyperlink>> groupedLinks = hyperlinks
            .GroupBy(h => h.SubAddress)
            .ToDictionary(g => g.Key, g => g.ToList());

        Dictionary<LiteratureInfo, List<Word.Hyperlink>> literatureWithLinks = rs.LiteratureList.Literature
            .Where(c => groupedLinks.ContainsKey(c.BookmarkName))
            .ToDictionary(
                c => c,
                c => groupedLinks[c.BookmarkName]
            );
        return literatureWithLinks;
    }

    public static bool UpdateLiteratureHyperLinks(Word.Document doc)
    {
        Dictionary<LiteratureInfo, List<Word.Hyperlink>> links = GetLiteratureLinks(doc);
        return UpdateHyperlinks(doc, links);
    }
    
    
    private static Dictionary<PictureInfo, List<Word.Hyperlink>> GetPictureLinks(Word.Document doc)
    {
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;

        List<Word.Hyperlink> hyperlinks = doc.Hyperlinks
            .Cast<Word.Hyperlink>()
            .Where(h => !string.IsNullOrEmpty(h.SubAddress))
            .ToList();

        Dictionary<string, List<Word.Hyperlink>> groupedLinks = hyperlinks
            .GroupBy(h => h.SubAddress)
            .ToDictionary(g => g.Key, g => g.ToList());

        Dictionary<PictureInfo, List<Word.Hyperlink>> pictureWithLinks = rs.PictureList.Pictures
            .Where(c => groupedLinks.ContainsKey(c.BookmarkName))
            .ToDictionary(
                c => c,
                c => groupedLinks[c.BookmarkName]
            );
        return pictureWithLinks;
    }

    public static bool UpdatePictureHyperLinks(Word.Document doc)
    {
        Dictionary<PictureInfo, List<Word.Hyperlink>> links = GetPictureLinks(doc);
        return UpdateHyperlinks(doc, links);
    }
    
    
    private static Dictionary<FormulaInfo, List<Word.Hyperlink>> GetFormulaLinks(Word.Document doc)
    {
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;

        List<Word.Hyperlink> hyperlinks = doc.Hyperlinks
            .Cast<Word.Hyperlink>()
            .Where(h => !string.IsNullOrEmpty(h.SubAddress))
            .ToList();

        Dictionary<string, List<Word.Hyperlink>> groupedLinks = hyperlinks
            .GroupBy(h => h.SubAddress)
            .ToDictionary(g => g.Key, g => g.ToList());

        Dictionary<FormulaInfo, List<Word.Hyperlink>> formulaWithLinks = rs.FormulaList.Formulas
            .Where(c => groupedLinks.ContainsKey(c.BookmarkName))
            .ToDictionary(
                c => c,
                c => groupedLinks[c.BookmarkName]
            );
        return formulaWithLinks;
    }

    public static bool UpdateFormulaHyperLinks(Word.Document doc)
    {
        Dictionary<FormulaInfo, List<Word.Hyperlink>> links = GetFormulaLinks(doc);
        return UpdateHyperlinks(doc, links);
    }
    
    
    private static Dictionary<TableInfo, List<Word.Hyperlink>> GetTableLinks(Word.Document doc)
    {
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;

        List<Word.Hyperlink> hyperlinks = doc.Hyperlinks
            .Cast<Word.Hyperlink>()
            .Where(h => !string.IsNullOrEmpty(h.SubAddress))
            .ToList();

        Dictionary<string, List<Word.Hyperlink>> groupedLinks = hyperlinks
            .GroupBy(h => h.SubAddress)
            .ToDictionary(g => g.Key, g => g.ToList());

        Dictionary<TableInfo, List<Word.Hyperlink>> tableWithLinks = rs.TableList.Tables
            .Where(c => groupedLinks.ContainsKey(c.BookmarkName))
            .ToDictionary(
                c => c,
                c => groupedLinks[c.BookmarkName]
            );
        return tableWithLinks;
    }

    public static bool UpdateTableHyperLinks(Word.Document doc)
    {
        Dictionary<TableInfo, List<Word.Hyperlink>> links = GetTableLinks(doc);
        return UpdateHyperlinks(doc, links);
    }
    

    public static List<ChapterInfo> FindNearChaptersByTemplate(Word.Document doc, SymbolType type)
    {
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;

        DocumentSettings docSettings = rs.DocumentSettings;
        int sCount = docSettings.GetSymbolCount(type);
        Word.Range sel = doc.ActiveWindow.Selection.Range;
        int position = sel.Start;


        List<ChapterInfo> nearest = rs.ChaptersList.Chapters
            .Where(c =>
            {
                var bookmark = doc.Bookmarks[c.BookmarkName];
                return bookmark != null && (int)bookmark.Range.Start <= position;
            })
            .GroupBy(c => c.StyleName)
            .Select(group => group
                .OrderByDescending(c => (int)doc.Bookmarks[c.BookmarkName].Range.Start)
                .First()
            )
            .OrderBy(c => (int)doc.Bookmarks[c.BookmarkName].Range.Start)
            .Take(sCount)
            .ToList();

        if (nearest.Count != sCount & sCount > 1)
        {
            throw new Exception("Количество глав не совпадает с шаблоном. Очевидно, что курсор стоит не в том месте!");
        }
        
        return nearest;
    }

    public static int CountSymbols(char symbol, string countString)
    {
        return countString.Count(s=> s == symbol);
    }

    public static string GetPrefix(LinkType linkType)
    {
        switch (linkType)
        {
            case LinkType.Literature:
                return "_lt_";
            case LinkType.Formula:
                return "_fm_";
            case LinkType.Picture:
                return "_pc_";
            case LinkType.Table:
                return "_tb_";
        }
        return string.Empty;
    }
    
    public static void ReplaceAllLinks(Word.Document doc, LinkType linkType, string bookmarkName)
    {
        
        List<Word.Hyperlink> hyperlinks = doc.Hyperlinks
            .Cast<Word.Hyperlink>()
            .Where(h => h.SubAddress.Contains(bookmarkName))
            .ToList();
        
                
                for (int i = hyperlinks.Count-1; i >=0 ; i--)
                {
                    Word.Hyperlink link = hyperlinks[i];
                    string newtext = string.Empty;
                    switch (linkType)
                    {
                        case LinkType.Literature:
                            newtext = string.Join(" ", "Здесь была ссылка на бывший источник литературы", link.TextToDisplay);
                            break;
                       
                        case LinkType.Formula:
                            newtext = string.Join(" ", "Здесь была ссылка на бывшую формулу", link.TextToDisplay);
                            break;
                        
                        case LinkType.Picture:
                            newtext = string.Join(" ", "Здесь была ссылка на бывший рисунок", link.TextToDisplay);
                            break;
                        
                        case LinkType.Table:
                            newtext = string.Join(" ", "Здесь была ссылка на бывшую таблицу", link.TextToDisplay);
                            break;
                    }

                    link.TextToDisplay = newtext;
                    link.Range.HighlightColorIndex = Word.WdColorIndex.wdYellow;
                    link.Delete();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(link);
                    
        }
        
    }
}