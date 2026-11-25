using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using XtendedRefs.ComAddIn.Application.Commands;
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

    public static void RescanDocument2(Word.Document doc, List<LevelStyleItem> styles)
    {
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;
        rs.ChaptersList.Chapters.Clear();
        rs.ChaptersList.Chapters = ScanChapters(rs.Document, rs.DocumentSettings.LevelStyles);
    }
    
    
    public static void RescanDocument(Word.Document doc, List<LevelStyleItem> styles)
    {
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;
        List<ChapterInfo> chapters = rs.ChaptersList.Chapters;

        var existingRanges = new HashSet<(int start, int end)>();
        var chapterPositions = new Dictionary<ChapterInfo, int>();
        
        
        for (int i = chapters.Count - 1; i >= 0; i--)
        {
            ChapterInfo ch = chapters[i];
            
            if (string.IsNullOrEmpty(ch.BookmarkName) || !doc.Bookmarks.Exists(ch.BookmarkName))
            {
                chapters.RemoveAt(i);
                continue;
            }

            Word.Range r = doc.Bookmarks[ch.BookmarkName].Range;
            existingRanges.Add((r.Start, r.End));
            chapterPositions[ch] = r.Start;
            var currentNumbers = ExtractNumbers(r.ListFormat.ListString);
            if (ch.NumberParts == null || !ch.NumberParts.SequenceEqual(currentNumbers))
            {
                ch.NumberParts = currentNumbers;
            }
        }
        
        CommonCommands.ShowInfo(chapters.Count.ToString());
        
        for (int i = 0; i < styles.Count; i++)
        {
            string styleName = styles[i].StyleName;
            Word.Style targetStyle;
            try
            {
                targetStyle = doc.Styles[styleName];
            }
            catch
            {
                continue;
            }

            Word.Range rng = doc.Content;
            Word.Find find = rng.Find;
            find.ClearFormatting();
            find.Text = "";
            find.set_Style(targetStyle);
            find.Forward = true;
            find.Wrap = Word.WdFindWrap.wdFindStop;
            while (find.Execute())
            {
                Word.Range paraRange = rng.Paragraphs.First.Range;

                (int Start, int End) key = (paraRange.Start, paraRange.End);
            
                if (!existingRanges.Contains(key))
                {
                    string bmName = "_ch_" + Guid.NewGuid().ToString("N");
                    doc.Bookmarks.Add(bmName, paraRange);

                    string cleanTitle = paraRange.Text?.TrimEnd('\r', '\a', '\n') ?? string.Empty;

                    var chInfo = new ChapterInfo(cleanTitle, styleName, i)
                    {
                        BookmarkName = bmName,
                        NumberParts = ExtractNumbers(paraRange.ListFormat.ListString)
                    };

                    chapters.Add(chInfo);
                    existingRanges.Add(key);
                    chapterPositions[chInfo] = paraRange.Start;
                }

                rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            }
        }
        
        CommonCommands.ShowInfo(chapters.Count.ToString());
        
        if (chapters.Count == 0)
        {
            return;
        }

        chapters.Sort((a, b) => chapterPositions[a].CompareTo(chapterPositions[b]));
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
        doc.Application.ScreenUpdating = false; // Отключаем прорисовку интерфейса
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
            MessageBox.Show($@"Ошибка при обновлении: {ex.Message}", @"Ошибка", MessageBoxButtons.OK,
                MessageBoxIcon.Error);
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
        int sCount = rs.DocumentSettings.GetSymbolCount(type);
        int position = doc.ActiveWindow.Selection.Range.Start;
        var chaptersBeforeCursor = new List<(ChapterInfo Info, int Start)>();
        foreach (var c in rs.ChaptersList.Chapters)
        {
            if (doc.Bookmarks.Exists(c.BookmarkName))
            {
                int start = doc.Bookmarks[c.BookmarkName].Range.Start;
                if (start <= position)
                {
                    chaptersBeforeCursor.Add((c, start));
                }
            }
        }

        chaptersBeforeCursor = chaptersBeforeCursor.OrderBy(x => x.Start).ToList();
        var hierarchy = new List<(ChapterInfo Info, int Start)>();
        foreach (var item in chaptersBeforeCursor)
        {
            int existingIndex = hierarchy.FindIndex(h => h.Info.StyleName == item.Info.StyleName);
            if (existingIndex != -1)
            {
                hierarchy.RemoveRange(existingIndex, hierarchy.Count - existingIndex);
            }

            hierarchy.Add(item);
        }

        List<ChapterInfo> nearest = hierarchy
            .Take(sCount-1)
            .Select(x => x.Info)
            .ToList();

        if (nearest.Count != sCount-1 && sCount >= 1)
        {
            throw new Exception("Количество глав не совпадает с шаблоном!");
        }

        return nearest;
    }


    public static int CountSymbols(char symbol, string countString)
    {
        return countString.Count(s => s == symbol);
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


        for (int i = hyperlinks.Count - 1; i >= 0; i--)
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

    public static List<string> GetParagraphStyleNames(Word.Document doc)
    {
        List<string> styles = new List<string>();

        foreach (Word.Style style in doc.Styles)
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