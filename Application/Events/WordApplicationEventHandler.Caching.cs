using System;
using System.Collections.Generic;
using System.IO;
using XtendedRefs.ComAddIn.Domain.Abstractions;
using XtendedRefs.ComAddIn.Domain.Models;
using XtendedRefs.ComAddIn.Domain.State;
using XtendedRefs.ComAddIn.Infrastructure.Documents;
using XtendedRefs.ComAddIn.Infrastructure.Logging;
using Word = Microsoft.Office.Interop.Word;

namespace XtendedRefs.ComAddIn.Application.Events;

internal sealed partial class WordApplicationEventHandler
{
    private void CacheSilently(Word.Document document)
    {
        bool hasFigures = GetFigureBookmarkNames(document).Count > 0;
        bool hasFormulas = false;
        for (int i = 1; i <= document.Bookmarks.Count; i++)
        {
            string name = document.Bookmarks[i].Name;
            if (name.StartsWith("Formula_", StringComparison.Ordinal))
                hasFormulas = true;

            if (hasFigures && hasFormulas)
                break;
        }
        if (hasFigures)
        {
            CacheFigures(document);
        }

        if (hasFormulas)
        {
            CacheFormulas(document);
        }
    }

    private void CacheFigures(Word.Document document)
    {
        string folder = Path.Combine(Path.GetTempPath(), "WordFigures");
        Directory.CreateDirectory(folder);
        int savedStart = _application.Selection.Start;
        int savedEnd = _application.Selection.End;

        int processedCount = 0;
        foreach (string bookmarkName in GetFigureBookmarkNames(document))
        {
            processedCount++;
            if (File.Exists(Path.Combine(folder, bookmarkName + ".emf")) ||
                File.Exists(Path.Combine(folder, bookmarkName + ".png")))
            {
                continue;
            }

            try
            {
                _clipboardEmfService.GenerateSingleFigureEmf(document, bookmarkName,
                    (bm, hash) => { Singleton<DocumentCommandUtilities>.Instance.SaveFigureHash(bm, hash); return true; });
            }
            catch (Exception ex)
            {
            }
        }

        try
        {
            document.Range(savedStart, savedEnd).Select();
        }
        catch (Exception ex)
        {
        }
        
    }

    private List<string> GetFigureBookmarkNames(Word.Document document)
    {
        List<string> result = new List<string>();

        try
        {
            
        }
        catch (Exception ex)
        {
        }

        if (result.Count > 0)
        {
            return result;
        }

        // Fallback для документов, где рисунки еще созданы старой схемой Figure_/FigureNum_.
        List<string> legacyNames = _bookmarkNavigator.GetNames(document, "Figure_", "FigureNum_");
        return legacyNames;
    }

    private void CacheFormulas(Word.Document document)
    {
        string folder = Path.Combine(Path.GetTempPath(), "WordFormulas");
        Directory.CreateDirectory(folder);
        int savedStart = _application.Selection.Start;
        int savedEnd = _application.Selection.End;

        List<Tuple<string, long>> formulaPositions = new List<Tuple<string, long>>();
        for (int i = 1; i <= document.Bookmarks.Count; i++)
        {
            Word.Bookmark bm = document.Bookmarks[i];
            if (bm.Name.StartsWith("Formula_", StringComparison.Ordinal))
            {
                try
                {
                    formulaPositions.Add(Tuple.Create(bm.Name, (long)bm.Range.Start));
                }
                catch (Exception ex)
                {
                }
            }
        }

        Dictionary<string, int> rowIndex = BuildFormulaRowIndex(document, formulaPositions);

        foreach (Tuple<string, long> fp in formulaPositions)
        {
            string emfPath = Path.Combine(folder, fp.Item1 + ".emf");
            if (File.Exists(emfPath))
            {
                continue;
            }
            
            try
            {
                Word.Table table = FindTableContainingPosition(document, fp.Item2);
                if (table == null)
                    continue;

                if (!rowIndex.TryGetValue(fp.Item1, out int rowNumber))
                    continue;

                Word.Range cellRange = table.Cell(rowNumber, 2).Range;
                if (cellRange.InlineShapes.Count > 0)
                {
                    _clipboardEmfService.CopyMathTypeToFile(cellRange.InlineShapes[1], emfPath);
                }
            }
            catch (Exception ex)
            {
            }
        }

        try
        {
            document.Range(savedStart, savedEnd).Select();
        }
        catch (Exception ex)
        {
        }
        
    }

    private static void CleanupFolder(string folder)
    {
        if (!Directory.Exists(folder))
        {
            return;
        }

        try
        {
            foreach (string file in Directory.GetFiles(folder))
            {
                try
                {
                    File.Delete(file);
                }
                catch (Exception ex)
                {
                }
            }
            Directory.Delete(folder);
        }
        catch (Exception ex)
        {
        }
    }

    private void ScheduleTask(int delayMs, Action action)
    {
        System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer { Interval = delayMs };
        _activeTimers.Add(timer);
        timer.Tick += (s, e) =>
        {
            timer.Stop();
            _activeTimers.Remove(timer);
            timer.Dispose();
            try
            {
                action();
            }
            catch (Exception ex)
            {
            }
        };
        timer.Start();
    }

    private static Word.Table FindTableContainingPosition(Word.Document document, long position)
    {
        foreach (Word.Table table in document.Tables)
        {
            if (position >= table.Range.Start && position <= table.Range.End)
            {
                return table;
            }
        }

        return null;
    }

    private static Dictionary<string, int> BuildFormulaRowIndex(Word.Document document, List<Tuple<string, long>> formulaPositions)
    {
        Dictionary<string, int> result = new Dictionary<string, int>(StringComparer.Ordinal);
        List<Tuple<string, long>> sorted = new List<Tuple<string, long>>(formulaPositions);
        sorted.Sort((a, b) => a.Item2.CompareTo(b.Item2));

        long currentTableStart = long.MinValue;
        int rowNumber = 0;

        foreach (Tuple<string, long> fp in sorted)
        {
            Word.Table table = FindTableContainingPosition(document, fp.Item2);
            if (table == null)
            {
                continue;
            }

            if (table.Range.Start != currentTableStart)
            {
                currentTableStart = table.Range.Start;
                rowNumber = 1;
            }
            else
            {
                rowNumber++;
            }

            result[fp.Item1] = rowNumber;
        }

        return result;
    }
}
