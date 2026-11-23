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
        AppLogger.Instance.Info("CacheSilently started");
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
            AppLogger.Instance.Info("CacheSilently: figures detected");
            CacheFigures(document);
        }

        if (hasFormulas)
        {
            AppLogger.Instance.Info("CacheSilently: formulas detected");
            CacheFormulas(document);
        }

        AppLogger.Instance.Info($"CacheSilently finished. hasFigures={hasFigures}, hasFormulas={hasFormulas}");
    }

    private void CacheFigures(Word.Document document)
    {
        AppLogger.Instance.Info("CacheFigures started");
        string folder = Path.Combine(Path.GetTempPath(), "WordFigures");
        Directory.CreateDirectory(folder);
        int savedStart = _application.Selection.Start;
        int savedEnd = _application.Selection.End;

        int processedCount = 0;
        foreach (string bookmarkName in GetFigureBookmarkNames(document))
        {
            processedCount++;
            AppLogger.Instance.Info("CacheFigures processing: " + bookmarkName);
            if (File.Exists(Path.Combine(folder, bookmarkName + ".emf")) ||
                File.Exists(Path.Combine(folder, bookmarkName + ".png")))
            {
                AppLogger.Instance.Info("CacheFigures skipped existing preview: " + bookmarkName);
                continue;
            }

            if (_mediaMap.TrySavePreview(document, bookmarkName, folder))
            {
                AppLogger.Instance.Info("CacheFigures restored preview from media map: " + bookmarkName);
                continue;
            }

            try
            {
                _clipboardEmfService.GenerateSingleFigureEmf(document, bookmarkName,
                    (bm, hash) => { Singleton<DocumentCommandUtilities>.Instance.SaveFigureHash(bm, hash); return true; });
                AppLogger.Instance.Info("CacheFigures generated preview: " + bookmarkName);
            }
            catch (Exception ex)
            {
                AppLogger.Instance.Exception("CacheFigures failed for " + bookmarkName, ex);
            }
        }

        try
        {
            document.Range(savedStart, savedEnd).Select();
        }
        catch (Exception ex)
        {
            AppLogger.Instance.Exception("CacheFigures failed to restore selection", ex);
        }

        AppLogger.Instance.Info("CacheFigures finished. processed=" + processedCount);
    }

    private List<string> GetFigureBookmarkNames(Word.Document document)
    {
        List<string> result = new List<string>();

        try
        {
            
        }
        catch (Exception ex)
        {
            AppLogger.Instance.Exception("GetFigureBookmarkNames failed to read NumberingState", ex);
        }

        if (result.Count > 0)
        {
            AppLogger.Instance.Info("GetFigureBookmarkNames state count=" + result.Count);
            return result;
        }

        // Fallback для документов, где рисунки еще созданы старой схемой Figure_/FigureNum_.
        List<string> legacyNames = _bookmarkNavigator.GetNames(document, "Figure_", "FigureNum_");
        AppLogger.Instance.Info("GetFigureBookmarkNames legacy count=" + legacyNames.Count);
        return legacyNames;
    }

    private void CacheFormulas(Word.Document document)
    {
        AppLogger.Instance.Info("CacheFormulas started");
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
                    AppLogger.Instance.Exception("CacheFormulas failed to read formula bookmark position: " + bm.Name, ex);
                }
            }
        }

        Dictionary<string, int> rowIndex = BuildFormulaRowIndex(document, formulaPositions);

        foreach (Tuple<string, long> fp in formulaPositions)
        {
            string emfPath = Path.Combine(folder, fp.Item1 + ".emf");
            if (File.Exists(emfPath))
            {
                AppLogger.Instance.Info("CacheFormulas skipped existing preview: " + fp.Item1);
                continue;
            }

            if (_mediaMap.TrySavePreview(document, fp.Item1, folder))
            {
                AppLogger.Instance.Info("CacheFormulas restored preview from media map: " + fp.Item1);
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
                    AppLogger.Instance.Info("CacheFormulas generated preview: " + fp.Item1);
                }
            }
            catch (Exception ex)
            {
                AppLogger.Instance.Exception("CacheFormulas failed for " + fp.Item1, ex);
            }
        }

        try
        {
            document.Range(savedStart, savedEnd).Select();
        }
        catch (Exception ex)
        {
            AppLogger.Instance.Exception("CacheFormulas failed to restore selection", ex);
        }

        AppLogger.Instance.Info("CacheFormulas finished. processed=" + formulaPositions.Count);
    }

    private static void CleanupFolder(string folder)
    {
        AppLogger.Instance.Info("CleanupFolder started: " + folder);
        if (!Directory.Exists(folder))
        {
            AppLogger.Instance.Info("CleanupFolder skipped missing folder: " + folder);
            return;
        }

        try
        {
            foreach (string file in Directory.GetFiles(folder))
            {
                try
                {
                    File.Delete(file);
                    AppLogger.Instance.Info("CleanupFolder deleted file: " + file);
                }
                catch (Exception ex)
                {
                    AppLogger.Instance.Exception("CleanupFolder failed to delete file: " + file, ex);
                }
            }

            Directory.Delete(folder);
            AppLogger.Instance.Info("CleanupFolder deleted folder: " + folder);
        }
        catch (Exception ex)
        {
            AppLogger.Instance.Exception("CleanupFolder failed: " + folder, ex);
        }
    }

    private void ScheduleTask(int delayMs, Action action)
    {
        AppLogger.Instance.Info("ScheduleTask created. delayMs=" + delayMs);
        System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer { Interval = delayMs };
        _activeTimers.Add(timer);
        timer.Tick += (s, e) =>
        {
            timer.Stop();
            _activeTimers.Remove(timer);
            timer.Dispose();
            AppLogger.Instance.Info("ScheduleTask executing. delayMs=" + delayMs);
            try
            {
                action();
                AppLogger.Instance.Info("ScheduleTask finished. delayMs=" + delayMs);
            }
            catch (Exception ex)
            {
                AppLogger.Instance.Exception("ScheduleTask failed. delayMs=" + delayMs, ex);
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
