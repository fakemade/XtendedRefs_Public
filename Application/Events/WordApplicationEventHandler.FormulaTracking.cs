using System;
using System.Collections.Generic;
using System.IO;
using XtendedRefs.ComAddIn.Infrastructure.Logging;
using Word = Microsoft.Office.Interop.Word;

namespace XtendedRefs.ComAddIn.Application.Events;

internal sealed partial class WordApplicationEventHandler
{
    private void OnWindowBeforeDoubleClick(Word.Selection selection, ref bool cancel)
    {
        AppLogger.Instance.Info("OnWindowBeforeDoubleClick started");
        try
        {
            if (selection.InlineShapes.Count > 0)
            {
                Word.InlineShape shape = selection.InlineShapes[1];
                if (shape.Type == Word.WdInlineShapeType.wdInlineShapeEmbeddedOLEObject)
                {
                    string classType = shape.OLEFormat.ClassType;
                    if (classType != null &&
                        (classType.StartsWith("Equation.", StringComparison.Ordinal) ||
                         string.Equals(classType, "Equation", StringComparison.Ordinal)))
                    {
                        AppLogger.Instance.Info($"MathType object opened. Start={selection.Start}, End={selection.End}, ClassType={classType}");
                        _isMathTypeOpen = true;
                        _lastFormulaStart = selection.Start;
                        _lastFormulaEnd = selection.End;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            AppLogger.Instance.Exception("OnWindowBeforeDoubleClick failed", ex);
        }
    }

    private void OnWindowSelectionChange(Word.Selection selection)
    {
        AppLogger.Instance.Info("OnWindowSelectionChange started");
        try
        {
            if (_isMathTypeOpen)
            {
                AppLogger.Instance.Info("MathType edit probably finished; scheduling preview update");
                _isMathTypeOpen = false;
                ScheduleTask(1000, UpdateFormulaEmfAfterEdit);
            }
        }
        catch (Exception ex)
        {
            AppLogger.Instance.Exception("OnWindowSelectionChange failed", ex);
        }
    }

    private void UpdateFormulaEmfAfterEdit()
    {
        AppLogger.Instance.Info("UpdateFormulaEmfAfterEdit started");
        try
        {
            Word.Document document = _application.ActiveDocument;
            if (document == null)
            {
                AppLogger.Instance.Warn("UpdateFormulaEmfAfterEdit skipped: active document is null");
                return;
            }

            string bookmarkName = _bookmarkNavigator.FindNameInRange(
                document,
                (int)Math.Max(0, (int)(_lastFormulaStart - 100)),
                (int)(_lastFormulaEnd + 100),
                "Formula_");
            if (string.IsNullOrEmpty(bookmarkName))
            {
                AppLogger.Instance.Warn("UpdateFormulaEmfAfterEdit skipped: formula bookmark not found");
                return;
            }
            AppLogger.Instance.Info("UpdateFormulaEmfAfterEdit bookmark: " + bookmarkName);

            string folder = Path.Combine(Path.GetTempPath(), "WordFormulas");
            Directory.CreateDirectory(folder);
            string emfPath = Path.Combine(folder, bookmarkName + ".emf");
            try
            {
                File.Delete(emfPath);
                AppLogger.Instance.Info("UpdateFormulaEmfAfterEdit deleted old preview: " + emfPath);
            }
            catch (Exception ex)
            {
                AppLogger.Instance.Exception("UpdateFormulaEmfAfterEdit failed to delete old preview: " + emfPath, ex);
            }

            long bmStart = document.Bookmarks[bookmarkName].Range.Start;
            Word.Table table = FindTableContainingPosition(document, bmStart);
            if (table == null)
            {
                AppLogger.Instance.Warn("UpdateFormulaEmfAfterEdit skipped: containing table not found");
                return;
            }

            List<Tuple<string, long>> formulaPositions = new List<Tuple<string, long>>();
            foreach (string name in _bookmarkNavigator.GetNames(document, "Formula_"))
            {
                try
                {
                    formulaPositions.Add(Tuple.Create(name, (long)document.Bookmarks[name].Range.Start));
                }
                catch (Exception ex)
                {
                    AppLogger.Instance.Exception("UpdateFormulaEmfAfterEdit failed to read formula bookmark position: " + name, ex);
                }
            }

            Dictionary<string, int> rowIndex = BuildFormulaRowIndex(document, formulaPositions);
            if (!rowIndex.TryGetValue(bookmarkName, out int rowNum))
            {
                AppLogger.Instance.Warn("UpdateFormulaEmfAfterEdit skipped: row index not found");
                return;
            }

            Word.Range cellRange = table.Cell(rowNum, 2).Range;
            if (cellRange.InlineShapes.Count == 0)
            {
                AppLogger.Instance.Warn("UpdateFormulaEmfAfterEdit skipped: no inline shapes in formula cell");
                return;
            }

            Word.Window window = null;
            bool prevShowFieldCodes = false;
            try
            {
                window = _application.ActiveWindow;
                if (window != null)
                {
                    prevShowFieldCodes = window.View.ShowFieldCodes;
                }
            }
            catch (Exception ex)
            {
                AppLogger.Instance.Exception("UpdateFormulaEmfAfterEdit failed to read active window field-code state", ex);
            }

            try
            {
                table.Range.Fields.Update();
                System.Windows.Forms.Application.DoEvents();
                _application.ScreenRefresh();
                System.Windows.Forms.Application.DoEvents();
                try
                {
                    if (window != null) window.View.ShowFieldCodes = false;
                }
                catch (Exception ex)
                {
                    AppLogger.Instance.Exception("UpdateFormulaEmfAfterEdit failed to hide field codes", ex);
                }

                _clipboardEmfService.CopyMathTypeToFile(cellRange.InlineShapes[1], emfPath);
                AppLogger.Instance.Info("UpdateFormulaEmfAfterEdit copied preview: " + emfPath);
            }
            finally
            {
                try
                {
                    if (window != null) window.View.ShowFieldCodes = prevShowFieldCodes;
                }
                catch (Exception ex)
                {
                    AppLogger.Instance.Exception("UpdateFormulaEmfAfterEdit failed to restore field-code state", ex);
                }
            }
        }
        catch (Exception ex)
        {
            AppLogger.Instance.Exception("UpdateFormulaEmfAfterEdit failed", ex);
        }

        AppLogger.Instance.Info("UpdateFormulaEmfAfterEdit finished");
    }
}
