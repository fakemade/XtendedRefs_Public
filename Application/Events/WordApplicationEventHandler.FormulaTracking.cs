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
                        _isMathTypeOpen = true;
                        _lastFormulaStart = selection.Start;
                        _lastFormulaEnd = selection.End;
                    }
                }
            }
        }
        catch (Exception ex)
        {
        }
    }

    private void OnWindowSelectionChange(Word.Selection selection)
    {
        try
        {
            if (_isMathTypeOpen)
            {
                _isMathTypeOpen = false;
                ScheduleTask(1000, UpdateFormulaEmfAfterEdit);
            }
        }
        catch (Exception ex)
        {
        }
    }

    private void UpdateFormulaEmfAfterEdit()
    {
        try
        {
            Word.Document document = _application.ActiveDocument;
            if (document == null)
            {
                return;
            }

            string bookmarkName = _bookmarkNavigator.FindNameInRange(
                document,
                (int)Math.Max(0, (int)(_lastFormulaStart - 100)),
                (int)(_lastFormulaEnd + 100),
                "Formula_");
            if (string.IsNullOrEmpty(bookmarkName))
            {
                return;
            }

            string folder = Path.Combine(Path.GetTempPath(), "WordFormulas");
            Directory.CreateDirectory(folder);
            string emfPath = Path.Combine(folder, bookmarkName + ".emf");
            try
            {
                File.Delete(emfPath);
            }
            catch (Exception ex)
            {
            }

            long bmStart = document.Bookmarks[bookmarkName].Range.Start;
            Word.Table table = FindTableContainingPosition(document, bmStart);
            if (table == null)
            {
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
                }
            }

            Dictionary<string, int> rowIndex = BuildFormulaRowIndex(document, formulaPositions);
            if (!rowIndex.TryGetValue(bookmarkName, out int rowNum))
            {
                return;
            }

            Word.Range cellRange = table.Cell(rowNum, 2).Range;
            if (cellRange.InlineShapes.Count == 0)
            {
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
                }

                _clipboardEmfService.CopyMathTypeToFile(cellRange.InlineShapes[1], emfPath);
            }
            finally
            {
                try
                {
                    if (window != null) window.View.ShowFieldCodes = prevShowFieldCodes;
                }
                catch (Exception ex)
                {
                }
            }
        }
        catch (Exception ex)
        {
        }
    }
}
