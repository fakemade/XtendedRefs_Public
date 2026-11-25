using System;
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
            }
        }
        catch (Exception ex)
        {
        }
    }
}
