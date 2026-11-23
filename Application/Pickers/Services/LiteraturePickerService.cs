using System;
using System.Collections.Generic;
using System.Windows.Interop;
using XtendedRefs.ComAddIn.Application.Contracts;
using XtendedRefs.ComAddIn.Application.Utilities;
using XtendedRefs.ComAddIn.Domain.Abstractions;
using XtendedRefs.ComAddIn.Domain.Models.NewModels;
using XtendedRefs.ComAddIn.Domain.State;
using XtendedRefs.ComAddIn.UI.Windows;
using Word = Microsoft.Office.Interop.Word;

namespace XtendedRefs.ComAddIn.Application.Pickers.Services;

public sealed class LiteraturePickerService : ILiteraturePickerService
{
    public LiteratureInfo PickBookmark(Word.Document document)
    {
        
        if (document == null)
        {
            throw new ArgumentNullException(nameof(document));
        }
        LiteraturePickerWindow window = new LiteraturePickerWindow(document);
        try
        {
            if (document.ActiveWindow != null)
            {
                new WindowInteropHelper(window).Owner = new IntPtr(document.ActiveWindow.Hwnd);
            }
        }
        catch
        {
            // The owner handle is optional: the dialog can still be shown without it.
        }

        bool? dialogResult = window.ShowDialog();
        return dialogResult == true
            ? window.SelectedBookmarkName
            : null;
    }

    public List<LiteratureInfo> BuildSourceList()
    {
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;
        rs.LiteratureList.Literature.Sort((a, b) =>
            DocumentUtilities.CompareParts(a.NumberParts, b.NumberParts)
        );
        return rs.LiteratureList.Literature;
    }
}
