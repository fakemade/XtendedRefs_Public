using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Interop;
using XtendedRefs.ComAddIn.Application.Contracts;
using XtendedRefs.ComAddIn.Domain.Abstractions;
using XtendedRefs.ComAddIn.Domain.Models.NewModels;
using XtendedRefs.ComAddIn.Domain.State;
using XtendedRefs.ComAddIn.UI.Windows;
using Word = Microsoft.Office.Interop.Word;

namespace XtendedRefs.ComAddIn.Application.Pickers.Services;

public sealed class ChapterPickerService : IChapterPickerService
{
    public ChapterInfo PickChapter(Word.Document document)
    {
        
        if (document == null)
        {
            throw new ArgumentNullException(nameof(document));
        }
        ChapterPickerWindow window = new ChapterPickerWindow(document);
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
            ? window.selectedChapter
            : null;
        
    }


    public List<ChapterInfo> BuildSourceList()
    {
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;
        return rs.ChaptersList.Chapters;

    }
}
