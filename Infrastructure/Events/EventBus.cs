using System;
using XtendedRefs.ComAddIn.Domain.Models.NewModels;
using Word = Microsoft.Office.Interop.Word;

namespace XtendedRefs.ComAddIn.Infrastructure.Events;

public static class EventBus<T> where T : DocumentAnchor
{
    public static event Func<Word.Document,bool>  OnHyperlinksUpdated;

    public static void Raise(Word.Document doc)
    {
        OnHyperlinksUpdated?.Invoke(doc);
    }
}