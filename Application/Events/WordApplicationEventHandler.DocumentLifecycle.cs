using System;
using System.Windows;
using XtendedRefs.ComAddIn.Application.Utilities;
using XtendedRefs.ComAddIn.Domain.Abstractions;
using XtendedRefs.ComAddIn.Domain.Documents;
using XtendedRefs.ComAddIn.Domain.Models.NewModels;
using XtendedRefs.ComAddIn.Domain.State;
using XtendedRefs.ComAddIn.Infrastructure.Persistence;
using Word = Microsoft.Office.Interop.Word;

namespace XtendedRefs.ComAddIn.Application.Events;

internal sealed partial class WordApplicationEventHandler
{
    private void OnDocumentOpen(Word.Document doc)
    {
        doc.Bookmarks.ShowHidden = true;
        Singleton<XmlProvider>.Initialize(() => new XmlProvider(doc,
            typeof(ChaptersList),
            typeof(LiteratureList),
            typeof(FormulaList),
            typeof(PictureList),
            typeof(TableList),
            typeof(DocumentSettings))
        );
        Singleton<RuntimeStore>.Initialize(() => new RuntimeStore(doc, Singleton<XmlProvider>.Instance));
        Singleton<RuntimeStore>.Instance.Init();
        //Scan document
        if(!Singleton<RuntimeStore>.Instance.DocumentSettings.IsScanned)
        {
            DocumentUtilities.ScanChapters(doc, Singleton<RuntimeStore>.Instance.DocumentSettings.LevelStyles);
            Singleton<RuntimeStore>.Instance.DocumentSettings.IsScanned = true;
        }

    }

    private void OnDocumentBeforeSave(Word.Document document, ref bool saveAsUI, ref bool cancel)
    {

        try
        {
            Singleton<RuntimeStore>.Instance.SaveBeforeClose();
        }
        catch (Exception ex)
        {

        }
        
    }

    private void OnDocumentBeforeClose(Word.Document doc, ref bool cancel)
    {

        try
        {
            if (doc.Saved == true)
            {

                Singleton<RuntimeStore>.Instance.SaveBeforeClose();
            }

            if (_application.Documents.Count <= 1)
            {

                // CleanupFolder(Path.Combine(Path.GetTempPath(), "WordFormulas"));
                // CleanupFolder(Path.Combine(Path.GetTempPath(), "WordFigures"));
                // CleanupFolder(Path.Combine(Path.GetTempPath(), "WordTables"));
            }
        }
        catch (Exception e)
        {
            MessageBox.Show(e.Message);
        }
    }

}
