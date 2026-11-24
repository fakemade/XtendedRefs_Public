using System;
using System.IO;
using System.Windows;
using XtendedRefs.ComAddIn.Application.Composition;
using XtendedRefs.ComAddIn.Domain.Abstractions;
using XtendedRefs.ComAddIn.Domain.Documents;
using XtendedRefs.ComAddIn.Domain.Models.NewModels;
using XtendedRefs.ComAddIn.Domain.State;
using XtendedRefs.ComAddIn.Infrastructure.Persistence;
using XtendedRefs.ComAddIn.Infrastructure.Documents;
using Word = Microsoft.Office.Interop.Word;

namespace XtendedRefs.ComAddIn.Application.Events;

internal sealed partial class WordApplicationEventHandler
{
    private void OnDocumentOpen(Word.Document document)
    {
        // Document-specific singletons are recreated for the active Word document.
        document.Bookmarks.ShowHidden = true;
        
        Singleton<DocumentCommandUtilities>.Initialize(() => new DocumentCommandUtilities(document));
        
        
        Singleton<XmlProvider>.Initialize(() => new XmlProvider(document,
        typeof(ChaptersList),
                         typeof(LiteratureList),
                         typeof(FormulaList),
                         typeof(PictureList),
                         typeof(TableList),
                         typeof(DocumentSettings))
        );
        
        Singleton<RuntimeStore>.Initialize(() => new RuntimeStore(document, Singleton<XmlProvider>.Instance));

        Singleton<RuntimeStore>.Instance.Init();

        


        Singleton<CommandsLocator>.Instance.SetVariables();


        try
        {
            Connect.Instance.setList(Singleton<RuntimeStore>.Instance.DocumentSettings.LevelStyles);
            CacheSilently(document);

        }
        catch (Exception e)
        {

            MessageBox.Show(e.Message);
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

    private void OnDocumentBeforeClose(Word.Document document, ref bool cancel)
    {

        try
        {
            if (document.Saved == true)
            {

                Singleton<RuntimeStore>.Instance.SaveBeforeClose();
            }

            if (_application.Documents.Count <= 1)
            {

                CleanupFolder(Path.Combine(Path.GetTempPath(), "WordFormulas"));
                CleanupFolder(Path.Combine(Path.GetTempPath(), "WordFigures"));
                CleanupFolder(Path.Combine(Path.GetTempPath(), "WordTables"));
            }
        }
        catch (Exception e)
        {
            MessageBox.Show(e.Message);
        }


    }

    private static string SafeDocumentName(Word.Document document)
    {
        try
        {
            return document?.FullName ?? "<null>";
        }
        catch
        {
            return "<unavailable>";
        }
    }
}
