using System.Collections.Generic;
using XtendedRefs.ComAddIn.Application.Utilities;
using XtendedRefs.ComAddIn.Domain.Models.NewModels;
using XtendedRefs.ComAddIn.Infrastructure.Events;
using XtendedRefs.ComAddIn.Infrastructure.OfficeCom;
using Word = Microsoft.Office.Interop.Word;

namespace XtendedRefs.ComAddIn.Application.Events;

internal sealed partial class WordApplicationEventHandler
{
    private readonly Word.Application _application;
    private volatile bool _isMathTypeOpen;

    public WordApplicationEventHandler(IRibbonUI ribbonUi)
    {
        _application = AddInContext.Application;
    }

    public void Subscribe()
    {
        _application.DocumentOpen += OnDocumentOpen;
        _application.DocumentBeforeSave += OnDocumentBeforeSave;
        _application.DocumentBeforeClose += OnDocumentBeforeClose;
        _application.WindowBeforeDoubleClick += OnWindowBeforeDoubleClick;
        _application.WindowSelectionChange += OnWindowSelectionChange;
        
        EventBus<ChapterInfo>.OnHyperlinksUpdated += DocumentUtilities.UpdateChapterHyperLinks;
        EventBus<ChapterInfo>.OnHyperlinksUpdated += DocumentUtilities.UpdateStyles;

        EventBus<LiteratureInfo>.OnHyperlinksUpdated += DocumentUtilities.UpdateLiteratureHyperLinks;
        EventBus<LiteratureInfo>.OnHyperlinksUpdated += DocumentUtilities.UpdateStyles;
        EventBus<LiteratureInfo>.OnHyperlinksUpdated += DocumentUtilities.UpdateLiteratureFields;
        
        EventBus<FormulaInfo>.OnHyperlinksUpdated += DocumentUtilities.UpdateFormulaHyperLinks;
        
        EventBus<PictureInfo>.OnHyperlinksUpdated += DocumentUtilities.UpdatePictureHyperLinks;
        EventBus<PictureInfo>.OnHyperlinksUpdated += DocumentUtilities.UpdateStyles;
        
        EventBus<TableInfo>.OnHyperlinksUpdated += DocumentUtilities.UpdateTableHyperLinks;
        EventBus<TableInfo>.OnHyperlinksUpdated += DocumentUtilities.UpdateStyles;
    }

    public void Unsubscribe()
    {
        try
        {
            _application.DocumentOpen -= OnDocumentOpen;
            _application.DocumentBeforeSave -= OnDocumentBeforeSave;
            _application.DocumentBeforeClose -= OnDocumentBeforeClose;
            _application.WindowBeforeDoubleClick -= OnWindowBeforeDoubleClick;
            _application.WindowSelectionChange -= OnWindowSelectionChange;
            
            EventBus<ChapterInfo>.OnHyperlinksUpdated -= DocumentUtilities.UpdateChapterHyperLinks;
            EventBus<ChapterInfo>.OnHyperlinksUpdated -= DocumentUtilities.UpdateStyles;

            EventBus<LiteratureInfo>.OnHyperlinksUpdated -= DocumentUtilities.UpdateLiteratureHyperLinks;
            EventBus<LiteratureInfo>.OnHyperlinksUpdated -= DocumentUtilities.UpdateStyles;
            EventBus<LiteratureInfo>.OnHyperlinksUpdated -= DocumentUtilities.UpdateLiteratureFields;
        
            EventBus<FormulaInfo>.OnHyperlinksUpdated -= DocumentUtilities.UpdateFormulaHyperLinks;
        
            EventBus<PictureInfo>.OnHyperlinksUpdated -= DocumentUtilities.UpdatePictureHyperLinks;
            EventBus<PictureInfo>.OnHyperlinksUpdated -= DocumentUtilities.UpdateStyles;
            
            EventBus<TableInfo>.OnHyperlinksUpdated -= DocumentUtilities.UpdateTableHyperLinks;
            EventBus<TableInfo>.OnHyperlinksUpdated -= DocumentUtilities.UpdateStyles;
            
        }
        catch (System.Exception ex)
        {
        }
    }
}
