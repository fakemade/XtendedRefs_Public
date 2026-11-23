using XtendedRefs.ComAddIn.Application.Composition;
using XtendedRefs.ComAddIn.Domain.Abstractions;
using XtendedRefs.ComAddIn.Domain.Models.NewModels;
using XtendedRefs.ComAddIn.Domain.State;
using Word = Microsoft.Office.Interop.Word;

namespace XtendedRefs.ComAddIn.Application.Commands.NewCommands;

public class ChapterCommands
{
    public void InsertChapterLink()
    {
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;
        Word.Document document = rs.Document;

        ServiceLocator srv  = Singleton<ServiceLocator>.Instance;
        
        
        
        ChapterInfo item = srv.ChapterPickerService.PickChapter(document);
        if (item == null)
        {
            return;
        }
        
        Word.Range targetRange = rs.Document.ActiveWindow.Selection.Range.Duplicate;
        targetRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);

        targetRange.Hyperlinks.Add(
            Anchor: targetRange,
            SubAddress: item.BookmarkName,
            TextToDisplay: item.Number,
            ScreenTip: item.Title
            );
        
    }
}