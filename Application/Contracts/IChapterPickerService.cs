using System.Collections.Generic;
using XtendedRefs.ComAddIn.Domain.Models.NewModels;
using Word = Microsoft.Office.Interop.Word;

namespace XtendedRefs.ComAddIn.Application.Contracts;

internal interface IChapterPickerService
{
    ChapterInfo PickChapter(Word.Document document);
    List<ChapterInfo> BuildSourceList();

}
