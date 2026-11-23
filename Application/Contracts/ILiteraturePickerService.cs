using System.Collections.Generic;
using XtendedRefs.ComAddIn.Domain.Models.NewModels;
using Word = Microsoft.Office.Interop.Word;

namespace XtendedRefs.ComAddIn.Application.Contracts;

internal interface ILiteraturePickerService
{
    LiteratureInfo PickBookmark(Word.Document document);
    List<LiteratureInfo> BuildSourceList();

}
