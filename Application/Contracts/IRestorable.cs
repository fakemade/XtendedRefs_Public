using Word = Microsoft.Office.Interop.Word;
namespace XtendedRefs.ComAddIn.Application.Contracts;

public interface IRestorable
{
    public void Restore(Word.Document document);
}