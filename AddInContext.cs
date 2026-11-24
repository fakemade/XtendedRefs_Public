using Extensibility;
using Word = Microsoft.Office.Interop.Word;

namespace XtendedRefs.ComAddIn;

internal static class AddInContext
{
    public static Word.Application Application { get; set; }

    public static object ComAddInInstance { get; set; }

    public static ext_ConnectMode ConnectMode { get; set; }

    public static bool IsWordAvailable => Application != null;
}
