using System.Collections.Generic;
using System.Runtime.InteropServices;
using Extensibility;
using XtendedRefs.ComAddIn.Application.Composition;
using XtendedRefs.ComAddIn.Infrastructure.OfficeCom;
using XtendedRefs.ComAddIn.Domain.Documents;
using WordApplicationEventHandler = XtendedRefs.ComAddIn.Application.Events.WordApplicationEventHandler;

namespace XtendedRefs.ComAddIn;

[ComVisible(true)]
[Guid("1A2D0B5D-3B55-4F1D-9C69-2D6A54FB2C44")]
[ProgId(ProgId)]
[ClassInterface(ClassInterfaceType.AutoDispatch)]
public sealed partial class Connect : IDTExtensibility2, IRibbonExtensibility
{
    private static Connect _instance;

    internal const string ProgId = "XtendedRefs.ComAddIn";
    private const string AddInFriendlyName = "XtendedRefs C# COM Add-in";
    private const string AddInDescription = "Надстройка для управления объектами и вставкой рисунков, таблиц, формул и их учет.";

    private IRibbonUI _ribbon;
    private WordApplicationEventHandler _wordEventHandler;
    private CommandsLocator _commands;

    private int _startIndex = 0;
    private const int PageSize = 5;
    private List<LevelStyleItem> _styles;
    

    public static Connect Instance => _instance;
}
