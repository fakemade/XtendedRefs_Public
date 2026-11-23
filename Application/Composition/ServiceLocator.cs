using XtendedRefs.ComAddIn.Application.Pickers.Services;

namespace XtendedRefs.ComAddIn.Application.Composition;

public sealed class ServiceLocator
{
    public LiteraturePickerService LiteraturePickerService { get; }

    public PicturePickerService PicturePickerService { get; }

    public FormulaPickerService FormulaPickerService { get; }

    public TablePickerService TablePickerService { get; }
    
    public ChapterPickerService ChapterPickerService { get; }
    

    public ServiceLocator()
    {
        LiteraturePickerService = new LiteraturePickerService();
        PicturePickerService = new PicturePickerService();
        FormulaPickerService = new FormulaPickerService();
        TablePickerService = new TablePickerService();
        ChapterPickerService = new ChapterPickerService();
    }
}
